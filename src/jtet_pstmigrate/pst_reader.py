"""PST extraction via libpst's `readpst` binary.

Why subprocess instead of a Python PST library?
  - libpst (readpst) has been the de-facto Linux PST reader for ~20 years.
    It handles the format edge cases — encrypted PSTs, ANSI vs Unicode,
    corrupted nodes — far better than any current Python binding.
  - The trade-off is that some MAPI properties (read state, importance,
    categories, flag status) don't survive the MIME conversion. For mail-only
    archive migration this is an acceptable loss.

Output layout we ask readpst for:
  --mode separate (-S)  one .eml file per message
  --output-mode mboxrd  no, we use -e to get .eml
  -e                    save messages as .eml files in folder hierarchy
  -D                    include items in 'Deleted Items'
  -t e                  emit only emails (no contacts/calendar/journal)
"""

from __future__ import annotations

import dataclasses
import email
import hashlib
import re
import shutil
import subprocess
import threading
import time
from collections.abc import Iterator
from email.policy import compat32
from pathlib import Path

from loguru import logger


class ReadpstError(RuntimeError):
    pass


@dataclasses.dataclass(slots=True, frozen=True)
class ExtractedMessage:
    """One .eml file produced by readpst."""

    file_path: Path
    folder_path: tuple[str, ...]   # ('Top of Personal Folders', 'Inbox', 'Sub')
    bytes_: int
    dedupe_key: str                 # 'imid:<...>' or 'sha256:<hex>'
    subject: str
    received: str | None
    message_id: str | None


@dataclasses.dataclass(slots=True, frozen=True)
class ExtractedAppointment:
    """One .ics file produced by ``readpst -t a``."""

    file_path: Path
    folder_path: tuple[str, ...]
    bytes_: int
    dedupe_key: str                 # 'uid:<...>' or 'sha256:<hex>'
    summary: str                    # SUMMARY field, truncated
    uid: str | None
    start: str | None               # raw DTSTART text, for logging only


@dataclasses.dataclass(slots=True, frozen=True)
class ExtractedContact:
    """One .vcf file produced by ``readpst -t c``."""

    file_path: Path
    folder_path: tuple[str, ...]
    bytes_: int
    dedupe_key: str                 # 'uid:<...>', 'fn-email:<...>', or 'sha256:<hex>'
    display_name: str               # FN field, truncated, for logs
    primary_email: str | None       # first EMAIL value, used in dedup fallback


def check_readpst(binary: Path) -> str:
    """Return the readpst version string, or raise."""
    exe = shutil.which(str(binary)) or str(binary)
    if not Path(exe).exists() and not shutil.which(exe):
        raise ReadpstError(f"readpst not found: {binary} (try `dnf install libpst`)")
    try:
        result = subprocess.run([exe, "-V"], capture_output=True, text=True, timeout=10)
    except (subprocess.SubprocessError, OSError) as e:
        raise ReadpstError(f"readpst not executable: {e}") from e
    out = (result.stdout + result.stderr).strip().splitlines()
    return out[0] if out else "(unknown version)"


def extract_pst(pst_path: Path, work_dir: Path, *, binary: Path = Path("readpst"), include_deleted: bool = False) -> Path:
    """Extract mail (.eml) from a PST. See ``_run_readpst`` for the contract."""
    return _run_readpst(
        pst_path, work_dir, binary=binary, include_deleted=include_deleted,
        type_codes="e", subdir_suffix="",
    )


def extract_pst_calendar(
    pst_path: Path, work_dir: Path, *, binary: Path = Path("readpst"),
    include_deleted: bool = False,
) -> Path:
    """Extract appointments (.ics) from a PST.

    Output goes into a separate sibling directory of the mail extraction
    (``<work_dir>/<pst_stem>__calendar``) so the two never overlap and so
    re-running ``import`` doesn't invalidate the calendar sentinel.
    """
    return _run_readpst(
        pst_path, work_dir, binary=binary, include_deleted=include_deleted,
        type_codes="a", subdir_suffix="__calendar",
    )


def extract_pst_contacts(
    pst_path: Path, work_dir: Path, *, binary: Path = Path("readpst"),
    include_deleted: bool = False,
) -> Path:
    """Extract contacts (.vcf) from a PST.

    See ``extract_pst_calendar`` for the directory-isolation rationale.
    """
    return _run_readpst(
        pst_path, work_dir, binary=binary, include_deleted=include_deleted,
        type_codes="c", subdir_suffix="__contacts",
    )


def _run_readpst(
    pst_path: Path,
    work_dir: Path,
    *,
    binary: Path,
    include_deleted: bool,
    type_codes: str,
    subdir_suffix: str,
) -> Path:
    """Run readpst with a specific ``-t`` selection, return the output dir.

    Resume-friendly: if a previous run already extracted this PST successfully
    for this type-code set (sentinel file present + matching source size + same
    type-codes), we reuse the existing work dir instead of re-running readpst,
    which can take hours on multi-GB PSTs.

    If the work dir exists but the sentinel is missing or recorded a different
    type-code set, we assume the previous extraction was for something else
    and start over (wipe + re-extract).
    """
    if not pst_path.exists():
        raise ReadpstError(f"PST not found: {pst_path}")

    out_dir = work_dir / (_safe_name(pst_path.stem) + subdir_suffix)
    sentinel = out_dir / ".extract_complete"
    src_size = pst_path.stat().st_size
    sentinel_payload = f"{src_size}\n{type_codes}\n"

    if sentinel.exists():
        recorded = sentinel.read_text(errors="replace").strip().splitlines()
        recorded_size = -1
        recorded_codes = ""
        if recorded:
            try:
                recorded_size = int(recorded[0])
            except ValueError:
                recorded_size = -1
            recorded_codes = recorded[1] if len(recorded) > 1 else "e"  # legacy mail extractions only had size
        if recorded_size == src_size and recorded_codes == type_codes:
            n_files, total_bytes = _measure_dir(out_dir)
            logger.bind(ctx="pst").info(
                "Reusing previous extraction of {} (-t {}): {} files / {:.2f} GB "
                "(skipping readpst — delete {} to force re-extract)",
                pst_path.name, type_codes, n_files, total_bytes / 1024**3, sentinel,
            )
            return out_dir
        logger.bind(ctx="pst").warning(
            "Sentinel for {} records ({} bytes / -t {}); re-extracting for ({} bytes / -t {})",
            pst_path.name, recorded_size, recorded_codes or "?", src_size, type_codes,
        )

    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    args = [
        str(binary),
        "-e",                  # one file per item, with proper extensions (.eml/.ics/.vcf)
        "-t", type_codes,
        "-o", str(out_dir),
    ]
    if include_deleted:
        args.append("-D")
    args.append(str(pst_path))

    src_bytes = src_size
    log = logger.bind(ctx="pst")
    log.info(
        "Extracting {} ({:.2f} GB) -> {}",
        pst_path.name, src_bytes / 1024**3, out_dir,
    )

    # readpst is silent during extraction. Spawn a watcher thread that
    # periodically reports the work-dir size + file count + estimated %
    # so the user knows the process is alive on multi-GB PSTs.
    stop_watcher = threading.Event()
    watcher = threading.Thread(
        target=_watch_extraction,
        args=(out_dir, src_bytes, pst_path.name, stop_watcher),
        name=f"extract-watch-{pst_path.stem[:20]}",
        daemon=True,
    )
    watcher.start()

    started = time.monotonic()
    try:
        result = subprocess.run(args, capture_output=True, text=True, timeout=60 * 60 * 6)
    except subprocess.TimeoutExpired as e:
        stop_watcher.set()
        raise ReadpstError(f"readpst timed out on {pst_path}") from e
    finally:
        stop_watcher.set()
        watcher.join(timeout=2.0)

    elapsed = time.monotonic() - started
    final_count, final_bytes = _measure_dir(out_dir)
    tail = (result.stderr or result.stdout or "").splitlines()[-20:]

    if result.returncode != 0:
        # readpst commonly aborts on folder names with characters Windows
        # forbids in directory names (", :, *, ?, <, >, |) e.g.:
        #   mk_separate_dir: Cannot create directory CSC-7" BUTTERFLY HOOKS
        # When that happens, most folders extracted successfully — only the
        # offending folder's contents are lost. Treat it as partial success
        # so we still upload the thousands of recoverable messages.
        if final_count > 0:
            log.warning(
                "readpst exit {} on {} but produced {} files / {:.2f} GB — "
                "treating as PARTIAL extraction (some messages lost). Last lines:\n  {}",
                result.returncode, pst_path.name, final_count,
                final_bytes / 1024**3, "\n  ".join(tail),
            )
        else:
            raise ReadpstError(
                f"readpst exit {result.returncode} on {pst_path.name} (no files produced):\n  "
                + "\n  ".join(tail)
            )
    else:
        log.info(
            "Extracted {} in {:.0f}s -> {} files / {:.2f} GB (source was {:.2f} GB)",
            pst_path.name, elapsed, final_count,
            final_bytes / 1024**3, src_bytes / 1024**3,
        )

    # Sentinel: marks this extraction as complete and records both source
    # size and the readpst -t selection, so a later restart can reuse it
    # only when extracting the same item types from the same PST file.
    sentinel.write_text(sentinel_payload, encoding="utf-8")
    return out_dir


def _watch_extraction(out_dir: Path, src_bytes: int, pst_name: str, stop: threading.Event) -> None:
    """Background heartbeat: log work-dir size every 30s while readpst runs."""
    log = logger.bind(ctx="pst")
    poll_interval = 30.0
    started = time.monotonic()
    last_bytes = 0
    while not stop.wait(poll_interval):
        try:
            count, written = _measure_dir(out_dir)
        except OSError:
            continue
        elapsed = time.monotonic() - started
        rate_mb_s = ((written - last_bytes) / 1024**2) / poll_interval
        last_bytes = written
        # readpst output is *roughly* the same size as the source MAPI store,
        # within ~10-20% either way. Treat the percentage as a hint, not gospel.
        pct = (written / src_bytes * 100) if src_bytes else 0
        log.info(
            "Extracting {}: {} files / {:.2f} GB written (~{:.0f}% of source) "
            "• {:.0f}s elapsed • {:.1f} MB/s",
            pst_name, count, written / 1024**3, pct, elapsed, rate_mb_s,
        )


def _measure_dir(path: Path) -> tuple[int, int]:
    """Return (file_count, total_bytes) under path. Tolerant of missing files."""
    count = 0
    total = 0
    if not path.exists():
        return 0, 0
    for p in path.rglob("*"):
        try:
            if p.is_file():
                count += 1
                total += p.stat().st_size
        except OSError:
            pass
    return count, total


def iter_messages(extracted_root: Path) -> Iterator[ExtractedMessage]:
    """Walk the extracted tree, yield one message per .eml file.

    readpst lays out:
        <out_dir>/<pst_stem>/<Folder>/<Sub>/12345.eml

    The first directory under out_dir is the PST display name; we strip it so
    folder_path starts at the meaningful root.
    """
    if not extracted_root.exists():
        return

    pst_dirs = [p for p in extracted_root.iterdir() if p.is_dir()]
    if not pst_dirs:
        return

    for pst_dir in pst_dirs:
        for eml in sorted(pst_dir.rglob("*.eml")):
            try:
                yield _build_message(eml, pst_dir)
            except Exception as e:
                logger.bind(ctx="pst").warning("Skipping unreadable {}: {}", eml, e)


# Read only enough bytes from each .eml to capture standard RFC-822 headers.
# 99%+ of real-world messages have all relevant headers (Message-ID, Subject,
# Date, From, To) within the first 16 KB. Reading the full file just to
# extract these is the dominant cost of iter_messages on large mailboxes
# (60 GB across 200k files = ~60 GB of pointless disk I/O per startup).
_HEADER_PEEK_BYTES = 16 * 1024


def _build_message(eml: Path, pst_dir: Path) -> ExtractedMessage:
    rel = eml.relative_to(pst_dir).parts[:-1]   # drop the filename
    file_size = eml.stat().st_size

    # Fast path: read just the header bytes; do NOT slurp the whole .eml.
    # The full bytes are read at upload time when we actually need them.
    with eml.open("rb") as f:
        head = f.read(_HEADER_PEEK_BYTES)
    parsed = email.message_from_bytes(head, policy=compat32)
    message_id = (parsed.get("Message-ID") or parsed.get("Message-Id") or "").strip("<> \t")
    subject = parsed.get("Subject", "") or ""
    received = parsed.get("Date") or None

    if message_id:
        dedupe_key = f"imid:{message_id.lower()}"
    else:
        # Rare fallback: no Message-ID header. We need a stable fingerprint,
        # so read the whole file to hash it.
        raw = eml.read_bytes()
        dedupe_key = f"sha256:{hashlib.sha256(raw).hexdigest()}"

    return ExtractedMessage(
        file_path=eml,
        folder_path=tuple(_clean_folder(p) for p in rel),
        bytes_=file_size,
        dedupe_key=dedupe_key,
        subject=subject[:500],
        received=received,
        message_id=message_id or None,
    )


def iter_appointments(extracted_root: Path) -> Iterator[ExtractedAppointment]:
    """Walk the calendar-extracted tree, yield one appointment per .ics file.

    readpst -t a -e lays out:
        <out_dir>/<pst_stem>/Calendar/12345.ics

    Files that don't parse as iCalendar are logged and skipped rather than
    aborting the whole run.
    """
    if not extracted_root.exists():
        return

    pst_dirs = [p for p in extracted_root.iterdir() if p.is_dir()]
    if not pst_dirs:
        return

    for pst_dir in pst_dirs:
        for ics in sorted(pst_dir.rglob("*.ics")):
            try:
                yield _build_appointment(ics, pst_dir)
            except Exception as e:
                logger.bind(ctx="pst").warning("Skipping unreadable {}: {}", ics, e)


def _build_appointment(ics: Path, pst_dir: Path) -> ExtractedAppointment:
    rel = ics.relative_to(pst_dir).parts[:-1]
    file_size = ics.stat().st_size

    raw = ics.read_bytes()
    summary = ""
    uid: str | None = None
    start: str | None = None
    # Parse just the lines we need without pulling in icalendar at index time.
    # The proper iCalendar parse happens at upload time. Here we only need a
    # stable dedup key and a human-readable summary for logs.
    for line in raw.splitlines():
        try:
            decoded = line.decode("utf-8", errors="replace")
        except Exception:
            continue
        # iCalendar folds long lines with leading whitespace; we don't unfold
        # because the fields we read are short and unfolded in practice for
        # readpst's output.
        if not summary and decoded.startswith("SUMMARY:"):
            summary = decoded[len("SUMMARY:"):].strip()
        elif not uid and decoded.startswith("UID:"):
            uid = decoded[len("UID:"):].strip()
        elif not start and (decoded.startswith("DTSTART:") or decoded.startswith("DTSTART;")):
            start = decoded.split(":", 1)[1].strip() if ":" in decoded else None
        if summary and uid and start:
            break

    dedupe_key = (
        f"uid:{uid.lower()}" if uid else f"sha256:{hashlib.sha256(raw).hexdigest()}"
    )

    return ExtractedAppointment(
        file_path=ics,
        folder_path=tuple(_clean_folder(p) for p in rel),
        bytes_=file_size,
        dedupe_key=dedupe_key,
        summary=(summary or "(no subject)")[:500],
        uid=uid,
        start=start,
    )


def iter_contacts(extracted_root: Path) -> Iterator[ExtractedContact]:
    """Walk the contacts-extracted tree, yield one contact per .vcf file.

    readpst -t c -e lays out:
        <out_dir>/<pst_stem>/Contacts/12345.vcf

    Files that don't parse as vCard are logged and skipped rather than
    aborting the whole run.
    """
    if not extracted_root.exists():
        return

    pst_dirs = [p for p in extracted_root.iterdir() if p.is_dir()]
    if not pst_dirs:
        return

    for pst_dir in pst_dirs:
        for vcf in sorted(pst_dir.rglob("*.vcf")):
            try:
                yield _build_contact(vcf, pst_dir)
            except Exception as e:
                logger.bind(ctx="pst").warning("Skipping unreadable {}: {}", vcf, e)


def _build_contact(vcf: Path, pst_dir: Path) -> ExtractedContact:
    rel = vcf.relative_to(pst_dir).parts[:-1]
    file_size = vcf.stat().st_size

    raw = vcf.read_bytes()
    fn = ""
    uid: str | None = None
    primary_email: str | None = None
    # Cheap line-scan for dedup metadata; the proper vCard parse happens
    # at upload time. vCard line folding (continuation = leading SP/TAB)
    # affects only NOTE and ADR in practice; the short fields we read
    # below are unfolded in readpst's output.
    for line in raw.splitlines():
        try:
            decoded = line.decode("utf-8", errors="replace")
        except Exception:
            continue
        # Property name can have parameters: "EMAIL;TYPE=INTERNET:x@y".
        if ":" not in decoded:
            continue
        head, value = decoded.split(":", 1)
        name = head.split(";", 1)[0].strip().upper()
        value = value.strip()
        if not value:
            continue
        if not fn and name == "FN":
            fn = value
        elif not uid and name == "UID":
            uid = value
        elif not primary_email and name == "EMAIL":
            primary_email = value
        if fn and uid and primary_email:
            break

    if uid:
        dedupe_key = f"uid:{uid.lower()}"
    elif fn or primary_email:
        # FN+email is unique enough in practice -- two contacts named
        # "John Smith" with the same primary email are almost certainly
        # the same person exported twice.
        dedupe_key = f"fn-email:{fn.lower()}|{(primary_email or '').lower()}"
    else:
        dedupe_key = f"sha256:{hashlib.sha256(raw).hexdigest()}"

    return ExtractedContact(
        file_path=vcf,
        folder_path=tuple(_clean_folder(p) for p in rel),
        bytes_=file_size,
        dedupe_key=dedupe_key,
        display_name=(fn or primary_email or "(no name)")[:500],
        primary_email=primary_email,
    )


_INVALID_FS = re.compile(r'[<>:"|?*\x00-\x1f]+')


def _safe_name(name: str) -> str:
    return _INVALID_FS.sub("_", name).strip(". ")


def _clean_folder(name: str) -> str:
    """Trim readpst's quirks from folder names.

    readpst sometimes appends counters or escapes characters; preserve the
    user-visible name as much as possible while keeping it Graph-safe.
    """
    n = name.strip().replace("\\", "_").replace("/", "_")
    return n[:255] or "Unnamed"
