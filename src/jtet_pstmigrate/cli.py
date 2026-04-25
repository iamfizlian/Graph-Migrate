"""CLI entry point — `pstmigrate ...` (or `python -m jtet_pstmigrate`)."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Annotated

import typer
from rich.console import Console
from rich.table import Table

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, MappingRow, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient
from jtet_pstmigrate.log import configure_logging
from jtet_pstmigrate.orchestrator import Orchestrator, load_mapping
from jtet_pstmigrate.pst_reader import check_readpst
from jtet_pstmigrate.state import StateStore

app = typer.Typer(
    name="pstmigrate",
    help="PST -> Microsoft 365 mailbox migration via Microsoft Graph.",
    no_args_is_help=True,
    pretty_exceptions_show_locals=False,
)
console = Console()


ConfigOpt = Annotated[Path | None, typer.Option("--config", "-c", help="TOML config file")]


def _load(config_path: Path | None) -> AppConfig:
    cfg = AppConfig.load(config_path)
    return expand_user_paths(cfg)


def _run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _filter_mapping(
    rows: list[MappingRow],
    *,
    mailboxes: list[str] | None,
    pst_names: list[str] | None,
    limit: int | None,
) -> list[MappingRow]:
    """Apply selection filters in order: mailbox match, PST filename match, then row-count limit.

    UPN matching is case-insensitive (M365 UPNs aren't case-sensitive).
    PST name matching uses substring on the file basename, also case-insensitive.
    """
    selected = list(rows)

    if mailboxes:
        wanted = {m.strip().lower() for m in mailboxes if m and m.strip()}
        selected = [r for r in selected if r.target_mailbox.lower() in wanted]
        unmatched = wanted - {r.target_mailbox.lower() for r in selected}
        if unmatched:
            console.print(
                f"[yellow]Warning:[/] no mapping rows for: {', '.join(sorted(unmatched))}"
            )

    if pst_names:
        needles = [p.strip().lower() for p in pst_names if p and p.strip()]
        selected = [
            r for r in selected
            if any(n in r.pst_path.name.lower() for n in needles)
        ]

    if limit is not None and limit > 0:
        selected = selected[:limit]

    return selected


def _print_selection(rows: list[MappingRow], heading: str = "Selected") -> None:
    if not rows:
        console.print("[red]No mapping rows match the selection.[/]")
        return
    table = Table(title=f"{heading} — {len(rows)} row(s)")
    table.add_column("#", justify="right")
    table.add_column("PST", overflow="fold")
    table.add_column("Size", justify="right")
    table.add_column("Mailbox", overflow="fold")
    table.add_column("Root folder")
    for i, r in enumerate(rows, start=1):
        size = "?"
        if r.pst_path.exists():
            gb = r.pst_path.stat().st_size / 1024**3
            size = f"{gb:0.2f} GB" if gb >= 1 else f"{r.pst_path.stat().st_size / 1024**2:0.0f} MB"
        table.add_row(str(i), r.pst_path.name, size, r.target_mailbox, r.target_root_folder or "(root)")
    console.print(table)


# Selection options shared between validate and import
MailboxOpt = Annotated[
    list[str] | None,
    typer.Option(
        "--mailbox", "-M",
        help="Filter to one or more target mailboxes (UPN). Repeat for multi-select.",
    ),
]
PstFilterOpt = Annotated[
    list[str] | None,
    typer.Option(
        "--pst", "-P",
        help="Filter by substring of the PST filename. Repeat for multi-select.",
    ),
]
LimitOpt = Annotated[
    int | None,
    typer.Option("--limit", "-n", help="After other filters, take only the first N rows."),
]
ListOnlyOpt = Annotated[
    bool,
    typer.Option("--list", help="Show what would run and exit (dry-run for selection)."),
]


@app.command()
def template(
    output: Annotated[Path, typer.Option("--output", "-o")] = Path("mapping.example.csv"),
    force: Annotated[bool, typer.Option("--force", "-f", help="Overwrite if the file already exists.")] = False,
) -> None:
    """Write a starter mapping CSV (defaults to mapping.example.csv to avoid clobbering a real mapping.csv)."""
    if output.exists() and not force:
        console.print(
            f"[red]{output} already exists.[/] Refusing to overwrite — pass [bold]--force[/] "
            f"if you really want to replace it."
        )
        raise typer.Exit(1)
    output.write_text(
        "PSTPath,TargetMailbox,TargetRootFolder\n"
        "/data/pst/jsmith.pst,jsmith@contoso.com,Imported PST\n"
        "/data/pst/bdoe.pst,bdoe@contoso.com,Imported PST\n",
        encoding="utf-8",
    )
    console.print(f"[green]Wrote {output}[/]. Edit it, then rename to mapping.csv (or pass it via -m).")


@app.command("init-config")
def init_config(
    output: Annotated[Path, typer.Option("--output", "-o")] = Path("config.toml"),
    force: Annotated[bool, typer.Option("--force", "-f", help="Overwrite if the file already exists.")] = False,
) -> None:
    """Write a starter TOML config file."""
    if output.exists() and not force:
        console.print(
            f"[red]{output} already exists.[/] Refusing to overwrite — pass [bold]--force[/] "
            f"if you really want to replace it."
        )
        raise typer.Exit(1)
    output.write_text(
        """log_level = "INFO"

# Each [[apps]] entry is one Entra app registration. Throughput scales ~linearly
# with the number of apps because Graph throttle buckets are per-app. Add more
# entries to push past per-app caps. Each app must be granted the same Graph
# permission (Mail.ReadWrite, application).
[[apps]]
name = "primary"
tenant_id = "contoso.onmicrosoft.com"
client_id = "00000000-0000-0000-0000-000000000000"
# Use ONE of these:
client_secret = "REPLACE_OR_USE_CERT"
# client_certificate_path = "/etc/pstmigrate/cert.pem"

# Uncomment to add a second app for higher aggregate throughput:
# [[apps]]
# name = "secondary"
# tenant_id = "contoso.onmicrosoft.com"
# client_id = "11111111-1111-1111-1111-111111111111"
# client_secret = "..."

[throttle]
max_retries = 8
initial_backoff_seconds = 1.0
max_backoff_seconds = 120.0

[migration]
workers_per_mailbox = 4
max_parallel_mailboxes = 8
target_root_folder = "Imported PST"
fail_fast = false

[paths]
state_dir = ".pstmigrate-state"
work_dir  = ".pstmigrate-work"
log_dir   = "logs"
readpst_binary = "readpst"
""",
        encoding="utf-8",
    )
    console.print(f"[green]Wrote {output}[/]")
    console.print(
        r"Edit the [bold]\[\[apps]][/] section(s), then run "
        "[bold]pstmigrate validate -c config.toml -m mapping.csv[/]"
    )


@app.command()
def validate(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    limit: LimitOpt = None,
) -> None:
    """Check prerequisites: readpst, config, mapping CSV, Graph token, mailboxes.

    Selection flags (mailbox/pst/limit) restrict the resolvability check to only
    the rows you intend to run, so you can pre-flight a single mailbox quickly.
    """
    cfg = _load(config)
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=f"validate_{_run_id()}")

    table = Table(title="Pre-flight Checks")
    table.add_column("Check")
    table.add_column("Result")

    # readpst
    try:
        ver = check_readpst(cfg.paths.readpst_binary)
        table.add_row("readpst available", f"[green]OK[/] {ver}")
    except Exception as e:
        table.add_row("readpst available", f"[red]FAIL[/] {e}")

    # mapping CSV
    rows: list[MappingRow] = []
    try:
        all_rows = load_mapping(mapping)
        rows = _filter_mapping(all_rows, mailboxes=mailbox, pst_names=pst, limit=limit)
        suffix = f" (filtered from {len(all_rows)})" if len(rows) != len(all_rows) else ""
        table.add_row("mapping CSV parses", f"[green]OK[/] {len(rows)} rows{suffix}")
    except Exception as e:
        table.add_row("mapping CSV parses", f"[red]FAIL[/] {e}")

    # PST files exist
    missing = [r.pst_path for r in rows if not r.pst_path.exists()]
    if missing:
        table.add_row("PST files exist", f"[red]FAIL[/] missing: {', '.join(str(p) for p in missing[:5])}")
    elif rows:
        total_gb = sum(r.pst_path.stat().st_size for r in rows) / 1024**3
        table.add_row("PST files exist", f"[green]OK[/] {len(rows)} files, {total_gb:0.2f} GB total")

    # Graph token + mailbox resolution — verify each app independently
    try:
        pool = AppPool(cfg.apps)
        per_app_ok: list[str] = []
        per_app_fail: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for name in pool.names:
                try:
                    graph.get("/me/$metadata" if False else "/$metadata", expect_status=(200,), app_id=name)
                    per_app_ok.append(name)
                except Exception as e:
                    per_app_fail.append(f"{name}: {e}")
        if per_app_fail:
            table.add_row("Graph token (per app)", f"[red]FAIL[/] {'; '.join(per_app_fail[:3])}")
        else:
            table.add_row("Graph token (per app)", f"[green]OK[/] {len(per_app_ok)} app(s): {', '.join(per_app_ok)}")

        # Try opening each unique mailbox using the first app.
        # We deliberately use a Mail.Read*-only endpoint here. /users/{upn}
        # would also work but requires User.Read.All, which we don't grant.
        # /users/{upn}/mailFolders/inbox needs only Mail.ReadWrite (Application)
        # AND surfaces the same failure modes we care about: the user doesn't
        # exist, the mailbox isn't provisioned, or the app's
        # ApplicationAccessPolicy doesn't include this mailbox.
        unique = sorted({r.target_mailbox for r in rows})
        unresolved: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for upn in unique:
                try:
                    graph.get(
                        f"/users/{upn}/mailFolders/inbox",
                        params={"$select": "id,displayName"},
                        app_id=pool.names[0],
                    )
                except Exception as e:
                    unresolved.append(f"{upn} ({e})")
        if unresolved:
            table.add_row(
                "Mailboxes accessible",
                f"[red]FAIL[/] {len(unresolved)}/{len(unique)}: {unresolved[0]}",
            )
        elif unique:
            table.add_row("Mailboxes accessible", f"[green]OK[/] {len(unique)} mailboxes")
    except Exception as e:
        table.add_row("Graph token + connectivity", f"[red]FAIL[/] {e}")

    console.print(table)


@app.command("import")
def run_import(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    limit: LimitOpt = None,
    list_only: ListOnlyOpt = False,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Run the migration.

    Selection examples:

      Single mailbox:    pstmigrate import -c c.toml -m m.csv -M emmak@contoso.onmicrosoft.com
      Several mailboxes: pstmigrate import -c c.toml -m m.csv -M a@x -M b@x -M c@x
      Just a canary:     pstmigrate import -c c.toml -m m.csv -n 1
      One PST file:      pstmigrate import -c c.toml -m m.csv -P jsmith.pst
      Preview selection: pstmigrate import -c c.toml -m m.csv -M a@x --list
    """
    cfg = _load(config)
    run_id = f"import_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(all_rows, mailboxes=mailbox, pst_names=pst, limit=limit)
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    if list_only or len(rows) != len(all_rows):
        _print_selection(
            rows,
            heading=("Dry-run selection" if list_only else "Selected rows (filtered)"),
        )
    if list_only:
        raise typer.Exit(0)

    total_gb = sum(r.pst_path.stat().st_size for r in rows if r.pst_path.exists()) / 1024**3
    pool = AppPool(cfg.apps)
    filter_note = (
        f"  selection = {len(rows)}/{len(all_rows)} rows (filtered)\n"
        if len(rows) != len(all_rows) else ""
    )
    console.print(
        f"\n[bold]About to import[/] {len(rows)} PSTs ({total_gb:0.2f} GB) "
        f"into {len({r.target_mailbox for r in rows})} mailboxes.\n"
        f"{filter_note}"
        f"  app pool = {len(pool)} ({', '.join(pool.names)})\n"
        f"  workers/mailbox = {cfg.migration.workers_per_mailbox}\n"
        f"  parallel mailboxes = {cfg.migration.max_parallel_mailboxes}\n"
        f"  state dir = {cfg.paths.state_dir}\n"
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    orch = Orchestrator(cfg, state, pool)
    reports = orch.run(rows)

    failed = sum(1 for r in reports if r.status != "done" or r.items_failed)
    raise typer.Exit(1 if failed else 0)


@app.command()
def status(
    config: ConfigOpt = None,
) -> None:
    """Show progress from prior runs (read-only)."""
    cfg = _load(config)
    db = cfg.paths.state_dir / "state.sqlite"
    if not db.exists():
        console.print(f"[yellow]No state DB at {db}[/]")
        raise typer.Exit(0)

    state = StateStore(db)
    runs = state.all_runs()
    if not runs:
        console.print("No runs recorded.")
        return

    table = Table(title="PST Runs")
    for col in ("Mailbox", "PST", "Status", "Total", "Done", "Failed", "Skipped", "Last Error"):
        table.add_column(col, overflow="fold")
    for r in runs:
        counts = state.counts_for_run(r["target_mailbox"], r["pst_path"])
        table.add_row(
            r["target_mailbox"],
            Path(r["pst_path"]).name,
            r["status"],
            str(r["items_total"]),
            str(counts["done"]),
            str(counts["failed"]),
            str(counts["skipped"]),
            (r["last_error"] or "")[:80],
        )
    console.print(table)

    breakdown = state.app_breakdown()
    if breakdown:
        per_app = Table(title="Per-app totals (across all PSTs)")
        per_app.add_column("App")
        per_app.add_column("Done", justify="right", style="green")
        per_app.add_column("Failed", justify="right", style="red")
        per_app.add_column("Skipped", justify="right", style="yellow")
        for app_name, statuses in sorted(breakdown.items()):
            per_app.add_row(
                app_name,
                str(statuses.get("done", 0)),
                str(statuses.get("failed", 0)),
                str(statuses.get("skipped", 0)),
            )
        console.print(per_app)


def main() -> None:
    try:
        app()
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted. Resume by re-running `import`.[/]")
        sys.exit(130)
