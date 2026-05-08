from pathlib import Path
from types import SimpleNamespace

from jtet_pstmigrate.config import MappingRow
from jtet_pstmigrate.orchestrator import Orchestrator
from jtet_pstmigrate.pst_reader import ExtractedMessage
from jtet_pstmigrate.state import StateStore


class FakePool:
    def pick(self) -> str:
        return "app-a"


class FakeFolders:
    def __init__(self) -> None:
        self.ensured: list[tuple[str, ...]] = []

    def ensure_path(self, folder_path: tuple[str, ...]) -> str:
        self.ensured.append(folder_path)
        return "folder-a"


class FakeUploader:
    def __init__(self) -> None:
        self.uploaded: list[ExtractedMessage] = []

    def upload(self, msg: ExtractedMessage, folder_id: str, *, app_id: str):
        self.uploaded.append(msg)
        return SimpleNamespace(graph_message_id=f"graph-{len(self.uploaded)}", bytes_uploaded=msg.bytes_)


def test_duplicate_message_rows_are_skipped_by_default(tmp_path: Path) -> None:
    state = StateStore(tmp_path / "state.sqlite")
    row = MappingRow(pst_path=tmp_path / "mail.pst", target_mailbox="alice@example.com")
    state.upsert_message(
        mailbox=row.target_mailbox,
        pst_path=str(row.pst_path),
        source_path=str(tmp_path / "work" / "original.eml"),
        dedupe_key="imid:duplicate",
        status="done",
    )
    duplicate = _message(tmp_path / "work" / "copy.eml")
    folders = FakeFolders()
    uploader = FakeUploader()
    orchestrator = Orchestrator(SimpleNamespace(), state, FakePool())

    outcome = orchestrator._upload_one(duplicate, row, folders, uploader)

    assert outcome == "skipped"
    assert not uploader.uploaded
    assert not folders.ensured
    assert state.counts_for_run(row.target_mailbox, str(row.pst_path))["skipped"] == 1


def test_import_skipped_duplicates_uploads_duplicate_message_rows(tmp_path: Path) -> None:
    state = StateStore(tmp_path / "state.sqlite")
    row = MappingRow(pst_path=tmp_path / "mail.pst", target_mailbox="alice@example.com")
    state.upsert_message(
        mailbox=row.target_mailbox,
        pst_path=str(row.pst_path),
        source_path=str(tmp_path / "work" / "original.eml"),
        dedupe_key="imid:duplicate",
        status="done",
    )
    duplicate = _message(tmp_path / "work" / "copy.eml")
    folders = FakeFolders()
    uploader = FakeUploader()
    orchestrator = Orchestrator(
        SimpleNamespace(),
        state,
        FakePool(),
        import_skipped_duplicates=True,
    )

    outcome = orchestrator._upload_one(duplicate, row, folders, uploader)

    assert outcome == "uploaded"
    assert uploader.uploaded == [duplicate]
    assert folders.ensured == [duplicate.folder_path]
    counts = state.counts_for_run(row.target_mailbox, str(row.pst_path))
    assert counts["done"] == 2
    assert counts["skipped"] == 0


def _message(path: Path) -> ExtractedMessage:
    return ExtractedMessage(
        file_path=path,
        folder_path=("Top of Personal Folders", "Inbox"),
        bytes_=123,
        dedupe_key="imid:duplicate",
        subject="duplicate",
        received=None,
        message_id="duplicate",
    )
