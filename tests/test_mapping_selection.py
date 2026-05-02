from pathlib import Path

from jtet_pstmigrate.mapping import load_mapping
from jtet_pstmigrate.selection import SelectionFilters, select_mapping


def test_load_mapping_parses_required_columns(tmp_path: Path) -> None:
    mapping = tmp_path / "mapping.csv"
    mapping.write_text(
        "PSTPath,TargetMailbox,TargetRootFolder\n"
        "/pst/alice.pst,Alice@Example.COM,Imported\n",
        encoding="utf-8",
    )

    rows = load_mapping(mapping)

    assert len(rows) == 1
    assert rows[0].pst_path == Path("/pst/alice.pst")
    assert rows[0].target_mailbox == "alice@example.com"
    assert rows[0].target_root_folder == "Imported"


def test_select_mapping_include_exclude_and_limit(tmp_path: Path) -> None:
    rows = load_mapping(_mapping_file(tmp_path))

    result = select_mapping(
        rows,
        SelectionFilters(
            pst_names=["archive"],
            exclude_mailboxes=["bob@example.com"],
            limit=1,
        ),
    )

    assert [row.target_mailbox for row in result.rows] == ["alice@example.com"]
    assert result.total_rows == 3
    assert result.filtered is True


def test_select_mapping_warns_for_unmatched_mailbox(tmp_path: Path) -> None:
    rows = load_mapping(_mapping_file(tmp_path))

    result = select_mapping(rows, SelectionFilters(mailboxes=["missing@example.com"]))

    assert result.rows == []
    assert result.warnings == ["no mapping rows for: missing@example.com"]


def _mapping_file(tmp_path: Path) -> Path:
    mapping = tmp_path / "mapping.csv"
    mapping.write_text(
        "PSTPath,TargetMailbox,TargetRootFolder\n"
        "/pst/alice-archive.pst,alice@example.com,Imported\n"
        "/pst/bob-archive.pst,bob@example.com,Imported\n"
        "/pst/cara-live.pst,cara@example.com,Imported\n",
        encoding="utf-8",
    )
    return mapping

