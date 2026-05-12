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
from jtet_pstmigrate.config import AppConfig, MappingRow
from jtet_pstmigrate.entra_setup import (
    PermissionPreset,
    create_migration_apps_with_device_login,
    write_config_for_created_apps,
)
from jtet_pstmigrate.graph_client import GraphClient
from jtet_pstmigrate.log import configure_logging
from jtet_pstmigrate.mapping import load_mapping
from jtet_pstmigrate.orchestrator import Orchestrator
from jtet_pstmigrate.pst_reader import check_readpst
from jtet_pstmigrate.selection import SelectionFilters, select_mapping
from jtet_pstmigrate.services import load_config
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
    return load_config(config_path)


def _run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _filter_mapping(
    rows: list[MappingRow],
    *,
    mailboxes: list[str] | None,
    pst_names: list[str] | None,
    exclude_mailboxes: list[str] | None = None,
    exclude_pst_names: list[str] | None = None,
    limit: int | None,
) -> list[MappingRow]:
    """Apply selection filters in order:

      1. ``mailboxes``  -- include only these (UPN exact match, case-insensitive)
      2. ``pst_names``  -- include only PSTs whose filename contains one of these
                           (substring, case-insensitive)
      3. ``exclude_mailboxes`` -- drop these UPNs (exact match, case-insensitive)
      4. ``exclude_pst_names`` -- drop PSTs matching these substrings
      5. ``limit`` -- after all filtering, take only the first N rows

    Excludes win over includes: ``-M user@x -X user@x`` returns nothing.
    Excludes work on their own too -- you don't have to pass any include
    filter, just ``-X user@x`` to "everything except this user".
    """
    result = select_mapping(
        rows,
        SelectionFilters(
            mailboxes=mailboxes,
            pst_names=pst_names,
            exclude_mailboxes=exclude_mailboxes,
            exclude_pst_names=exclude_pst_names,
            limit=limit,
        ),
    )
    for warning in result.warnings:
        console.print(f"[yellow]Warning:[/] {warning}")
    return result.rows


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
ExcludeMailboxOpt = Annotated[
    list[str] | None,
    typer.Option(
        "--exclude-mailbox", "-X",
        help=(
            "Skip these target mailboxes (UPN). Repeat for multi-select. "
            "Useful when bulk-running everyone EXCEPT users you've already "
            "finished. Applied after --mailbox / --pst includes."
        ),
    ),
]
ExcludePstFilterOpt = Annotated[
    list[str] | None,
    typer.Option(
        "--exclude-pst",
        help=(
            "Skip PSTs whose filename contains any of these substrings. "
            "Repeat for multi-select."
        ),
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


@app.command("setup-entra")
def setup_entra(
    tenant: Annotated[str, typer.Option("--tenant", "-t", help="Tenant ID or domain")] = ...,
    config: Annotated[Path, typer.Option("--config", "-c", help="Config file to write")] = Path("config.toml"),
    apps: Annotated[int, typer.Option("--apps", help="Number of app registrations to create")] = 1,
    app_prefix: Annotated[str, typer.Option("--app-prefix", help="Display-name prefix")] = "pstmigrate",
    secret_days: Annotated[int, typer.Option("--secret-days", help="Client secret lifetime in days")] = 180,
    permission_preset: Annotated[
        PermissionPreset,
        typer.Option("--permission-preset", help="Graph permission preset: full or mail"),
    ] = "full",
    force: Annotated[bool, typer.Option("--force", "-f", help="Overwrite existing config file")] = False,
) -> None:
    """Create Entra app registrations and write their credentials to config.toml."""
    if config.exists() and not force:
        console.print(
            f"[red]{config} already exists.[/] Refusing to overwrite — pass [bold]--force[/] "
            f"if you want setup to replace it."
        )
        raise typer.Exit(1)

    def show_device_flow(flow: dict) -> None:
        console.print("")
        console.print("[yellow]Admin sign-in required[/]")
        console.print(f"Open: [green]{flow.get('verification_uri')}[/]")
        console.print(f"Code: [bold green]{flow.get('user_code')}[/]")
        console.print("")

    def show_event(message: str) -> None:
        console.print(f"[cyan]{message}[/]")

    created_apps = create_migration_apps_with_device_login(
        tenant_id=tenant,
        app_prefix=app_prefix,
        app_count=apps,
        secret_lifetime_days=secret_days,
        permission_preset=permission_preset,
        device_flow_callback=show_device_flow,
        event_callback=show_event,
    )
    write_config_for_created_apps(config, created_apps)

    table = Table(title="Created Entra Apps")
    table.add_column("Name")
    table.add_column("Client ID")
    table.add_column("Secret Expires")
    table.add_column("Permissions")
    for created in created_apps:
        table.add_row(
            created.name,
            created.client_id,
            created.secret_expires_at,
            ", ".join(created.permissions),
        )
    console.print(table)
    console.print(f"[green]Wrote {config}[/]")
    console.print("[yellow]Client secrets are stored in the config file. Protect it like a credential.[/]")


@app.command()
def validate(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    limit: LimitOpt = None,
) -> None:
    """Check prerequisites: readpst, config, mapping CSV, Graph token, mailboxes.

    Selection flags (mailbox/pst/exclude-mailbox/exclude-pst/limit) restrict the
    resolvability check to only the rows you intend to run, so you can pre-flight
    a single mailbox quickly or skip mailboxes you've already finished.
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
        rows = _filter_mapping(
            all_rows,
            mailboxes=mailbox,
            pst_names=pst,
            exclude_mailboxes=exclude_mailbox,
            exclude_pst_names=exclude_pst,
            limit=limit,
        )
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
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    limit: LimitOpt = None,
    list_only: ListOnlyOpt = False,
    import_skipped_duplicates: Annotated[
        bool,
        typer.Option(
            "--import-skipped-duplicates",
            help=(
                "Upload source messages whose Message-ID already has a completed row. "
                "Use with --mailbox/--pst for remediation re-runs that should import "
                "copies previously recorded as skipped duplicates."
            ),
        ),
    ] = False,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Run the migration.

    Selection examples:

      Single mailbox:    pstmigrate import -c c.toml -m m.csv -M emmak@contoso.onmicrosoft.com
      Several mailboxes: pstmigrate import -c c.toml -m m.csv -M a@x -M b@x -M c@x
      Skip finished:     pstmigrate import -c c.toml -m m.csv -X tinad@x -X allyson@x
      Just a canary:     pstmigrate import -c c.toml -m m.csv -n 1
      One PST file:      pstmigrate import -c c.toml -m m.csv -P jsmith.pst
      Preview selection: pstmigrate import -c c.toml -m m.csv -M a@x --list
      Import skipped duplicate copies for one mailbox:
                          pstmigrate import -c c.toml -m m.csv -M a@x --import-skipped-duplicates
    """
    cfg = _load(config)
    run_id = f"import_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=limit,
    )
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
        f"  duplicate mode = {'import skipped duplicates' if import_skipped_duplicates else 'skip duplicates'}\n"
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    orch = Orchestrator(cfg, state, pool, import_skipped_duplicates=import_skipped_duplicates)
    reports = orch.run(rows)

    failed = sum(1 for r in reports if r.status != "done" or r.items_failed)
    raise typer.Exit(1 if failed else 0)


@app.command("import-calendar")
def run_import_calendar(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    limit: LimitOpt = None,
    list_only: ListOnlyOpt = False,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Import calendar appointments from PSTs into each mailbox's default calendar.

    This is a separate pass from ``import`` (mail). It uses the same mapping
    CSV but extracts only appointments via ``readpst -t a`` into a sibling
    work directory, so re-running it doesn't invalidate or trigger a re-run
    of the mail extraction. State for calendar items lives in the
    ``non_mail_items`` table, which is independent of the ``messages``
    dedupe table -- so a calendar UID and a mail Message-ID can never
    collide and a 'done' calendar event won't be re-uploaded if you also
    re-run ``import``.

    All events go into the user's default calendar; sub-calendar structure
    inside the PST (custom calendars the user kept) is flattened. Recurring
    series become single-occurrence events for v1; the original RRULE text
    is preserved in the event body so no data is silently dropped.
    """
    cfg = _load(config)
    run_id = f"import-calendar_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=limit,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    if list_only or len(rows) != len(all_rows):
        _print_selection(
            rows,
            heading=("Dry-run selection (calendar)" if list_only else "Selected rows (calendar, filtered)"),
        )
    if list_only:
        raise typer.Exit(0)

    pool = AppPool(cfg.apps)
    filter_note = (
        f"  selection = {len(rows)}/{len(all_rows)} rows (filtered)\n"
        if len(rows) != len(all_rows) else ""
    )
    console.print(
        f"\n[bold]About to import calendar items[/] from {len(rows)} PSTs "
        f"into {len({r.target_mailbox for r in rows})} mailbox calendar(s).\n"
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
    reports = orch.run_calendar(rows)

    failed = sum(1 for r in reports if r.status != "done" or r.items_failed)
    raise typer.Exit(1 if failed else 0)


@app.command("purge-calendar")
def run_purge_calendar(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Delete previously-imported calendar events and clear their state.

    Use this when a calendar-import bug needs to be re-run cleanly: after
    purge, the next ``import-calendar`` re-uploads every event from
    scratch with the corrected code path. Mail state is untouched.

    The purge is keyed on (mailbox, pst_path) just like ``import-calendar``,
    so you can scope to a single mailbox/PST while you iterate on a fix
    and leave other mailboxes alone.
    """
    cfg = _load(config)
    run_id = f"purge-calendar_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=None,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    pool = AppPool(cfg.apps)

    # Show the user what we're about to wipe before doing it.
    total = 0
    for row in rows:
        items = state.list_done_items(row.target_mailbox, str(row.pst_path), "event")
        if items:
            console.print(
                f"  {row.target_mailbox} | {row.pst_path.name}: "
                f"[yellow]{len(items)}[/] event(s) to purge"
            )
            total += len(items)

    if total == 0:
        console.print("[green]Nothing to purge.[/] No 'done' calendar rows for the selected scope.")
        raise typer.Exit(0)

    console.print(
        f"\n[bold red]About to DELETE {total} calendar event(s) from Graph[/] "
        f"and remove their state rows.\n"
        f"This is irreversible -- the events will be gone from the destination "
        f"mailbox and ``import-calendar`` will re-create them from the PST extract."
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    from jtet_pstmigrate.orchestrator import Orchestrator
    orch = Orchestrator(cfg, state, pool)
    deleted, missing, errors = orch.purge_calendar(rows)

    console.print(
        f"\n[bold]Purge complete:[/] "
        f"deleted=[green]{deleted}[/]  "
        f"already-gone=[yellow]{missing}[/]  "
        f"errors=[red]{errors}[/]"
    )
    raise typer.Exit(1 if errors else 0)


@app.command("import-contacts")
def run_import_contacts(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    limit: LimitOpt = None,
    list_only: ListOnlyOpt = False,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Import contacts from PSTs into each mailbox's default contact folder.

    Separate pass from ``import`` (mail) and ``import-calendar``. Uses the
    same mapping CSV but extracts only contacts via ``readpst -t c`` into a
    sibling work directory (``__contacts``), so re-running it doesn't
    invalidate or trigger the mail/calendar extractions. State for contact
    items lives in the ``non_mail_items`` table with ``item_type='contact'``,
    keyed by vCard UID (or FN+email when no UID is present).

    All contacts go into the user's default contact folder; sub-folder
    structure inside the PST is flattened. Required Graph permission:
    ``Contacts.ReadWrite`` (Application) on every app in the pool, with
    admin consent. A 403 here means a worker app is missing that grant.
    """
    cfg = _load(config)
    run_id = f"import-contacts_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=limit,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    if list_only or len(rows) != len(all_rows):
        _print_selection(
            rows,
            heading=("Dry-run selection (contacts)" if list_only else "Selected rows (contacts, filtered)"),
        )
    if list_only:
        raise typer.Exit(0)

    pool = AppPool(cfg.apps)
    filter_note = (
        f"  selection = {len(rows)}/{len(all_rows)} rows (filtered)\n"
        if len(rows) != len(all_rows) else ""
    )
    console.print(
        f"\n[bold]About to import contacts[/] from {len(rows)} PSTs "
        f"into {len({r.target_mailbox for r in rows})} mailbox contact folder(s).\n"
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
    reports = orch.run_contacts(rows)

    failed = sum(1 for r in reports if r.status != "done" or r.items_failed)
    raise typer.Exit(1 if failed else 0)


@app.command("purge-contacts")
def run_purge_contacts(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Delete previously-imported contacts and clear their state.

    Use this when a contacts-import bug needs a clean re-run: after purge,
    the next ``import-contacts`` re-uploads every contact from scratch with
    the corrected code path. Mail and calendar state are untouched.
    """
    cfg = _load(config)
    run_id = f"purge-contacts_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=None,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    pool = AppPool(cfg.apps)

    total = 0
    for row in rows:
        items = state.list_done_items(row.target_mailbox, str(row.pst_path), "contact")
        if items:
            console.print(
                f"  {row.target_mailbox} | {row.pst_path.name}: "
                f"[yellow]{len(items)}[/] contact(s) to purge"
            )
            total += len(items)

    if total == 0:
        console.print("[green]Nothing to purge.[/] No 'done' contact rows for the selected scope.")
        raise typer.Exit(0)

    console.print(
        f"\n[bold red]About to DELETE {total} contact(s) from Graph[/] "
        f"and remove their state rows.\n"
        f"This is irreversible -- the contacts will be gone from the destination "
        f"mailbox and ``import-contacts`` will re-create them from the PST extract."
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    from jtet_pstmigrate.orchestrator import Orchestrator
    orch = Orchestrator(cfg, state, pool)
    deleted, missing, errors = orch.purge_contacts(rows)

    console.print(
        f"\n[bold]Purge complete:[/] "
        f"deleted=[green]{deleted}[/]  "
        f"already-gone=[yellow]{missing}[/]  "
        f"errors=[red]{errors}[/]"
    )
    raise typer.Exit(1 if errors else 0)


@app.command("purge-mail")
def run_purge_mail(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Wipe every mail folder + message in the selected mailboxes (except Deleted Items).

    Walks ``/users/{upn}/mailFolders`` top-down for each target
    mailbox. For every folder it tries ``DELETE
    /users/{upn}/mailFolders/{id}``: on success Graph cascades the
    entire subtree -- every message and every child folder beneath
    that folder -- in a single round-trip. For "distinguished" folders
    that Exchange refuses to delete (Inbox, Sent Items, Drafts,
    Outbox, Junk Email, Conversation History, ...), it falls back to
    recursing into their children and draining their own messages
    with parallel per-message DELETE.

    Net result: after this command runs, the only folders left are
    the ones Exchange protects (and Deleted Items), and they are all
    empty. The mailbox is ready for a clean re-import that recreates
    whatever folder structure your PST had.

    Does NOT consult the local state database. The destination
    mailbox is the sole source of truth, so this works equally well
    after ``reset-state``, after a botched run, or against a mailbox
    you never imported into with this tool. Calendar and contact
    items are untouched.

    Mailboxes are de-duplicated, so multiple PST rows for the same
    UPN only wipe the mailbox once.

    Required Graph permission: ``Mail.ReadWrite`` (Application) on
    every app in the pool -- already required by ``import``, so this
    command needs no extra admin grants.

    Bulk usage (everyone except the pilot mailboxes you've already
    finished):

      pstmigrate purge-mail -c c.toml -m m.csv \
          -X tinad@x -X allysonp@x -X debbiep@x -y
    """
    cfg = _load(config)
    run_id = f"purge-mail_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=None,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    pool = AppPool(cfg.apps)

    # De-dup mailboxes for the confirmation prompt; the orchestrator
    # also de-dups internally.
    seen: set[str] = set()
    unique_mailboxes: list[str] = []
    for row in rows:
        key = row.target_mailbox.lower()
        if key not in seen:
            seen.add(key)
            unique_mailboxes.append(row.target_mailbox)

    console.print(
        f"\n[bold red]About to wipe all mail folders + messages[/] "
        f"(except Deleted Items) from {len(unique_mailboxes)} "
        f"mailbox(es):"
    )
    for mb in unique_mailboxes:
        console.print(f"  - {mb}")
    console.print(
        f"\n  Custom folders are cascade-deleted (folder + all "
        f"contents in one round-trip).\n"
        f"  Distinguished folders (Inbox / Sent Items / Drafts / "
        f"Outbox / Junk Email / etc.) cannot be deleted; their "
        f"messages are drained instead.\n"
        f"  app pool = {len(pool)} ({', '.join(pool.names)})\n"
        f"  workers/mailbox = {cfg.migration.workers_per_mailbox}\n"
        f"  parallel mailboxes = {cfg.migration.max_parallel_mailboxes}\n"
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    from jtet_pstmigrate.orchestrator import Orchestrator
    orch = Orchestrator(cfg, state, pool)
    deleted, missing, errors = orch.purge_mail(rows)

    console.print(
        f"\n[bold]Mail purge complete:[/] "
        f"deleted=[green]{deleted}[/]  "
        f"already-gone=[yellow]{missing}[/]  "
        f"errors=[red]{errors}[/]"
    )
    raise typer.Exit(1 if errors else 0)


@app.command("import-all")
def run_import_all(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    limit: LimitOpt = None,
    list_only: ListOnlyOpt = False,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Run mail + calendar + contacts back-to-back for the selected mailboxes.

    Equivalent to running ``import`` then ``import-calendar`` then
    ``import-contacts`` with the same mapping and filters, but with one
    confirmation prompt and one consolidated exit code.

    The three phases share state but run independently: a mail failure on
    one mailbox does not skip that mailbox's calendar or contacts. Each
    phase is also idempotent against the existing dedup tables, so running
    this against a partly-imported mailbox safely no-ops the parts that
    already finished.

    Use this for production/bulk runs once the per-pass commands have
    been verified on a pilot mailbox. For iterative debugging keep using
    the dedicated subcommands -- those produce the same state, just with
    smaller blast radius if something goes wrong.

    Bulk-run example, skipping users that are already done:

      pstmigrate import-all -c c.toml -m m.csv \
          -X tinad@contoso.onmicrosoft.com \
          -X allysonp@contoso.onmicrosoft.com -y
    """
    cfg = _load(config)
    run_id = f"import-all_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=limit,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    if list_only or len(rows) != len(all_rows):
        _print_selection(
            rows,
            heading=("Dry-run selection (mail + calendar + contacts)"
                     if list_only else
                     "Selected rows (mail + calendar + contacts, filtered)"),
        )
    if list_only:
        raise typer.Exit(0)

    pool = AppPool(cfg.apps)
    filter_note = (
        f"  selection = {len(rows)}/{len(all_rows)} rows (filtered)\n"
        if len(rows) != len(all_rows) else ""
    )
    console.print(
        f"\n[bold]About to import MAIL + CALENDAR + CONTACTS[/] "
        f"from {len(rows)} PSTs into "
        f"{len({r.target_mailbox for r in rows})} mailbox(es).\n"
        f"{filter_note}"
        f"  app pool = {len(pool)} ({', '.join(pool.names)})\n"
        f"  workers/mailbox = {cfg.migration.workers_per_mailbox}\n"
        f"  parallel mailboxes = {cfg.migration.max_parallel_mailboxes}\n"
        f"  state dir = {cfg.paths.state_dir}\n"
        f"\n"
        f"Required Graph Application permissions on every app in the pool:\n"
        f"  - Mail.ReadWrite      (mail phase)\n"
        f"  - Calendars.ReadWrite (calendar phase)\n"
        f"  - Contacts.ReadWrite  (contacts phase)\n"
        f"A 403 in any phase usually means a worker app is missing the\n"
        f"corresponding grant. The other phases will still run.\n"
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")
    orch = Orchestrator(cfg, state, pool)

    # Phases run sequentially because each builds its own GraphClient and
    # ThreadPoolExecutor; running them concurrently would just multiply
    # per-mailbox throttling pressure without finishing any faster on a
    # single tenant. We deliberately do NOT short-circuit the next phase
    # when the previous one has failures -- mail and calendar/contacts
    # are independent in the source PST, and if one phase has bad data
    # we still want the others to land.
    console.print("\n[bold cyan]Phase 1/3 -- mail[/]")
    mail_reports = orch.run(rows)

    console.print("\n[bold cyan]Phase 2/3 -- calendar[/]")
    cal_reports = orch.run_calendar(rows)

    console.print("\n[bold cyan]Phase 3/3 -- contacts[/]")
    con_reports = orch.run_contacts(rows)

    def _failed(reports: list) -> int:
        return sum(1 for r in reports if r.status != "done" or r.items_failed)

    failed = _failed(mail_reports) + _failed(cal_reports) + _failed(con_reports)

    console.print(
        f"\n[bold]All phases complete.[/]  "
        f"mail-failed=[red]{_failed(mail_reports)}[/]  "
        f"calendar-failed=[red]{_failed(cal_reports)}[/]  "
        f"contacts-failed=[red]{_failed(con_reports)}[/]"
    )
    raise typer.Exit(1 if failed else 0)


@app.command("reset-state")
def run_reset_state(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m")] = ...,  # type: ignore[assignment]
    mailbox: MailboxOpt = None,
    pst: PstFilterOpt = None,
    exclude_mailbox: ExcludeMailboxOpt = None,
    exclude_pst: ExcludePstFilterOpt = None,
    yes: Annotated[bool, typer.Option("--yes", "-y", help="Skip confirmation")] = False,
) -> None:
    """Clear local dedup/run state for the selected scope. Does NOT touch Graph.

    Wipes the matching rows from these tables in ``state.sqlite``:
      - ``messages``        (mail dedup keyed by Message-ID)
      - ``non_mail_items``  (calendar/contacts dedup keyed by UID/FN+email)
      - ``pst_runs``        (per-mailbox-per-PST run history)
      - ``folder_map``      (cached Graph folder IDs per mailbox)

    [bold red]DANGER:[/] this only clears LOCAL state. Items already
    uploaded to Graph stay where they are. Running ``import`` /
    ``import-calendar`` / ``import-contacts`` / ``import-all`` afterwards
    WILL upload everything again, which means duplicate emails, calendar
    events, and contacts in the destination mailbox -- the dedup tables
    are how we prevent that, and you just emptied them.

    Intended for: starting from scratch after manually clearing the
    destination mailboxes (mailbox reset, delete-and-recreate, manual
    purge in OWA, etc.). For a safe reset that also removes items from
    Graph, use ``purge-calendar`` / ``purge-contacts`` instead -- there's
    no equivalent for mail because mail-purge of large mailboxes is
    expensive (one DELETE per message); if you need that, do it on the
    Exchange side and then run this to clear local tracking.

    Default scope is every row in the mapping CSV. Narrow with
    ``--mailbox`` / ``--pst`` to reset specific users, or use
    ``--exclude-mailbox`` / ``--exclude-pst`` to reset everyone EXCEPT
    a few -- handy when you've finished a couple of pilot users and
    want to clear local state for the rest before a bulk re-run.
    """
    cfg = _load(config)
    run_id = f"reset-state_{_run_id()}"
    configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=run_id)

    all_rows = load_mapping(mapping)
    if not all_rows:
        console.print("[red]Empty mapping CSV[/]")
        raise typer.Exit(1)

    rows = _filter_mapping(
        all_rows,
        mailboxes=mailbox,
        pst_names=pst,
        exclude_mailboxes=exclude_mailbox,
        exclude_pst_names=exclude_pst,
        limit=None,
    )
    if not rows:
        console.print("[red]Selection produced 0 rows. Nothing to do.[/]")
        raise typer.Exit(1)

    state = StateStore(cfg.paths.state_dir / "state.sqlite")

    # Show the user what we're about to wipe before doing it.
    total = {"messages": 0, "non_mail_items": 0, "pst_runs": 0, "folder_map": 0}
    affected_mailboxes: set[str] = set()
    per_row_counts: list[tuple[str, str, dict[str, int]]] = []
    for row in rows:
        counts = state.count_scope(row.target_mailbox, str(row.pst_path))
        if any(counts.values()):
            per_row_counts.append((row.target_mailbox, row.pst_path.name, counts))
        for k, v in counts.items():
            total[k] += v
        affected_mailboxes.add(row.target_mailbox)

    folder_total = sum(state.count_folder_map(m) for m in affected_mailboxes)
    total["folder_map"] = folder_total

    if not any(total.values()):
        console.print("[green]Nothing to reset.[/] No state rows for the selected scope.")
        raise typer.Exit(0)

    if per_row_counts:
        table = Table(title="Rows to clear (per mapping row)")
        table.add_column("Mailbox", overflow="fold")
        table.add_column("PST", overflow="fold")
        table.add_column("messages", justify="right")
        table.add_column("non_mail_items", justify="right")
        table.add_column("pst_runs", justify="right")
        for mbx, pst_name, c in per_row_counts:
            table.add_row(
                mbx, pst_name,
                str(c["messages"]), str(c["non_mail_items"]), str(c["pst_runs"]),
            )
        console.print(table)
    if folder_total:
        console.print(
            f"\nfolder_map cache: [yellow]{folder_total}[/] row(s) "
            f"across {len(affected_mailboxes)} mailbox(es) (safe to drop -- "
            f"folder lookups will re-resolve via Graph on the next import)."
        )

    console.print(
        f"\n[bold red]About to WIPE local state[/] for "
        f"{len(rows)} mapping row(s) "
        f"({len(affected_mailboxes)} mailbox(es)):\n"
        f"  messages       = {total['messages']}\n"
        f"  non_mail_items = {total['non_mail_items']}\n"
        f"  pst_runs       = {total['pst_runs']}\n"
        f"  folder_map     = {total['folder_map']}\n"
        f"\n"
        f"[red]Items already uploaded to Graph remain in the destination "
        f"mailboxes.[/] Re-running an import after this WILL create "
        f"duplicates unless you've separately cleared the destination side."
    )
    if not yes and not typer.confirm("Proceed?", default=False):
        raise typer.Exit(0)

    cleared = {"messages": 0, "non_mail_items": 0, "pst_runs": 0, "folder_map": 0}
    for row in rows:
        result = state.clear_scope(row.target_mailbox, str(row.pst_path))
        for k, v in result.items():
            cleared[k] += v
    for m in affected_mailboxes:
        cleared["folder_map"] += state.clear_folder_map(m)

    console.print(
        f"\n[bold]Reset complete:[/]  "
        f"messages=[yellow]{cleared['messages']}[/]  "
        f"non_mail_items=[yellow]{cleared['non_mail_items']}[/]  "
        f"pst_runs=[yellow]{cleared['pst_runs']}[/]  "
        f"folder_map=[yellow]{cleared['folder_map']}[/]"
    )


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


@app.command()
def web(
    config: ConfigOpt = None,
    mapping: Annotated[Path, typer.Option("--mapping", "-m", help="Mapping CSV file")] = Path("mapping.csv"),
    host: Annotated[str, typer.Option("--host", help="Bind address")] = "127.0.0.1",
    port: Annotated[int, typer.Option("--port", help="Bind port")] = 8765,
) -> None:
    """Start the local web admin GUI."""
    try:
        import uvicorn

        from jtet_pstmigrate.webapp import create_app
    except Exception as e:
        console.print(f"[red]Web UI dependencies are not available:[/] {e}")
        console.print("Install them with: [bold]pip install -e .[web][/]")
        raise typer.Exit(1) from e

    if mapping and not mapping.exists():
        console.print(f"[yellow]Warning:[/] mapping file does not exist yet: {mapping}")

    console.print(f"[green]Starting pstmigrate web UI[/] at http://{host}:{port}")
    # The browser polls job panels frequently while work is running; suppress
    # Uvicorn access logs so the terminal stays useful for migration logs.
    uvicorn.run(create_app(config_path=config, mapping_path=mapping), host=host, port=port, access_log=False)


def main() -> None:
    try:
        app()
    except KeyboardInterrupt:
        console.print("\n[yellow]Interrupted. Resume by re-running `import`.[/]")
        sys.exit(130)
