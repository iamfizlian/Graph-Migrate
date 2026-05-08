"""Local FastAPI web UI for pstmigrate."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Annotated, Any

from jinja2 import DictLoader, Environment, select_autoescape

from jtet_pstmigrate.jobs import JobKind, JobManager, JobSpec
from jtet_pstmigrate.mapping import load_mapping
from jtet_pstmigrate.reports import StateQueries
from jtet_pstmigrate.selection import SelectionFilters, select_mapping
from jtet_pstmigrate.services import load_config
from jtet_pstmigrate.state import StateStore

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover
    import tomli as tomllib

try:
    from fastapi import FastAPI, Form
    from fastapi.responses import HTMLResponse, RedirectResponse
    from fastapi.staticfiles import StaticFiles
except ImportError as e:  # pragma: no cover - exercised by the CLI command.
    raise RuntimeError(
        "The web UI dependencies are not installed. Run `pip install -e .[web]`."
    ) from e


TEMPLATES: dict[str, str] = {
    "layout.html": """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>pstmigrate</title>
  <script src="/static/htmx.min.js"></script>
  <style>
    :root { color-scheme: light; --ink:#18212f; --muted:#627083; --line:#d9e1ea; --bg:#f5f7fa; --panel:#fff; --accent:#176b87; --warn:#a15c00; --bad:#b42318; --good:#087443; }
    * { box-sizing: border-box; }
    body { margin:0; font:14px/1.45 system-ui, -apple-system, Segoe UI, sans-serif; color:var(--ink); background:var(--bg); }
    header { display:flex; align-items:center; gap:18px; padding:12px 20px; background:#122033; color:white; }
    header strong { font-size:16px; }
    nav { display:flex; gap:4px; flex-wrap:wrap; }
    nav a { color:#d9e8f0; text-decoration:none; padding:7px 10px; border-radius:6px; }
    nav a:hover { background:#203651; color:white; }
    main { max-width:1180px; margin:0 auto; padding:22px; }
    h1 { font-size:24px; margin:0 0 16px; }
    h2 { font-size:17px; margin:0 0 12px; }
    .grid { display:grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap:12px; }
    .panel, .card { background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:14px; }
    .panel { margin-bottom:16px; }
    .metric { font-size:24px; font-weight:700; }
    .muted { color:var(--muted); }
    .ok { color:var(--good); font-weight:700; }
    .bad { color:var(--bad); font-weight:700; }
    .warn { color:var(--warn); font-weight:700; }
    table { width:100%; border-collapse:collapse; background:var(--panel); border:1px solid var(--line); border-radius:8px; overflow:hidden; }
    th, td { padding:9px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; }
    th { background:#eaf0f6; font-size:12px; text-transform:uppercase; color:#435366; }
    tr:last-child td { border-bottom:0; }
    label { display:block; font-weight:650; margin:10px 0 4px; }
    input, select, textarea { width:100%; padding:8px 9px; border:1px solid #bac6d4; border-radius:6px; font:inherit; background:white; }
    textarea { min-height:90px; resize:vertical; }
    button, .button { display:inline-flex; align-items:center; justify-content:center; min-height:36px; border:1px solid #0f5d77; border-radius:6px; background:var(--accent); color:white; padding:7px 12px; text-decoration:none; font-weight:650; cursor:pointer; }
    button.secondary, .button.secondary { background:white; color:var(--accent); }
    .row { display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap:12px; }
    .actions { display:flex; gap:8px; flex-wrap:wrap; margin-top:12px; }
    .progress-wrap { margin:12px 0; }
    .progress-bar { height:18px; border-radius:999px; overflow:hidden; background:#d9e1ea; border:1px solid #bac6d4; }
    .progress-fill { height:100%; background:linear-gradient(90deg, #176b87, #25a18e); transition:width .25s ease; }
    .event-list { margin:8px 0 0; padding-left:20px; max-height:260px; overflow:auto; }
    .event-list li { margin:3px 0; }
    pre { white-space:pre-wrap; overflow:auto; background:#0f1720; color:#e8eef5; padding:12px; border-radius:8px; max-height:520px; }
    @media (max-width: 760px) { .grid, .row { grid-template-columns:1fr; } main { padding:14px; } }
  </style>
</head>
<body>
  <header>
    <strong>pstmigrate</strong>
    <nav>
      <a href="/">Dashboard</a><a href="/config">Config</a><a href="/mapping">Mapping</a>
      <a href="/validate">Validate</a><a href="/run">Run</a><a href="/status">Status</a>
      <a href="/logs">Logs</a><a href="/tools">Tools</a>
    </nav>
  </header>
  <main>{% block body %}{% endblock %}</main>
</body>
</html>
""",
    "dashboard.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Dashboard</h1>
<section class="grid">
  <div class="card"><div class="muted">Runs</div><div class="metric">{{ totals.runs }}</div></div>
  <div class="card"><div class="muted">Done</div><div class="metric ok">{{ totals.done }}</div></div>
  <div class="card"><div class="muted">Failed</div><div class="metric bad">{{ totals.failed }}</div></div>
  <div class="card"><div class="muted">Skipped</div><div class="metric warn">{{ totals.skipped }}</div></div>
</section>
<section class="panel" style="margin-top:16px">
  <h2>Recent Jobs</h2>
  <div hx-get="/partials/jobs" hx-trigger="every 3s">{{ jobs_html|safe }}</div>
</section>
<section class="panel">
  <h2>Configuration</h2>
  <p><strong>Config:</strong> {{ config_path or "(default/env)" }}</p>
  <p><strong>Mapping:</strong> {{ mapping_path or "(not set)" }}</p>
</section>
{% endblock %}
""",
    "jobs.html": """
<table>
  <thead><tr><th>Job</th><th>Kind</th><th>Status</th><th>Phase</th><th>Rows</th><th>Items</th><th>Error</th></tr></thead>
  <tbody>
  {% for job in jobs %}
    {% set counts = job.item_counts %}
    <tr>
      <td><a href="/jobs/{{ job.job_id }}">{{ job.job_id }}</a></td>
      <td>{{ job.kind }}</td>
      <td class="{{ 'ok' if job.status == 'done' else 'bad' if job.status == 'failed' else 'warn' if job.status == 'blocked' else '' }}">{{ job.status }}</td>
      <td>{{ job.phase }}</td>
      <td>{{ job.selected_rows }}</td>
      <td>{{ counts.uploaded }} up / {{ counts.failed }} fail / {{ counts.skipped }} skip</td>
      <td>{{ job.last_error or "" }}</td>
    </tr>
  {% else %}
    <tr><td colspan="7" class="muted">No jobs yet.</td></tr>
  {% endfor %}
  </tbody>
</table>
""",
    "config.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Config</h1>
<section class="panel">
  <p><strong>Path:</strong> {{ config_path or "(default/env)" }}</p>
  {% if config_error %}<p class="bad">{{ config_error }}</p>{% endif %}
  <p><strong>Apps:</strong> {{ app_count }} | <strong>Workers/mailbox:</strong> {{ workers }} | <strong>Parallel mailboxes:</strong> {{ parallel }}</p>
</section>

<section class="panel">
  <h2>Microsoft Graph App</h2>
  <form method="post" action="/config">
    <div class="row">
      <div><label>App name</label><input name="app_name" value="{{ form.app_name }}"></div>
      <div><label>Tenant ID or domain</label><input name="tenant_id" value="{{ form.tenant_id }}" required></div>
    </div>
    <label>Client ID</label>
    <input name="client_id" value="{{ form.client_id }}" required>
    <div class="row">
      <div>
        <label>Client secret</label>
        <input name="client_secret" type="password" value="" placeholder="{{ 'Leave blank to keep existing secret' if form.has_secret else 'Paste secret value' }}">
      </div>
      <div>
        <label>Certificate path</label>
        <input name="client_certificate_path" value="{{ form.client_certificate_path }}" placeholder="/path/to/cert.pem">
      </div>
    </div>
    <div class="row">
      <div><label>Workers per mailbox</label><input name="workers_per_mailbox" type="number" min="1" max="16" value="{{ form.workers_per_mailbox }}"></div>
      <div><label>Parallel mailboxes</label><input name="max_parallel_mailboxes" type="number" min="1" max="64" value="{{ form.max_parallel_mailboxes }}"></div>
    </div>
    <div class="row">
      <div><label>State directory</label><input name="state_dir" value="{{ form.state_dir }}"></div>
      <div><label>Work directory</label><input name="work_dir" value="{{ form.work_dir }}"></div>
    </div>
    <div class="row">
      <div><label>Log directory</label><input name="log_dir" value="{{ form.log_dir }}"></div>
      <div><label>readpst binary</label><input name="readpst_binary" value="{{ form.readpst_binary }}"></div>
    </div>
    <div class="actions"><button>Save Config</button></div>
  </form>
</section>

<section class="panel">
  <h2>Current File</h2>
<pre>{{ config_text }}</pre>
</section>
{% endblock %}
""",
    "mapping.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Mapping</h1>
<section class="panel">
  <form method="get" action="/mapping">
    <div class="row">
      <div><label>Mailbox contains/exact UPN</label><input name="mailbox" value="{{ mailbox }}"></div>
      <div><label>PST filename contains</label><input name="pst" value="{{ pst }}"></div>
    </div>
    <div class="actions"><button>Preview</button><a class="button secondary" href="/mapping">Clear</a></div>
  </form>
</section>
<table>
  <thead><tr><th>#</th><th>PST</th><th>Exists</th><th>Mailbox</th><th>Root Folder</th></tr></thead>
  <tbody>
  {% for row in rows %}
    <tr><td>{{ loop.index }}</td><td>{{ row.pst_path }}</td><td>{{ "yes" if row.pst_path.exists() else "no" }}</td><td>{{ row.target_mailbox }}</td><td>{{ row.target_root_folder or "(root)" }}</td></tr>
  {% else %}
    <tr><td colspan="5" class="muted">No rows.</td></tr>
  {% endfor %}
  </tbody>
</table>
{% endblock %}
""",
    "validate.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Validate</h1>
<section class="panel">
  <form method="post" action="/jobs">
    <input type="hidden" name="kind" value="validate">
    {% include "mapping_select.html" %}
    <div class="actions"><button>Run Validation</button></div>
  </form>
</section>
{{ jobs_html|safe }}
{% endblock %}
""",
    "run.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Run</h1>
<section class="panel">
  <form method="post" action="/jobs">
    <label>Operation</label>
    <select name="kind">
      <option value="import-mail">Import mail</option>
      <option value="import-calendar">Import calendar</option>
      <option value="import-contacts">Import contacts</option>
      <option value="import-all">Import all</option>
    </select>
    {% include "mapping_select.html" %}
    <label><input type="checkbox" name="import_skipped_duplicates" value="yes" style="width:auto"> Import skipped duplicate mail copies (remediation re-run).</label>
    <p class="muted">Applies to mail imports only. Exact source rows already marked done are still skipped.</p>
    <div class="actions"><button>Start Import</button></div>
  </form>
</section>
{{ jobs_html|safe }}
{% endblock %}
""",
    "tools.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Tools</h1>
<section class="panel">
  <form method="post" action="/jobs">
    <label>Destructive Operation</label>
    <select name="kind">
      <option value="purge-calendar">Purge imported calendar items</option>
      <option value="purge-contacts">Purge imported contacts</option>
      <option value="purge-mail">Purge mailbox mail</option>
      <option value="reset-state">Reset local state</option>
    </select>
    {% include "mapping_select.html" %}
    <label><input type="checkbox" name="confirmed" value="yes" style="width:auto"> I understand this operation is destructive.</label>
    <div class="actions"><button>Run Tool</button></div>
  </form>
</section>
{{ jobs_html|safe }}
{% endblock %}
""",
    "status.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Status</h1>
<table>
  <thead><tr><th>Mailbox</th><th>PST</th><th>Status</th><th>Total</th><th>Done</th><th>Failed</th><th>Skipped</th><th>Last Error</th></tr></thead>
  <tbody>
  {% for row in rows %}
    <tr><td>{{ row.mailbox }}</td><td>{{ row.pst }}</td><td>{{ row.status }}</td><td>{{ row.total }}</td><td>{{ row.done }}</td><td>{{ row.failed }}</td><td>{{ row.skipped }}</td><td>{{ row.last_error }}</td></tr>
  {% else %}
    <tr><td colspan="8" class="muted">No state recorded.</td></tr>
  {% endfor %}
  </tbody>
</table>
{% endblock %}
""",
    "logs.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Logs</h1>
<table>
  <thead><tr><th>Name</th><th>Size</th><th>Modified</th></tr></thead>
  <tbody>
  {% for log in logs %}
    <tr><td><a href="/logs/{{ log.name }}">{{ log.name }}</a></td><td>{{ log.size }}</td><td>{{ log.mtime }}</td></tr>
  {% else %}
    <tr><td colspan="3" class="muted">No logs found.</td></tr>
  {% endfor %}
  </tbody>
</table>
{% endblock %}
""",
    "log_detail.html": """
{% extends "layout.html" %}
{% block body %}
<h1>{{ name }}</h1>
{% if error %}<p class="bad">{{ error }}</p>{% endif %}
<section class="grid">
  <div class="card"><div class="muted">Lines</div><div class="metric">{{ summary.total }}</div></div>
  <div class="card"><div class="muted">Errors</div><div class="metric bad">{{ summary.errors }}</div></div>
  <div class="card"><div class="muted">Warnings</div><div class="metric warn">{{ summary.warnings }}</div></div>
  <div class="card"><div class="muted">Info</div><div class="metric ok">{{ summary.info }}</div></div>
</section>
<section class="panel" style="margin-top:16px">
  <table>
    <thead><tr><th>Time</th><th>Level</th><th>Context</th><th>Message</th><th>Details</th></tr></thead>
    <tbody>
    {% for row in rows %}
      <tr>
        <td>{{ row.time }}</td>
        <td class="{{ row.level_class }}">{{ row.level }}</td>
        <td>{{ row.context }}</td>
        <td>{{ row.message }}</td>
        <td>{% if row.details %}<details><summary>view</summary><pre>{{ row.details }}</pre></details>{% endif %}</td>
      </tr>
    {% else %}
      <tr><td colspan="5" class="muted">No log lines found.</td></tr>
    {% endfor %}
    </tbody>
  </table>
</section>
{% if raw_text %}
<section class="panel">
  <h2>Raw Tail</h2>
  <pre>{{ raw_text }}</pre>
</section>
{% endif %}
{% endblock %}
""",
    "job_detail.html": """
{% extends "layout.html" %}
{% block body %}
<h1>Job {{ job.job_id }}</h1>
<section class="panel" hx-get="/jobs/{{ job.job_id }}/panel" hx-trigger="every 2s">
  {% include "job_panel.html" %}
</section>
{% endblock %}
""",
    "job_panel.html": """
<p><strong>Kind:</strong> {{ job.kind }} | <strong>Status:</strong> {{ job.status }} | <strong>Phase:</strong> {{ job.phase }} | <strong>Rows:</strong> {{ job.selected_rows }}</p>
<p><strong>Current activity:</strong> {{ job.activity }}</p>
{% if job.progress_total %}
<div class="progress-wrap" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="{{ job.progress_percent }}">
  <div class="progress-bar"><div class="progress-fill" style="width: {{ job.progress_percent }}%"></div></div>
  <p class="muted">{{ job.progress_current }}/{{ job.progress_total }} items processed ({{ job.progress_percent }}%) — {{ job.progress_uploaded }} uploaded, {{ job.progress_skipped }} skipped, {{ job.progress_failed }} failed, {{ job.progress_cancelled }} cancelled</p>
</div>
{% endif %}
{% if job.status in ["queued", "running"] %}
<form method="post" action="/jobs/{{ job.job_id }}/cancel" onsubmit="return confirm('Stop this job after in-flight work finishes?');">
  <button class="secondary" type="submit">Stop job</button>
</form>
{% endif %}
{% if job.last_error %}<p class="bad">{{ job.last_error }}</p>{% endif %}
{% if job.events %}
<h2>Live activity</h2>
<ul class="event-list">
{% for event in job.events|reverse %}
  <li><span class="muted">{{ event.elapsed }}</span> {{ event.message }}</li>
{% endfor %}
</ul>
{% endif %}
{% if job.validation %}
<table><thead><tr><th>Check</th><th>Result</th><th>Message</th></tr></thead><tbody>
{% for check in job.validation.checks %}
<tr><td>{{ check.name }}</td><td class="{{ 'ok' if check.ok else 'bad' }}">{{ "OK" if check.ok else "FAIL" }}</td><td>{{ check.message }}</td></tr>
{% endfor %}
</tbody></table>
{% endif %}
{% if job.reports %}
<table><thead><tr><th>Mailbox</th><th>PST</th><th>Status</th><th>Total</th><th>Uploaded</th><th>Skipped</th><th>Failed</th><th>Error</th></tr></thead><tbody>
{% for report in job.reports %}
<tr><td>{{ report.mailbox }}</td><td>{{ report.pst_path.name }}</td><td>{{ report.status }}</td><td>{{ report.items_total }}</td><td>{{ report.items_uploaded }}</td><td>{{ report.items_skipped }}</td><td>{{ report.items_failed }}</td><td>{{ report.last_error or "" }}</td></tr>
{% endfor %}
</tbody></table>
{% endif %}
{% if job.result %}<pre>{{ job.result }}</pre>{% endif %}
""",
    "mapping_select.html": """
<div class="row">
  <div>
    <label>Mailbox</label>
    <select name="mailbox">
      <option value="">All mailboxes ({{ mapping_options.mailboxes|length }})</option>
      {% for mailbox in mapping_options.mailboxes %}
        <option value="{{ mailbox }}">{{ mailbox }}</option>
      {% endfor %}
    </select>
  </div>
  <div>
    <label>PST</label>
    <select name="pst">
      <option value="">All PSTs ({{ mapping_options.psts|length }})</option>
      {% for pst in mapping_options.psts %}
        <option value="{{ pst }}">{{ pst }}</option>
      {% endfor %}
    </select>
  </div>
</div>
<p class="muted">Selections come from {{ mapping_path or "mapping.csv" }}. Choose both a mailbox and PST only when you need to target one mapping row.</p>
{% if mapping_options.rows %}
<table style="margin-top:10px">
  <thead><tr><th>Mailbox</th><th>PST</th><th>Exists</th></tr></thead>
  <tbody>
  {% for row in mapping_options.rows[:8] %}
    <tr><td>{{ row.target_mailbox }}</td><td>{{ row.pst_path.name }}</td><td>{{ "yes" if row.pst_path.exists() else "no" }}</td></tr>
  {% endfor %}
  </tbody>
</table>
{% else %}
<p class="warn">No mapping rows loaded yet.</p>
{% endif %}
""",
}


def create_app(config_path: Path | None = None, mapping_path: Path | None = None) -> FastAPI:
    config_path = config_path or Path("config.toml")
    app = FastAPI(title="pstmigrate")
    app.state.config_path = config_path
    app.state.mapping_path = mapping_path
    app.state.jobs = JobManager()
    static_dir = Path(__file__).with_name("web_static")
    app.mount("/static", StaticFiles(directory=static_dir), name="static")

    env = Environment(
        loader=DictLoader(TEMPLATES),
        autoescape=select_autoescape(["html"]),
    )

    def render(template_name: str, **context: Any) -> HTMLResponse:
        context.setdefault("config_path", config_path)
        context.setdefault("mapping_path", mapping_path)
        return HTMLResponse(env.get_template(template_name).render(**context))

    def state_queries() -> StateQueries | None:
        try:
            cfg = load_config(config_path)
        except Exception:
            return None
        db = cfg.paths.state_dir / "state.sqlite"
        if not db.exists():
            return None
        return StateQueries(StateStore(db))

    def jobs_html() -> str:
        return env.get_template("jobs.html").render(jobs=app.state.jobs.latest())

    def mapping_options() -> dict[str, Any]:
        rows = _load_mapping_for_ui(mapping_path)
        return {
            "rows": rows,
            "mailboxes": sorted({row.target_mailbox for row in rows}),
            "psts": sorted({row.pst_path.name for row in rows}),
        }

    @app.get("/", response_class=HTMLResponse)
    async def dashboard() -> HTMLResponse:
        queries = state_queries()
        totals = queries.dashboard_totals() if queries else {"runs": 0, "done": 0, "failed": 0, "skipped": 0}
        return render("dashboard.html", totals=totals, jobs_html=jobs_html())

    @app.get("/partials/jobs", response_class=HTMLResponse)
    async def partial_jobs() -> HTMLResponse:
        return HTMLResponse(jobs_html())

    @app.get("/config", response_class=HTMLResponse)
    async def config_page() -> HTMLResponse:
        form = _config_form_defaults(config_path)
        try:
            cfg = load_config(config_path)
        except Exception as e:
            return render(
                "config.html",
                config_text=_mask_config(config_path.read_text(encoding="utf-8")) if config_path.exists() else "",
                config_error=f"Config is not ready: {e}",
                app_count=0,
                workers="?",
                parallel="?",
                form=form,
            )
        text = ""
        if config_path and config_path.exists():
            text = _mask_config(config_path.read_text(encoding="utf-8"))
        return render(
            "config.html",
            config_text=text or "Loaded from defaults/environment.",
            config_error="",
            app_count=len(cfg.apps),
            workers=cfg.migration.workers_per_mailbox,
            parallel=cfg.migration.max_parallel_mailboxes,
            form=form,
        )

    @app.post("/config")
    async def save_config(
        app_name: Annotated[str, Form()] = "primary",
        tenant_id: Annotated[str, Form()] = "",
        client_id: Annotated[str, Form()] = "",
        client_secret: Annotated[str, Form()] = "",
        client_certificate_path: Annotated[str, Form()] = "",
        workers_per_mailbox: Annotated[int, Form()] = 2,
        max_parallel_mailboxes: Annotated[int, Form()] = 4,
        state_dir: Annotated[str, Form()] = ".pstmigrate-state",
        work_dir: Annotated[str, Form()] = ".pstmigrate-work",
        log_dir: Annotated[str, Form()] = "logs",
        readpst_binary: Annotated[str, Form()] = "readpst",
    ) -> RedirectResponse:
        existing = _config_form_defaults(config_path)
        secret = client_secret.strip()
        if not secret and existing["has_secret"]:
            secret = existing["client_secret"]
        _write_config(
            config_path=config_path,
            app_name=app_name,
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=secret,
            client_certificate_path=client_certificate_path,
            workers_per_mailbox=workers_per_mailbox,
            max_parallel_mailboxes=max_parallel_mailboxes,
            state_dir=state_dir,
            work_dir=work_dir,
            log_dir=log_dir,
            readpst_binary=readpst_binary,
        )
        return RedirectResponse("/config", status_code=303)

    @app.get("/mapping", response_class=HTMLResponse)
    async def mapping_page(mailbox: str = "", pst: str = "") -> HTMLResponse:
        rows = load_mapping(mapping_path) if mapping_path else []
        filters = SelectionFilters(
            mailboxes=[mailbox] if mailbox else None,
            pst_names=[pst] if pst else None,
        )
        selected = select_mapping(rows, filters).rows
        return render("mapping.html", rows=selected, mailbox=mailbox, pst=pst)

    @app.get("/validate", response_class=HTMLResponse)
    async def validate_page() -> HTMLResponse:
        return render("validate.html", jobs_html=jobs_html(), mapping_options=mapping_options())

    @app.get("/run", response_class=HTMLResponse)
    async def run_page() -> HTMLResponse:
        return render("run.html", jobs_html=jobs_html(), mapping_options=mapping_options())

    @app.get("/tools", response_class=HTMLResponse)
    async def tools_page() -> HTMLResponse:
        return render("tools.html", jobs_html=jobs_html(), mapping_options=mapping_options())

    @app.get("/status", response_class=HTMLResponse)
    async def status_page() -> HTMLResponse:
        queries = state_queries()
        rows = queries.run_rows() if queries else []
        return render("status.html", rows=rows)

    @app.get("/logs", response_class=HTMLResponse)
    async def logs_page() -> HTMLResponse:
        try:
            cfg = load_config(config_path)
        except Exception as e:
            return render("log_detail.html", name="Logs", error=f"Config is not ready: {e}", **_empty_log_view())
        logs = []
        if cfg.paths.log_dir.exists():
            for path in sorted(cfg.paths.log_dir.glob("*.jsonl"), key=lambda p: p.stat().st_mtime, reverse=True):
                stat = path.stat()
                logs.append({"name": path.name, "size": stat.st_size, "mtime": int(stat.st_mtime)})
        return render("logs.html", logs=logs)

    @app.get("/logs/{name}", response_class=HTMLResponse)
    async def log_detail(name: str) -> HTMLResponse:
        try:
            cfg = load_config(config_path)
        except Exception as e:
            return render("log_detail.html", name=name, error=f"Config is not ready: {e}", **_empty_log_view())
        path = (cfg.paths.log_dir / name).resolve()
        if path.parent != cfg.paths.log_dir.resolve() or not path.exists():
            return HTMLResponse("Log not found", status_code=404)
        return render("log_detail.html", name=name, error="", **_read_log_view(path))

    @app.post("/jobs")
    async def submit_job(
        kind: Annotated[JobKind, Form()],
        mailbox: Annotated[str, Form()] = "",
        pst: Annotated[str, Form()] = "",
        confirmed: Annotated[str | None, Form()] = None,
        import_skipped_duplicates: Annotated[str | None, Form()] = None,
    ) -> RedirectResponse:
        spec = JobSpec(
            kind=kind,
            config_path=config_path,
            mapping_path=mapping_path,
            filters=SelectionFilters(
                mailboxes=[mailbox] if mailbox else None,
                pst_names=[pst] if pst else None,
            ),
            confirmed=confirmed == "yes",
            import_skipped_duplicates=import_skipped_duplicates == "yes",
        )
        record = app.state.jobs.submit(spec)
        return RedirectResponse(f"/jobs/{record.job_id}", status_code=303)

    @app.post("/jobs/{job_id}/cancel")
    async def cancel_job(job_id: str) -> RedirectResponse:
        app.state.jobs.cancel(job_id)
        return RedirectResponse(f"/jobs/{job_id}", status_code=303)

    @app.get("/jobs/{job_id}", response_class=HTMLResponse)
    async def job_detail(job_id: str) -> HTMLResponse:
        job = app.state.jobs.get(job_id)
        if job is None:
            return HTMLResponse("Job not found", status_code=404)
        return render("job_detail.html", job=job)

    @app.get("/jobs/{job_id}/panel", response_class=HTMLResponse)
    async def job_panel(job_id: str) -> HTMLResponse:
        job = app.state.jobs.get(job_id)
        if job is None:
            # HTMX keeps polling job panels while a tab is open. After a server
            # restart the in-memory job list is gone, so old tabs otherwise spam
            # 404s forever. Status 286 is HTMX's documented "stop polling" code.
            return HTMLResponse(
                '<p class="muted">Job is no longer available in this server session. '
                'Refresh the Jobs list to view current jobs.</p>',
                status_code=286,
            )
        status_code = 286 if job.status in {"done", "failed", "blocked", "cancelled"} else 200
        return HTMLResponse(env.get_template("job_panel.html").render(job=job), status_code=status_code)

    return app


def _mask_config(text: str) -> str:
    masked: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith(("client_secret", "client_certificate_password")):
            key = line.split("=", 1)[0]
            masked.append(f'{key}= "********"')
        else:
            masked.append(line)
    return "\n".join(masked)


def _load_mapping_for_ui(mapping_path: Path | None) -> list:
    if mapping_path is None or not mapping_path.exists():
        return []
    try:
        return load_mapping(mapping_path)
    except Exception:
        return []


def _empty_log_view() -> dict[str, Any]:
    return {
        "summary": {"total": 0, "errors": 0, "warnings": 0, "info": 0},
        "rows": [],
        "raw_text": "",
    }


def _read_log_view(path: Path, *, max_lines: int = 500) -> dict[str, Any]:
    lines = path.read_text(errors="replace").splitlines()
    tail = lines[-max_lines:]
    rows = [_parse_log_line(line) for line in tail]
    summary = {
        "total": len(lines),
        "errors": sum(1 for row in rows if row["level"] in {"ERROR", "CRITICAL"}),
        "warnings": sum(1 for row in rows if row["level"] == "WARNING"),
        "info": sum(1 for row in rows if row["level"] == "INFO"),
    }
    return {"summary": summary, "rows": rows, "raw_text": ""}


def _parse_log_line(line: str) -> dict[str, str]:
    try:
        payload = json.loads(line)
    except json.JSONDecodeError:
        return {
            "time": "",
            "level": "TEXT",
            "level_class": "",
            "context": "",
            "message": line,
            "details": "",
        }

    record = payload.get("record") if isinstance(payload, dict) else None
    if not isinstance(record, dict):
        return {
            "time": "",
            "level": "JSON",
            "level_class": "",
            "context": "",
            "message": str(payload),
            "details": "",
        }

    level = _log_level_name(record.get("level"))
    extra = record.get("extra") if isinstance(record.get("extra"), dict) else {}
    exception = record.get("exception")
    details: dict[str, Any] = {}
    if exception:
        details["exception"] = exception
    if extra:
        details["extra"] = extra

    return {
        "time": _log_time(record.get("time")),
        "level": level,
        "level_class": _log_level_class(level),
        "context": str(extra.get("ctx") or ""),
        "message": str(record.get("message") or payload.get("text") or ""),
        "details": json.dumps(details, indent=2, default=str) if details else "",
    }


def _log_level_name(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("name") or "").upper()
    return str(value or "").upper()


def _log_level_class(level: str) -> str:
    if level in {"ERROR", "CRITICAL"}:
        return "bad"
    if level == "WARNING":
        return "warn"
    if level == "INFO":
        return "ok"
    return ""


def _log_time(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("repr") or value.get("timestamp")
    text = str(value or "")
    return text[:19].replace("T", " ")


def _config_form_defaults(config_path: Path) -> dict[str, Any]:
    defaults: dict[str, Any] = {
        "app_name": "primary",
        "tenant_id": "",
        "client_id": "",
        "client_secret": "",
        "has_secret": False,
        "client_certificate_path": "",
        "workers_per_mailbox": 2,
        "max_parallel_mailboxes": 4,
        "state_dir": ".pstmigrate-state",
        "work_dir": ".pstmigrate-work",
        "log_dir": "logs",
        "readpst_binary": "readpst",
    }
    if not config_path.exists():
        return defaults
    try:
        data = tomllib.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return defaults
    apps = data.get("apps") or []
    if apps:
        app = apps[0]
        defaults["app_name"] = app.get("name") or "primary"
        defaults["tenant_id"] = app.get("tenant_id") or ""
        defaults["client_id"] = app.get("client_id") or ""
        defaults["client_secret"] = app.get("client_secret") or ""
        defaults["has_secret"] = bool(defaults["client_secret"])
        defaults["client_certificate_path"] = str(app.get("client_certificate_path") or "")
    migration = data.get("migration") or {}
    defaults["workers_per_mailbox"] = migration.get("workers_per_mailbox", defaults["workers_per_mailbox"])
    defaults["max_parallel_mailboxes"] = migration.get(
        "max_parallel_mailboxes", defaults["max_parallel_mailboxes"]
    )
    paths = data.get("paths") or {}
    defaults["state_dir"] = str(paths.get("state_dir", defaults["state_dir"]))
    defaults["work_dir"] = str(paths.get("work_dir", defaults["work_dir"]))
    defaults["log_dir"] = str(paths.get("log_dir", defaults["log_dir"]))
    defaults["readpst_binary"] = str(paths.get("readpst_binary", defaults["readpst_binary"]))
    return defaults


def _write_config(
    *,
    config_path: Path,
    app_name: str,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    client_certificate_path: str,
    workers_per_mailbox: int,
    max_parallel_mailboxes: int,
    state_dir: str,
    work_dir: str,
    log_dir: str,
    readpst_binary: str,
) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    auth_lines = [
        "[[apps]]",
        f'name = "{_toml_escape(app_name.strip() or "primary")}"',
        f'tenant_id = "{_toml_escape(tenant_id.strip())}"',
        f'client_id = "{_toml_escape(client_id.strip())}"',
    ]
    if client_certificate_path.strip():
        auth_lines.append(
            f'client_certificate_path = "{_toml_escape(client_certificate_path.strip())}"'
        )
    else:
        auth_lines.append(f'client_secret = "{_toml_escape(client_secret.strip())}"')
    body = f"""log_level = "INFO"

{chr(10).join(auth_lines)}

[throttle]
max_retries = 8
initial_backoff_seconds = 1.0
max_backoff_seconds = 120.0

[migration]
workers_per_mailbox = {workers_per_mailbox}
max_parallel_mailboxes = {max_parallel_mailboxes}
target_root_folder = "Imported PST"
fail_fast = false

[paths]
state_dir = "{_toml_escape(state_dir.strip() or ".pstmigrate-state")}"
work_dir = "{_toml_escape(work_dir.strip() or ".pstmigrate-work")}"
log_dir = "{_toml_escape(log_dir.strip() or "logs")}"
readpst_binary = "{_toml_escape(readpst_binary.strip() or "readpst")}"
"""
    config_path.write_text(body, encoding="utf-8")


def _toml_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')
