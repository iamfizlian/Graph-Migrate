# pstmigrate — PST → Microsoft 365 via Graph

Imports `.pst` files into Exchange Online mailboxes through the Microsoft Graph
API. Replaces the Outlook/MAPI/COM approach (`PS-Migrate/pst-import.ps1`) with
something that runs headless on Linux, uses app-only auth, and parallelizes
across mailboxes without an Outlook process.

## Why this instead of the PowerShell + Outlook script?

The columns below compare the **old** `PS-Migrate/pst-import.ps1` (left)
against **this tool, `pstmigrate`** (right). The install instructions in the
next section are for `pstmigrate` — that's what runs on Linux.

| Aspect | Old: `pst-import.ps1` (PowerShell + Outlook COM) | New: `pstmigrate` (this) |
|---|---|---|
| Host OS | Windows only, Outlook installed, signed-in admin profile | Linux, macOS, or Windows; no Outlook needed |
| Auth | Interactive admin sign-in, Full Access on each mailbox | App-only cert/secret, `Mail.ReadWrite` (application) |
| Parallelism | Single-threaded, one PST at a time | N mailboxes × M workers, multi-app pool |
| Throttling | EWS/MAPI per-tenant, opaque | Graph published limits, honours `Retry-After` |
| Resume | Per-folder + per-item file logs | SQLite, atomic, queryable |
| What's lost | Nothing (full MAPI fidelity) | Read state, importance, categories, flag status |

The fidelity loss is real but small for archive PSTs that are predominantly
mail. If you need calendar/contacts migrated too, that's out of scope here.

## Install

`pstmigrate` needs two things on the host:

1. The `readpst` binary from the [`libpst`](https://www.five-ten-sg.com/libpst/)
   project, on `PATH` (used to extract `.eml` files from each PST).
2. Python 3.10+ with this package installed.

Graph itself is a REST API and is fully cross-platform — the only OS-sensitive
piece is `readpst`. The PSTs don't have to be on the same machine as
`pstmigrate`; they just need to be reachable as regular file paths.

Pick a platform:

- **[Windows — native](#windows--native-everything-runs-in-powershell)** —
  everything runs in PowerShell, no Linux. Best if your PSTs already live on
  this Windows server (your `D:\PST\…` case).
- **[Windows — WSL2 + Ubuntu](#windows--wsl2--ubuntu-alternative)** —
  alternative if you don't want MSYS2 on the box.
- **[Linux (Fedora / RHEL / Debian / Ubuntu)](#linux)**
- **[macOS](#macos)**

---

### Windows — native (everything runs in PowerShell)

Tested target: Windows 10 21H2 / Windows 11 / Windows Server 2019 & 2022,
PowerShell 5.1 or 7.x. You'll need local admin to install Python and MSYS2.

The end state is: a Python virtual env in
`C:\Tools\Graph-Migrate\.venv`, with `readpst.exe` on PATH coming from MSYS2
(`C:\msys64\ucrt64\bin\readpst.exe`).

#### 1. Install Python 3.12 (5 min)

Download the **Windows installer (64-bit)** from
<https://www.python.org/downloads/windows/>.

In the installer:

- ✅ tick **"Add python.exe to PATH"** (this is the most common skipped step).
- Click **Install Now**.

Open a **new** PowerShell window (the PATH change won't be visible in shells
opened before the install) and verify:

```powershell
python --version
# expected: Python 3.12.x

python -m pip --version
# expected: pip 24.x ... (python 3.12)
```

If `python` is not recognised, re-run the installer and choose **Modify →
Next → tick "Add Python to environment variables"**.

#### 2. Install Git for Windows (2 min)

Download from <https://git-scm.com/download/win> and accept the defaults.
Verify in a new PowerShell:

```powershell
git --version
# expected: git version 2.x.x.windows.x
```

(If you already have the project copied as a folder, you can skip Git.)

#### 3. Install MSYS2 and the `libpst` package — this is what gives you `readpst.exe` (10 min)

`libpst` has no official Windows installer; the most reliable way to get a
maintained Windows build is via MSYS2's package manager.

1. Download the installer from <https://www.msys2.org/> ("MSYS2 Installer"
   button at the top — file name is `msys2-x86_64-YYYYMMDD.exe`). Run it and
   accept the defaults; it installs to `C:\msys64`.
2. From the Start menu, launch **"MSYS2 UCRT64"** (not MSYS, not MINGW64 —
   specifically **UCRT64**). A black terminal window opens.
3. Update the package database and core packages:
   ```bash
   pacman -Syu
   ```
   When it asks to close the terminal, type `Y`, close the window, and
   re-open **MSYS2 UCRT64** from the Start menu.
4. Run the update again to finish, then install `libpst`:
   ```bash
   pacman -Syu
   pacman -S --noconfirm mingw-w64-ucrt-x86_64-libpst
   ```
5. Verify inside the MSYS2 window:
   ```bash
   readpst -V
   # expected: ReadPST / LibPST v0.6.x
   ```
   You can close the MSYS2 window now — we won't need it again. The binary
   you just installed lives at `C:\msys64\ucrt64\bin\readpst.exe`.

#### 4. Add `readpst.exe` to your Windows PATH (2 min)

So PowerShell can find it without going back into MSYS2.

GUI path (works on every Windows version):

1. Press <kbd>Win</kbd> and type **"Edit the system environment variables"** → Enter.
2. Click **Environment Variables…**
3. Under **System variables** (lower box), select **Path** → **Edit…**
4. Click **New**, paste:
   ```
   C:\msys64\ucrt64\bin
   ```
5. **OK** → **OK** → **OK**.

Or, equivalently, from an **elevated** PowerShell (one-liner):

```powershell
[Environment]::SetEnvironmentVariable(
  'Path',
  [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';C:\msys64\ucrt64\bin',
  'Machine'
)
```

Either way, **close every PowerShell window and open a fresh one**, then
verify:

```powershell
where.exe readpst
# expected: C:\msys64\ucrt64\bin\readpst.exe

readpst -V
# expected: ReadPST / LibPST v0.6.x
```

If `where.exe readpst` returns nothing, the new PowerShell window didn't
inherit the PATH change — log out and back in (or reboot) and try again.

#### 5. Allow PowerShell to run venv activation scripts (30 sec, one-time)

By default Windows blocks the script that activates a Python venv. Loosen it
for the current user — this is the standard developer setting:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
# answer Y when prompted
```

#### 6. Get the project onto the box (1 min)

Pick a folder you have write access to. `C:\Tools` is a good default:

```powershell
mkdir C:\Tools -Force | Out-Null
cd C:\Tools

# Either clone with git:
git clone <your repo url> JTET-Email-Migration
cd JTET-Email-Migration\Graph-Migrate

# Or, if you copied the folder over, just cd into it:
# cd C:\Tools\JTET-Email-Migration\Graph-Migrate
```

#### 7. Create the Python virtual environment and install `pstmigrate` (3 min)

```powershell
cd C:\Tools\JTET-Email-Migration\Graph-Migrate

python -m venv .venv
.\.venv\Scripts\Activate.ps1
# your prompt should now be prefixed with (.venv)

python -m pip install --upgrade pip
pip install -e .

pstmigrate --help
# expected: a usage table listing the import / validate / status / template / init-config commands
```

If `pstmigrate` is not recognised after `pip install -e .`, run
`python -m jtet_pstmigrate --help` instead — same tool, fully-qualified
module path.

#### 8. Configure and test against your tenant (5 min)

> **Before this step**, do the
> **[One-time Microsoft 365 setup](#one-time-microsoft-365-setup--register-the-entra-app)**
> below to register an Entra app and get the three values you need:
> `tenant_id`, `client_id`, and `client_secret`. Without them
> `config.toml` is just a template and `validate` will fail with a Graph
> token error.

```powershell
# This repo already ships with a pre-built mapping.csv (your 13 PSTs) and a
# starter config.toml. Use those — DO NOT run `pstmigrate template` or
# `pstmigrate init-config` against the same paths or you'll overwrite them.
# (Both commands now refuse to overwrite by default.)

# Open config.toml in Notepad and paste the three values from the M365 setup
# section into the [[apps]] block:
#   tenant_id     = "<Directory (tenant) ID>"
#   client_id     = "<Application (client) ID>"
#   client_secret = "<the Value column from Certificates & secrets>"
notepad config.toml

# Pre-flight a single mailbox first
pstmigrate validate -c config.toml -m mapping.csv -M emmak@jteatono365.onmicrosoft.com
```

> First-time scratch setup (only if you don't have config.toml or mapping.csv):
> `pstmigrate init-config -o config.toml` writes a starter TOML;
> `pstmigrate template -o mapping.example.csv` writes a 2-row sample CSV
> (rename to `mapping.csv` after editing). Both refuse to overwrite an
> existing file unless you pass `--force`.

Expected output: a "Pre-flight Checks" table with green `OK` rows for
`readpst available`, `mapping CSV parses`, `PST files exist`, `Graph token
(per app)`, and `Mailboxes resolvable`.

#### 9. Day-of-migration run

```powershell
# Activate the venv (every new PowerShell window needs this once)
cd C:\Tools\JTET-Email-Migration\Graph-Migrate
.\.venv\Scripts\Activate.ps1

# Canary one mailbox end-to-end first
pstmigrate import -c config.toml -m mapping.csv -M emmak@jteatono365.onmicrosoft.com

# Then run the whole mapping
pstmigrate import -c config.toml -m mapping.csv

# Watch progress at any point from a second PowerShell window
pstmigrate status -c config.toml
```

State, work files and logs are written under the project folder by default:

```
C:\Tools\JTET-Email-Migration\Graph-Migrate\
  .pstmigrate-state\state.sqlite        ← dedupe + resume DB; do NOT delete
  .pstmigrate-work\<pst-stem>\…         ← extracted .eml tree (deletable after success)
  logs\import_YYYYMMDD_HHMMSS.jsonl     ← structured log per run
```

To run multi-day jobs without a babysitting a PowerShell window, either:

- Run inside `tmux` from MSYS2 / Git Bash, **or**
- Wrap the command in a scheduled task / a `Start-Process -WindowStyle Hidden`,
  **or**
- Just leave the window open — the migration is fully resumable, so even if
  the box reboots you can re-run the same `pstmigrate import` command and it
  will pick up exactly where it stopped.

#### Native-Windows troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `pstmigrate: command not found` after `pip install -e .` | venv not activated, or pip installed against system Python | Re-activate: `.\.venv\Scripts\Activate.ps1`, then `pip install -e .` again |
| `where.exe readpst` returns nothing | New PowerShell window didn't inherit PATH | Open a brand-new PowerShell window; if still nothing, log out + back in |
| Activation fails: `running scripts is disabled on this system` | PS execution policy too strict | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Defender / SmartScreen blocks `readpst.exe` | Antivirus heuristic on an unsigned binary | Add `C:\msys64\ucrt64\bin\readpst.exe` to AV exclusions |
| `pstmigrate validate` says Graph token FAIL | Wrong tenant/client id, secret expired, or admin consent missing | Re-check `config.toml`, then in Entra → API permissions, click **Grant admin consent** |
| Hangs on a huge PST during extraction | `readpst` is single-threaded per file; expected on 20+ GB PSTs | Wait — extraction runs once, then upload begins; subsequent runs reuse the extracted tree |

---

### Windows — WSL2 + Ubuntu (alternative)

Use this if MSYS2 is blocked by your security policy, or you prefer a Linux
toolchain. Your PSTs stay on `D:\PST\…` and WSL exposes them as `/mnt/d/PST/…`
with no copy.

```powershell
# From an elevated PowerShell, install WSL2 + Ubuntu (one-time). Reboot when prompted.
wsl --install -d Ubuntu
```

Launch **"Ubuntu"** from the Start menu, finish first-time user setup, then
inside the Ubuntu shell:

```bash
sudo apt update
sudo apt install -y pst-utils python3-venv python3-pip git
readpst -V        # confirm libpst is on PATH

# Get the project (clone, or cp -r from /mnt/c/...)
cd ~/JTET-Email-Migration/Graph-Migrate
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
pstmigrate --help

# Use the WSL-style mapping that points at /mnt/d/PST/...
pstmigrate validate -c config.toml -m mapping.wsl.csv
pstmigrate import   -c config.toml -m mapping.wsl.csv -M emmak@jteatono365.onmicrosoft.com
```

WSL path notes:

- `mapping.csv` (Windows paths like `D:\PST\…`) is for the **native** Windows
  path above. `mapping.wsl.csv` (paths like `/mnt/d/PST/…`) is the one to use
  under WSL.
- Reading PSTs from `/mnt/d/…` is fine but the NTFS bridge is slower than
  ext4. If a particular PST is huge (> ~20 GB) and you have spare Linux-side
  disk, copy it into WSL first (`cp /mnt/d/PST/big.pst ~/pst/`) and update
  the mapping row.
- Run long jobs inside `tmux` or `screen` so closing the terminal doesn't
  kill the import.

---

### Linux

Fedora / RHEL:

```bash
sudo dnf install -y libpst python3-pip
readpst -V
```

Debian / Ubuntu:

```bash
sudo apt install -y pst-utils python3-venv python3-pip
readpst -V
```

Then the common Python install:

```bash
cd Graph-Migrate
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
pstmigrate --help
```

### macOS

```bash
brew install libpst python
readpst -V

cd Graph-Migrate
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
pstmigrate --help
```

## One-time Microsoft 365 setup — register the Entra app

This is where the `tenant_id`, `client_id`, and `client_secret` values you
paste into `config.toml` come from. You need to do this **once per tenant**,
not once per mailbox. Requires Global Admin (or Application Administrator +
Privileged Role Administrator) in the **destination** tenant.

### 1. Register the app

1. Sign in to <https://entra.microsoft.com> as a Global Admin.
2. Left nav → **Applications** → **App registrations** → **+ New registration**.
3. Fill in:
   - **Name**: `pstmigrate` (anything you like — it's just a label)
   - **Supported account types**: **Accounts in this organizational directory only (Single tenant)**
   - **Redirect URI**: leave blank (we use app-only auth, no redirect)
4. Click **Register**.

You'll land on the app's **Overview** page. **Copy these two values into a
notepad — you'll paste them into `config.toml` shortly:**

| Field on Overview page | Goes into `config.toml` as |
|---|---|
| **Application (client) ID** | `client_id` |
| **Directory (tenant) ID** | `tenant_id` |

### 2. Grant the Mail.ReadWrite Graph permission

1. On the app's left nav → **API permissions** → **+ Add a permission**.
2. Choose **Microsoft Graph** → **Application permissions** (NOT Delegated).
3. In the search box type `Mail.ReadWrite`, tick the box, click **Add permissions**.
4. Back on the API permissions page, click **✓ Grant admin consent for \<tenant\>** at the top. The Status column should change to a green **"Granted for \<tenant\>"**. *This step is mandatory — without it, every Graph call returns 403.*

> If you also want to read calendar/contacts later, repeat with
> `Calendars.ReadWrite` and `Contacts.ReadWrite`. For mail-only migration,
> just `Mail.ReadWrite`.

### 3. Create a client secret

1. App's left nav → **Certificates & secrets** → **Client secrets** tab → **+ New client secret**.
2. Description: `pstmigrate`. Expires: pick the shortest window that covers your migration window (6 months is fine; 24 months is the max).
3. Click **Add**.
4. **IMMEDIATELY copy the *Value* column** (not the Secret ID — the Value, the longer string). It will only be shown this one time. This is your `client_secret` for `config.toml`.

If you'd rather use a certificate (cleaner for production, no expiring secrets), upload a public-key cert under the **Certificates** tab on the same page and reference its private key path via `client_certificate_path` in `config.toml` instead of `client_secret`.

### 4. (Recommended) Scope the app to only the mailboxes you migrate into

By default the `Mail.ReadWrite` application permission grants the app
read/write access to **every mailbox in the tenant**. You almost certainly
don't want that. Apply an
[Application Access Policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access)
to scope the app to only the migration target mailboxes:

```powershell
# In an Exchange Online PowerShell session (Connect-ExchangeOnline)

# 1. Create a mail-enabled security group containing all destination UPNs
New-DistributionGroup -Name "PSTMigrate-Targets" -Type Security -Members @(
  "allysonp@jteatono365.onmicrosoft.com",
  "debbiep@jteatono365.onmicrosoft.com",
  "emmak@jteatono365.onmicrosoft.com"
  # ... etc, all 13 destination mailboxes
)

# 2. Restrict the app to only those mailboxes
New-ApplicationAccessPolicy `
  -AppId <Application (client) ID from step 1> `
  -PolicyScopeGroupId PSTMigrate-Targets@jteatono365.onmicrosoft.com `
  -AccessRight RestrictAccess `
  -Description "PST migration scope"

# 3. Verify — the first should return Granted, the second Denied
Test-ApplicationAccessPolicy -Identity emmak@jteatono365.onmicrosoft.com   -AppId <client-id>
Test-ApplicationAccessPolicy -Identity ceo@jteaton.com                     -AppId <client-id>
```

Policies take ~1 hour to propagate fully — give it time before you run
`pstmigrate validate` if you just created one.

### 5. Plug the values into config.toml

After step 1–3 you have:

```
tenant_id     = Directory (tenant) ID    from app Overview page
client_id     = Application (client) ID  from app Overview page
client_secret = Value column             from Certificates & secrets → Client secrets
```

Open `config.toml` (created earlier by `pstmigrate init-config`) and paste
them into the `[[apps]]` block. Then run:

```powershell
pstmigrate validate -c config.toml -m mapping.csv
```

A green `Graph token (per app)` row confirms the values are correct.

### Want higher throughput? Repeat steps 1–3 N times.

Each additional Entra app registration gets its own Graph throttle bucket.
Adding 3 apps to `config.toml` triples the aggregate request rate against the
tenant. See **[Multi-app pool](#multi-app-pool--getting-past-per-app-throttles)**
below for details.

## Usage

```bash
# 1. Generate config + mapping templates
pstmigrate init-config -o config.toml
pstmigrate template    -o mapping.csv

# 2. Edit config.toml (tenant_id, client_id, secret/cert) and mapping.csv

# 3. Pre-flight: validate config, mapping, mailbox access
pstmigrate validate -c config.toml -m mapping.csv

# 4. Run the import (re-runs are safe — completed messages are skipped)
pstmigrate import -c config.toml -m mapping.csv

# 5. Inspect status of past runs
pstmigrate status -c config.toml

# Optional: start the local web admin GUI
pstmigrate web -c config.toml -m mapping.csv
```

## Local web GUI

The CLI remains the canonical automation surface, but the package also includes
a local-only web admin UI for day-of-migration operations:

### Windows PowerShell

From the Windows migration machine:

```powershell
cd C:\Tools\JTET-Email-Migration\Graph-Migrate

# Activate the virtual environment created during install
.\.venv\Scripts\Activate.ps1

# Install the optional web dependencies once
pip install -e ".[web]"

# Start the GUI
pstmigrate web -c config.toml -m mapping.csv --host 127.0.0.1 --port 8765
```

Then open this address in Edge/Chrome on that same Windows machine:

```text
http://127.0.0.1:8765
```

Leave the PowerShell window open while using the GUI. Stop it with
<kbd>Ctrl</kbd>+<kbd>C</kbd>.

### Linux / WSL / macOS

```bash
cd Graph-Migrate
source .venv/bin/activate
pip install -e ".[web]"
pstmigrate web -c config.toml -m mapping.csv --host 127.0.0.1 --port 8765
```

Open <http://127.0.0.1:8765>. The GUI reads the same `config.toml`,
`mapping.csv`, logs, work directory, and SQLite state DB as the CLI. It does
not store any separate Graph credentials.

If `config.toml` does not exist yet, open the **Config** tab and fill in the
tenant ID/domain, client ID, client secret or certificate path, concurrency,
and local paths. Saving the form creates `config.toml`; existing secrets are
masked in the browser and preserved when the secret field is left blank.

The GUI includes:

- Dashboard and prior-run status from `.pstmigrate-state/state.sqlite`
- Config and mapping preview pages
- Validation, import, import-all, purge, and reset-state job forms
- A run-page checkbox for mail remediation re-runs that imports previously skipped duplicate message copies
- One active mutating job at a time
- Explicit confirmation for destructive tools
- Logs browser for JSONL run logs

### Selecting a subset of mailboxes / PSTs

Both `validate` and `import` accept selection flags. Filters are applied in
order: mailbox match → PST filename match → row-count limit. UPN matching is
case-insensitive.

```bash
# One mailbox (canary the smallest user first)
pstmigrate import -c config.toml -m mapping.csv -M emmak@contoso.onmicrosoft.com

# Several mailboxes (repeat -M)
pstmigrate import -c config.toml -m mapping.csv \
  -M allysonp@contoso.onmicrosoft.com \
  -M emmak@contoso.onmicrosoft.com \
  -M tinad@contoso.onmicrosoft.com

# Just the first row of the mapping (smoke test)
pstmigrate import -c config.toml -m mapping.csv -n 1

# Pick a row by source PST filename substring
pstmigrate import -c config.toml -m mapping.csv -P jsmith4.23.pst

# Preview what would run without doing anything
pstmigrate import -c config.toml -m mapping.csv -M emmak@contoso.onmicrosoft.com --list

# Pre-flight a single mailbox
pstmigrate validate -c config.toml -m mapping.csv -M emmak@contoso.onmicrosoft.com
```

Resume is per-(mailbox, PST), so it is safe to migrate a single mailbox now,
then come back later and run the full mapping — the canary's already-imported
messages will be skipped.

If you need to remediate a mailbox by importing duplicate message copies that
were previously recorded as skipped because another source message had the same
Message-ID, scope the re-run to that mailbox/PST and add
`--import-skipped-duplicates`:

```bash
pstmigrate import -c config.toml -m mapping.csv \
  -M emmak@contoso.onmicrosoft.com \
  --import-skipped-duplicates
```

Exact source rows that are already marked done are still skipped; the flag only
changes handling for separate source `.eml` rows whose dedupe key already has a
completed import in the target mailbox.

## Post-import remediation: clearing stuck "draft" flags

PSTs imported through the Microsoft Graph mail API (this tool, plus every
other Graph-based migrator we've tested) land each message with the MAPI
`MSGFLAG_UNSENT` bit (0x08) of `PR_MESSAGE_FLAGS` set. In Outlook on the
Web that surfaces as a "Draft" badge on every imported item and inflates
the Drafts folder with hundreds of historical messages. Content is
intact — it's a metadata-only problem — but the UX is wrong.

We've empirically verified (commits `4cf2171`, `bfcea7f`, `0dc779d`)
that **neither Microsoft Graph nor EWS can clear this flag** on Exchange
Online: every documented write path (PATCH `isDraft`, PATCH the extended
property, `/copy`, EWS `SetItemField`, `DeleteItemField`, multi-property
updates, `CreateItem` from MIME) returns Success while leaving the bit
unchanged. The cloud's MAPI store rejects the writes silently at a layer
beneath the public API frontends.

Outlook desktop talks to Exchange Online over **MAPI/HTTP**, a separate
protocol from Graph and EWS, and on that path the property write is
accepted. `Fix-DraftsViaOutlook.ps1` automates the fix.

### One-time setup

Grant the admin account that's signed in to Outlook FullAccess on every
imported mailbox, with auto-mapping disabled so the mailboxes don't
attach to the profile:

```powershell
Connect-ExchangeOnline
$admin = "your-admin@jteatono365.onmicrosoft.com"
Import-Csv .\mapping.csv | ForEach-Object {
    Add-MailboxPermission -Identity $_.TargetMailbox -User $admin `
        -AccessRights FullAccess -InheritanceType All `
        -AutoMapping $false
}
```

`-AutoMapping $false` matters; with it `$true` Outlook auto-attaches
all 13 mailboxes to the admin's profile, slowing startup and cluttering
the folder pane. We open them programmatically.

### Fix one mailbox at a time first

```powershell
.\Fix-DraftsViaOutlook.ps1 `
    -Mailbox allysonp@jteatono365.onmicrosoft.com -DryRun

.\Fix-DraftsViaOutlook.ps1 `
    -Mailbox allysonp@jteatono365.onmicrosoft.com

.\.venv\Scripts\python.exe _inspect_dates.py -c config.toml `
    --mailbox allysonp@jteatono365.onmicrosoft.com --folder sentitems --top 5
```

Expected after the second command: every row shows `draft=no` and
`msgFlg=0x0411` (UNSENT bit cleared). Throughput is roughly 25 items/sec,
so a mailbox with 485 stuck drafts finishes in well under a minute.

### Then everything

```powershell
.\Fix-DraftsViaOutlook.ps1 -MappingFile .\mapping.csv
```

The script walks every mail folder in each mailbox except the well-known
**Drafts**, **Deleted Items**, **Outbox**, and **Junk Email** folders
(legitimate drafts there should stay drafts). It uses an Outlook
`Items.Restrict` DASL filter so it only pulls items that actually have
the UNSENT bit set, rather than iterating every message.

### Diagnostic helpers (kept for posterity)

These scripts proved Graph and EWS could not solve the problem and are
preserved as evidence / future reference:

- `_fix_drafts.py` — Graph PATCH attempt (returns 200 OK, no effect).
- `_test_copy_strategy.py` — Graph `/copy` test (new copy is also a draft).
- `_fix_drafts_ews.py` — EWS UpdateItem batch (reports Success, no effect).
- `_test_ews_strategies.py` — five EWS write paths in one diagnostic;
  table output for each.
- `_test_ews_recreate.py` — EWS CreateItem from MIME into a non-Drafts
  folder (new item is also a draft).
- `_inspect_dates.py` — read-only inspector that prints `isDraft`,
  `isRead`, `PR_MESSAGE_FLAGS`, and date columns. Use this after any
  fix attempt to confirm whether the bit actually moved.

## How it works

```
mapping.csv ──┐
              v
         ┌──────────────────────────────────────┐
         │  Orchestrator (ThreadPool: mailboxes)│
         └─────┬────────────────────────────────┘
               │  per (PST, mailbox)
               v
       ┌────────────────────┐    ┌──────────────────────┐
       │ readpst → .eml/dir │ -> │ MIME enumerator      │
       └────────────────────┘    └────────┬─────────────┘
                                          │ per message
                                          v
                                ┌─────────────────────────┐
                                │ workers (per mailbox):  │
                                │  - dedupe via SQLite    │
                                │  - ensure folder path   │
                                │  - POST MIME or         │
                                │     create + attach     │
                                │  - record outcome       │
                                └────────────┬────────────┘
                                             v
                                ┌─────────────────────────┐
                                │ Graph (with Retry-After,│
                                │  jittered backoff)      │
                                └─────────────────────────┘
```

State lives in `.pstmigrate-state/state.sqlite`. It records every message's
dedupe key (RFC 5322 `Message-ID` when present, else SHA-256 of the bytes), so
re-runs are idempotent and concurrent safe across worker threads.

## Tuning

- **Throughput plateau** is per-mailbox-per-app. Going past
  `workers_per_mailbox=4` against a single app usually just earns more 429s.
  Scale by adding mailboxes in parallel (`max_parallel_mailboxes`) AND by
  adding more app registrations (see "Multi-app pool" below).
- **First run on a populated mailbox** spends time creating folder hierarchy.
  Subsequent runs hit the cached folder IDs in SQLite — no list-and-find calls.
- **Retry budget** (`max_retries`) is per request. With the default backoff
  ladder the longest possible wait between attempts is ~120s, total worst-case
  budget per request is ~10 minutes.

## Multi-app pool — getting past per-app throttles

Microsoft Graph computes throttle budgets per `(app_id, target_resource)`.
That means **two separate app registrations granted the same permission get
two independent throttle budgets** against the same mailbox. You can stack
this trick:

1. In Entra → App registrations, create N apps (e.g. `pstmigrate-1`,
   `pstmigrate-2`, `pstmigrate-3`).
2. Grant each one `Mail.ReadWrite` (application) and admin consent.
3. Add a `[[apps]]` block per app in `config.toml`:
   ```toml
   [[apps]]
   name = "primary"
   tenant_id = "contoso.onmicrosoft.com"
   client_id = "..."
   client_secret = "..."

   [[apps]]
   name = "secondary"
   tenant_id = "contoso.onmicrosoft.com"
   client_id = "..."
   client_secret = "..."
   ```
4. Run `pstmigrate validate` — it tests every app independently.

The orchestrator round-robins one app per message and pins it for that
message's full request chain (folder ensure + create + attachments) so each
upload's stats and throttle budget are coherent. Aggregate throughput scales
~linearly with the pool size until you hit the per-mailbox ceiling, which
Microsoft will lift on request for migrations.

The `status` command shows a per-app rollup so you can verify load is
actually balancing across the pool.

> Per-app stats also surface in the import summary table — if one app is
> taking all the 429s while others are idle, you know to investigate
> (usually a not-fully-consented app, or a tenant Conditional Access policy
> applied to one app and not the others).

## Things this does *not* do (be honest)

- Migrate calendar items, contacts, tasks, journal entries, notes.
- Preserve read/unread state, importance, flag/follow-up, categories,
  custom MAPI named properties, retention tags.
- Handle password-protected PSTs without you supplying the password to readpst
  separately (open issue: the binary doesn't accept passwords on the
  command line — protected PSTs need to be unlocked beforehand).
- Repair corrupted PSTs (run `scanpst.exe` on Windows first).
- Migrate from PSTs > ~50 GB efficiently — readpst will succeed but the
  extraction phase becomes the bottleneck. Split very large PSTs first.

## Output artefacts per run

```
.pstmigrate-state/
  state.sqlite                  # dedupe + resume + folder cache
.pstmigrate-work/
  <pst-stem>/...                # readpst extracted .eml tree (deletable after success)
logs/
  import_20260424_204500.jsonl  # structured per-event log
```

## License

`jtet-pstmigrate` is released under the [Apache License, Version 2.0](LICENSE).
See [NOTICE](NOTICE) for third-party attribution and a note on the GPL boundary
with the external `readpst` binary, which this project invokes but does not
bundle or redistribute.
