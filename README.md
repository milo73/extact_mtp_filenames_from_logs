# extract_mtp_filenames

Parse a PlanetPress Workflow log file and produce a list of MailToPost jobs, mapping each original mail subject name to its final PDF name and outcome.

## How it works

Each `WPROC: MailToPost` block in the log is parsed for three things:

| Log content | What is captured |
|-------------|-----------------|
| `[0002] %{Bestandsnaam}` | Original mail filename (e.g. `BPT-00-00-33267563-104257861`) |
| `File sent` path containing `MailToPost\ToPrinter\C` | Final PDF filename including C4-/C5- prefix (e.g. `C5-BPT-00-00-010P62GEUBQ1PC6.pdf`) |
| `[0021] 1: Document naar Uitval` | Indicates the job was moved to the Uitval folder instead |

## Branches

| Branch | Implementation |
|--------|---------------|
| `main` | Python |
| `feature/powershell-port` | PowerShell (no Python required) |

---

## Python (`main`)

### Files

| File | Purpose |
|------|---------|
| `extract_mtp_filenames.py` | Core parser – can be used standalone |
| `run_daily.py` | Daily runner – finds yesterday's log, parses it, sends e-mail |
| `run_all.py` | Processes all log files in a folder |
| `run_daily.bat` | Windows Task Scheduler entry point |
| `mount_share.ps1` | Connects and maps the Azure file share |
| `config.ini` | SMTP settings, recipients, log folder path |

### Usage

```bash
# Parse a single log file
python extract_mtp_filenames.py <logfile>

# Also export to CSV
python extract_mtp_filenames.py <logfile> --csv output.csv

# Process all log files in a folder
python run_all.py F:\PlanetPress\Logs

# Filter by date range
python run_all.py F:\PlanetPress\Logs --from 2026-03-01 --to 2026-03-31 --csv march.csv

# Manual rerun of the daily mailer for a specific date
python run_daily.py --date 2026-03-25
```

### Requirements

Python 3.10+ (no third-party dependencies).

---

## PowerShell (`feature/powershell-port`)

### Files

| File | Purpose |
|------|---------|
| `extract_mtp_filenames.ps1` | Core parser – defines `Parse-MtpLog`; can be used standalone |
| `run_daily.ps1` | Daily runner – finds yesterday's log, parses it, sends e-mail |
| `run_all.ps1` | Processes all log files in a folder |
| `run_daily.bat` | Windows Task Scheduler entry point |
| `mount_share.ps1` | Connects and maps the Azure file share |
| `config.ini` | SMTP settings, recipients, log folder path (shared with Python) |

### Usage

```powershell
# Parse a single log file
.\extract_mtp_filenames.ps1 -LogFile D:\Logs\ppw20260311.log

# Also export to CSV
.\extract_mtp_filenames.ps1 -LogFile D:\Logs\ppw20260311.log -CsvOutput out.csv

# Process all log files in a folder
.\run_all.ps1 -LogFolder F:\PlanetPress\Logs

# Filter by date range
.\run_all.ps1 -LogFolder F:\PlanetPress\Logs -FromDate 2026-03-01 -ToDate 2026-03-31 -CsvOutput march.csv

# Manual rerun of the daily mailer for a specific date
.\run_daily.ps1 -Date 2026-03-25
```

### Requirements

Windows PowerShell 5.1+ or PowerShell 7+. No additional modules needed.

---

## Configuration (`config.ini`)

Used by both implementations.

```ini
[paths]
log_folder = F:\PlanetPress\Logs

[smtp]
host     = smtp.example.com
port     = 587
use_tls  = true
username = sender@example.com
password = secret
from     = sender@example.com

[email]
recipients = recipient1@example.com, recipient2@example.com
subject    = MailToPost verwerking {date}
```

> **Office 365:** set `host = smtp.office365.com`, `port = 587`, `use_tls = true`.  
> **Local relay (no auth):** leave `username` and `password` empty, set `use_tls = false`.

---

## Daily automated run (Task Scheduler)

### 1. Set the Azure storage key as a system environment variable

In an elevated PowerShell prompt on the server:

```powershell
[System.Environment]::SetEnvironmentVariable('AZURE_STORAGE_KEY', '<your-key>', 'Machine')
Restart-Service Schedule
```

### 2. Update `LOGDIR` in `run_daily.bat`

Edit the `LOGDIR` variable to point to where you want the run log written.

### 3. Create the scheduled task

Open **Task Scheduler** and create a new task:

| Setting | Value |
|---------|-------|
| General | Run whether user is logged on or not |
| Trigger | Daily, e.g. 06:00 AM |
| Action – Program | `C:\Windows\System32\cmd.exe` |
| Action – Arguments | `/c "F:\Planetpress\MailToPost\Scripts\run_daily.bat"` |
| Settings | Restart on failure: every 5 min, up to 3 times |

Or from an elevated PowerShell prompt:

```powershell
$action   = New-ScheduledTaskAction -Execute 'cmd.exe' `
              -Argument '/c "F:\Planetpress\MailToPost\Scripts\run_daily.bat"'
$trigger  = New-ScheduledTaskTrigger -Daily -At '06:00AM'
$settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName 'MailToPost Daily Report' `
  -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
```

---

## E-mail output

Each daily run sends an e-mail containing:
- A summary line: **Totaal / Verwerkt / Uitval**
- An HTML table with one row per job; Uitval rows highlighted in red
- A plain-text fallback for mail clients that do not render HTML
