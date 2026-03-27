# extract_mtp_filenames

Parse a PlanetPress Workflow log file and produce a list of MailToPost jobs, mapping each original mail subject name to its final PDF name and outcome.

## How it works

Each `WPROC: MailToPost` block in the log is parsed for three things:

| Step | What is captured |
|------|-----------------|
| `[0002] %{Bestandsnaam}` | Original mail filename (e.g. `BPT-00-00-33267563-104257861`) |
| `[0058] File sent` | Final PDF filename including C4-/C5- prefix (e.g. `C5-BPT-00-00-010P62GEUBQ1PC6.pdf`) |
| `[0021] 1: Document naar Uitval` | Indicates the job was moved to the Uitval folder instead |

## Files

| File | Purpose |
|------|---------|
| `extract_mtp_filenames.py` | Core parser – can be used standalone |
| `run_daily.py` | Daily runner – finds yesterday's log, parses it, sends e-mail |
| `run_daily.bat` | Windows Task Scheduler entry point |
| `config.ini` | SMTP settings, recipients, log folder path |

## Manual usage

```bash
# Print table to stdout
python extract_mtp_filenames.py <logfile>

# Also export to CSV
python extract_mtp_filenames.py <logfile> --csv output.csv
```

## Daily automated run

### 1. Configure `config.ini`

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

### 2. Update paths in `run_daily.bat`

Edit the `PYTHON`, `SCRIPT`, `CONFIG`, and `LOGDIR` variables at the top of `run_daily.bat` to match your server layout.

### 3. Create the scheduled task

Open **Task Scheduler** and create a new task:

| Setting | Value |
|---------|-------|
| General | Run whether user is logged on or not |
| Trigger | Daily, e.g. 06:00 AM |
| Action – Program | `C:\Windows\System32\cmd.exe` |
| Action – Arguments | `/c "F:\Planetpress\MailToPost\Scripts\run_daily.bat"` |
| Settings | Restart on failure: every 5 min, up to 3 times |

Or create it from an elevated PowerShell prompt:

```powershell
$action  = New-ScheduledTaskAction -Execute 'cmd.exe' `
             -Argument '/c "F:\Planetpress\MailToPost\Scripts\run_daily.bat"'
$trigger = New-ScheduledTaskTrigger -Daily -At '06:00AM'
$settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName 'MailToPost Daily Report' `
  -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
```

### Manual rerun for a specific date

```bash
python run_daily.py --date 2026-03-25
```

## Example e-mail output

The e-mail contains a summary line (Totaal / Verwerkt / Uitval) followed by an HTML table. Uitval rows are highlighted in red.

## Requirements

Python 3.10+ (no third-party dependencies).
