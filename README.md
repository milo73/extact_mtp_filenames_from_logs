# extract_mtp_filenames

Parse a PlanetPress Workflow log file and produce a list of MailToPost jobs, mapping each original mail subject name to its new PDF name and outcome.

## How it works

Each `WPROC: MailToPost` block in the log is parsed for three things:

| Step | What is captured |
|------|-----------------|
| `[0002] %{Bestandsnaam}` | Original mail filename (e.g. `BPT-00-00-33267563-104257861`) |
| `[0020] %{Bestandsnaam}` | New PDF base name (e.g. `BPT-00-00-010P62GEUBQ1PC6`) |
| `[0021] 1: Document naar Uitval` | Indicates the job was moved to the Uitval folder instead |

## Usage

```bash
# Print table to stdout
python extract_mtp_filenames.py <logfile>

# Also export to CSV
python extract_mtp_filenames.py <logfile> --csv output.csv
```

## Example output

```
Original mail subject         New PDF name               Outcome          Start time   Thread ID
----------------------------  -------------------------  ---------------  -----------  ---------
BPT-00-00-33267563-104257861  BPT-00-00-010P62GEUBQ1PC6  Sent to printer  10:43:06 AM  12824
BPT-00-00-104306103           BPT-00-00-010P62GG8LNZK0F  Uitval           10:43:14 AM  12824
```

## Requirements

Python 3.10+ (no third-party dependencies).
