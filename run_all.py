"""
run_all.py

Process all PlanetPress Workflow log files in a folder.
Files must match the naming pattern ppwYYYYMMDD.log.

Produces a combined table across all log files, sorted by date,
with an extra Date column prepended.

Usage:
    python run_all.py <log_folder> [--csv <output.csv>] [--from YYYY-MM-DD] [--to YYYY-MM-DD]
"""

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from extract_mtp_filenames import parse_log, MtpEntry


RE_LOG_FILENAME = re.compile(r"^ppw(\d{4})(\d{2})(\d{2})\.log$", re.IGNORECASE)

HEADER = ["Date", "Original mail subject", "Final PDF name", "Outcome", "Start time", "Thread ID"]


@dataclass
class DatedEntry:
    log_date: date
    entry: MtpEntry


def find_log_files(folder: Path, from_date: date | None, to_date: date | None) -> list[tuple[date, Path]]:
    results = []
    for f in folder.iterdir():
        m = RE_LOG_FILENAME.match(f.name)
        if not m:
            continue
        log_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if from_date and log_date < from_date:
            continue
        if to_date and log_date > to_date:
            continue
        results.append((log_date, f))
    return sorted(results)


def collect_entries(log_files: list[tuple[date, Path]]) -> list[DatedEntry]:
    all_entries = []
    for log_date, path in log_files:
        entries = parse_log(str(path))
        for e in entries:
            all_entries.append(DatedEntry(log_date=log_date, entry=e))
    return all_entries


def print_table(dated_entries: list[DatedEntry]) -> None:
    col_widths = [len(h) for h in HEADER]
    rows = []
    for de in dated_entries:
        e = de.entry
        row = [
            de.log_date.strftime("%Y-%m-%d"),
            e.original_name or "(none)",
            e.final_pdf_name or "(none)",
            e.outcome(),
            e.start_time,
            e.thread_id,
        ]
        rows.append(row)
        for i, cell in enumerate(row):
            col_widths[i] = max(col_widths[i], len(cell))

    fmt = "  ".join(f"{{:<{w}}}" for w in col_widths)
    sep = "  ".join("-" * w for w in col_widths)

    print(fmt.format(*HEADER))
    print(sep)
    for row in rows:
        print(fmt.format(*row))


def write_csv(dated_entries: list[DatedEntry], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADER)
        for de in dated_entries:
            e = de.entry
            writer.writerow([
                de.log_date.strftime("%Y-%m-%d"),
                e.original_name or "",
                e.final_pdf_name or "",
                e.outcome(),
                e.start_time,
                e.thread_id,
            ])
    print(f"\nCSV written to: {path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Process all ppwYYYYMMDD.log files in a folder."
    )
    parser.add_argument("log_folder", help="Folder containing the log files")
    parser.add_argument("--csv", metavar="OUTPUT_CSV", help="Write results to a CSV file")
    parser.add_argument("--from", dest="from_date", metavar="YYYY-MM-DD",
                        help="Only include logs on or after this date")
    parser.add_argument("--to", dest="to_date", metavar="YYYY-MM-DD",
                        help="Only include logs on or before this date")
    args = parser.parse_args()

    folder = Path(args.log_folder)
    if not folder.is_dir():
        print(f"Folder not found: {folder}", file=sys.stderr)
        sys.exit(1)

    from_date = date.fromisoformat(args.from_date) if args.from_date else None
    to_date   = date.fromisoformat(args.to_date)   if args.to_date   else None

    log_files = find_log_files(folder, from_date, to_date)
    if not log_files:
        print("No matching log files found.", file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(log_files)} log file(s).\n")

    dated_entries = collect_entries(log_files)
    if not dated_entries:
        print("No MailToPost entries found across all log files.", file=sys.stderr)
        sys.exit(1)

    uitval    = sum(1 for de in dated_entries if de.entry.uitval)
    processed = len(dated_entries) - uitval
    print(f"Totaal: {len(dated_entries)}  |  Verwerkt: {processed}  |  Uitval: {uitval}\n")

    print_table(dated_entries)

    if args.csv:
        write_csv(dated_entries, args.csv)


if __name__ == "__main__":
    main()
