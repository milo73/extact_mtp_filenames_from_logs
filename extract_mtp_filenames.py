"""
extract_mtp_filenames.py

Parse a PlanetPress Workflow log file and produce a list of
MailToPost process entries showing:
  - original mail subject name  (from step [0002] %{Bestandsnaam})
  - final PDF name              (from "File sent" line containing MailToPost\ToPrinter\C)
  - outcome: sent to printer, or moved to Uitval

Usage:
    python extract_mtp_filenames.py <logfile> [--csv <output.csv>]

Output (stdout):
    Tab-separated table; optionally also written as CSV.
"""

import re
import sys
import csv
import argparse
from dataclasses import dataclass
from typing import Optional


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

# Start of a new WPROC block
RE_WPROC_START = re.compile(
    r"^WPROC:\s+MailToPost\s+\(thread id:\s*(\d+)\)\s+-\s+(.+)$"
)

# Step [0002]: original filename
RE_STEP2_NAME = re.compile(
    r"\[0002\].*?%\{Bestandsnaam\}\s+is set to\s+\"(.+?)\""
)

# Final PDF filename: "File sent" line whose path contains MailToPost\ToPrinter\C
RE_PRINTER_FILE = re.compile(
    r"File sent\s*:.*MailToPost[/\\]ToPrinter[/\\](C[^/\\,]+\.pdf),\s*size:"
)

# Uitval indicator: any line that stores/sets the %{Uitval} variable
RE_UITVAL = re.compile(
    r"%\{Uitval\}"
)

# End of a WPROC block
RE_WPROC_END = re.compile(
    r"^WPROC:\s+MailToPost\s+\(thread id:\s*(\d+)\)\s+complete"
)


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

@dataclass
class MtpEntry:
    thread_id: str
    start_time: str
    original_name: Optional[str] = None
    final_pdf_name: Optional[str] = None
    uitval: bool = False

    def outcome(self) -> str:
        if self.uitval:
            return "Uitval"
        if self.final_pdf_name:
            return "Sent to printer"
        return "Unknown"


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

def parse_log(path: str) -> list[MtpEntry]:
    entries: list[MtpEntry] = []
    current: Optional[MtpEntry] = None

    with open(path, encoding="utf-8", errors="replace") as fh:
        for line in fh:
            line = line.rstrip("\n")

            # New WPROC block
            m = RE_WPROC_START.match(line)
            if m:
                current = MtpEntry(thread_id=m.group(1), start_time=m.group(2).strip())
                entries.append(current)
                continue

            if current is None:
                continue

            # Step [0002]: original name
            if not current.original_name:
                m = RE_STEP2_NAME.search(line)
                if m:
                    current.original_name = m.group(1)
                    continue

            # ToPrinter line: final PDF filename sent to printer
            if not current.final_pdf_name:
                m = RE_PRINTER_FILE.search(line)
                if m:
                    current.final_pdf_name = m.group(1)
                    continue

            # Step [0021]: Uitval
            if RE_UITVAL.search(line):
                current.uitval = True
                continue

            # End of block – reset pointer
            if RE_WPROC_END.match(line):
                current = None

    return entries


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

HEADER = ["Original mail subject", "Final PDF name", "Outcome", "Start time", "Thread ID"]


def print_table(entries: list[MtpEntry]) -> None:
    col_widths = [len(h) for h in HEADER]
    rows = []
    for e in entries:
        row = [
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


def write_csv(entries: list[MtpEntry], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(HEADER)
        for e in entries:
            writer.writerow([
                e.original_name or "",
                e.final_pdf_name or "",
                e.outcome(),
                e.start_time,
                e.thread_id,
            ])
    print(f"\nCSV written to: {path}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract MailToPost filenames from a PlanetPress Workflow log."
    )
    parser.add_argument("logfile", help="Path to the .log file")
    parser.add_argument("--csv", metavar="OUTPUT_CSV", help="Also write results to a CSV file")
    args = parser.parse_args()

    entries = parse_log(args.logfile)

    if not entries:
        print("No MailToPost entries found in the log.", file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(entries)} MailToPost process(es).\n")
    print_table(entries)

    if args.csv:
        write_csv(entries, args.csv)


if __name__ == "__main__":
    main()
