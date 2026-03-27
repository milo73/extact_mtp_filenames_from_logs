"""
run_daily.py

Daily runner for extract_mtp_filenames.
Intended to be called by Windows Task Scheduler each morning.

  1. Reads config.ini for log folder, SMTP settings, and recipients.
  2. Locates yesterday's log file  (ppwYYYYMMDD.log).
  3. Parses the log with extract_mtp_filenames.parse_log().
  4. Sends an HTML summary e-mail to the configured recipients.

Usage:
    python run_daily.py [--config config.ini] [--date YYYY-MM-DD]

    --config   Path to config file (default: config.ini next to this script)
    --date     Process a specific date instead of yesterday (for manual reruns)
"""

import argparse
import configparser
import os
import smtplib
import sys
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# Import the parser from the sibling module
sys.path.insert(0, str(Path(__file__).parent))
from extract_mtp_filenames import parse_log, MtpEntry


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def log_path_for(log_folder: str, target_date: date) -> Path:
    filename = target_date.strftime("ppw%Y%m%d.log")
    return Path(log_folder) / filename


def build_html(entries: list[MtpEntry], target_date: date, log_file: Path) -> str:
    date_str = target_date.strftime("%d-%m-%Y")
    total     = len(entries)
    uitval    = sum(1 for e in entries if e.uitval)
    processed = total - uitval

    rows_html = ""
    for e in entries:
        outcome_style = 'color:#c0392b;font-weight:bold' if e.uitval else 'color:#27ae60'
        rows_html += (
            f"<tr>"
            f"<td>{e.original_name or '(none)'}</td>"
            f"<td>{e.final_pdf_name or '(none)'}</td>"
            f"<td style='{outcome_style}'>{e.outcome()}</td>"
            f"<td>{e.start_time}</td>"
            f"</tr>\n"
        )

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body       {{ font-family: Arial, sans-serif; font-size: 13px; color: #333; }}
  h2         {{ color: #2c3e50; }}
  .summary   {{ margin-bottom: 16px; }}
  table      {{ border-collapse: collapse; width: 100%; }}
  th         {{ background: #2c3e50; color: #fff; padding: 7px 10px; text-align: left; }}
  td         {{ padding: 6px 10px; border-bottom: 1px solid #ddd; }}
  tr:hover td{{ background: #f5f5f5; }}
  .footer    {{ margin-top: 12px; font-size: 11px; color: #999; }}
</style>
</head>
<body>
<h2>MailToPost verwerking &ndash; {date_str}</h2>
<div class="summary">
  <strong>Totaal:</strong> {total} &nbsp;|&nbsp;
  <strong style="color:#27ae60">Verwerkt:</strong> {processed} &nbsp;|&nbsp;
  <strong style="color:#c0392b">Uitval:</strong> {uitval}
</div>
<table>
  <thead>
    <tr>
      <th>Originele mail</th>
      <th>Definitieve PDF naam</th>
      <th>Uitkomst</th>
      <th>Starttijd</th>
    </tr>
  </thead>
  <tbody>
{rows_html}  </tbody>
</table>
<div class="footer">Logbestand: {log_file}</div>
</body>
</html>"""


def build_plain(entries: list[MtpEntry], target_date: date, log_file: Path) -> str:
    date_str  = target_date.strftime("%d-%m-%Y")
    uitval    = sum(1 for e in entries if e.uitval)
    processed = len(entries) - uitval

    lines = [
        f"MailToPost verwerking - {date_str}",
        f"Totaal: {len(entries)}  |  Verwerkt: {processed}  |  Uitval: {uitval}",
        "",
        f"{'Originele mail':<35} {'Definitieve PDF naam':<40} {'Uitkomst':<16} Starttijd",
        "-" * 105,
    ]
    for e in entries:
        lines.append(
            f"{(e.original_name or '(none)'):<35} "
            f"{(e.final_pdf_name or '(none)'):<40} "
            f"{e.outcome():<16} "
            f"{e.start_time}"
        )
    lines.append("")
    lines.append(f"Logbestand: {log_file}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# E-mail
# ---------------------------------------------------------------------------

def send_mail(cfg: configparser.ConfigParser,
              subject: str,
              html_body: str,
              plain_body: str) -> None:

    recipients = [r.strip() for r in cfg["email"]["recipients"].split(",") if r.strip()]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = cfg["smtp"]["from"]
    msg["To"]      = ", ".join(recipients)

    msg.attach(MIMEText(plain_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body,  "html",  "utf-8"))

    host    = cfg["smtp"]["host"]
    port    = int(cfg["smtp"]["port"])
    use_tls = cfg["smtp"].getboolean("use_tls")
    user    = cfg["smtp"].get("username", "")
    passwd  = cfg["smtp"].get("password", "")

    if use_tls:
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
    else:
        server = smtplib.SMTP(host, port)

    if user:
        server.login(user, passwd)

    server.sendmail(cfg["smtp"]["from"], recipients, msg.as_bytes())
    server.quit()
    print(f"E-mail verzonden naar: {', '.join(recipients)}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    script_dir = Path(__file__).parent

    parser = argparse.ArgumentParser(
        description="Daily MailToPost log runner – parses yesterday's log and e-mails results."
    )
    parser.add_argument(
        "--config", default=str(script_dir / "config.ini"),
        help="Path to config.ini (default: config.ini next to this script)"
    )
    parser.add_argument(
        "--date", metavar="YYYY-MM-DD",
        help="Process a specific date instead of yesterday"
    )
    args = parser.parse_args()

    # Load config
    cfg = configparser.ConfigParser()
    if not cfg.read(args.config):
        print(f"Config file not found: {args.config}", file=sys.stderr)
        sys.exit(1)

    # Determine target date
    if args.date:
        target_date = date.fromisoformat(args.date)
    else:
        target_date = date.today() - timedelta(days=1)

    # Locate log file
    log_file = log_path_for(cfg["paths"]["log_folder"], target_date)
    if not log_file.exists():
        print(f"Log file not found: {log_file}", file=sys.stderr)
        sys.exit(1)

    print(f"Verwerken: {log_file}")

    # Parse
    entries = parse_log(str(log_file))
    if not entries:
        print("Geen MailToPost-vermeldingen gevonden in het logbestand.", file=sys.stderr)
        sys.exit(1)

    print(f"{len(entries)} processen gevonden.")

    # Build e-mail content
    date_str = target_date.strftime("%d-%m-%Y")
    subject  = cfg["email"]["subject"].format(date=date_str)
    html     = build_html(entries, target_date, log_file)
    plain    = build_plain(entries, target_date, log_file)

    # Send
    send_mail(cfg, subject, html, plain)


if __name__ == "__main__":
    main()
