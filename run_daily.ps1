<#
.SYNOPSIS
    Daily runner: parses yesterday's log and e-mails an HTML summary.

.DESCRIPTION
    Reads config.ini, locates yesterday's ppwYYYYMMDD.log, parses it
    using Parse-MtpLog from extract_mtp_filenames.ps1, and sends an
    HTML summary e-mail to the configured recipients.

.PARAMETER ConfigFile
    Path to config.ini. Defaults to config.ini next to this script.

.PARAMETER Date
    Process a specific date (YYYY-MM-DD) instead of yesterday.
    Useful for manual reruns.

.EXAMPLE
    .\run_daily.ps1
    .\run_daily.ps1 -Date 2026-03-25
#>
param(
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config.ini'),
    [string]$Date
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source the parser
. (Join-Path $PSScriptRoot 'extract_mtp_filenames.ps1')

# ---------------------------------------------------------------------------
# Read config.ini
# ---------------------------------------------------------------------------
function Read-IniFile([string]$Path) {
    $ini = @{}
    $section = ''
    foreach ($line in Get-Content $Path) {
        $line = $line.Trim()
        if ($line -match '^\[(.+)\]$') {
            $section = $matches[1]
            $ini[$section] = @{}
        } elseif ($line -match '^([^=;#]+)=(.*)$') {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            if ($section) { $ini[$section][$key] = $val }
        }
    }
    return $ini
}

if (-not (Test-Path $ConfigFile)) {
    Write-Error "Config file not found: $ConfigFile"
    exit 1
}
$cfg = Read-IniFile $ConfigFile

# ---------------------------------------------------------------------------
# Determine target date and log file
# ---------------------------------------------------------------------------
$targetDate = if ($Date) { [datetime]::ParseExact($Date, 'yyyy-MM-dd', $null) }
              else        { (Get-Date).AddDays(-1) }

$logFileName = $targetDate.ToString('ppwyyyyyMMdd') -replace 'ppw', 'ppw'
$logFileName = 'ppw' + $targetDate.ToString('yyyyMMdd') + '.log'
$logFile     = Join-Path $cfg['paths']['log_folder'] $logFileName

if (-not (Test-Path $logFile)) {
    Write-Error "Log file not found: $logFile"
    exit 1
}

Write-Host "Verwerken: $logFile"

# ---------------------------------------------------------------------------
# Parse
# ---------------------------------------------------------------------------
$entries = @(Parse-MtpLog -Path $logFile)

if ($entries.Count -eq 0) {
    Write-Error 'Geen MailToPost-vermeldingen gevonden in het logbestand.'
    exit 1
}

$uitval    = @($entries | Where-Object { $_.Outcome -eq 'Uitval' }).Count
$processed = $entries.Count - $uitval
Write-Host "$($entries.Count) processen gevonden  |  Verwerkt: $processed  |  Uitval: $uitval"

# ---------------------------------------------------------------------------
# Build HTML e-mail body
# ---------------------------------------------------------------------------
$dateStr = $targetDate.ToString('dd-MM-yyyy')

$rowsHtml = foreach ($e in $entries) {
    $style = if ($e.Outcome -eq 'Uitval') { 'color:#c0392b;font-weight:bold' } else { 'color:#27ae60' }
    "<tr><td>$($e.OriginalName)</td><td>$($e.FinalPdfName)</td><td style='$style'>$($e.Outcome)</td><td>$($e.StartTime)</td></tr>"
}

$htmlBody = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body       { font-family: Arial, sans-serif; font-size: 13px; color: #333; }
  h2         { color: #2c3e50; }
  .summary   { margin-bottom: 16px; }
  table      { border-collapse: collapse; width: 100%; }
  th         { background: #2c3e50; color: #fff; padding: 7px 10px; text-align: left; }
  td         { padding: 6px 10px; border-bottom: 1px solid #ddd; }
  tr:hover td{ background: #f5f5f5; }
  .footer    { margin-top: 12px; font-size: 11px; color: #999; }
</style>
</head>
<body>
<h2>MailToPost verwerking &ndash; $dateStr</h2>
<div class="summary">
  <strong>Totaal:</strong> $($entries.Count) &nbsp;|&nbsp;
  <strong style="color:#27ae60">Verwerkt:</strong> $processed &nbsp;|&nbsp;
  <strong style="color:#c0392b">Uitval:</strong> $uitval
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
$($rowsHtml -join "`n")
  </tbody>
</table>
<div class="footer">Logbestand: $logFile</div>
</body>
</html>
"@

$plainBody = @"
MailToPost verwerking - $dateStr
Totaal: $($entries.Count)  |  Verwerkt: $processed  |  Uitval: $uitval

$('Originele mail'.PadRight(35)) $('Definitieve PDF naam'.PadRight(40)) $('Uitkomst'.PadRight(16)) Starttijd
$('-' * 105)
$(($entries | ForEach-Object { "$($_.OriginalName.PadRight(35)) $($_.FinalPdfName.PadRight(40)) $($_.Outcome.PadRight(16)) $($_.StartTime)" }) -join "`n")

Logbestand: $logFile
"@

# ---------------------------------------------------------------------------
# Send e-mail
# ---------------------------------------------------------------------------
$smtpCfg   = $cfg['smtp']
$emailCfg  = $cfg['email']
$recipients = $emailCfg['recipients'] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
$subject    = $emailCfg['subject'] -replace '\{date\}', $dateStr

$msg           = New-Object System.Net.Mail.MailMessage
$msg.From      = $smtpCfg['from']
$msg.Subject   = $subject
$msg.IsBodyHtml = $false

foreach ($r in $recipients) { $msg.To.Add($r) }

$plainView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($plainBody, [System.Text.Encoding]::UTF8, 'text/plain')
$htmlView  = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($htmlBody,  [System.Text.Encoding]::UTF8, 'text/html')
$msg.AlternateViews.Add($plainView)
$msg.AlternateViews.Add($htmlView)

$smtp = New-Object System.Net.Mail.SmtpClient($smtpCfg['host'], [int]$smtpCfg['port'])
$smtp.EnableSsl = [System.Convert]::ToBoolean($smtpCfg['use_tls'])
if ($smtpCfg['username']) {
    $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpCfg['username'], $smtpCfg['password'])
}

$smtp.Send($msg)
$msg.Dispose()

Write-Host "E-mail verzonden naar: $($recipients -join ', ')"
