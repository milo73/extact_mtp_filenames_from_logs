<#
.SYNOPSIS
    Parse a PlanetPress Workflow log file and list all MailToPost jobs.

.DESCRIPTION
    Extracts per-job data from each WPROC: MailToPost block:
      - Original mail filename   ([0002] %{Bestandsnaam})
      - Final PDF sent to printer (File sent line containing MailToPost\ToPrinter\C)
      - Outcome: Sent to printer | Uitval | Unknown

.PARAMETER LogFile
    Path to the log file to parse.

.PARAMETER CsvOutput
    Optional. Path to write CSV output.

.EXAMPLE
    .\extract_mtp_filenames.ps1 -LogFile D:\Logs\ppw20260311.log
    .\extract_mtp_filenames.ps1 -LogFile D:\Logs\ppw20260311.log -CsvOutput out.csv
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$LogFile,

    [string]$CsvOutput
)

# ---------------------------------------------------------------------------
# Shared parse function (dot-sourced by run_daily.ps1 and run_all.ps1)
# ---------------------------------------------------------------------------
function Parse-MtpLog {
    param([string]$Path)

    $reWprocStart   = [regex]'^WPROC:\s+MailToPost\s+\(thread id:\s*(\d+)\)\s+-\s+(.+)$'
    $reStep2Name    = [regex]'\[0002\].*?%\{Bestandsnaam\}\s+is set to\s+"(.+?)"'
    $rePrinterFile  = [regex]'File sent\s*:.*MailToPost[/\\]ToPrinter[/\\](C[^/\\,]+\.pdf),\s*size:'
    $reUitval       = [regex]'%\{Uitval\}'
    $reWprocEnd     = [regex]'^WPROC:\s+MailToPost\s+\(thread id:\s*(\d+)\)\s+complete'

    $entries = [System.Collections.Generic.List[hashtable]]::new()
    $current = $null

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        $m = $reWprocStart.Match($line)
        if ($m.Success) {
            $current = @{
                ThreadId     = $m.Groups[1].Value
                StartTime    = $m.Groups[2].Value.Trim()
                OriginalName = $null
                FinalPdfName = $null
                Uitval       = $false
            }
            $entries.Add($current)
            continue
        }

        if ($null -eq $current) { continue }

        if (-not $current.OriginalName) {
            $m = $reStep2Name.Match($line)
            if ($m.Success) { $current.OriginalName = $m.Groups[1].Value; continue }
        }

        if (-not $current.FinalPdfName) {
            $m = $rePrinterFile.Match($line)
            if ($m.Success) { $current.FinalPdfName = $m.Groups[1].Value; continue }
        }

        if ($reUitval.IsMatch($line)) { $current.Uitval = $true; continue }

        if ($reWprocEnd.IsMatch($line)) { $current = $null }
    }

    # Return PSCustomObjects with computed Outcome
    foreach ($e in $entries) {
        $outcome = if ($e.Uitval) { 'Uitval' }
                   elseif ($e.FinalPdfName) { 'Sent to printer' }
                   else { 'Unknown' }
        [PSCustomObject]@{
            OriginalName = if ($e.OriginalName) { $e.OriginalName } else { '(none)' }
            FinalPdfName = if ($e.FinalPdfName) { $e.FinalPdfName } else { '(none)' }
            Outcome      = $outcome
            StartTime    = $e.StartTime
            ThreadId     = $e.ThreadId
        }
    }
}

# ---------------------------------------------------------------------------
# Standalone runner
# ---------------------------------------------------------------------------
if (-not (Test-Path $LogFile)) {
    Write-Error "Log file not found: $LogFile"
    exit 1
}

$results = Parse-MtpLog -Path $LogFile

if (-not $results) {
    Write-Error "No MailToPost entries found in: $LogFile"
    exit 1
}

Write-Host "Found $(@($results).Count) MailToPost process(es).`n"
$results | Format-Table -AutoSize

if ($CsvOutput) {
    $results | Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV written to: $CsvOutput"
}
