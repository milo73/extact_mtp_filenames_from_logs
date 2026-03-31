<#
.SYNOPSIS
    Process all ppwYYYYMMDD.log files in a folder.

.DESCRIPTION
    Scans a folder for files matching ppwYYYYMMDD.log, parses each one,
    and outputs a combined table sorted by date with an extra Date column.

.PARAMETER LogFolder
    Folder containing the log files.

.PARAMETER CsvOutput
    Optional. Path to write combined CSV output.

.PARAMETER FromDate
    Only include log files on or after this date (YYYY-MM-DD).

.PARAMETER ToDate
    Only include log files on or before this date (YYYY-MM-DD).

.EXAMPLE
    .\run_all.ps1 -LogFolder F:\PlanetPress\Logs
    .\run_all.ps1 -LogFolder F:\PlanetPress\Logs -FromDate 2026-03-01 -ToDate 2026-03-31
    .\run_all.ps1 -LogFolder F:\PlanetPress\Logs -CsvOutput march_2026.csv
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$LogFolder,

    [string]$CsvOutput,
    [string]$FromDate,
    [string]$ToDate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source the parser
. (Join-Path $PSScriptRoot 'extract_mtp_filenames.ps1')

# ---------------------------------------------------------------------------
# Find matching log files
# ---------------------------------------------------------------------------
$fromDt = if ($FromDate) { [datetime]::ParseExact($FromDate, 'yyyy-MM-dd', $null) } else { $null }
$toDt   = if ($ToDate)   { [datetime]::ParseExact($ToDate,   'yyyy-MM-dd', $null) } else { $null }

$logFiles = Get-ChildItem -Path $LogFolder -Filter 'ppw????????.log' |
    Where-Object {
        $_.Name -match '^ppw(\d{4})(\d{2})(\d{2})\.log$'
        $d = [datetime]::ParseExact($matches[0] -replace 'ppw|\.log', '', $null, 'yyyyMMdd')
        (-not $fromDt -or $d -ge $fromDt) -and (-not $toDt -or $d -le $toDt)
    } |
    Sort-Object Name

if (-not $logFiles) {
    Write-Error 'No matching log files found.'
    exit 1
}

Write-Host "Found $($logFiles.Count) log file(s).`n"

# ---------------------------------------------------------------------------
# Parse all files and collect results with a Date column
# ---------------------------------------------------------------------------
$allEntries = foreach ($file in $logFiles) {
    $file.Name -match '^ppw(\d{8})\.log$' | Out-Null
    $logDate = [datetime]::ParseExact($matches[1], 'yyyyMMdd', $null)

    $entries = @(Parse-MtpLog -Path $file.FullName)
    foreach ($e in $entries) {
        [PSCustomObject]@{
            Date         = $logDate.ToString('yyyy-MM-dd')
            OriginalName = $e.OriginalName
            FinalPdfName = $e.FinalPdfName
            Outcome      = $e.Outcome
            StartTime    = $e.StartTime
            ThreadId     = $e.ThreadId
        }
    }
}

$allEntries = @($allEntries)

if ($allEntries.Count -eq 0) {
    Write-Error 'No MailToPost entries found across all log files.'
    exit 1
}

$uitval    = @($allEntries | Where-Object { $_.Outcome -eq 'Uitval' }).Count
$processed = $allEntries.Count - $uitval
Write-Host "Totaal: $($allEntries.Count)  |  Verwerkt: $processed  |  Uitval: $uitval`n"

$allEntries | Format-Table -AutoSize

if ($CsvOutput) {
    $allEntries | Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV written to: $CsvOutput"
}
