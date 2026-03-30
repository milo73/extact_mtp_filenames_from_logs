# mount_share.ps1
# Ensures the Azure file share is reachable and mapped as drive Z:

$storageAccount = "stddifilepweu02.file.core.windows.net"
$sharePath       = "\\$storageAccount\ddi-data-p-01"

$result = Test-NetConnection -ComputerName $storageAccount -Port 445
if (-not $result.TcpTestSucceeded) {
    Write-Error "Cannot reach $storageAccount on port 445. Check firewall or VPN."
    exit 1
}

# Register credentials (cmdkey persists across reboots)
cmd.exe /C "cmdkey /add:`"$storageAccount`" /user:`"localhost\stddifilepweu02`" /pass:`"$env:AZURE_STORAGE_KEY`""

# Map drive Z: if not already present
if (-not (Test-Path 'Z:\')) {
    New-PSDrive -Name Z -PSProvider FileSystem -Root $sharePath -Persist | Out-Null
    Write-Host "Drive Z: mapped."
} else {
    Write-Host "Drive Z: already mapped."
}
