# Ensure the script runs with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrative privileges. Please run as administrator."
    exit 1
}

# Log the reboot operation
$logFile = "C:\reboot-log.txt"
"Starting reboot process at $(Get-Date)" | Out-File -FilePath $logFile -Append

# Reboot the machine
try {
    Restart-Computer -Force
} catch {
    "Error during reboot: $_" | Out-File -FilePath $logFile -Append
    exit 2
}

# Log success (if somehow reachable post-reboot, useful for debugging)
"Reboot triggered successfully at $(Get-Date)" | Out-File -FilePath $logFile -Append
