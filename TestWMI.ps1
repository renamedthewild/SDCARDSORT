$LogFile = "C:\Users\LukeWilden\RURAL IT SOLUTIONS\DRONE - Documents\R2F\TestLog.txt"
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    Write-Host "$timestamp - $Message"
}
Write-Log "Starting WMI test..."
Register-WmiEvent -Query "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_LogicalDisk' AND TargetInstance.DriveType=2" -SourceIdentifier "TestSDCard" -Action {
    $drive = $Event.SourceEventArgs.NewEvent.TargetInstance
    Write-Log "Detected drive: $($drive.DeviceID), DriveType: $($drive.DriveType)"
}
Write-Log "Waiting for SD card. Press Ctrl+C to stop."
while ($true) { Start-Sleep -Seconds 3600 }