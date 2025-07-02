# PowerShell script to copy Micro SD DCIM folder to OneDrive folder with prompts on SD card insertion
# Phase 1 Baseline: Detects SD card, prompts, copies, renames, ejects, logs, syncs to OneDrive

# Configuration variables
$LocalDestination = "C:\Users\LukeWilden\RURAL IT SOLUTIONS\DRONE - Documents\R2F\LUKE WILDEN PILOT UPLOAD"
$LogFile = "C:\Users\LukeWilden\RURAL IT SOLUTIONS\DRONE - Documents\R2F\CopyLog.txt"
$ConfigFile = "C:\Users\LukeWilden\RURAL IT SOLUTIONS\DRONE - Documents\R2F\LastBunNumber.txt"
$FirstRunFile = "C:\Users\LukeWilden\RURAL IT SOLUTIONS\DRONE - Documents\R2F\LastRunDate.txt"

# Define DeviceEject class once at script start
$code = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class DeviceEject {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(
        IntPtr hDevice,
        uint dwIoControlCode,
        IntPtr lpInBuffer,
        uint nInBufferSize,
        IntPtr lpOutBuffer,
        uint nOutBufferSize,
        out uint lpBytesReturned,
        IntPtr lpOverlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    private const uint GENERIC_READ = 0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;
    private const uint FILE_SHARE_READ = 0x00000001;
    private const uint FILE_SHARE_WRITE = 0x00000002;
    private const uint OPEN_EXISTING = 3;
    private const uint FSCTL_LOCK_VOLUME = 0x90018;
    private const uint FSCTL_DISMOUNT_VOLUME = 0x90020;
    private const uint IOCTL_STORAGE_EJECT_MEDIA = 0x2D4808;

    public static bool EjectDrive(string driveLetter) {
        string path = @"\\.\" + driveLetter.TrimEnd('\\');
        IntPtr hVolume = CreateFile(
            path,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero);

        if (hVolume == IntPtr.Zero || hVolume.ToInt64() == -1) {
            throw new Exception("Failed to open drive: " + Marshal.GetLastWin32Error());
        }

        uint bytesReturned;
        bool result;

        result = DeviceIoControl(hVolume, FSCTL_LOCK_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
        if (!result) {
            CloseHandle(hVolume);
            throw new Exception("Failed to lock volume: " + Marshal.GetLastWin32Error());
        }

        result = DeviceIoControl(hVolume, FSCTL_DISMOUNT_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
        if (!result) {
            CloseHandle(hVolume);
            throw new Exception("Failed to dismount volume: " + Marshal.GetLastWin32Error());
        }

        result = DeviceIoControl(hVolume, IOCTL_STORAGE_EJECT_MEDIA, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
        CloseHandle(hVolume);

        if (!result) {
            throw new Exception("Failed to eject media: " + Marshal.GetLastWin32Error());
        }

        return true;
    }
}
"@
if (-not ([Type]::GetType('DeviceEject'))) {
    Add-Type -TypeDefinition $code -Language CSharp
}

# Function to log messages
function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    Write-Host "$timestamp - $Message"
}

# Function to safely eject SD card with retries
function Eject-SDCard {
    param(
        $DriveLetter,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 7
    )
    $attempt = 0
    # Initial delay to allow processes to release locks
    Start-Sleep -Seconds 5
    while ($attempt -lt $MaxRetries) {
        $attempt++
        try {
            if (-not (Test-Path "$DriveLetter\")) {
                Write-Log "Drive $DriveLetter no longer accessible, assuming already ejected"
                return $true
            }
            [DeviceEject]::EjectDrive($DriveLetter)
            Start-Sleep -Seconds 2
            if (-not (Test-Path "$DriveLetter\")) {
                Write-Log "Successfully ejected SD card ($DriveLetter) on attempt $attempt"
                return $true
            }
            Write-Log "Ejection attempt $attempt failed for $DriveLetter"
        }
        catch {
            Write-Log "Error during ejection attempt $attempt"
            Write-Log $_.Exception.Message
        }
        if ($attempt -lt $MaxRetries) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
    Write-Log "Failed to eject SD card ($DriveLetter) after $MaxRetries attempts. Please remove manually."
    return $false
}

# Function to rename DCIM folder with retries
function Rename-DCIMFolder {
    param(
        $CurrentPath,
        $NewName,
        $DriveLetter,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 1
    )
    $attempt = 0
    $newPath = Join-Path $DriveLetter $NewName
    while ($attempt -lt $MaxRetries) {
        $attempt++
        try {
            Rename-Item -Path $CurrentPath -NewName $NewName -ErrorAction Stop
            Start-Sleep -Milliseconds 500
            if ((Test-Path $newPath) -and -not (Test-Path $CurrentPath)) {
                Write-Log "Successfully renamed DCIM folder to $NewName on attempt $attempt"
                return $true
            }
            Write-Log "Rename verification failed on attempt $attempt"
        }
        catch {
            Write-Log "Error renaming DCIM folder on attempt $attempt"
            Write-Log $_.Exception.Message
        }
        if ($attempt -lt $MaxRetries) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
    Write-Log "Failed to rename DCIM folder to $NewName after $MaxRetries attempts"
    return $false
}

# Function to get SD card
function Get-SDCard {
    param($DriveLetter)
    $drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 -and $_.DeviceID -eq $DriveLetter }
    foreach ($drive in $drives) {
        if (Get-ChildItem -Path "$($drive.DeviceID)\" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq "DCIM" }) {
            Write-Log "Found DCIM folder on drive $($drive.DeviceID)"
            return $drive.DeviceID
        }
    }
    Write-Log "No DCIM folder found on drive $DriveLetter"
    return $null
}

# Function to process SD card
function Process-SDCard {
    param($DriveLetter)
    $sdDrive = Get-SDCard -DriveLetter $DriveLetter
    if (-not $sdDrive) {
        return
    }
    Write-Log "Processing SD card at $sdDrive"

    $lastBunNumber = if (Test-Path $ConfigFile) { Get-Content $ConfigFile } else { "" }

    Write-Host "Select Data Type:"
    Write-Host "1. Drone SD Card"
    Write-Host "2. 360 SD Card"
    $dataTypeChoice = Read-Host "Enter 1 or 2"
    $dataType = if ($dataTypeChoice -eq "1") { "DRONE" } else { "360" }

    do {
        Write-Host "Enter Site BUN Number (e.g., 350 0000, default: $lastBunNumber, will be saved for next run):"
        $bunNumber = Read-Host
        if ([string]::IsNullOrWhiteSpace($bunNumber)) { $bunNumber = $lastBunNumber }
    } while ($bunNumber -notmatch "^[\d\s]+$")
    $bunNumber | Out-File -FilePath $ConfigFile

    Write-Host "Enter Site Name (e.g., millicent):"
    $siteName = Read-Host
    if ([string]::IsNullOrWhiteSpace($siteName)) { $siteName = "UnknownSite" }

    $folderName = "$bunNumber $siteName"

    $currentDate = Get-Date
    $dailyFolderName = $currentDate.ToString("dddd dd MM yyyy").ToUpper()
    $dailyFolderPath = Join-Path $LocalDestination $dailyFolderName
    $isFirstRunToday = $true

    if (Test-Path $FirstRunFile) {
        $lastRunDate = Get-Content $FirstRunFile
        if ($lastRunDate -eq $currentDate.ToString("yyyy-MM-dd")) {
            $isFirstRunToday = $false
        }
    }
    $currentDate.ToString("yyyy-MM-dd") | Out-File -FilePath $FirstRunFile

    if ($isFirstRunToday -and -not (Test-Path $dailyFolderPath)) {
        New-Item -ItemType Directory -Path $dailyFolderPath -Force
        Write-Log "Created daily folder: $dailyFolderName"
    }

    try {
        $dcmiFolder = Get-ChildItem -Path "$sdDrive\" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq "DCIM" }
        if (-not $dcmiFolder) {
            Write-Log "DCIM folder not found on SD card"
            return
        }
        $dcmiPath = $dcmiFolder.FullName

        $localFolder = Join-Path $dailyFolderPath "$folderName\$dataType"
        New-Item -ItemType Directory -Path $localFolder -Force

        Write-Log "Copying DCIM folder from SD card to local drive"
        $files = Get-ChildItem -Path $dcmiPath -Recurse -File
        $totalFiles = $files.Count
        $currentFile = 0

        foreach ($file in $files) {
            $currentFile++
            $percentComplete = [math]::Round(($currentFile / $totalFiles) * 100, 2)
            Write-Progress -Activity "Copying DCIM folder from SD card" -Status "File $currentFile of $totalFiles" -PercentComplete $percentComplete

            $relativePath = $file.FullName.Substring($dcmiPath.Length)
            $destPath = Join-Path $localFolder $relativePath
            $destDir = Split-Path $destPath -Parent

            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force
            }

            Copy-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
        }
        Write-Log "Local copy completed"
        Write-Progress -Activity "Copying DCIM folder from SD card" -Completed

        $newDcmiName = "$bunNumber $siteName"
        if (-not (Rename-DCIMFolder -CurrentPath $dcmiPath -NewName $newDcmiName -DriveLetter $sdDrive)) {
            Write-Log "Aborting ejection due to rename failure"
            return
        }

        Eject-SDCard -DriveLetter $sdDrive
    }
    catch {
        Write-Log "Error processing SD card"
        Write-Log $_.Exception.Message
    }
}

# Ensure log and config directories exist
$logDir = Split-Path $LogFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force
}

# Main script
try {
    Write-Log "Script started. Polling for SD cards every 3 seconds..."

    $processedDrives = @{}
    while ($true) {
        $drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 }
        $currentDrives = $drives | Select-Object -ExpandProperty DeviceID
        foreach ($drive in $drives) {
            if (-not $processedDrives.ContainsKey($drive.DeviceID)) {
                Write-Log "Polling detected drive: $($drive.DeviceID)"
                Process-SDCard -DriveLetter $drive.DeviceID
                $processedDrives[$drive.DeviceID] = $true
            }
        }
        $keysToRemove = @($processedDrives.Keys | Where-Object { $_ -notin $currentDrives })
        foreach ($key in $keysToRemove) {
            $processedDrives.Remove($key)
        }
        Start-Sleep -Seconds 3
    }
}
catch {
    Write-Log "Script error"
    Write-Log $_.Exception.Message
}