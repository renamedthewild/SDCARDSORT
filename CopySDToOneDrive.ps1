# PowerShell script to copy Micro SD DCIM folder to OneDrive folder with prompts on SD card insertion
# Phase 3: Prefills site name from BUN list with duplicate handling, auto-detects Drone/360 SD Card, creates date-based folders for 360, uses configurable paths

# Load settings from Settings.json
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SettingsFile = Join-Path $ScriptDir "Settings.json"
try {
    if (Test-Path $SettingsFile) {
        $Settings = Get-Content $SettingsFile | ConvertFrom-Json
        # Log raw settings for debugging
        Write-Host "Raw Settings.json content:"
        Get-Content $SettingsFile | Write-Host
        # Replace %USERNAME% with user profile path
        $Settings | Get-Member -MemberType NoteProperty | ForEach-Object {
            $value = $Settings.$($_.Name)
            if ($value -is [string]) {
                $Settings.$($_.Name) = Join-Path $env:USERPROFILE ($value -replace "^%USERNAME%\\", "")
            }
        }
        $LocalDestination = $Settings.LocalDestination
        $LogFile = $Settings.LogFile
        $LastBunNumberFile = $Settings.LastBunNumberFile
        $LastRunDateFile = $Settings.LastRunDateFile
        $BunSiteMapFile = $Settings.BunSiteMapFile
        # Validate paths
        foreach ($path in @($LocalDestination, $LogFile, $LastBunNumberFile, $LastRunDateFile, $BunSiteMapFile)) {
            if ([string]::IsNullOrWhiteSpace($path)) { throw "Empty path detected in Settings.json" }
            if ($path -match "[<>|?*]") { throw "Invalid characters in path: $path" }
        }
    }
    else {
        throw "Settings.json not found at $SettingsFile"
    }
}
catch {
    Write-Host "Error loading Settings.json: $($_.Exception.Message)"
    Write-Host "Using default paths in $env:USERPROFILE\DroneSDProcessor"
    $DefaultDir = Join-Path $env:USERPROFILE "DroneSDProcessor"
    if (-not (Test-Path $DefaultDir)) {
        New-Item -ItemType Directory -Path $DefaultDir -Force | Out-Null
    }
    $LocalDestination = Join-Path $DefaultDir "PilotUploads"
    $LogFile = Join-Path $DefaultDir "CopyLog.txt"
    $LastBunNumberFile = Join-Path $DefaultDir "LastBunNumber.txt"
    $LastRunDateFile = Join-Path $DefaultDir "LastRunDate.txt"
    $BunSiteMapFile = Join-Path $DefaultDir "BunSiteMap.csv"
}

# Define DeviceEject class
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
    $logDir = Split-Path $LogFile -Parent
    try {
        if (-not (Test-Path $logDir)) {
            Write-Host "Creating log directory: $logDir"
            New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
        }
        "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
        Write-Host "$timestamp - $Message"
    }
    catch {
        Write-Host "Error writing to log file $LogFile : $($_.Exception.Message)"
    }
}

# Log initial paths for debugging
Write-Log "Script started with LogFile: $LogFile"
Write-Log "LocalDestination: $LocalDestination"
Write-Log "BunSiteMapFile: $BunSiteMapFile"

# Function to safely eject SD card with retries
function Eject-SDCard {
    param(
        $DriveLetter,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 7
    )
    $attempt = 0
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

# Function to extract BUN number from a drone filename
function Get-BunNumberFromFile {
    param($DcimPath)
    try {
        $file = Get-ChildItem -Path $DcimPath -File -Recurse | Where-Object { $_.Name -match "V_-.*-(\d+)-" } | Select-Object -First 1
        if ($file -and $file.Name -match "-(\d+)-") {
            return $Matches[1]
        }
    }
    catch {
        Write-Log "Error extracting BUN number from file: $($_.Exception.Message)"
    }
    return $null
}

# Function to normalize BUN number to XXX XXXX format
function Normalize-BunNumber {
    param($BunNumber)
    $cleanBun = $BunNumber -replace "\D", ""
    if ($cleanBun -match "^(\d{3})(\d{4})$") {
        return "$($Matches[1]) $($Matches[2])"
    }
    return $cleanBun
}

# Function to detect Drone SD Card based on filenames
function Detect-DroneSDCard {
    param($DcimPath)
    try {
        $file = Get-ChildItem -Path $DcimPath -File -Recurse | Where-Object { $_.Name -match "V_-(TVT|TOWER|HD|CH|CD|C)-" } | Select-Object -First 1
        return $null -ne $file
    }
    catch {
        Write-Log "Error detecting Drone SD Card: $($_.Exception.Message)"
        return $false
    }
}

# Function to detect 360 SD Card based on .insv files
function Detect-360SDCard {
    param($DcimPath)
    try {
        $file = Get-ChildItem -Path $DcimPath -File -Recurse | Where-Object { $_.Extension -ieq ".insv" } | Select-Object -First 1
        return $null -ne $file
    }
    catch {
        Write-Log "Error detecting 360 SD Card: $($_.Exception.Message)"
        return $false
    }
}

# Function to get unique creation dates for 360 files
function Get-360FileDates {
    param($DcimPath)
    try {
        $files = Get-ChildItem -Path $DcimPath -File -Recurse | Where-Object { $_.Extension -ieq ".insv" }
        $dates = $files | Select-Object -ExpandProperty CreationTime | ForEach-Object { $_.ToString("dd MM yyyy") } | Sort-Object -Unique
        return $dates
    }
    catch {
        Write-Log "Error getting 360 file dates: $($_.Exception.Message)"
        return @((Get-Date).ToString("dd MM yyyy"))
    }
}

# Function to get most recent creation date for 360 DCIM renaming
function Get-360MostRecentDate {
    param($DcimPath)
    try {
        $file = Get-ChildItem -Path $DcimPath -File -Recurse | Where-Object { $_.Extension -ieq ".insv" } | Sort-Object CreationTime -Descending | Select-Object -First 1
        if ($file) {
            return $file.CreationTime.ToString("dd MM yyyy")
        }
        return (Get-Date).ToString("dd MM yyyy")
    }
    catch {
        Write-Log "Error getting most recent 360 file date: $($_.Exception.Message)"
        return (Get-Date).ToString("dd MM yyyy")
    }
}

# Function to get site name from BUN number
function Get-SiteNameFromBun {
    param($BunNumber)
    try {
        if (Test-Path $BunSiteMapFile) {
            $cleanBun = $BunNumber -replace "\D", ""
            $map = Import-Csv -Path $BunSiteMapFile
            $sites = $map | Where-Object { $_.BUNNumber -eq $cleanBun } | Select-Object -ExpandProperty SiteName
            if ($sites.Count -eq 0) {
                Write-Log "No site found for BUN $cleanBun"
                return $null
            }
            elseif ($sites.Count -eq 1) {
                return $sites
            }
            else {
                Write-Host "Duplicate BUN $cleanBun found. Select a site:"
                $index = 1
                foreach ($site in $sites) {
                    Write-Host "$index. $site"
                    $index++
                }
                Write-Host "$index. Enter new site name"
                $choice = Read-Host "Enter 1-$index"
                if ($choice -eq $index) {
                    $newSite = Read-Host "Enter new Site Name (e.g., millicent)"
                    return $newSite
                }
                elseif ($choice -match "^\d+$" -and $choice -ge 1 -and $choice -lt $index) {
                    return $sites[$choice - 1]
                }
                else {
                    Write-Log "Invalid selection for duplicate BUN $cleanBun. Defaulting to manual input."
                    return $null
                }
            }
        }
        else {
            Write-Log "BUN site map file not found at $BunSiteMapFile"
        }
    }
    catch {
        Write-Log "Error reading BUN site map: $($_.Exception.Message)"
    }
    return $null
}

# Function to sort files into VT SPHERES or DT TOWER SURVEY, placing files in root
function Sort-DroneFiles {
    param(
        $RawFolder,
        $VtSpheresFolder,
        $DtTowerFolder
    )
    try {
        $files = Get-ChildItem -Path $RawFolder -File -Recurse
        foreach ($file in $files) {
            $fileName = $file.Name
            # Handle potential filename conflicts
            $destVtPath = Join-Path $VtSpheresFolder $fileName
            $destDtPath = Join-Path $DtTowerFolder $fileName
            $destRawPath = Join-Path $RawFolder $fileName
            $counter = 1
            $originalFileName = $fileName
            while (($fileName -match "V_-TVT-" -and (Test-Path $destVtPath)) -or ($fileName -match "V_-(TOWER|HD|CH|CD|C)-" -and (Test-Path $destDtPath)) -or (-not ($fileName -match "V_-(TVT|TOWER|HD|CH|CD|C)-") -and (Test-Path $destRawPath))) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($originalFileName)
                $extension = [System.IO.Path]::GetExtension($originalFileName)
                $fileName = "$baseName_$counter$extension"
                $destVtPath = Join-Path $VtSpheresFolder $fileName
                $destDtPath = Join-Path $DtTowerFolder $fileName
                $destRawPath = Join-Path $RawFolder $fileName
                $counter++
            }

            if ($file.Name -match "V_-TVT-") {
                Move-Item -Path $file.FullName -Destination $destVtPath -ErrorAction Stop
                Write-Log "Moved $($file.Name) to VT SPHERES root"
            }
            elseif ($file.Name -match "V_-(TOWER|HD|CH|CD|C)-") {
                Move-Item -Path $file.FullName -Destination $destDtPath -ErrorAction Stop
                Write-Log "Moved $($file.Name) to DT TOWER SURVEY root"
            }
            else {
                Move-Item -Path $file.FullName -Destination $destRawPath -ErrorAction Stop
                Write-Log "Moved $($file.Name) to DRONERAW root"
            }
        }
        Write-Log "File sorting completed"
    }
    catch {
        Write-Log "Error sorting files: $($_.Exception.Message)"
    }
}

# Function to process SD card
function Process-SDCard {
    param($DriveLetter)
    $sdDrive = Get-SDCard -DriveLetter $DriveLetter
    if (-not $sdDrive) {
        return
    }
    Write-Log "Processing SD card at $sdDrive"

    $lastBunNumber = if (Test-Path $LastBunNumberFile) { Get-Content $LastBunNumberFile } else { "" }
    $dcmiFolder = Get-ChildItem -Path "$sdDrive\" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq "DCIM" }
    if (-not $dcmiFolder) {
        Write-Log "DCIM folder not found on SD card"
        return
    }
    $dcmiPath = $dcmiFolder.FullName

    # Auto-detect 360 or Drone SD Card
    $is360 = Detect-360SDCard -DcimPath $dcmiPath
    $isDrone = if (-not $is360) { Detect-DroneSDCard -DcimPath $dcmiPath } else { $false }
    if ($is360) {
        Write-Host "360 SD Card detected. Press Enter to confirm or enter 1/2 for Drone/360:"
        $dataTypeChoice = Read-Host
        if ([string]::IsNullOrWhiteSpace($dataTypeChoice)) { $dataTypeChoice = "2" }
    }
    elseif ($isDrone) {
        Write-Host "Drone SD Card detected. Press Enter to confirm or enter 1/2 for Drone/360:"
        $dataTypeChoice = Read-Host
        if ([string]::IsNullOrWhiteSpace($dataTypeChoice)) { $dataTypeChoice = "1" }
    }
    else {
        Write-Host "Select Data Type:"
        Write-Host "1. Drone SD Card"
        Write-Host "2. 360 SD Card"
        $dataTypeChoice = Read-Host "Enter 1 or 2"
    }
    $dataType = if ($dataTypeChoice -eq "1") { "DRONE" } else { "360" }

    try {
        if (-not (Test-Path $dcmiPath)) {
            Write-Log "DCIM folder not found on SD card"
            return
        }

        $currentDate = Get-Date
        $dailyFolderName = $currentDate.ToString("dddd dd MM yyyy").ToUpper()
        $dailyFolderPath = Join-Path $LocalDestination $dailyFolderName
        $isFirstRunToday = $true

        if (Test-Path $LastRunDateFile) {
            $lastRunDate = Get-Content $LastRunDateFile
            if ($lastRunDate -eq $currentDate.ToString("yyyy-MM-dd")) {
                $isFirstRunToday = $false
            }
        }
        $currentDate.ToString("yyyy-MM-dd") | Out-File -FilePath $LastRunDateFile -Force

        if ($isFirstRunToday -and -not (Test-Path $dailyFolderPath)) {
            New-Item -ItemType Directory -Path $dailyFolderPath -Force | Out-Null
            Write-Log "Created daily folder: $dailyFolderName"
        }

        if ($dataType -eq "DRONE") {
            # Drone SD Card processing
            $suggestedBun = Get-BunNumberFromFile -DcimPath $dcmiPath
            if ($suggestedBun) {
                Write-Host "Detected BUN Number: $suggestedBun. Press Enter to confirm or type a new BUN Number (e.g., 350 0000, default: $lastBunNumber):"
                $bunNumber = Read-Host
                if ([string]::IsNullOrWhiteSpace($bunNumber)) { $bunNumber = $suggestedBun }
            }
            else {
                Write-Host "No BUN Number detected. Enter Site BUN Number (e.g., 350 0000, default: $lastBunNumber):"
                $bunNumber = Read-Host
                if ([string]::IsNullOrWhiteSpace($bunNumber)) { $bunNumber = $lastBunNumber }
            }
            $formattedBun = Normalize-BunNumber -BunNumber $bunNumber
            $bunNumber | Out-File -FilePath $LastBunNumberFile -Force

            # Prefill site name from BUN number
            $suggestedSite = Get-SiteNameFromBun -BunNumber $bunNumber
            if ($suggestedSite) {
                Write-Host "Detected Site Name: $suggestedSite. Press Enter to confirm or type a new Site Name (e.g., millicent):"
                $siteName = Read-Host
                if ([string]::IsNullOrWhiteSpace($siteName)) { $siteName = $suggestedSite }
            }
            else {
                Write-Host "Enter Site Name (e.g., millicent):"
                $siteName = Read-Host
                if ([string]::IsNullOrWhiteSpace($siteName)) { $siteName = "UnknownSite" }
            }

            $folderName = "$formattedBun $siteName"
            $siteFolder = Join-Path $dailyFolderPath $folderName
            $rawFolder = Join-Path $siteFolder "DRONERAW"
            $vtSpheresFolder = Join-Path $siteFolder "VT SPHERES"
            $dtTowerFolder = Join-Path $siteFolder "DT TOWER SURVEY"
            $cdtInstaFolder = Join-Path $siteFolder "CDT INSTA VIDEO"
            $cdtShelterFolder = Join-Path $siteFolder "CDT SHELTER (SCOPE 1 & 2 ONLY)"
            $iphoneCivilFolder = Join-Path $siteFolder "IPHONE CIVIL"
            New-Item -ItemType Directory -Path $rawFolder, $vtSpheresFolder, $dtTowerFolder, $cdtInstaFolder, $cdtShelterFolder, $iphoneCivilFolder -Force | Out-Null

            Write-Log "Copying DCIM folder from Drone SD card to $rawFolder"
            $files = Get-ChildItem -Path $dcmiPath -Recurse -File
            $totalFiles = $files.Count
            $currentFile = 0

            foreach ($file in $files) {
                $currentFile++
                $percentComplete = [math]::Round(($currentFile / $totalFiles) * 100, 2)
                Write-Progress -Activity "Copying DCIM folder from Drone SD card" -Status "File $currentFile of $totalFiles" -PercentComplete $percentComplete

                $relativePath = $file.FullName.Substring($dcmiPath.Length)
                $destPath = Join-Path $rawFolder $relativePath
                $destDir = Split-Path $destPath -Parent

                if (-not (Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }

                Copy-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
            }
            Write-Progress -Activity "Copying DCIM folder from Drone SD card" -Completed
            Write-Log "Drone SD card copy completed"

            $newDcmiName = "$formattedBun $siteName"
            if (-not (Rename-DCIMFolder -CurrentPath $dcmiPath -NewName $newDcmiName -DriveLetter $sdDrive)) {
                Write-Log "Aborting ejection due to rename failure"
                return
            }

            Eject-SDCard -DriveLetter $sdDrive

            # Sort drone files after SD card is ejected
            Sort-DroneFiles -RawFolder $rawFolder -VtSpheresFolder $vtSpheresFolder -DtTowerFolder $dtTowerFolder
        }
        else {
            # 360 SD Card processing
            $fileDates = Get-360FileDates -DcimPath $dcmiPath
            $mostRecentDate = Get-360MostRecentDate -DcimPath $dcmiPath

            Write-Log "Copying DCIM folder from 360 SD card to date-based folders"
            $files = Get-ChildItem -Path $dcmiPath -Recurse -File
            $totalFiles = $files.Count
            $currentFile = 0

            foreach ($file in $files) {
                $currentFile++
                $percentComplete = [math]::Round(($currentFile / $totalFiles) * 100, 2)
                Write-Progress -Activity "Copying DCIM folder from 360 SD card" -Status "File $currentFile of $totalFiles" -PercentComplete $percentComplete

                $fileDate = $file.CreationTime.ToString("dd MM yyyy")
                $dateFolderName = "$fileDate 360"
                $dateFolder = Join-Path $dailyFolderPath $dateFolderName
                if (-not (Test-Path $dateFolder)) {
                    New-Item -ItemType Directory -Path $dateFolder -Force | Out-Null
                    Write-Log "Created 360 folder: $dateFolderName"
                }

                $relativePath = $file.FullName.Substring($dcmiPath.Length)
                $destPath = Join-Path $dateFolder $relativePath
                $destDir = Split-Path $destPath -Parent
                if (-not (Test-Path $destDir)) {
                    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                }

                # Handle filename conflicts
                $destFileName = [System.IO.Path]::GetFileName($destPath)
                $counter = 1
                while (Test-Path $destPath) {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($destFileName)
                    $extension = [System.IO.Path]::GetExtension($destFileName)
                    $destFileName = "$baseName_$counter$extension"
                    $destPath = Join-Path $destDir $destFileName
                    $counter++
                }

                Copy-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
            }
            Write-Progress -Activity "Copying DCIM folder from 360 SD card" -Completed
            Write-Log "360 SD card copy completed"

            $newDcmiName = "$mostRecentDate 360"
            if (-not (Rename-DCIMFolder -CurrentPath $dcmiPath -NewName $newDcmiName -DriveLetter $sdDrive)) {
                Write-Log "Aborting ejection due to rename failure"
                return
            }

            Eject-SDCard -DriveLetter $sdDrive
        }
    }
    catch {
        Write-Log "Error processing SD card: $($_.Exception.Message)"
    }
}

# Ensure log and config directories exist
$logDir = Split-Path $LogFile -Parent
try {
    if (-not (Test-Path $logDir)) {
        Write-Log "Creating log directory: $logDir"
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Host "Error creating log directory $logDir : $($_.Exception.Message)"
    exit
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