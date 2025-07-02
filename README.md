
# Drone SD Card Processor

A PowerShell script to automate copying, sorting, and organizing files from Drone and 360 camera SD cards, with configurable paths and site name prefilling for Drone files.

## Features
- **Detects SD Cards**: Polls for SD card insertion every 3 seconds.
- **Auto-Detects Drone/360 SD Cards**: Identifies Drone files (e.g., `V_-TVT-`) or 360 files (`.insv`) and defaults to the appropriate mode.
- **Drone Workflow**:
  - Prefills BUN number and site name from filenames and `BunSiteMap.csv`.
  - Copies `DCIM` to `DRONERAW`, sorts files to `VT SPHERES`, `DT TOWER SURVEY`, or `DRONERAW`.
  - Creates folders: `DRONERAW`, `VT SPHERES`, `DT TOWER SURVEY`, `CDT INSTA VIDEO`, `CDT SHELTER (SCOPE 1 & 2 ONLY)`, `IPHONE CIVIL`.
  - Renames `DCIM` to `<BUN SiteName>` (e.g., `300 0409 test2`).
- **360 Workflow**:
  - Copies `.insv` files to date-based folders (e.g., `MONDAY 23 06 2025\23 06 2025 360`) based on file creation dates.
  - Merges files from multiple SD cards with the same creation date.
  - Renames `DCIM` to `<Most Recent Date> 360` (e.g., `23 06 2025 360`).
- **OneDrive Sync**: Copies files to a configurable output folder for SharePoint syncing.
- **Logging**: Records actions to a configurable log file.

## Prerequisites
- Windows 10/11 with PowerShell 5.1 or later.
- OneDrive configured for file syncing.
- SD card reader with Drone or 360 camera files.

## Setup
1. **Clone or Download**:
   ```bash
   git clone https://github.com/your-username/drone-sd-processor.git
   ```
   Or download the ZIP and extract.

2. **Create Settings.json**:
   - Copy `Settings.json.template` to `Settings.json` in the script directory.
   - Edit paths, replacing `%USERNAME%` with your user profile path (e.g., `C:\Users\YourUser`):
     ```json
     {
       "LocalDestination": "YourUser\\RURAL IT SOLUTIONS\\DRONE - Documents\\R2F\\PilotUploads",
       "LogFile": "YourUser\\RURAL IT SOLUTIONS\\DRONE - Documents\\R2F\\CopyLog.txt",
       "LastBunNumberFile": "YourUser\\RURAL IT SOLUTIONS\\DRONE - Documents\\R2F\\LastBunNumber.txt",
       "LastRunDateFile": "YourUser\\RURAL IT SOLUTIONS\\DRONE - Documents\\R2F\\LastRunDate.txt",
       "BunSiteMapFile": "YourUser\\RURAL IT SOLUTIONS\\DRONE - Documents\\R2F\\BunSiteMap.csv"
     }
     ```

3. **Create BunSiteMap.csv**:
   - Copy `BunSiteMap.csv.template` to the path specified in `Settings.json`.
   - Add BUN number to site name mappings (BUNs without spaces):
     ```csv
     BUNNumber,SiteName
     3400502,SiteA
     3500995,SiteB
     ```

4. **Run the Script**:
   - Open PowerShell, navigate to the script directory:
     ```powershell
     cd path\to\drone-sd-processor
     ```
   - Run:
     ```powershell
     .\CopySDToOneDrive.ps1
     ```
   - Insert an SD card and follow prompts.

## Usage
- **Drone SD Card**:
  - Prompts: Data Type (auto-detects Drone), BUN Number (auto-filled), Site Name (auto-filled from `BunSiteMap.csv`).
  - Output: `LocalDestination\<DAY DD MM YYYY>\<BUN SiteName>\DRONERAW`, `VT SPHERES`, etc.
- **360 SD Card**:
  - Prompts: Data Type (auto-detects 360).
  - Output: `LocalDestination\<DAY DD MM YYYY>\<DD MM YYYY> 360` for each file creation date.
- **Sync**: Files sync to SharePoint via OneDrive.

## Troubleshooting
- **Settings.json Missing**: Create from template and verify paths.
- **BunSiteMap.csv Errors**: Ensure correct CSV format and numeric BUNs.
- **Ejection Issues**: Close File Explorer or pause/resume OneDrive:
  ```powershell
  & "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" /pause
  & "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" /resume
  ```
- Check `CopyLog.txt` for errors.

## Contributing
- Fork the repository, make changes, and submit a pull request.
- Report issues or suggest features via GitHub Issues.

## License
[MIT License](LICENSE)
