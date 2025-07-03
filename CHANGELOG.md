
# Changelog for CopySDToOneDrive.ps1

## Version 1.0.0 (2025-07-03)

### Overview
The `CopySDToOneDrive.ps1` script automates the copying of media files from Micro SD cards to a OneDrive folder structure for drone and 360 camera data. It supports automatic detection of Drone and 360 SD cards, processes files based on user-selected options, and ensures robust error handling and logging. Developed as part of **Drone Phase 3**, the script has been refined through multiple iterations to address user feedback and ensure reliability.

### Initial Features (Phase 3 Baseline)
- **Configuration**: Loads paths from `Settings.json` for `LocalDestination`, `LogFile`, `LastBunNumberFile`, `LastRunDateFile`, and `BunSiteMapFile`.
- **SD Card Detection**: Auto-detects Drone SD cards (based on filenames like `V_-(TVT|TOWER|HD|CH|CD|C)-`) and 360 SD cards (based on `.insv` files).
- **360 Basic Logic**: Copies `.insv` files to `<DAY DD MM YYYY>\<DD MM YYYY> 360`, renames `DCIM` to `<Date> 360_X` on conflict.
- **360 Smart Logic**: Prompts for options (1. Use last Drone Site, 2. Lookup site, 3. 360 DATA Dump), creates folder structure (`DRONERAW`, `VT SPHERES`, `DT TOWER SURVEY`, `CDT INSTA VIDEO`, `CDT SHELTER (SCOPE 1 & 2 ONLY)`, `IPHONE CIVIL`), and copies `.insv`/`.lrv` files.
- **Drone Logic**: Copies files to `DRONERAW`, sorts to `VT SPHERES` or `DT TOWER SURVEY` based on filename patterns.
- **Logging**: Writes detailed logs to `CopyLog.txt` for debugging and tracking.
- **Ejection**: Safely ejects SD cards with retries using a C# `DeviceEject` class.
- **BUN/Site Handling**: Uses `BunSiteMap.csv` for site lookup, supports duplicate BUN resolution, and saves last BUN for reuse.

### Issues and Fixes

#### Issue 1: Double Prompt for Option 2 (Site Lookup)
- **Problem**: When selecting option 2 in 360 smart logic, the script prompted twice for “1. Use last Drone Site, 2. Lookup site, 3. 360 DATA Dump” (logged at 2025-07-03 22:14:01 and 22:14:09).
- **Fix**: Removed redundant prompt in `Process-360Smart`, passing `$DataTypeChoice` directly from `Process-SDCard` (artifact version ID `aa83cad4-ef61-4351-821d-c2f1e5d69785`).

#### Issue 2: Unexpected Merge Prompt
- **Problem**: The script prompted to merge files into existing `CDT INSTA VIDEO` or `CDT SHELTER (SCOPE 1 & 2 ONLY)` folders even when newly created (logged at 2025-07-03 22:17:14).
- **Fix**: Moved folder existence checks (`$mergeInsta`, `$mergeShelter`) before folder creation, prompting only if folders pre-exist with files (artifact version ID `aa83cad4-ef61-4351-821d-c2f1e5d69785`).

#### Issue 3: No Folder Selection for `.insv` Files
- **Problem**: The script did not prompt for folder selection (`CDT INSTA VIDEO` or `CDT SHELTER`) for each `.insv` file, defaulting to `CDT INSTA VIDEO` when merging (logged at 2025-07-03 22:17:14).
- **Fix**: Added per-file prompts for folder selection unless merging into an existing folder with files (artifact version ID `aa83cad4-ef61-4351-821d-c2f1e5d69785`).

#### Issue 4: Syntax Errors in `Process-360Smart`
- **Problem**: Parsing errors due to missing braces, unclosed strings, and a typo (`Lilliputian`) in `Process-360Smart` (logged errors at lines 404, 471, 527, 531, 582, 583).
- **Fix**: Corrected brace matching, completed `try`/`catch` blocks, fixed string termination, and removed typo (artifact version ID `1351f41e-21ff-42e9-b422-763275a76875`).

#### Issue 5: Sequential Folder Selection Prompts
- **Problem**: Folder selection for `.insv` files was prompted per file before copying, requiring the technician to stay at the computer.
- **Fix**: Modified `Process-360Smart` to collect all folder selections upfront using a `$fileDestinations` hashtable, then copy files in one loop (artifact version ID `85584c06-a3d9-4e63-899e-2cb3ff258ac0`).

#### Issue 6: Missing Site Name in Option 1 Prompt
- **Problem**: Option 1 prompt showed only the BUN (e.g., “Using last Drone BUN: 3500009”) without the site name (logged at 2025-07-03 22:55:46).
- **Fix**: Added `Get-SiteNameFromBun` call to display both BUN and site name (e.g., “Using last Drone BUN and site: 350 0009 Millicent”) in `Process-360Smart` (artifact version ID `89d04cfa-4e09-4984-8ef8-8cf4c235d4bb`).

#### Issue 7: Ampersand Parsing Errors
- **Problem**: Unescaped `&` in `"CDT SHELTER (SCOPE 1 & 2 ONLY)"` caused parsing errors at lines 581, 617, 645, and 946.
- **Fix**: Introduced `$shelterFolderName = 'CDT SHELTER (SCOPE 1 "&" 2 ONLY)'` and used it consistently in `Process-360Smart` and `Process-SDCard` for folder paths and prompts (artifact version ID `f03aebc5-8fdf-4a7c-bad3-d236b2ded980`).

### Final Features
- **360 Smart Logic**:
  - Prompts for options (1. Use last Drone Site, 2. Lookup site, 3. 360 DATA Dump).
  - Option 1 displays last BUN and site name (e.g., “350 0009 Millicent”).
  - Option 2 supports site lookup with `BunSiteMap.csv` (e.g., search “mil” for Millicent).
  - Collects `.insv` folder selections (`CDT INSTA VIDEO` or `CDT SHELTER (SCOPE 1 & 2 ONLY)`) upfront, allowing technicians to answer all prompts and leave.
  - Merges files into existing folders only if they contain files, with user confirmation.
  - Creates folder structure: `<DAY DD MM YYYY>\<BUN SiteName>\{DRONERAW, VT SPHERES, DT TOWER SURVEY, CDT INSTA VIDEO, CDT SHELTER (SCOPE 1 & 2 ONLY), IPHONE CIVIL}`.
  - Renames `DCIM` to `<BUN SiteName> 360` or `<BUN SiteName> 360_X` on conflict.
- **360 Basic Logic**: Copies `.insv` files to `<DAY DD MM YYYY>\<DD MM YYYY> 360`, renames `DCIM` to `<Date> 360_X`.
- **Drone Logic**: Copies files to `DRONERAW`, sorts to `VT SPHERES` or `DT TOWER SURVEY` based on filename patterns.
- **Robustness**: Handles errors with `try`/`catch`, logs to `CopyLog.txt`, supports UTF-8 without BOM, and safely ejects SD cards.
- **Usability**: Streamlined prompts for technician efficiency, with clear BUN/site display and upfront file selection.

### Notes
- **Tested**: Successfully tested on 2025-07-03 for 360 smart logic (options 1, 2, 3) and Drone logic, with confirmed functionality for BUN/site prompts, folder selection, merging, and SD card ejection.
- **Dependencies**: Requires `Settings.json` and `BunSiteMap.csv` in `C:\Users\<Username>\RURAL IT SOLUTIONS\DRONE - Documents\R2F\`.
- **Repository**: Ready for push to `https://github.com/renamedthewild/SDCARDSORT.git`.

### Authors
- Luke Wilden (Project Lead)
- Grok 3 (xAI) (Script Development and Debugging)

### Last Updated
- 2025-07-03 11:25 PM AEST
