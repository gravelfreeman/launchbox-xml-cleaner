# LaunchBox XML Cleaner

<img width="1251" height="244" alt="{40ACCBC6-210E-418A-9D6C-5E27B0B647A7}" src="https://github.com/user-attachments/assets/2bda972b-f0d2-4131-a853-53ea51c123ad" />  

  
<img width="1100" height="377" alt="{BF0708CE-CE90-4F95-A5DA-26260FA5FF22}" src="https://github.com/user-attachments/assets/7b342d97-2d48-4fdf-8eef-c596379ecce0" />  

***
  
This script cleans a LaunchBox platform XML so it only keeps entries for games that actually exist in a local roms folder.

It is especially useful when a MAME platform XML was created from `MAME Arcade Full Set...` import but your local ROM folder only contains part of that set, and you want the XML reduced to the games you actually have while still preserving additional applications (clones) from merged romsets.

LaunchBox's `Scan For Removed MAME Roms...` feature removes clone entries from merged sets because those clone files are stored inside the parent `rom.zip`.

Although it was designed around this MAME workflow, it can also be used with other LaunchBox platform XML files.

## What It Does

The cleaner scans the current ROM folder and compares it against the provided LaunchBox XML.

It keeps a game when a local ROM exists for that game, including:
- the main `Game` `ApplicationPath`
- any linked `AdditionalApplication` entries for the same `GameID`

After the main game pass, it removes orphaned metadata entries that no longer belong to a kept game:
- `<AdditionalApplication>`
- `<GameControllerSupport>`
- `<AlternateName>`

## How to Use

1. Paste a `<Platform>.xml` file in it's corresponding roms folder
2. Paste `_LB_XML_Cleaner.ps1` and `_LB_XML_Drop.cmd` files in the roms folder
3. Drag and drop the `<Platform>.xml` file onto the `_LB_XML_Drop.cmd` launcher

### Command line

If you prefer using command line, run the PowerShell script directly and point it to the XML file and ROM folder you want to validate.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -PreviewOnly -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -OutputPath "C:\Path\Cleaned.xml" -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\_LB_XML_Cleaner.ps1" -InPlace -XmlPath "C:\Path\Platform.xml" -RomsPath "C:\Path\Roms"
```

### Flags
- `-XmlPath`: LaunchBox XML file to clean. Can also be passed as the first positional argument.
- `-RomsPath`: Folder containing the ROMs or CHDs to scan.
- `-PreviewOnly`: Run the cleanup without writing any file.
- `-OutputPath`: Write to a specific output path instead of replacing the source XML.
- `-InPlace`: Replace the source XML directly. If `-OutputPath` is not provided, the default behavior already performs a safe backup-and-replace workflow.
- `-RomExtensions`: File extensions treated as ROMs. Default: `.zip`, `.7z`, `.chd`.

The script never modifies or deletes files or folders.  
It only reads the folder content and writes a new XML based on local roms.
