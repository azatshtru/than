# Backup / Phone Folder Diff Tool

A simple, fast GUI tool to compare two sources (local backup folders and/or a connected phone over MTP) and see which files and folders are the same, different, or only present on one side.

## Files

- `backup_diff_gui.py` — main GUI application
- `mtp_helper.ps1` — PowerShell helper used for phone (MTP) access/enumeration on Windows
- `run_diff_tool.bat` — Windows launcher

## Requirements

- Windows
- Python 3 installed (uses built-in `tkinter`, no extra packages required)
- PowerShell (for phone/MTP support, included with Windows)

## Run

### Option 1: Double-click
- Double-click `run_diff_tool.bat`

### Option 2: Command line
From this folder:

```powershell
python backup_diff_gui.py
```

## How to use

1. For **Side A** and **Side B**, choose a source type:
   - **Local folder**: browse to a folder (e.g. `OnePlus Nord CE Backup 20260603`).
   - **Phone (MTP)**: click **Refresh devices**, select your phone and storage (e.g. `Internal shared storage`), and optionally enter a subfolder (e.g. `DCIM`, `Pictures`, or `Documents\WhatsApp`).
2. Click **Compare**.
3. Browse results in the table:
   - **Kind** shows whether the item is a **File** or **Folder**.
   - **Same** / **Different** / **Only in A / Only in B**
   - Folders are marked **Different** if any compared contents under them differ or are missing on one side.
4. Use **Status**, **Kind** (Files/Folders), and **Search** filters to narrow results.
5. Double-click a row (or click **View Diff / Details**) for details:
   - **Folders**: summary of matching/missing/different items under that folder.
   - **Files**: for local files, shows a text diff (or hashes/info for binary/large files). For phone/MTP files, shows metadata/details (no content copying).

## Comparison behavior (fast mode)

- **No copying from phone/MTP:** the tool indexes phone contents via the MTP/Shell API and does not copy phone files to a temp folder for hashing/diffing.
- **Local files (same path):** compared by size and SHA-256 content hash.
- **Phone/MTP files (same path):** compared by metadata available without copying (name/path and size). If path and size match, they are treated as **Same** in fast mode (content not byte-verified). Details will note this.
- **Same base name, different extension (or missing extension on one side):** if two files are in the same folder and have the same base name but different extensions (e.g. `IMG-20260706-WA0000` vs `IMG-20260706-WA0000.jpg`), the tool checks their contents; if the contents are the same (SHA-256), they are deemed **Same** and shown as `pathA <-> pathB`. This content check is performed when both files are locally accessible (fast, no phone copying). If content cannot be verified without copying, they remain unpaired.
- **Phone/MTP naming:** the tool reads actual MTP file-name metadata (with fallbacks) to better handle cases where display names hide extensions or dotfiles (e.g. `.nomedia`).

## Phone (MTP) notes

- Your phone must be connected and unlocked, with file transfer (MTP) enabled so it appears in File Explorer/This PC.
- If your phone or storage does not appear, make sure it is unlocked and in file-transfer mode, then click **Refresh devices**.
- MTP enumeration can be slower than local folders, but comparison avoids copying file contents from the phone for speed. If a device/provider does not expose certain files over MTP, they cannot be compared.
