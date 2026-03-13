# FileOrganizer

**Modern PowerShell File Organizer** – GUI + CLI – Move or Copy files by category with dry-run, logging, and full configurability.

## Features
- Beautiful GUI (Out-GridView multi-select, confirmation dialogs, progress bar)
- Fully configurable categories via `config.json`
- Move **or** Copy
- Safe dry-run preview
- Automatic duplicate renaming (`file_1.ext`)
- ESC key to cancel
- Remembers last used folders
- Full timestamped logging
- Works on PowerShell 5.1 and 7+
- CLI support for scripts and automation

## Quick Start (GUI – recommended)
1. Download the repository
2. Run `FileOrganizer.ps1` (double-click or right-click → Run with PowerShell)
3. Follow the friendly prompts

## CLI Examples
```powershell
# Move to single folder with type subfolders
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized"

# Copy + dry-run first
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -Action Copy -DryRun

# Separate folder per type
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -SeparateTypes
