# FileOrganizer

Modern PowerShell file organizer with GUI and CLI modes. Move or copy files by category with dry-run preview, duplicate handling, and full configurability.

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://learn.microsoft.com/en-us/powershell/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)](https://www.microsoft.com/windows)

## Features

- **GUI Mode** — Out-GridView multi-select, confirmation dialogs, progress bar
- **CLI Mode** — Scriptable command-line automation
- **Move or Copy** — Choose action per operation
- **Dry-Run Preview** — See what will happen before committing
- **Automatic Duplicate Renaming** — Appends `_1`, `_2` etc. to avoid overwrites
- **Configurable Categories** — Custom file type mapping via `config.json`
- **Separate Type Folders** — Option to organize into per-type subfolders
- **Comprehensive Logging** — Timestamped logs for every operation
- **Remembers Last Used** — Saves source and destination paths

## Quick Start

```powershell
# GUI mode (recommended)
.\FileOrganizer.ps1

# CLI mode - organize Downloads
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized"
```

## CLI Examples

```powershell
# Copy instead of move
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -Action Copy

# Dry-run first to preview
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -DryRun

# Separate folders per file type
.\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -SeparateTypes
```

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+

## Configuration

Edit `config.json` to customize file categories and default destinations:

```json
{
  "categories": {
    "Images": ["jpg", "png", "gif", "bmp", "svg"],
    "Documents": ["pdf", "doc", "docx", "txt", "md"],
    "Videos": ["mp4", "avi", "mkv", "mov"],
    "Code": ["ps1", "py", "js", "ts", "cpp", "cs"]
  }
}
```

## Project Structure

```
├── FileOrganizer.ps1        # Main script (GUI + CLI)
├── config.json              # Categories and settings
├── lastused.json            # Saved paths (auto-created)
├── logs/                    # Operation logs (auto-created)
├── LICENSE
└── README.md
```

## License

MIT
