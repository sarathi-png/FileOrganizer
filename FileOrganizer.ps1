# FileOrganizer.ps1 - Modern File Organizer with GUI & CLI
# A professional-grade Windows utility for organizing files by type
# 
# Default config.json (auto-created if missing):
# <#
# {
#   "categories": {
#     "Word Files": ["doc", "docx", "docm", "dotx"],
#     "PowerPoint Files": ["ppt", "pptx", "pptm"],
#     "Excel Files": ["xls", "xlsx", "xlsm", "csv"],
#     "PDF Files": ["pdf"],
#     "Images": ["jpg", "jpeg", "png", "gif", "bmp", "svg", "webp", "tiff"],
#     "Videos": ["mp4", "avi", "mkv", "mov", "wmv", "flv"],
#     "Executables": ["exe", "msi", "bat", "cmd", "ps1"],
#     "Archives": ["zip", "rar", "7z", "tar", "gz"],
#     "Audio": ["mp3", "wav", "flac", "ogg", "aac", "wma"],
#     "Code": ["ps1", "py", "js", "ts", "cpp", "c", "h", "cs", "java", "rb", "go", "rs"],
#     "Text": ["txt", "md", "log", "rtf", "json", "xml", "yaml", "yml", "ini", "cfg"]
#   },
#   "excludePatterns": ["Thumbs.db", ".DS_Store", "desktop.ini", "*.tmp", "~$*"],
#   "createTypeSubfolder": true,
#   "defaultAction": "Move",
#   "defaultDepth": 100
# }
# #>
# 
# Usage:
#   .\FileOrganizer.ps1                                    # Interactive GUI mode
#   .\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -DryRun
#   .\FileOrganizer.ps1 -SourcePath "C:\Src" -DestPath "D:\Dest" -Action Copy -Depth 2
#   .\FileOrganizer.ps1 -CategoriesFile "custom.json"
#   Get-Help .\FileOrganizer.ps1 -Detailed

#Requires -Version 5.1

# ============================================================
# IMPORTANT: param() MUST come first before any executable code
# ============================================================

[CmdletBinding()]
param (
    [Parameter(HelpMessage="Source folder path to scan")]
    [string]$SourcePath,
    
    [Parameter(HelpMessage="Destination folder path")]
    [string]$DestPath,
    
    [Parameter(HelpMessage="Scan depth: 0=Main only, 1=Subfolders, 2=SubSubfolders, 100=All")]
    [ValidateRange(0, 100)]
    [int]$Depth = 100,
    
    [Parameter(HelpMessage="Action to perform: Move or Copy")]
    [ValidateSet("Move", "Copy")]
    [string]$Action = "Move",
    
    [Parameter(HelpMessage="Simulate operation without moving/copying files")]
    [switch]$DryRun,
    
    [Parameter(HelpMessage="Path to custom categories JSON file")]
    [string]$CategoriesFile,
    
    [Parameter(HelpMessage="Run without looping for additional tasks")]
    [switch]$NoLoop,
    
    [Parameter(HelpMessage="Use separate destination folder per file type")]
    [switch]$SeparateDestinations,
    
    [Parameter(HelpMessage="Show this help message")]
    [switch]$Help
)

# ============================================================
# Now we can have executable code after param()
# ============================================================

# Early error handling - don't let script silently fail
$ErrorActionPreference = "Continue"
$Script:StartupError = $null

# Check execution policy
try {
    $execPolicy = Get-ExecutionPolicy -ErrorAction SilentlyContinue
    if ($execPolicy -eq "Restricted") {
        Write-Host "WARNING: Execution Policy is Restricted." -ForegroundColor Yellow
        Write-Host "Run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    }
} catch { }

# Fix for $PSScriptRoot being empty when run via right-click "Run with PowerShell"
$Script:ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Script:ScriptPath) { $Script:ScriptPath = Get-Location }

# Pre-load Windows Forms (needed for GUI dialogs)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch {
    $errMsg = $_.Exception.Message
    $Script:StartupError = "Failed to load Windows Forms - $errMsg"
}

#region Global Variables
$Script:Config = $null
$Script:LastUsed = $null
$Script:ScriptName = "FileOrganizer"
$Script:LogFile = $null
$Script:IsElevated = $false
$Script:PSVersion = $PSVersionTable.PSVersion.Major
#endregion

# Check for startup errors
if ($Script:StartupError) {
    Write-Host "ERROR: $Script:StartupError" -ForegroundColor Red
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
    exit 1
}

# Show startup banner
Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "   FileOrganizer v2.0 - Loading..." -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Script Path: $Script:ScriptPath" -ForegroundColor Gray
Write-Host "PowerShell: $($Script:PSVersion)" -ForegroundColor Gray
Write-Host ""

#region Comment-Based Help
<#
.SYNOPSIS
    Organizes files by type into destination folders.

.DESCRIPTION
    FileOrganizer is a professional-grade Windows utility that scans a source folder,
    categorizes files by type (Word, PDF, Images, etc.), and moves or copies them to
    destination folders. Supports both CLI and interactive GUI modes.

    Features:
    - CLI and Interactive GUI hybrid operation
    - Configurable file categories via JSON
    - Move or Copy operations
    - Dry-Run mode for safe testing
    - Modern Out-GridView selection
    - Remembers last used folders
    - Comprehensive logging

.PARAMETER SourcePath
    Source folder to scan. If not provided, launches interactive folder browser.

.PARAMETER DestPath
    Destination folder for organized files. If not provided, launches interactive browser.

.PARAMETER Depth
    Scan depth: 0=Main folder only, 1=One level deep, 2=Two levels, 100=All (default).

.PARAMETER Action
    Operation type: "Move" (default) or "Copy". Use Copy for safer testing.

.PARAMETER DryRun
    Simulates the operation without actually moving/copying files. Shows what would happen.

.PARAMETER CategoriesFile
    Path to custom JSON file with file type categories.

.PARAMETER NoLoop
    Runs single operation without prompting for additional tasks.

.PARAMETER SeparateDestinations
    Use separate destination folder per file type (prompts for each).

.EXAMPLE
    .\FileOrganizer.ps1
    
    Launches interactive GUI mode with folder browsers and selections.

.EXAMPLE
    .\FileOrganizer.ps1 -SourcePath "C:\Downloads" -DestPath "D:\Organized" -DryRun
    
    Scans Downloads, shows what would be organized to D:\Organized without making changes.

.EXAMPLE
    .\FileOrganizer.ps1 -SourcePath "C:\Files" -Action Copy -Depth 2
    
    Copies files from C:\Files (including subfolders 2 levels deep) with interactive destination.

.NOTES
    Author: FileOrganizer Team
    Version: 2.0.0
    Requires: PowerShell 5.1+
#>
#endregion

#region Initialization Functions

function Initialize-Environment {
    <#
    .SYNOPSIS
        Initializes the script environment - logging, config, paths
    #>
    [CmdletBinding()]
    param()
    
    # Set up log file path
    $logDate = Get-Date -Format "yyyy-MM-dd"
    $Script:LogFile = Join-Path $Script:ScriptPath "$($Script:ScriptName)_$logDate.log"
    
    # Create script path if it doesn't exist (for when running from other locations)
    if (-not $Script:ScriptPath) {
        $Script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    
    # Load configuration
    Load-Config -CategoriesFile $CategoriesFile
    
    # Load last used paths
    Load-LastUsed
    
    # Check PowerShell version for feature compatibility
    if ($Script:PSVersion -lt 5) {
        Write-Warning "PowerShell version 5.1 or higher recommended. Current: $($Script:PSVersion)"
    }
    
    Write-Log "INFO" "FileOrganizer started - PS Version: $($Script:PSVersion)"
}

function Load-Config {
    <#
    .SYNOPSIS
        Loads configuration from JSON file or creates default
    #>
    [CmdletBinding()]
    param(
        [string]$CategoriesFile
    )
    
    # Determine config file path
    if ($CategoriesFile) {
        $configPath = $CategoriesFile
    } else {
        $configPath = Join-Path $Script:ScriptPath "config.json"
    }
    
    # Default configuration
    $defaultConfig = @{
        categories = @{
            "Word Files" = @("doc", "docx", "docm", "dotx")
            "PowerPoint Files" = @("ppt", "pptx", "pptm")
            "Excel Files" = @("xls", "xlsx", "xlsm", "csv")
            "PDF Files" = @("pdf")
            "Images" = @("jpg", "jpeg", "png", "gif", "bmp", "svg", "webp", "tiff")
            "Videos" = @("mp4", "avi", "mkv", "mov", "wmv", "flv")
            "Executables" = @("exe", "msi", "bat", "cmd")
            "Archives" = @("zip", "rar", "7z", "tar", "gz")
            "Audio" = @("mp3", "wav", "flac", "ogg", "aac", "wma")
            "Code" = @("ps1", "py", "js", "ts", "cpp", "c", "h", "cs", "java", "rb", "go", "rs")
            "Text" = @("txt", "md", "log", "rtf", "json", "xml", "yaml", "yml", "ini", "cfg")
        }
        excludePatterns = @("Thumbs.db", ".DS_Store", "desktop.ini", "*.tmp", "~$*")
        createTypeSubfolder = $true
        defaultAction = "Move"
        defaultDepth = 100
    }
    
    # Load existing or create default config
    if (Test-Path $configPath) {
        try {
            $configContent = Get-Content $configPath -Raw -ErrorAction Stop
            # Handle PS 5.1 vs PS 6+ compatibility
            if ($Script:PSVersion -ge 6) {
                $Script:Config = $configContent | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            } else {
                $loadedConfig = $configContent | ConvertFrom-Json -ErrorAction Stop
                # Convert PSObject to Hashtable for PS 5.1 compatibility
                $Script:Config = @{}
                $loadedConfig.PSObject.Properties | ForEach-Object {
                    $Script:Config[$_.Name] = $_.Value
                }
            }
            Write-Log "INFO" "Configuration loaded from: $configPath"
        } catch {
            $errMsg = $_.Exception.Message
            Write-Warning "Failed to load config, using defaults - $errMsg"
            $Script:Config = $defaultConfig
            Save-Config -Config $Script:Config -Path $configPath
        }
    } else {
        # Create default config file
        $Script:Config = $defaultConfig
        try {
            Save-Config -Config $Script:Config -Path $configPath
            Write-Host "Created default config: $configPath" -ForegroundColor Cyan
        } catch {
            $errMsg = $_.Exception.Message
            Write-Warning "Could not create config file - $errMsg"
        }
    }
    
    # Ensure required keys exist
    if (-not $Script:Config.categories) { $Script:Config.categories = $defaultConfig.categories }
    if (-not $Script:Config.excludePatterns) { $Script:Config.excludePatterns = $defaultConfig.excludePatterns }
    if (-not $Script:Config.createTypeSubfolder) { $Script:Config.createTypeSubfolder = $true }
}

function Save-Config {
    <#
    .SYNOPSIS
        Saves configuration to JSON file
    #>
    [CmdletBinding()]
    param(
        [hashtable]$Config,
        [string]$Path
    )
    
    try {
        $Config | ConvertTo-Json -Depth 10 | Set-Content $Path -Force -ErrorAction Stop
    } catch {
        $errMsg = $_.Exception.Message
        Write-Warning "Failed to save config - $errMsg"
    }
}

function Load-LastUsed {
    <#
    .SYNOPSIS
        Loads last used source/destination paths
    #>
    [CmdletBinding()]
    param()
    
    $lastUsedPath = Join-Path $Script:ScriptPath "lastused.json"
    
    if (Test-Path $lastUsedPath) {
        try {
            $content = Get-Content $lastUsedPath -Raw
            # Handle PS 5.1 vs PS 6+ compatibility
            if ($Script:PSVersion -ge 6) {
                $Script:LastUsed = $content | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            } else {
                $loaded = $content | ConvertFrom-Json -ErrorAction Stop
                # Convert PSObject to Hashtable for PS 5.1 compatibility
                $Script:LastUsed = @{}
                $loaded.PSObject.Properties | ForEach-Object {
                    $Script:LastUsed[$_.Name] = $_.Value
                }
            }
        } catch {
            $Script:LastUsed = @{ sourcePath = ""; destPath = "" }
        }
    } else {
        $Script:LastUsed = @{ sourcePath = ""; destPath = "" }
    }
}

function Save-LastUsed {
    <#
    .SYNOPSIS
        Saves last used paths for future sessions
    #>
    [CmdletBinding()]
    param(
        [string]$SourcePath,
        [string]$DestPath
    )
    
    $lastUsedPath = Join-Path $Script:ScriptPath "lastused.json"
    
    $Script:LastUsed.sourcePath = $SourcePath
    $Script:LastUsed.destPath = $DestPath
    
    try {
        $Script:LastUsed | ConvertTo-Json | Set-Content $lastUsedPath -Force
    } catch {
        $errMsg = $_.Exception.Message
        Write-Warning "Could not save last used paths - $errMsg"
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes timestamped log entry to file
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Level,
        
        [Parameter(Mandatory)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    try {
        Add-Content -Path $Script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {
        # Silently continue if logging fails
    }
    
    # Also output to console based on level
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
        "INFO"  { Write-Verbose $logEntry }
        "DRYRUN" { Write-Host $logEntry -ForegroundColor Cyan }
    }
}

#endregion

#region GUI Helper Functions

function Test-OutGridViewAvailable {
    <#
    .SYNOPSIS
        Checks if Out-GridView is available (Windows only)
    #>
    if ($IsWindows -or (-not $IsLinux -and -not $IsMacOS)) {
        return $true
    }
    return $false
}

function Select-SourceFolder {
    <#
    .SYNOPSIS
        Shows folder browser dialog for source selection
    #>
    [CmdletBinding()]
    param(
        [string]$InitialPath
    )
    
    # Use last used path if available and no initial path provided
    if (-not $InitialPath -and $Script:LastUsed.sourcePath) {
        if (Test-Path $Script:LastUsed.sourcePath) {
            $InitialPath = $Script:LastUsed.sourcePath
        }
    }
    
    # Use Windows Forms folder browser
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select Source Folder to Scan"
    $dialog.ShowNewFolderButton = $false
    
    if ($InitialPath -and (Test-Path $InitialPath)) {
        $dialog.SelectedPath = $InitialPath
    }
    
    $result = $dialog.ShowDialog()
    
    if ($result -eq "OK") {
        Write-Log "INFO" "Source folder selected: $($dialog.SelectedPath)"
        return $dialog.SelectedPath
    }
    
    return $null
}

function Select-DestinationFolder {
    <#
    .SYNOPSIS
        Shows folder browser dialog for destination selection
    #>
    [CmdletBinding()]
    param(
        [string]$Description = "Select Destination Folder",
        [string]$InitialPath
    )
    
    # Use last used path if available
    if (-not $InitialPath -and $Script:LastUsed.destPath) {
        if (Test-Path $Script:LastUsed.destPath) {
            $InitialPath = $Script:LastUsed.destPath
        }
    }
    
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    
    if ($InitialPath -and (Test-Path $InitialPath)) {
        $dialog.SelectedPath = $InitialPath
    }
    
    $result = $dialog.ShowDialog()
    
    if ($result -eq "OK") {
        Write-Log "INFO" "Destination folder selected: $($dialog.SelectedPath)"
        return $dialog.SelectedPath
    }
    
    return $null
}

function Select-DepthGUI {
    <#
    .SYNOPSIS
        Shows depth selection via Out-GridView or console
    #>
    [CmdletBinding()]
    param()
    
    $depthOptions = @(
        @{ Name = "Main Folder Only"; Value = 0; Description = "Scan only the main folder (depth=0)" }
        @{ Name = "Subfolders"; Value = 1; Description = "Scan main folder + one level subfolders" }
        @{ Name = "SubSubfolders"; Value = 2; Description = "Scan main + 2 levels of subfolders" }
        @{ Name = "All Subfolders"; Value = 100; Description = "Scan all folders (recursive)" }
    )
    
    # Try Out-GridView first (Windows only)
    if (Test-OutGridViewAvailable) {
        try {
            $selected = $depthOptions | Out-GridView -Title "Select Scan Depth" -OutputMode Single
            if ($selected) {
                return $selected.Value
            }
        } catch {
            # Fall back to console
        }
    }
    
    # Console fallback
    Write-Host "`nScan Depth:" -ForegroundColor Cyan
    Write-Host "  1 - Main Folder Only    (0)"
    Write-Host "  2 - Subfolders         (1)"
    Write-Host "  3 - SubSubfolders      (2)"
    Write-Host "  4 - All Subfolders     (100 - default)" -ForegroundColor Yellow
    
    $choice = Read-Host "Enter choice (1-4, default 4)"
    
    switch ($choice) {
        "1" { return 0 }
        "2" { return 1 }
        "3" { return 2 }
        default { return 100 }
    }
}

function Show-TypeSelector {
    <#
    .SYNOPSIS
        Shows file type selection with Out-GridView checkboxes
    #>
    [CmdletBinding()]
    param(
        [hashtable]$TypeCounts
    )
    
    # Build selection objects for Out-GridView
    $selectionList = @()
    foreach ($type in $TypeCounts.Keys | Sort-Object) {
        $selectionList += [PSCustomObject]@{
            Type = $type
            Count = $TypeCounts[$type]
            Selected = $true  # Default all selected
        }
    }
    
    # Try Out-GridView with checkboxes
    if (Test-OutGridViewAvailable) {
        try {
            $selected = $selectionList | Out-GridView -Title "Select File Types to Organize (All Selected by Default)" -OutputMode Multiple
            
            if ($selected) {
                return $selected.Type
            }
        } catch {
            # Fall back to console
        }
    }
    
    # Console fallback with numbered selection
    Write-Host "`nDetected File Types:" -ForegroundColor Green
    $indexMap = @{}
    $i = 1
    
    foreach ($type in $TypeCounts.Keys | Sort-Object) {
        Write-Host "  $i - $type ($($TypeCounts[$type]) files)" -ForegroundColor White
        $indexMap[$i] = $type
        $i++
    }
    
    Write-Host "`nEnter numbers separated by commas (e.g., 1,3,5) or press Enter for ALL" -ForegroundColor Cyan
    $choice = Read-Host "Selection"
    
    $selectedTypes = @()
    
    if ([string]::IsNullOrWhiteSpace($choice)) {
        # All selected
        return $TypeCounts.Keys
    } else {
        $numbers = $choice.Split(",") | ForEach-Object { $_.Trim() }
        foreach ($n in $numbers) {
            if ($indexMap.ContainsKey([int]$n)) {
                $selectedTypes += $indexMap[[int]$n]
            }
        }
        
        if ($selectedTypes.Count -eq 0) {
            throw "No valid selections made"
        }
        
        return $selectedTypes
    }
}

function Select-ActionMode {
    <#
    .SYNOPSIS
        Shows Move/Copy selection dialog
    #>
    [CmdletBinding()]
    param()
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(400, 180)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(350, 30)
    $label.Text = "Select Operation Mode:"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    
    $moveRadio = New-Object System.Windows.Forms.RadioButton
    $moveRadio.Location = New-Object System.Drawing.Point(30, 60)
    $moveRadio.Size = New-Object System.Drawing.Size(150, 25)
    $moveRadio.Text = "Move (default)"
    $moveRadio.Checked = $true
    
    $copyRadio = New-Object System.Windows.Forms.RadioButton
    $copyRadio.Location = New-Object System.Drawing.Point(200, 60)
    $copyRadio.Size = New-Object System.Drawing.Size(150, 25)
    $copyRadio.Text = "Copy (safer)"
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(150, 110)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = "OK"
    
    $form.Controls.AddRange(@($label, $moveRadio, $copyRadio, $okButton))
    $form.AcceptButton = $okButton
    
    $result = $form.ShowDialog()
    
    if ($result -eq "OK") {
        if ($copyRadio.Checked) {
            return "Copy"
        }
    }
    
    return "Move"
}

function Confirm-Summary {
    <#
    .SYNOPSIS
        Shows rich confirmation MessageBox with operation summary
    #>
    [CmdletBinding()]
    param(
        [string]$SourcePath,
        [string[]]$SelectedTypes,
        [hashtable]$TypeCounts,
        [string]$Action,
        [hashtable]$Destinations,
        [bool]$DryRun
    )
    
    # Build summary message
    $totalFiles = 0
    foreach ($type in $SelectedTypes) {
        $totalFiles += $TypeCounts[$type]
    }
    
    $summary = @"
Source Folder: $SourcePath

Selected Types ($($SelectedTypes.Count) categories):
"@
    
    foreach ($type in $SelectedTypes | Sort-Object) {
        $count = $TypeCounts[$type]
        $dest = $Destinations[$type]
        $summary += "`n  • $type : $count files`n    → $dest"
    }
    
    $summary += @"

Operation: $Action $(if($DryRun){"(DRY RUN - No files will be modified)"})
Total Files: $totalFiles

Proceed with this operation?
"@
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $buttons = if ($DryRun) { "YesNo" } else { "YesNo" }
    $icon = if ($DryRun) { "Question" } else { "Question" }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $summary,
        "Confirm File Organization",
        $buttons,
        $icon
    )
    
    return ($result -eq "Yes")
}

function Select-DestinationMode {
    <#
    .SYNOPSIS
        Shows destination mode selection (single vs separate)
    #>
    [CmdletBinding()]
    param()
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Destination Mode"
    $form.Size = New-Object System.Drawing.Size(450, 200)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(400, 30)
    $label.Text = "Select Destination Mode:"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    
    $singleRadio = New-Object System.Windows.Forms.RadioButton
    $singleRadio.Location = New-Object System.Drawing.Point(30, 60)
    $singleRadio.Size = New-Object System.Drawing.Size(380, 25)
    $singleRadio.Text = "Single Destination (all types to one folder)"
    $singleRadio.Checked = $true
    
    $separateRadio = New-Object System.Windows.Forms.RadioButton
    $separateRadio.Location = New-Object System.Drawing.Point(30, 90)
    $separateRadio.Size = New-Object System.Drawing.Size(380, 25)
    $separateRadio.Text = "Separate Destination (pick folder per file type)"
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(175, 130)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = "OK"
    
    $form.Controls.AddRange(@($label, $singleRadio, $separateRadio, $okButton))
    $form.AcceptButton = $okButton
    
    $result = $form.ShowDialog()
    
    if ($result -eq "OK") {
        if ($separateRadio.Checked) {
            return "Separate"
        }
    }
    
    return "Single"
}

#endregion

#region Core Processing Functions

function Get-FileType {
    <#
    .SYNOPSIS
        Determines file category based on extension
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Extension
    )
    
    $ext = $Extension.TrimStart('.').ToLower()
    
    # Handle PS 5.1 (PSCustomObject) vs PS 6+ (Hashtable)
    $categories = $Script:Config.categories
    if ($categories -is [System.Collections.Hashtable]) {
        $categoryList = $categories.Keys
    } else {
        # PS 5.1 - convert PSObject to collection
        $categoryList = $categories.PSObject.Properties.Name
    }
    
    foreach ($category in $categoryList) {
        if ($categories -is [System.Collections.Hashtable]) {
            $extensions = $categories[$category]
        } else {
            $extensions = $categories.$category
        }
        if ($extensions -contains $ext) {
            return $category
        }
    }
    
    return "Others"
}

function Test-ShouldExclude {
    <#
    .SYNOPSIS
        Checks if file matches any exclude patterns
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FileName
    )
    
    foreach ($pattern in $Script:Config.excludePatterns) {
        # Handle wildcard patterns
        if ($pattern.Contains("*")) {
            if ($FileName -like $pattern) {
                return $true
            }
        }
        # Handle prefix patterns like ~$ (temporary Office files)
        elseif ($pattern.StartsWith("~")) {
            if ($FileName.StartsWith($pattern)) {
                return $true
            }
        }
        # Exact match
        else {
            if ($FileName -eq $pattern) {
                return $true
            }
        }
    }
    
    return $false
}

function Scan-SourceFiles {
    <#
    .SYNOPSIS
        Scans source folder and returns file list with metadata
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,
        
        [Parameter(Mandatory)]
        [int]$Depth
    )
    
    Write-Host "`nScanning files..." -ForegroundColor Cyan
    Write-Log "INFO" "Scanning source: $SourcePath with depth: $Depth"
    
    $files = @()
    
    try {
        # PowerShell 7+ has -Depth parameter
        if ($Script:PSVersion -ge 7) {
            $files = Get-ChildItem -Path $SourcePath -File -Recurse -Depth $Depth -ErrorAction SilentlyContinue
        } else {
            # PowerShell 5.1 - simulate depth limit
            if ($Depth -ge 100) {
                $files = Get-ChildItem -Path $SourcePath -File -Recurse -ErrorAction SilentlyContinue
            } else {
                # Manual depth filtering for PS 5.1
                $files = Get-ChildItem -Path $SourcePath -File -Recurse -ErrorAction SilentlyContinue | 
                    Where-Object { $_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count -le 
                    ($SourcePath.Split([IO.Path]::DirectorySeparatorChar).Count + $Depth) }
            }
        }
    } catch {
        $errMsg = $_.Exception.Message
        Write-Log "ERROR" "Error scanning source - $errMsg"
        throw "Failed to scan source folder - $errMsg"
    }
    
    # Filter out excluded files
    $filteredFiles = @()
    foreach ($file in $files) {
        if (-not (Test-ShouldExclude -FileName $file.Name)) {
            $filteredFiles += $file
        }
    }
    
    Write-Host "Found $($filteredFiles.Count) files (excluding $($files.Count - $filteredFiles.Count) matches)" -ForegroundColor Green
    
    return $filteredFiles
}

function Get-TypeSummary {
    <#
    .SYNOPSIS
        Counts files by category
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Files
    )
    
    $typeCount = @{}
    
    foreach ($file in $Files) {
        $typeName = Get-FileType -Extension $file.Extension
        
        if (-not $typeCount.ContainsKey($typeName)) {
            $typeCount[$typeName] = 0
        }
        $typeCount[$typeName]++
    }
    
    return $typeCount
}

function Get-UniqueFileName {
    <#
    .SYNOPSIS
        Generates unique filename to avoid conflicts
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory)]
        [string]$FileName
    )
    
    $fullPath = Join-Path $DestinationPath $FileName
    
    if (-not (Test-Path $fullPath)) {
        return $FileName
    }
    
    # File exists, generate unique name
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [IO.Path]::GetExtension($FileName)
    
    $counter = 1
    do {
        $newName = "${baseName}_${counter}${extension}"
        $fullPath = Join-Path $DestinationPath $newName
        $counter++
    } while (Test-Path $fullPath)
    
    return $newName
}

function Check-Cancel {
    <#
    .SYNOPSIS
        Checks if user pressed Escape to cancel operation
    #>
    [CmdletBinding()]
    param()
    
    if ($Script:PSVersion -ge 7 -and [Console]::IsInputRedirected) {
        return $false
    }
    
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq "Escape") {
            Write-Host "`nOperation Cancelled by user." -ForegroundColor Yellow
            Write-Log "WARN" "Operation cancelled by user"
            return $true
        }
    }
    return $false
}

#endregion

#region File Operation Functions

function Process-Files {
    <#
    .SYNOPSIS
        Main file processing loop with progress and cancellation
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Files,
        
        [Parameter(Mandatory)]
        [string[]]$SelectedTypes,
        
        [Parameter(Mandatory)]
        [hashtable]$TypeDestinationMap,
        
        [Parameter(Mandatory)]
        [string]$Action,
        
        [bool]$DryRun,
        
        [bool]$CreateTypeSubfolder
    )
    
    $processedCount = 0
    $errorCount = 0
    $skippedCount = 0
    $total = $Files.Count
    $index = 0
    
    $operationVerb = if ($Action -eq "Move") { "Moving" } else { "Copying" }
    $logAction = if ($DryRun) { "DRYRUN" } else { "INFO" }
    
    Write-Host "`n$operationVerb files..." -ForegroundColor $(if($DryRun){"Cyan"}else{"Yellow"})
    
    foreach ($file in $Files) {
        # Check for cancellation
        if (Check-Cancel) {
            Write-Host "`nOperation cancelled. Stopping..." -ForegroundColor Yellow
            break
        }
        
        $index++
        $percent = [math]::Round(($index / $total) * 100)
        Write-Progress -Activity "$operationVerb Files ($Action)" -Status "$index of $total files" -PercentComplete $percent
        
        # Determine file type
        $typeName = Get-FileType -Extension $file.Extension
        
        # Skip if not in selected types
        if ($SelectedTypes -notcontains $typeName) {
            $skippedCount++
            continue
        }
        
        # Get destination path
        $baseDest = $TypeDestinationMap[$typeName]
        
        # Create type subfolder if enabled
        if ($CreateTypeSubfolder) {
            $destPath = Join-Path $baseDest $typeName
        } else {
            $destPath = $baseDest
        }
        
        # Create destination if it doesn't exist
        if (-not (Test-Path $destPath)) {
            try {
                New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                Write-Log "INFO" "Created directory: $destPath"
            } catch {
                $errMsg = $_.Exception.Message
                Write-Log "ERROR" "Failed to create directory $destPath - $errMsg"
                $errorCount++
                continue
            }
        }
        
        # Get unique filename
        $newFileName = Get-UniqueFileName -DestinationPath $destPath -FileName $file.Name
        $newFullPath = Join-Path $destPath $newFileName
        
        # Perform the operation
        if ($DryRun) {
            # Dry run - just log what would happen
            Write-Log "DRYRUN" "Would $Action '$($file.FullName)' -> '$newFullPath'"
            $processedCount++
        } else {
            # Actual move or copy
            try {
                if ($Action -eq "Move") {
                    Move-Item -Path $file.FullName -Destination $newFullPath -Force -ErrorAction Stop
                } else {
                    Copy-Item -Path $file.FullName -Destination $newFullPath -Force -ErrorAction Stop
                }
                
                Write-Log "INFO" "$Action completed: $($file.Name) -> $newFullPath"
                $processedCount++
            } catch {
                $errMsg = $_.Exception.Message
                $fileName = $file.Name
                Write-Log "ERROR" "Failed to $Action '$fileName' - $errMsg"
                $errorCount++
            }
        }
    }
    
    Write-Progress -Activity "$operationVerb Files" -Completed
    
    # Return summary
    return @{
        Processed = $processedCount
        Errors = $errorCount
        Skipped = $skippedCount
    }
}

#endregion

#region Interactive Mode

function Start-InteractiveMode {
    <#
    .SYNOPSIS
        Runs the interactive GUI-based workflow
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "`n=== FileOrganizer - Interactive Mode ===" -ForegroundColor Cyan
    Write-Host "Press ESC anytime to cancel operation`n" -ForegroundColor Gray
    
    # Give user a moment to see the startup
    Start-Sleep -Milliseconds 500
    
    # Step 1: Select source
    Write-Host "Step 1: Select Source Folder (a dialog will open)" -ForegroundColor Yellow
    Write-Host ">>> Click Browse and select a folder <<<" -ForegroundColor Magenta
    $sourceFolder = Select-SourceFolder
    
    if (-not $sourceFolder) {
        Write-Host "No source folder selected. Exiting." -ForegroundColor Yellow
        return
    }
    
    # Step 2: Select depth
    Write-Host "`nStep 2: Select Scan Depth" -ForegroundColor Yellow
    $depth = Select-DepthGUI
    Write-Host "Using depth: $depth" -ForegroundColor Gray
    
    # Step 3: Scan files
    $files = Scan-SourcePath -SourcePath $sourceFolder -Depth $depth
    
    if ($files.Count -eq 0) {
        Write-Host "No files found in source folder." -ForegroundColor Yellow
        return
    }
    
    # Step 4: Get type summary
    $typeCounts = Get-TypeSummary -Files $files
    
    # Step 5: Select types
    Write-Host "`nStep 3: Select File Types" -ForegroundColor Yellow
    $selectedTypes = Show-TypeSelector -TypeCounts $typeCounts
    
    if ($selectedTypes.Count -eq 0) {
        Write-Host "No file types selected. Exiting." -ForegroundColor Yellow
        return
    }
    
    # Step 6: Select action (Move/Copy)
    Write-Host "`nStep 4: Select Operation" -ForegroundColor Yellow
    $action = Select-ActionMode
    Write-Host "Operation: $action" -ForegroundColor Gray
    
    # Step 7: Destination mode
    Write-Host "`nStep 5: Select Destination Mode" -ForegroundColor Yellow
    $destMode = Select-DestinationMode
    
    $typeDestinationMap = @{}
    
    if ($destMode -eq "Separate") {
        # Separate destination per type
        foreach ($type in $selectedTypes) {
            Write-Host "`nSelect destination for: $type" -ForegroundColor Cyan
            $dest = Select-DestinationFolder -Description "Select destination for $type"
            
            if (-not $dest) {
                Write-Host "Destination not selected. Using source folder." -ForegroundColor Yellow
                $dest = $sourceFolder
            }
            
            $typeDestinationMap[$type] = $dest
        }
    } else {
        # Single destination
        Write-Host "`nSelect single destination folder" -ForegroundColor Cyan
        $singleDest = Select-DestinationFolder -Description "Select destination for all files"
        
        if (-not $singleDest) {
            Write-Host "No destination selected. Using source folder." -ForegroundColor Yellow
            $singleDest = $sourceFolder
        }
        
        foreach ($type in $selectedTypes) {
            $typeDestinationMap[$type] = $singleDest
        }
    }
    
    # Step 8: DryRun option
    Write-Host "`nStep 6: Dry Run Mode?" -ForegroundColor Yellow
    $dryRunChoice = Read-Host "Run as Dry Run (simulate only)? y/n (default: n)"
    $dryRun = ($dryRunChoice -eq "y")
    
    if ($dryRun) {
        Write-Host "`n*** DRY RUN MODE - No files will be modified ***" -ForegroundColor Cyan
    }
    
    # Step 9: Confirm summary
    Write-Host "`nStep 7: Confirm Operation" -ForegroundColor Yellow
    $confirmed = Confirm-Summary -SourcePath $sourceFolder -SelectedTypes $selectedTypes -TypeCounts $typeCounts -Action $action -Destinations $typeDestinationMap -DryRun $dryRun
    
    if (-not $confirmed) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    # Step 10: Process files
    $result = Process-Files -Files $files -SelectedTypes $selectedTypes -TypeDestinationMap $typeDestinationMap -Action $action -DryRun $dryRun -CreateTypeSubfolder $Script:Config.createTypeSubfolder
    
    # Step 11: Show results
    Write-Host "`n=== Operation Complete ===" -ForegroundColor Green
    Write-Host "Files processed: $($result.Processed)" -ForegroundColor White
    
    if ($result.Errors -gt 0) {
        Write-Host "Errors: $($result.Errors)" -ForegroundColor Red
    }
    
    if ($result.Skipped -gt 0) {
        Write-Host "Skipped: $($result.Skipped)" -ForegroundColor Yellow
    }
    
    Write-Log "INFO" "Operation completed - Processed: $($result.Processed), Errors: $($result.Errors), Skipped: $($result.Skipped)"
    
    # Save last used paths on success
    if ($result.Processed -gt 0 -and -not $dryRun) {
        $firstDest = ($typeDestinationMap.Values | Select-Object -First 1)
        Save-LastUsed -SourcePath $sourceFolder -DestPath $firstDest
    }
}

function Scan-SourcePath {
    <#
    .SYNOPSIS
        Wrapper for scanning with user feedback
    #>
    [CmdletBinding()]
    param(
        [string]$SourcePath,
        [int]$Depth
    )
    
    return Scan-SourceFiles -SourcePath $SourcePath -Depth $Depth
}

function Start-CLIMode {
    <#
    .SYNOPSIS
        Runs the command-line interface mode with provided parameters
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "`n=== FileOrganizer - CLI Mode ===" -ForegroundColor Cyan
    
    # Validate source path
    if (-not $SourcePath) {
        Write-Error "SourcePath is required in CLI mode. Use -SourcePath parameter."
        return
    }
    
    if (-not (Test-Path $SourcePath)) {
        Write-Error "Source path does not exist: $SourcePath"
        return
    }
    
    # Set default destination if not provided
    if (-not $DestPath) {
        $DestPath = $SourcePath
        Write-Host "No destination specified, using source path: $DestPath" -ForegroundColor Yellow
    }
    
    # Scan files
    $files = Scan-SourceFiles -SourcePath $SourcePath -Depth $Depth
    
    if ($files.Count -eq 0) {
        Write-Host "No files found." -ForegroundColor Yellow
        return
    }
    
    # Get type summary
    $typeCounts = Get-TypeSummary -Files $files
    
    # Show summary
    Write-Host "`nFile Type Summary:" -ForegroundColor Green
    foreach ($type in ($typeCounts.Keys | Sort-Object)) {
        Write-Host "  $type : $($typeCounts[$type]) files" -ForegroundColor White
    }
    
    # Determine selected types (all by default in CLI)
    $selectedTypes = $typeCounts.Keys
    
    # Set up destination map
    $typeDestinationMap = @{}
    
    if ($SeparateDestinations) {
        Write-Host "`nSeparate destinations requested but not implemented in CLI mode." -ForegroundColor Yellow
        Write-Host "Using single destination: $DestPath" -ForegroundColor Yellow
    }
    
    foreach ($type in $selectedTypes) {
        $typeDestinationMap[$type] = $DestPath
    }
    
    # Show dry run status
    if ($DryRun) {
        Write-Host "`n*** DRY RUN MODE - No files will be modified ***" -ForegroundColor Cyan
    }
    
    # Process files
    $result = Process-Files -Files $files -SelectedTypes $selectedTypes -TypeDestinationMap $typeDestinationMap -Action $Action -DryRun $DryRun -CreateTypeSubfolder $Script:Config.createTypeSubfolder
    
    # Show results
    Write-Host "`n=== Operation Complete ===" -ForegroundColor Green
    Write-Host "Files processed: $($result.Processed)" -ForegroundColor White
    
    if ($result.Errors -gt 0) {
        Write-Host "Errors: $($result.Errors)" -ForegroundColor Red
    }
    
    Write-Log "INFO" "CLI operation completed - Processed: $($result.Processed), Errors: $($result.Errors)"
    
    # Save last used paths
    if ($result.Processed -gt 0 -and -not $DryRun) {
        Save-LastUsed -SourcePath $SourcePath -DestPath $DestPath
    }
}

#endregion

#region Main Entry Point

function Main {
    <#
    .SYNOPSIS
        Main entry point for the script
    #>
    [CmdletBinding()]
    param()
    
    # Initialize environment
    Initialize-Environment
    
    # Check if Help requested
    if ($Help) {
        Get-Help $MyInvocation.MyCommand.Path -Detailed
        return
    }
    
    # Determine mode: CLI or Interactive
    $isCLIMode = $SourcePath -and $DestPath
    
    if ($isCLIMode) {
        # Run in CLI mode
        Start-CLIMode
    } else {
        # Run in Interactive mode
        Start-InteractiveMode
        
        # Loop for additional tasks unless -NoLoop specified
        if (-not $NoLoop) {
            while ($true) {
                Write-Host "`n" -NoNewline
                $again = Read-Host "Do another task? y/n"
                
                if ($again -ne "y") {
                    break
                }
                
                Start-InteractiveMode
            }
        }
    }
    
    Write-Host "`nGoodbye!" -ForegroundColor Cyan
    Write-Log "INFO" "FileOrganizer ended"
    
    # Pause to keep window open (unless running with -NoExit or in CI)
    if (-not ([Environment]::GetCommandLineArgs() -contains '-NoExit')) {
        Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
        try { Read-Host } catch { Start-Sleep -Seconds 2 }
    }
}

# Run main function
try {
    Main
} catch {
    $errMsg = $_.Exception.Message
    $stackTrace = $_.ScriptStackTrace
    Write-Host "`n" -ForegroundColor Red
    Write-Host "FATAL ERROR: $errMsg" -ForegroundColor Red
    Write-Host "Stack Trace: $stackTrace" -ForegroundColor Red
    Write-Host "`nPress Enter to exit..." -ForegroundColor Yellow
    try { Read-Host } catch { Start-Sleep -Seconds 5 }
    exit 1
}

#endregion
