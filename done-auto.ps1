## Setup script by Doakyz, QuackWalks and A.I. for use with Done Set 3

# ---------------------------------------------------------
# PREVENT SLEEP / LOCK THROTTLING
# ---------------------------------------------------------
$code = @"
    [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
"@

if (-not ("Win32.Win32" -as [type])) {
    Add-Type -MemberDefinition $code -Name Win32 -Namespace Win32
}

# Flags for the API:
# ES_SYSTEM_REQUIRED (0x01) = Forces the system to be in the working state by resetting the system idle timer.
# ES_DISPLAY_REQUIRED (0x02) = Forces the display to be on (optional, but prevents lock screen throttling).
# ES_CONTINUOUS (0x80000000) = Informs the system that the state being set should remain in effect until the next call.
$ES_SYSTEM_REQUIRED = [uint32]0x00000001
$ES_DISPLAY_REQUIRED = [uint32]0x00000002
$ES_CONTINUOUS = [uint32]"0x80000000"

# Activate "Insomnia Mode"
Write-Host "Setting power execution state to prevent sleep..." -ForegroundColor DarkGray
[Win32.Win32]::SetThreadExecutionState($ES_SYSTEM_REQUIRED -bor $ES_DISPLAY_REQUIRED -bor $ES_CONTINUOUS)

try {
    $scriptLogPath = Join-Path $PSScriptRoot "setup_log.txt"
    Start-Transcript -Path $scriptLogPath -Append

    # Check for Administrator privileges
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Warning: Script is not running as Administrator." -ForegroundColor Yellow
        Write-Host "Preventing sleep and preserving file attributes works best with Admin rights."
    }

# Initialize GUI resources
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName presentationframework
Add-Type -AssemblyName microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

# Required for use with web SSL sites
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Load necessary modules
Import-Module Microsoft.PowerShell.Management
Import-Module Microsoft.PowerShell.Utility

# Set Window Title
$Host.UI.RawUI.WindowTitle = "Done Set 3 Auto Setup"

# --- Script Configuration ---
$script:ZipFiles = @{
    BaseSet = "Done Set 3.zip"
    ConfigsPlus = "Configs for Plus Model.zip"
    ConfigsV4 = "Configs for V4 Model.zip"
    SensibleArrangement = "Sensible Console Arrangement.zip"
    Manuals = "Manuals.zip"
    Cheats = "Cheats.zip"
    Imgs2DBoxAndScreenshot = "Imgs (2D Box and Screenshot).zip"
    Imgs2DBox = "Imgs (2D Box).zip"
    ImgsMiyooMix = "Imgs (Miyoo Mix).zip"
    PS1Addon = "PS1 Addon for 256gb SD Cards.zip"
}
# --- End Script Configuration ---

# --- Pre-Validation ---
$baseSetPath = Join-Path $PSScriptRoot $script:ZipFiles.BaseSet
if (-not (Test-Path $baseSetPath)) {
    Write-Host "Error: The base set file '$($script:ZipFiles.BaseSet)' was not found." -ForegroundColor Red
    Write-Host "Please ensure the zip file is in: $PSScriptRoot"
    Read-Host "Press Enter to exit"
    exit
}

# Function to display GUI folder selection dialog
function Show-FolderDialog {
    param(
        [string]$initialDirectory
    )
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select a directory"
    $folderDialog.RootFolder = "MyComputer"
    if (-not [string]::IsNullOrWhiteSpace($initialDirectory)) {
        $folderDialog.SelectedPath = $initialDirectory
    }
    $result = $folderDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderDialog.SelectedPath
    } else {
        return $null
    }
}

# Function to prompt for yes or no answers
function Prompt-YesNoQuestion {
    while ($true) {
        Write-Host "Y or N ?"
        $choice = Read-Host

        if ($choice -eq 'y') {
            return $true
        } elseif ($choice -eq 'n') {
            return $false
        } else {
            Write-Host "Invalid choice. Please enter 'y' or 'n'."
        }
    }
}

# Get the extraction location using the GUI folder selection dialog
Write-Host "Select the extraction path."
$extractionPath = Show-FolderDialog -initialDirectory ([System.Environment]::GetFolderPath('Desktop'))
if (-not $extractionPath) {
    Write-Host "No folder selected. Exiting script."
    exit
}

# Check if the extraction path exists, if not, create it
if (-not (Test-Path -Path $extractionPath)) {
    New-Item -ItemType Directory -Path $extractionPath | Out-Null
}

# Check for potential path length issues
if ($extractionPath.Length -gt 80) {
    Write-Host "Warning: The selected path is long ($($extractionPath.Length) characters)." -ForegroundColor Yellow
    Write-Host "Deep folder structures might exceed the Windows 260 character limit."
    Write-Host "It is recommended to extract to a shorter path (e.g., root of a drive)."
    if (-not (Prompt-YesNoQuestion)) { exit }
}

# Check for write access to the destination
$testFilePath = Join-Path $extractionPath "test_write_permission.tmp"
try {
    "test" | Out-File -FilePath $testFilePath -ErrorAction Stop
    Remove-Item -Path $testFilePath -ErrorAction SilentlyContinue
} catch {
    Write-Host "Error: Cannot write to the destination path: $extractionPath" -ForegroundColor Red
    Write-Host "Please check if the drive is read-only or locked."
    Read-Host "Press Enter to exit"
    exit
}

# Function to prompt for additional zip files based on the storage capacity
function Prompt-AdditionalZipFiles {
    $additionalZipFiles = @()

    # Configs
    Write-Host "Do you want optimizations and overlays? (recommended)"
    if (Prompt-YesNoQuestion) {
        # Prompt for model type
        $validChoice = $false
        while (-not $validChoice) {
            Write-Host "What Miyoo Mini model do you have?"
            Write-Host "1. Plus, v2 or v3 model"
            Write-Host "2. v4 model"
            $modelChoice = Read-Host
            switch ($modelChoice) {
                1 {
                    $configsZipFile = $script:ZipFiles.ConfigsPlus
                    $validChoice = $true
                }
                2 {
                    $configsZipFile = $script:ZipFiles.ConfigsV4
                    $validChoice = $true
                }
                default {
                    Write-Host "Invalid choice. Please enter '1' or '2'."
                }
            }
        }

        if ($configsZipFile -ne $null) {
            $additionalZipFiles += $configsZipFile
        }
    }

    # Sensible Console Arrangement
    Write-Host "Would you like a sensible arrangement of consoles (non-alphabetical)?"
    if (Prompt-YesNoQuestion) {
        $sensibleArrangementZipFile = $script:ZipFiles.SensibleArrangement
        if ($sensibleArrangementZipFile -ne $null) {
            $additionalZipFiles += $sensibleArrangementZipFile
        }
    }

    # Manuals
    Write-Host "Do you want game manuals?"
    if (Prompt-YesNoQuestion) {
        $manualsZipFile = $script:ZipFiles.Manuals
        if ($manualsZipFile -ne $null) {
            $additionalZipFiles += $manualsZipFile
        }
    }

    # Cheats
    Write-Host "Do you want cheats?"
    if (Prompt-YesNoQuestion) {
        $cheatsZipFile = $script:ZipFiles.Cheats
        $additionalZipFiles += $cheatsZipFile
    }

    # Thumbnails
    $validChoice = $false
    while (-not $validChoice) {
        Write-Host "Select your thumbnail option"
        Write-Host "1. 2D Box and Screenshot"
        Write-Host "2. 2D Box"
        Write-Host "3. Miyoo Mix"
        Write-Host "4. None"
        $pictureChoice = Read-Host
        switch ($pictureChoice) {
            1 { $additionalZipFiles += $script:ZipFiles.Imgs2DBoxAndScreenshot; $validChoice = $true }
            2 { $additionalZipFiles += $script:ZipFiles.Imgs2DBox; $validChoice = $true }
            3 { $additionalZipFiles += $script:ZipFiles.ImgsMiyooMix; $validChoice = $true }
            4 { $validChoice = $true }
            default { Write-Host "Invalid choice. Please enter '1', '2', '3' or '4'." }
        }
    }

    # PS1 Addon for 256GB SD Cards
    Write-Host "Would you like to install the PS1 addon for 256GB SD cards?"
    if (Prompt-YesNoQuestion) {
        $ps1AddonZipFile = $script:ZipFiles.PS1Addon
        if ($ps1AddonZipFile -ne $null) {
            $additionalZipFiles += $ps1AddonZipFile
        }
    }

    return $additionalZipFiles
}

# Function to update progress bar with asterisks
function Update-ProgressBar {
    param (
        [int]$percentage
    )
    $progressBar = "*" * ([math]::Round($percentage / 5))  # Each asterisk represents 5%
    Write-Host -NoNewline -Object "`r[$progressBar] $percentage% complete"
    $Host.UI.RawUI.WindowTitle = "Done Set 3 Setup - $percentage% Complete"
}

function Prompt-RobocopyProgress {
    while ($true) {
        Write-Host "" # Add a blank line for readability
        Write-Host "How would you like to monitor the file copy progress?" -ForegroundColor Yellow
        Write-Host "1. Show progress in this terminal (slower, provides feedback)."
        Write-Host "2. Write progress to a log file (faster, console will be mostly silent)."
        Write-Host "3. Silent copy (fastest, no progress shown)."
        $choice = Read-Host "Please enter 1, 2, or 3"
        switch ($choice) {
            '1' { return "terminal" }
            '2' { return "log" }
            '3' { return "silent" }
            default { Write-Host "Invalid choice." -ForegroundColor Red }
        }
    }
}

# Ask the user if they want to extract Done Set 3
Write-Host "Do you want to extract Done Set 3 Roms and BIOS?"
$extractDoneSet3 = Prompt-YesNoQuestion

$baseZipFile = $script:ZipFiles.BaseSet
$zipFilePaths = @()

if ($extractDoneSet3) {
    $zipFilePaths += $baseZipFile
}

# Get additional zip files to be extracted
$additionalZipFiles = Prompt-AdditionalZipFiles
$zipFilePaths += $additionalZipFiles

# Validate that selected zip files actually exist
$existingZipPaths = @()
$missingZipFiles = @()
foreach ($zipFile in $zipFilePaths) {
    if (Test-Path (Join-Path $PSScriptRoot $zipFile)) {
        $existingZipPaths += $zipFile
    } else {
        $missingZipFiles += $zipFile
    }
}
if ($missingZipFiles.Count -gt 0) {
    Write-Host "Warning: The following selected files are missing and will be skipped:" -ForegroundColor Yellow
    foreach ($missing in $missingZipFiles) { Write-Host "- $missing" -ForegroundColor Yellow }
    Write-Host ""
}
$zipFilePaths = $existingZipPaths
if ($zipFilePaths.Count -eq 0) { Write-Host "No valid zip files found. Exiting."; exit }

# Prompt user for robocopy progress preference
$userProgressChoice = Prompt-RobocopyProgress

# Prompt the user for confirmation to start the extraction
Write-Host ""
Write-Host "Please review your selections before proceeding with the extraction."
Write-Host "Zip files to be extracted:"
foreach ($zipFilePath in $zipFilePaths) {
    Write-Host "- $zipFilePath"
}
Write-Host "Extraction path: $extractionPath"
Write-Host "Copy progress mode: $userProgressChoice"

$confirmationPrompt = "Do you want to proceed with the extraction? (Y/N)"
$proceedWithExtraction = $null

while ($true) {
    $choice = Read-Host -Prompt $confirmationPrompt

    if ($choice -eq 'y') {
        $proceedWithExtraction = $true
        break  # Exit the loop to start extraction
    } elseif ($choice -eq 'n') {
        Write-Host "Extraction aborted by user."
        exit
    } else {
        Write-Host "Invalid choice. Please enter 'y' or 'n'."
    }
}

# Define a temporary path for extraction
$tempExtractionPath = Join-Path $PSScriptRoot "temp_extraction"

# --- Disk Space Check ---
$driveRoot = Split-Path $PSScriptRoot -Qualifier
$driveLetter = $driveRoot.TrimEnd(':')
$driveInfo = Get-PSDrive $driveLetter -ErrorAction SilentlyContinue

if ($driveInfo) {
    $estimatedSize = 0
    foreach ($zip in $zipFilePaths) {
        $p = Join-Path $PSScriptRoot $zip
        if (Test-Path $p) { $estimatedSize += (Get-Item $p).Length }
    }
    # Estimate 2.5x size for extraction overhead + original zip
    $requiredSpace = $estimatedSize * 2.5

    if ($driveInfo.Free -lt $requiredSpace) {
        $freeGB = [math]::Round($driveInfo.Free / 1GB, 2)
        $reqGB = [math]::Round($requiredSpace / 1GB, 2)
        Write-Host "Warning: Low disk space on $driveRoot." -ForegroundColor Yellow
        Write-Host "Available: $freeGB GB. Estimated Required: $reqGB GB." -ForegroundColor Yellow
        Write-Host "Do you want to continue anyway?"
        if (-not (Prompt-YesNoQuestion)) { exit }
    }
}

# Check if destination drive is FAT32 (Recommended for Miyoo Mini)
try {
    $destDrive = Split-Path $extractionPath -Qualifier
    if ($destDrive) {
        $destDriveLetter = $destDrive.TrimEnd(':')
        $volumeInfo = Get-Volume -DriveLetter $destDriveLetter -ErrorAction Stop
        if ($volumeInfo.FileSystem -ne "FAT32") {
            Write-Host "Warning: The destination drive ($destDrive) is formatted as $($volumeInfo.FileSystem)." -ForegroundColor Yellow
            Write-Host "Miyoo Mini and similar devices typically require FAT32."
            Write-Host "Do you want to continue anyway?"
            if (-not (Prompt-YesNoQuestion)) { exit }
        }
    }
} catch {}

# Create the temporary directory if it doesn't exist, and clear it if it does
if (Test-Path $tempExtractionPath) {
    Remove-Item -Path $tempExtractionPath -Recurse -Force
}
New-Item -ItemType Directory -Path $tempExtractionPath | Out-Null

# Perform the extraction with the -Force parameter to overwrite existing files
try {
    $extractionSuccessful = $false
    # Extract selected zip files
    $totalFiles = $zipFilePaths.Count
    $extractedFiles = 0

    # Check for 7-Zip
    $7zPath = $null
    $potential7zPaths = @("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe")
    # Check Registry for 7-Zip
    try {
        $regPath = "HKLM:\SOFTWARE\7-Zip"
        if (Test-Path $regPath) {
            $regProp = Get-ItemProperty -Path $regPath -Name "Path" -ErrorAction SilentlyContinue
            if ($regProp.Path) { $potential7zPaths += Join-Path $regProp.Path "7z.exe" }
        }
    } catch {}
    foreach ($p in $potential7zPaths) { if (Test-Path $p) { $7zPath = $p; break } }
    if (-not $7zPath -and (Get-Command "7z.exe" -ErrorAction SilentlyContinue)) { $7zPath = "7z.exe" }
    
    if ($7zPath) { Write-Host "7-Zip detected. Using it for faster extraction." -ForegroundColor Cyan }

    foreach ($zipFileName in $zipFilePaths) {
        $fullZipPath = Join-Path $PSScriptRoot $zipFileName
        if (Test-Path $fullZipPath) {
            Unblock-File -Path $fullZipPath -ErrorAction SilentlyContinue
            
            if ($7zPath) {
                Write-Host "Extracting $zipFileName with 7-Zip..."
                try {
                    & $7zPath x $fullZipPath "-o$tempExtractionPath" -y -bso0 -bsp1
                    if ($LASTEXITCODE -ne 0) { throw "7-Zip extraction failed for $zipFileName" }
                } catch {
                    Write-Host "7-Zip extraction failed. Falling back to standard extraction." -ForegroundColor Yellow
                    Expand-Archive -Path $fullZipPath -DestinationPath $tempExtractionPath -Force -ErrorAction Stop
                }
            } else {
                Write-Host "Extracting $zipFileName (Standard Windows extraction)..."
                Expand-Archive -Path $fullZipPath -DestinationPath $tempExtractionPath -Force -ErrorAction Stop
            }
            
            $extractedFiles++
            $percentage = [math]::Round(($extractedFiles / $totalFiles) * 100)
            Update-ProgressBar -percentage $percentage
        } else {
            Write-Host "File not found: $fullZipPath"
        }
    }

    Write-Host ""
    Write-Host "Extraction to temporary location complete."

    # --- Robocopy Command Construction ---
    $robocopyLogPath = Join-Path $PSScriptRoot "robocopy_log.txt"
    
    $roboArgs = @($tempExtractionPath, $extractionPath, "/MOVE", "/S", "/E", "/DCOPY:DA", "/COPY:DAT", "/MT:4", "/R:3", "/W:5", "/XX")

    switch ($userProgressChoice) {
        'terminal' {
            # Default args are fine
        }
        'log' {
            $roboArgs += "/LOG+:$robocopyLogPath"
            $roboArgs += "/NP"
            Write-Host "Robocopy progress will be logged to: $robocopyLogPath" -ForegroundColor Cyan
        }
        'silent' {
            $roboArgs += "/NFL"
            $roboArgs += "/NDL"
            $roboArgs += "/NP"
        }
    }
    # --- End Robocopy Command Construction ---

    Write-Host "Starting Robocopy to move files to the final destination: $extractionPath"
    & robocopy $roboArgs
    $exitCode = $LASTEXITCODE
    
    # Interpret the Robocopy exit code. Codes below 8 indicate success.
    if ($exitCode -lt 8) {
        Write-Host "File move process completed successfully." -ForegroundColor Green
        $extractionSuccessful = $true

        if ($zipFilePaths -contains $script:ZipFiles.Cheats) {
            $configFilePath = Join-Path $extractionPath "RetroArch\.retroarch\retroarch.cfg"
            if (Test-Path $configFilePath) {
                (Get-Content $configFilePath) -replace 'quick_menu_show_cheats = "false"', 'quick_menu_show_cheats = "true"' | Set-Content $configFilePath -Encoding UTF8
                Write-Host "Modified $configFilePath to enable cheats."
            } else {
                Write-Host "Config file not found: $configFilePath"
            }
        }
    } else {
        Write-Host "Error: Robocopy failed with exit code $exitCode." -ForegroundColor Red
        if ($exitCode -band 16) { Write-Host "- Fatal error encountered." -ForegroundColor Red }
        if ($exitCode -band 8) { Write-Host "- Retry limit exceeded (some files were locked)." -ForegroundColor Red }
        if ($exitCode -band 4) { Write-Host "- Mismatched files or attributes detected." -ForegroundColor Yellow }
        if ($exitCode -band 2) { Write-Host "- Extra files detected in destination." -ForegroundColor Yellow }
        if ($exitCode -band 1) { Write-Host "- Some files were copied successfully." -ForegroundColor Green }
        Write-Host "Please check the log file (if enabled) or the destination folder."
    }
} catch {
    Write-Host "An error occurred during extraction: $_"
} finally {
    # Clean up the temporary directory
    if ($extractionSuccessful) {
        Write-Host "Cleaning up temporary files..."
        if (Test-Path $tempExtractionPath) {
            Remove-Item -Path $tempExtractionPath -Recurse -Force
        }
        Write-Host "Cleanup complete."
    } else {
        Write-Host "Process did not complete successfully." -ForegroundColor Yellow
        Write-Host "Temporary files are located at: $tempExtractionPath" -ForegroundColor Yellow
        Write-Host "Do you want to delete these temporary files? (Y/N)"
        if (Prompt-YesNoQuestion) {
            Remove-Item -Path $tempExtractionPath -Recurse -Force
            Write-Host "Temporary files deleted."
        } else {
            Write-Host "Temporary files kept."
        }
    }
    Stop-Transcript -ErrorAction SilentlyContinue
}

} catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Script Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    # Deactivate "Insomnia Mode"
    Write-Host ""
    Write-Host "Restoring system power settings to default." -ForegroundColor DarkGray
    [Win32.Win32]::SetThreadExecutionState($ES_CONTINUOUS)
    
    # Notify user that the process is complete and prompt for manual exit
    Write-Host "Press Enter to exit."
    Read-Host
}
