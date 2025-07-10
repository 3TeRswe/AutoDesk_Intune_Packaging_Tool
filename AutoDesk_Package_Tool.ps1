# Enhanced AutoDesk Package Tool
# Version: 2.0
# Enhanced with comprehensive error handling, validation, and logging
#
# PARAMETERS:
# -BasePath: Base directory for Intune files (default: C:\Intune)
# -AutoDeskTempPath: Path where AutoDesk packages are deployed (default: C:\Windows\Temp\AutoDesk)
# -ForceDownload: Force download of IntuneWinAppUtil.exe even if it already exists
# -MaxFileAgeDays: Maximum age in days before redownloading IntuneWinAppUtil.exe (default: 30)
#
# EXAMPLES:
# .\AutoDesk_Package_Tool.ps1
# .\AutoDesk_Package_Tool.ps1 -ForceDownload
# .\AutoDesk_Package_Tool.ps1 -MaxFileAgeDays 7
# .\AutoDesk_Package_Tool.ps1 -BasePath "D:\Intune" -ForceDownload

param(
    [string]$BasePath = "C:\Intune",
    [string]$AutoDeskTempPath = "C:\Windows\Temp\AutoDesk",
    [switch]$ForceDownload,
    [int]$MaxFileAgeDays = 30
)

# Global variables
$Script:LogFile = Join-Path $env:TEMP "AutoDesk_Package_Tool_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

#region Helper Functions

function Write-LogMessage {
    param(
        [string]$Message, 
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] ${Level}: $Message"
    
    # Write to log file
    try {
        $logEntry | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8
    } catch {
        # If logging fails, continue silently
    }
    
    # Write to console unless suppressed
    if (!$NoConsole) {
        $color = switch ($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
        Write-Host $logEntry -ForegroundColor $color
    }
}

function Test-Prerequisites {
    Write-LogMessage "=== CHECKING PREREQUISITES ===" "INFO"
    $allGood = $true
    
    # Check if running as administrator
    Write-LogMessage "Checking administrator privileges..." "INFO"
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-LogMessage "Script must be run as Administrator" "ERROR"
        Write-LogMessage "Please run PowerShell as Administrator and try again." "ERROR"
        $allGood = $false
    } else {
        Write-LogMessage "Administrator privileges confirmed" "SUCCESS"
    }
    
    # Check if AutoDesk temp path exists
    Write-LogMessage "Checking AutoDesk temp path: $AutoDeskTempPath" "INFO"
    if (!(Test-Path $AutoDeskTempPath)) {
        Write-LogMessage "AutoDesk temp path not found: $AutoDeskTempPath" "ERROR"
        Write-LogMessage "Please ensure your AutoDesk package was deployed to the correct location." "ERROR"
        $allGood = $false
    } else {
        Write-LogMessage "AutoDesk temp path found" "SUCCESS"
    }
    
    # Check PowerShell version
    Write-LogMessage "Checking PowerShell version..." "INFO"
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-LogMessage "PowerShell 5.0 or higher is required. Current version: $($PSVersionTable.PSVersion)" "WARNING"
    } else {
        Write-LogMessage "PowerShell version: $($PSVersionTable.PSVersion) - OK" "SUCCESS"
    }
    
    # Check internet connectivity
    Write-LogMessage "Testing internet connectivity..." "INFO"
    try {
        $testConnection = Test-NetConnection -ComputerName "aka.ms" -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
        if ($testConnection) {
            Write-LogMessage "Internet connectivity confirmed" "SUCCESS"
        } else {
            Write-LogMessage "Internet connectivity test failed" "WARNING"
        }
    } catch {
        Write-LogMessage "Internet connectivity test failed: $($_.Exception.Message)" "WARNING"
        Write-LogMessage "Script will attempt to use cached IntuneWinAppUtil.exe if available" "INFO"
    }
    
    return $allGood
}

function New-RequiredFolders {
    param([string]$BasePath)
    
    Write-LogMessage "=== CREATING REQUIRED FOLDERS ===" "INFO"
    $folders = @("App", "Source", "Output", "SourceFiles")
    
    foreach ($folder in $folders) {
        $path = Join-Path $BasePath $folder
        try {
            if (!(Test-Path $path)) {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
                Write-LogMessage "Created folder: $path" "SUCCESS"
            } else {
                Write-LogMessage "Folder already exists: $path" "INFO"
            }
        } catch {
            Write-LogMessage "Failed to create folder $path`: $($_.Exception.Message)" "ERROR"
            return $false
        }
    }
    return $true
}

function Get-IntuneWinAppUtil {
    param(
        [string]$BasePath,
        [switch]$ForceDownload,
        [int]$MaxFileAgeDays
    )
    
    Write-LogMessage "=== ACQUIRING INTUNEWINAPPUTIL.EXE ===" "INFO"
    
    $intuneToolPath = Join-Path $BasePath "App\IntuneWinAppUtil.exe"
    
    # Check if file already exists
    if (Test-Path $intuneToolPath) {
        $fileInfo = Get-Item $intuneToolPath
        $fileAge = (Get-Date) - $fileInfo.LastWriteTime
        $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        
        Write-LogMessage "Found existing IntuneWinAppUtil.exe:" "INFO"
        Write-LogMessage "  File size: $fileSizeMB MB" "INFO"
        Write-LogMessage "  Created: $($fileInfo.CreationTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
        Write-LogMessage "  Modified: $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
        Write-LogMessage "  Age: $($fileAge.Days) days old" "INFO"
        
        # Validate the existing file
        $isValid = $false
        
        try {
            # Check if it's a valid executable
            if ($fileInfo.Length -gt 500KB) {
                $fileHeader = [System.IO.File]::ReadAllBytes($intuneToolPath) | Select-Object -First 2
                $isExecutable = ($fileHeader[0] -eq 0x4D -and $fileHeader[1] -eq 0x5A) # MZ header
                
                if ($isExecutable) {
                    # Try to get version info
                    try {
                        $versionInfo = $fileInfo.VersionInfo
                        if ($versionInfo.ProductName -like "*Intune*" -or $versionInfo.ProductName -like "*Content*") {
                            Write-LogMessage "  Product: $($versionInfo.ProductName)" "INFO"
                            Write-LogMessage "  Version: $($versionInfo.FileVersion)" "INFO"
                            $isValid = $true
                        } else {
                            Write-LogMessage "  Product name doesn't match expected Intune tool" "WARNING"
                        }
                    } catch {
                        # If we can't read version info, but it's an executable of reasonable size, assume it's valid
                        Write-LogMessage "  Cannot read version info, but appears to be valid executable" "INFO"
                        $isValid = $true
                    }
                } else {
                    Write-LogMessage "  File is not a valid executable" "WARNING"
                }
            } else {
                Write-LogMessage "  File is too small ($fileSizeMB MB)" "WARNING"
            }
        } catch {
            Write-LogMessage "  Error validating file: $($_.Exception.Message)" "WARNING"
        }
        
        # Decide whether to use existing file or download new one
        if ($ForceDownload) {
            Write-LogMessage "FORCED DOWNLOAD: -ForceDownload parameter specified, downloading fresh copy..." "WARNING"
            try {
                Remove-Item $intuneToolPath -Force
                Write-LogMessage "Removed existing file for forced download" "INFO"
            } catch {
                Write-LogMessage "Could not remove existing file: $($_.Exception.Message)" "ERROR"
                return $null
            }
        } elseif (!$isValid) {
            Write-LogMessage "INVALID FILE: Existing file appears corrupted or invalid, downloading fresh copy..." "WARNING"
            try {
                Remove-Item $intuneToolPath -Force
                Write-LogMessage "Removed invalid file" "INFO"
            } catch {
                Write-LogMessage "Could not remove invalid file: $($_.Exception.Message)" "ERROR"
            }
        } elseif ($fileAge.Days -gt $MaxFileAgeDays) {
            Write-LogMessage "OLD FILE: File is $($fileAge.Days) days old (max allowed: $MaxFileAgeDays days)" "WARNING"
            Write-LogMessage "Downloading fresh copy to ensure latest version..." "INFO"
            try {
                Remove-Item $intuneToolPath -Force
                Write-LogMessage "Removed old file" "INFO"
            } catch {
                Write-LogMessage "Could not remove old file: $($_.Exception.Message)" "WARNING"
                Write-LogMessage "Using existing file despite age" "INFO"
                return $intuneToolPath
            }
        } else {
            Write-LogMessage "USING EXISTING FILE: File is valid and recent (less than $MaxFileAgeDays days old)" "SUCCESS"
            Write-LogMessage "Skipping download process" "SUCCESS"
            Write-LogMessage "Location: $intuneToolPath" "INFO"
            
            # Offer option to force download if user wants
            Write-Host ""
            Write-Host "Found valid IntuneWinAppUtil.exe ($fileSizeMB MB, $($fileAge.Days) days old)" -ForegroundColor Green
            Write-Host "Do you want to download a fresh copy anyway? (y/N)" -ForegroundColor Yellow -NoNewline
            $userChoice = Read-Host " "
            
            if ($userChoice -eq "Y" -or $userChoice -eq "y") {
                Write-LogMessage "User requested fresh download" "INFO"
                try {
                    Remove-Item $intuneToolPath -Force
                    Write-LogMessage "Removed existing file for user-requested download" "INFO"
                } catch {
                    Write-LogMessage "Could not remove existing file: $($_.Exception.Message)" "ERROR"
                    return $intuneToolPath
                }
            } else {
                Write-LogMessage "User chose to keep existing file" "INFO"
                return $intuneToolPath
            }
        }
    } else {
        Write-LogMessage "IntuneWinAppUtil.exe not found, will download fresh copy" "INFO"
    }
    
    # Multiple URL sources for better reliability
    $downloadSources = @(
        @{
            Name = "GitHub Releases (Primary)"
            Url = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest/download/IntuneWinAppUtil.exe"
            Description = "Direct download from GitHub releases"
        },
        @{
            Name = "Microsoft aka.ms (Fallback)"
            Url = "https://aka.ms/intunewinapputil"
            Description = "Official Microsoft redirect URL"
        }
    )
    
    # Try each download source
    foreach ($source in $downloadSources) {
        Write-LogMessage "Trying $($source.Name): $($source.Description)" "INFO"
        
        # Method 1: Standard Invoke-WebRequest
        try {
            Write-LogMessage "Downloading from: $($source.Url)" "INFO"
            $progressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            
            Invoke-WebRequest -Uri $source.Url -OutFile $intuneToolPath -UseBasicParsing -TimeoutSec 120 -UserAgent "Mozilla/5.0"
            
            $ProgressPreference = $progressPreference
            
            # Verify download
            if (Test-Path $intuneToolPath) {
                $fileInfo = Get-Item $intuneToolPath
                $downloadedSize = [math]::Round($fileInfo.Length / 1MB, 2)
                
                # Check if it's actually an executable
                if ($fileInfo.Length -gt 10) {
                    $fileHeader = [System.IO.File]::ReadAllBytes($intuneToolPath) | Select-Object -First 2
                    $isExecutable = ($fileHeader[0] -eq 0x4D -and $fileHeader[1] -eq 0x5A) # MZ header
                } else {
                    $isExecutable = $false
                }
                
                if ($fileInfo.Length -gt 500KB -and $isExecutable) {
                    Write-LogMessage "Download completed successfully from $($source.Name) ($downloadedSize MB)" "SUCCESS"
                    
                    # Try to get version info
                    try {
                        $versionInfo = (Get-Item $intuneToolPath).VersionInfo
                        if ($versionInfo.ProductName) {
                            Write-LogMessage "Product: $($versionInfo.ProductName) v$($versionInfo.FileVersion)" "INFO"
                        }
                    } catch {
                        Write-LogMessage "Could not read version info (file may still be valid)" "INFO"
                    }
                    
                    return $intuneToolPath
                } else {
                    Write-LogMessage "Downloaded file appears invalid from $($source.Name) (Size: $downloadedSize MB, IsExe: $isExecutable)" "WARNING"
                    
                    # Show first 100 chars if it's not an executable (might be HTML error page)
                    if (!$isExecutable -and $fileInfo.Length -lt 10KB) {
                        try {
                            $content = Get-Content $intuneToolPath -Raw -ErrorAction SilentlyContinue
                            if ($content -and $content.Length -gt 0) {
                                $preview = $content.Substring(0, [Math]::Min(200, $content.Length))
                                Write-LogMessage "File content preview: $preview" "WARNING"
                            }
                        } catch { }
                    }
                    
                    Remove-Item $intuneToolPath -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-LogMessage "Failed to download from $($source.Name): $($_.Exception.Message)" "WARNING"
        }
        
        # Method 2: WebClient (fallback for same source)
        if (!(Test-Path $intuneToolPath)) {
            try {
                Write-LogMessage "Trying WebClient method for $($source.Name)..." "INFO"
                $webClient = New-Object System.Net.WebClient
                $webClient.Headers.Add("User-Agent", "Mozilla/5.0")
                $webClient.DownloadFile($source.Url, $intuneToolPath)
                $webClient.Dispose()
                
                if (Test-Path $intuneToolPath) {
                    $fileInfo = Get-Item $intuneToolPath
                    $downloadedSize = [math]::Round($fileInfo.Length / 1MB, 2)
                    
                    if ($fileInfo.Length -gt 500KB) {
                        Write-LogMessage "Download completed with WebClient from $($source.Name) ($downloadedSize MB)" "SUCCESS"
                        return $intuneToolPath
                    } else {
                        Write-LogMessage "Downloaded file too small with WebClient from $($source.Name)" "WARNING"
                        Remove-Item $intuneToolPath -Force -ErrorAction SilentlyContinue
                    }
                }
            } catch {
                Write-LogMessage "WebClient method failed for $($source.Name): $($_.Exception.Message)" "WARNING"
            }
        }
    }
    
    # Try downloading from GitHub ZIP as last resort
    try {
        Write-LogMessage "Attempting to extract from GitHub source ZIP..." "INFO"
        $githubZipUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/archive/refs/heads/master.zip"
        $tempZip = Join-Path $env:TEMP "IntuneWinAppUtil_source.zip"
        $tempExtract = Join-Path $env:TEMP "IntuneWinAppUtil_extract"
        
        # Download ZIP
        Invoke-WebRequest -Uri $githubZipUrl -OutFile $tempZip -UseBasicParsing -TimeoutSec 60
        
        if (Test-Path $tempZip) {
            # Extract ZIP
            if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
            Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
            
            # Look for IntuneWinAppUtil.exe in extracted files
            $exeFiles = Get-ChildItem -Path $tempExtract -Filter "IntuneWinAppUtil.exe" -Recurse -ErrorAction SilentlyContinue
            
            if ($exeFiles.Count -gt 0) {
                $sourceExe = $exeFiles[0].FullName
                Copy-Item $sourceExe $intuneToolPath -Force
                
                if (Test-Path $intuneToolPath) {
                    $fileInfo = Get-Item $intuneToolPath
                    $downloadedSize = [math]::Round($fileInfo.Length / 1MB, 2)
                    Write-LogMessage "Extracted IntuneWinAppUtil.exe from GitHub source ($downloadedSize MB)" "SUCCESS"
                    
                    # Cleanup
                    Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
                    Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
                    
                    return $intuneToolPath
                }
            } else {
                Write-LogMessage "IntuneWinAppUtil.exe not found in GitHub source ZIP" "WARNING"
            }
            
            # Cleanup
            Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
            Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-LogMessage "GitHub ZIP extraction failed: $($_.Exception.Message)" "WARNING"
    }
    
    # All methods failed - offer manual download
    Write-LogMessage "All automatic download methods failed" "ERROR"
    Write-LogMessage "MANUAL DOWNLOAD OPTIONS:" "ERROR"
    Write-LogMessage "Option 1 - Direct GitHub download:" "ERROR"
    Write-LogMessage "  Go to: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest" "ERROR"
    Write-LogMessage "  Download: IntuneWinAppUtil.exe" "ERROR"
    Write-LogMessage "Option 2 - Microsoft redirect:" "ERROR"
    Write-LogMessage "  Go to: https://aka.ms/intunewinapputil" "ERROR"
    Write-LogMessage "  Save file as: $intuneToolPath" "ERROR"
    
    # Check if user wants to wait for manual download
    Write-Host ""
    Write-Host "Would you like to pause so you can download manually? (Y/N)" -ForegroundColor Yellow -NoNewline
    $userChoice = Read-Host " "
    
    if ($userChoice -eq "Y" -or $userChoice -eq "y") {
        Write-Host ""
        Write-Host "Please download IntuneWinAppUtil.exe manually:" -ForegroundColor Cyan
        Write-Host "Best option: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/releases/latest" -ForegroundColor White
        Write-Host "Save to: $intuneToolPath" -ForegroundColor White
        Write-Host ""
        Read-Host "Press Enter after you've downloaded the file"
        
        # Check if user downloaded it
        if (Test-Path $intuneToolPath) {
            $fileInfo = Get-Item $intuneToolPath
            if ($fileInfo.Length -gt 500KB) {
                Write-LogMessage "Manual download detected and validated" "SUCCESS"
                return $intuneToolPath
            } else {
                Write-LogMessage "Manual download detected but file appears too small" "ERROR"
            }
        } else {
            Write-LogMessage "Manual download not detected" "ERROR"
        }
    }
    
    return $null
}

function Get-AutoDeskFolders {
    param([string]$AutoDeskTempPath)
    
    Write-LogMessage "=== SCANNING FOR AUTODESK APPLICATIONS ===" "INFO"
    
    try {
        # Enhanced pattern matching for AutoDesk applications
        $patterns = @("Revit*", "Auto*", "*CAD*", "Maya*", "3dsMax*", "Inventor*", "Fusion*")
        $tempFolders = @()
        
        foreach ($pattern in $patterns) {
            $foundFolders = Get-ChildItem $AutoDeskTempPath -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $pattern }
            if ($foundFolders) {
                $tempFolders += $foundFolders
                Write-LogMessage "Found $($foundFolders.Count) folder(s) matching pattern '$pattern'" "INFO"
            }
        }
        
        # Remove duplicates
        $tempFolders = $tempFolders | Sort-Object Name -Unique
        
        if ($tempFolders.Count -eq 0) {
            Write-LogMessage "No matching AutoDesk folders found in $AutoDeskTempPath" "WARNING"
            
            # Show what folders ARE available
            $allFolders = Get-ChildItem $AutoDeskTempPath -Directory -ErrorAction SilentlyContinue
            if ($allFolders.Count -gt 0) {
                Write-LogMessage "Available folders in $AutoDeskTempPath`:" "INFO"
                $allFolders | ForEach-Object { Write-LogMessage "  - $($_.Name)" "INFO" }
            } else {
                Write-LogMessage "No folders found in $AutoDeskTempPath" "WARNING"
            }
            return $null
        }
        
        Write-LogMessage "Found $($tempFolders.Count) AutoDesk application folder(s)" "SUCCESS"
        return $tempFolders
    } catch {
        Write-LogMessage "Error accessing AutoDesk temp folder: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Select-AutoDeskFolder {
    param([array]$Folders)
    
    Write-LogMessage "=== APPLICATION SELECTION ===" "INFO"
    Write-Host ""
    Write-Host "Available AutoDesk applications:" -ForegroundColor Cyan
    Write-Host "=================================" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $Folders.Count; $i++) {
        try {
            $folderInfo = Get-ChildItem $Folders[$i].FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
            $fileCount = $folderInfo.Count
            $sizeMB = if ($folderInfo.Sum) { [math]::Round($folderInfo.Sum / 1MB, 2) } else { 0 }
            $createdDate = $Folders[$i].CreationTime.ToString("yyyy-MM-dd HH:mm")
            
            Write-Host "$($i+1): $($Folders[$i].Name)" -ForegroundColor White
            Write-Host "     Files: $fileCount | Size: $sizeMB MB | Created: $createdDate" -ForegroundColor Gray
        } catch {
            Write-Host "$($i+1): $($Folders[$i].Name)" -ForegroundColor White
            Write-Host "     (Unable to read folder details)" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
    do {
        $selection = Read-Host "Select application number (1-$($Folders.Count))"
        $selectionInt = $selection -as [int]
        if (!$selectionInt -or $selectionInt -lt 1 -or $selectionInt -gt $Folders.Count) {
            Write-Host "Invalid selection. Please enter a number between 1 and $($Folders.Count)." -ForegroundColor Red
        }
    } while (!$selectionInt -or $selectionInt -lt 1 -or $selectionInt -gt $Folders.Count)
    
    $selectedFolder = $Folders[$selectionInt - 1]
    Write-LogMessage "Selected application: $($selectedFolder.Name)" "SUCCESS"
    return $selectedFolder
}

function Test-AutoDeskPackage {
    param([string]$SourcePath)
    
    Write-LogMessage "=== VALIDATING AUTODESK PACKAGE ===" "INFO"
    
    # Define required files with their purposes
    $requiredFiles = @(
        @{Path = "Summary.txt"; Description = "Application summary and metadata"; Critical = $true},
        @{Path = "image\Installer.exe"; Description = "AutoDesk installer executable"; Critical = $true},
        @{Path = "image\Collection.xml"; Description = "Installation configuration"; Critical = $true}
    )
    
    $missingCritical = @()
    $missingOptional = @()
    
    foreach ($file in $requiredFiles) {
        $fullPath = Join-Path $SourcePath $file.Path
        if (!(Test-Path $fullPath)) {
            if ($file.Critical) {
                $missingCritical += "$($file.Description) ($($file.Path))"
                Write-LogMessage "CRITICAL FILE MISSING: $($file.Path)" "ERROR"
            } else {
                $missingOptional += "$($file.Description) ($($file.Path))"
                Write-LogMessage "Optional file missing: $($file.Path)" "WARNING"
            }
        } else {
            $fileSize = [math]::Round((Get-Item $fullPath).Length / 1KB, 2)
            Write-LogMessage "Found: $($file.Path) ($fileSize KB)" "SUCCESS"
        }
    }
    
    if ($missingCritical.Count -gt 0) {
        Write-LogMessage "Package validation FAILED - critical files missing:" "ERROR"
        $missingCritical | ForEach-Object { Write-LogMessage "  - $_" "ERROR" }
        Write-LogMessage "This folder does not contain a valid AutoDesk deployment package." "ERROR"
        return $false
    }
    
    if ($missingOptional.Count -gt 0) {
        Write-LogMessage "Package validation completed with warnings:" "WARNING"
        $missingOptional | ForEach-Object { Write-LogMessage "  - $_" "WARNING" }
    } else {
        Write-LogMessage "Package validation completed successfully - all files present" "SUCCESS"
    }
    
    return $true
}

function Get-SummaryInfo {
    param([string]$SourcePath)
    
    Write-LogMessage "=== PARSING APPLICATION INFORMATION ===" "INFO"
    $summaryPath = Join-Path $SourcePath "Summary.txt"
    
    try {
        $summaryContent = Get-Content $summaryPath -Raw -ErrorAction Stop
        Write-LogMessage "Successfully loaded Summary.txt ($($summaryContent.Length) characters)" "SUCCESS"
    } catch {
        Write-LogMessage "Failed to read Summary.txt: $($_.Exception.Message)" "ERROR"
        return $null
    }
    
    # Initialize variables with defaults
    $programName = "Unknown"
    $buildNumber = "0.0.0.0"
    $productCode = "Unknown"
    $installerVersion = "0.0.0.0"
    
    # Enhanced parsing for actual product name (not deployment name)
    Write-LogMessage "Parsing actual product name..." "INFO"
    
    # Look for the actual product name in various sections
    $productNamePatterns = @(
        "(?ms)^([^:\r\n]+)\s*[\r\n]+.*?Product Code:\s*\{[A-F0-9\-]+\}",  # Product name above Product Code
        "(?ms)^([^:\r\n]+)\s*[\r\n]+.*?Build number:\s*[\d\.]+",          # Product name above Build number
        "(?ms)^Deployment:\s*(.+?)\s*[\r\n]+.*?^([^:\r\n]+)\s*[\r\n]+.*?Product Code:",  # Second line after Deployment
        "(?ms)^Product:\s*(.+?)\s*$",                                      # Direct Product field
        "(?ms)^Application:\s*(.+?)\s*$"                                   # Direct Application field
    )
    
    foreach ($pattern in $productNamePatterns) {
        if ($summaryContent -match $pattern) {
            $candidateName = $matches[1].Trim()
            
            # Skip if it looks like a deployment name (contains underscores and version numbers)
            if ($candidateName -notmatch "_\d+\.\d+" -and $candidateName -ne "Deployment" -and $candidateName.Length -gt 3) {
                $programName = $candidateName
                Write-LogMessage "Found product name: $programName" "SUCCESS"
                break
            } elseif ($matches.Count -gt 2 -and $matches[2]) {
                # Try the second match if available
                $candidateName = $matches[2].Trim()
                if ($candidateName -notmatch "_\d+\.\d+" -and $candidateName -ne "Deployment" -and $candidateName.Length -gt 3) {
                    $programName = $candidateName
                    Write-LogMessage "Found product name (second match): $programName" "SUCCESS"
                    break
                }
            }
        }
    }
    
    # If still not found, try to extract from lines containing common AutoDesk product names
    if ($programName -eq "Unknown") {
        Write-LogMessage "Trying alternative product name detection..." "INFO"
        $lines = $summaryContent -split "`n"
        foreach ($line in $lines) {
            $line = $line.Trim()
            # Look for lines containing known AutoDesk products
            if ($line -match "(Revit|AutoCAD|Maya|3ds Max|Inventor|Fusion|Civil 3D|Plant 3D)\s*\d{4}" -and $line -notmatch "_") {
                $programName = $line
                Write-LogMessage "Found product name from line scan: $programName" "SUCCESS"
                break
            }
        }
    }
    
    # Last resort: Use deployment name but clean it up
    if ($programName -eq "Unknown") {
        Write-LogMessage "Using deployment name as fallback..." "WARNING"
        if ($summaryContent -match "(?ms)^Deployment:\s*(.+?)\s*$") {
            $deploymentName = $matches[1].Trim()
            # Try to clean up deployment name (remove version suffixes)
            if ($deploymentName -match "^([^_]+)_\d+") {
                $programName = $matches[1] -replace "_", " "
                Write-LogMessage "Cleaned deployment name: $programName" "INFO"
            } else {
                $programName = $deploymentName
            }
        }
    }
    
    # Parse build number with validation
    Write-LogMessage "Parsing build number..." "INFO"
    if ($summaryContent -match "(?ms)Build number:\s*([\d\.]+)") {
        $buildNumber = $matches[1].Trim()
        try {
            # Validate version format
            [version]$buildNumber | Out-Null
            Write-LogMessage "Found valid build number: $buildNumber" "SUCCESS"
        } catch {
            Write-LogMessage "Build number format may be invalid: $buildNumber" "WARNING"
            # Keep the original value anyway
        }
    } else {
        Write-LogMessage "Could not find build number in Summary.txt" "WARNING"
    }
    
    # Parse product code
    Write-LogMessage "Parsing product code..." "INFO"
    if ($summaryContent -match "(?ms)Product Code:\s*(\{[A-F0-9\-]+\})") {
        $productCode = $matches[1].Trim()
        Write-LogMessage "Found product code: $productCode" "SUCCESS"
    } else {
        Write-LogMessage "Could not find product code in Summary.txt" "WARNING"
    }
    
    # Parse installer version
    Write-LogMessage "Parsing installer version..." "INFO"
    if ($summaryContent -match "(?ms)Autodesk Installer[\r\n]+.*?Build number:\s*([\d\.]+)") {
        $installerVersion = $matches[1].Trim()
        Write-LogMessage "Found installer version: $installerVersion" "SUCCESS"
    } else {
        Write-LogMessage "Could not find installer version in Summary.txt" "WARNING"
    }
    
    # Create and return info object
    $summaryInfo = @{
        ProgramName = $programName
        BuildNumber = $buildNumber
        ProductCode = $productCode
        InstallerVersion = $installerVersion
        SourcePath = $SourcePath
    }
    
    # Display summary
    Write-LogMessage "=== PARSED INFORMATION SUMMARY ===" "INFO"
    Write-LogMessage "Product Name     : $programName" "INFO"
    Write-LogMessage "Build Number     : $buildNumber" "INFO"
    Write-LogMessage "Product Code     : $productCode" "INFO"
    Write-LogMessage "Installer Version: $installerVersion" "INFO"
    
    # Validate that we got the essential information
    if ($productCode -eq "Unknown" -or $buildNumber -eq "0.0.0.0") {
        Write-LogMessage "WARNING: Missing essential information for Intune deployment" "WARNING"
        Write-LogMessage "This may affect detection rule accuracy" "WARNING"
        
        # Show first 10 lines for debugging
        Write-LogMessage "First 10 lines of Summary.txt for debugging:" "INFO"
        $summaryLines = $summaryContent -split "`n" | Select-Object -First 10
        $summaryLines | ForEach-Object { Write-LogMessage "  $_" "INFO" }
    }
    
    return $summaryInfo
}

function New-ApplicationPackage {
    param(
        [object]$SelectedFolder,
        [string]$BasePath
    )
    
    Write-LogMessage "=== CREATING APPLICATION PACKAGE ===" "INFO"
    
    $sourcePath = $SelectedFolder.FullName
    $zipName = "$($SelectedFolder.Name).zip"
    $zipDestination = Join-Path "$BasePath\Source" $zipName
    
    Write-LogMessage "Source folder: $sourcePath" "INFO"
    Write-LogMessage "Package destination: $zipDestination" "INFO"
    
    # Remove existing zip if present
    if (Test-Path $zipDestination) {
        try {
            $existingSize = [math]::Round((Get-Item $zipDestination).Length / 1MB, 2)
            Remove-Item $zipDestination -Force
            Write-LogMessage "Removed existing package file ($existingSize MB)" "INFO"
        } catch {
            Write-LogMessage "Could not remove existing zip file: $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
    
    try {
        # Analyze source folder
        Write-LogMessage "Analyzing source folder..." "INFO"
        $sourceFiles = Get-ChildItem -Path $sourcePath -Recurse -File -ErrorAction Stop
        $fileCount = $sourceFiles.Count
        $totalSizeMB = [math]::Round(($sourceFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 2)
        $totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)
        
        Write-LogMessage "Found $fileCount files totaling $totalSizeMB MB ($totalSizeGB GB)" "INFO"
        
        # Check for very large packages and offer alternatives
        if ($totalSizeGB -gt 8) {
            Write-LogMessage "EXTREMELY LARGE PACKAGE DETECTED ($totalSizeGB GB)" "WARNING"
            Write-LogMessage "This may cause compression issues and deployment problems" "WARNING"
            
            Write-Host ""
            Write-Host "WARNING: Very large AutoDesk package detected ($totalSizeGB GB)" -ForegroundColor Red
            Write-Host "This may cause issues with:" -ForegroundColor Yellow
            Write-Host "  - ZIP compression (may fail or take hours)" -ForegroundColor Yellow  
            Write-Host "  - Intune upload limits" -ForegroundColor Yellow
            Write-Host "  - Client download times" -ForegroundColor Yellow
            Write-Host "  - Disk space requirements" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Recommended alternatives:" -ForegroundColor Cyan
            Write-Host "  1. Use direct installer deployment (skip packaging)" -ForegroundColor White
            Write-Host "  2. Create smaller deployment packages" -ForegroundColor White
            Write-Host "  3. Use network installation source" -ForegroundColor White
            Write-Host ""
            Write-Host "Do you want to continue with packaging anyway? (y/N)" -ForegroundColor Yellow -NoNewline
            $userChoice = Read-Host " "
            
            if ($userChoice -ne "Y" -and $userChoice -ne "y") {
                Write-LogMessage "User cancelled large package creation" "INFO"
                return $null
            }
            
            Write-LogMessage "User chose to proceed with large package creation" "WARNING"
        } elseif ($totalSizeGB -gt 4) {
            Write-LogMessage "Large package detected ($totalSizeGB GB) - using enhanced compression method" "WARNING"
        }
        
        # Choose compression method based on size
        $compressionMethod = "Standard"
        if ($totalSizeGB -gt 4) {
            $compressionMethod = "Enhanced"
        }
        
        Write-LogMessage "Using $compressionMethod compression method for $totalSizeGB GB package" "INFO"
        
        # Create the package with appropriate method
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        if ($compressionMethod -eq "Enhanced") {
            # Use alternative compression for large files
            Write-LogMessage "Starting enhanced compression (this may take 10-30 minutes)..." "INFO"
            
            # Try System.IO.Compression for better large file handling
            try {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                
                # Create zip with better memory management
                $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
                
                Write-LogMessage "Creating ZIP archive with System.IO.Compression..." "INFO"
                [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcePath, $zipDestination, $compressionLevel, $false)
                
                Write-LogMessage "Enhanced compression completed" "SUCCESS"
                
            } catch {
                Write-LogMessage "Enhanced compression failed: $($_.Exception.Message)" "WARNING"
                Write-LogMessage "Falling back to chunk-based compression..." "INFO"
                
                # Fallback: Create smaller chunks and combine
                $success = New-ChunkedPackage -SourcePath $sourcePath -DestinationPath $zipDestination
                if (!$success) {
                    throw "All compression methods failed for large package"
                }
            }
        } else {
            # Standard compression for smaller packages
            Write-LogMessage "Starting standard compression..." "INFO"
            Compress-Archive -Path "$sourcePath\*" -DestinationPath $zipDestination -CompressionLevel Optimal -Force
        }
        
        $stopwatch.Stop()
        
        # Validate the created package
        if (Test-Path $zipDestination) {
            $packageSize = [math]::Round((Get-Item $zipDestination).Length / 1MB, 2)
            $packageSizeGB = [math]::Round($packageSize / 1024, 2)
            $compressionRatio = [math]::Round((1 - $packageSize/$totalSizeMB) * 100, 1)
            $timeElapsed = $stopwatch.Elapsed.TotalMinutes.ToString('F1')
            
            Write-LogMessage "Package created successfully!" "SUCCESS"
            Write-LogMessage "Compression time: $timeElapsed minutes" "SUCCESS"
            Write-LogMessage "Original size: $totalSizeMB MB ($totalSizeGB GB)" "INFO"
            Write-LogMessage "Package size: $packageSize MB ($packageSizeGB GB)" "INFO"
            Write-LogMessage "Compression ratio: $compressionRatio%" "INFO"
            
            # Warning for very large final packages
            if ($packageSizeGB -gt 8) {
                Write-LogMessage "WARNING: Final package is very large ($packageSizeGB GB)" "WARNING"
                Write-LogMessage "This may exceed Intune upload limits or cause deployment issues" "WARNING"
            }
            
            return $zipDestination
        } else {
            throw "Package file was not created"
        }
    } catch {
        Write-LogMessage "Failed to create package: $($_.Exception.Message)" "ERROR"
        
        # Provide specific guidance for large file issues
        if ($_.Exception.Message -like "*data stream*" -or $_.Exception.Message -like "*too long*" -or $_.Exception.Message -like "*Dataströmmen*") {
            Write-LogMessage "This appears to be a large file compression issue" "ERROR"
            Write-LogMessage "Suggestions:" "ERROR"
            Write-LogMessage "1. Try reducing package size by removing unnecessary files" "ERROR"
            Write-LogMessage "2. Use alternative deployment method (direct installer)" "ERROR"
            Write-LogMessage "3. Split into multiple smaller packages" "ERROR"
            Write-LogMessage "4. Use network installation source instead of packaging" "ERROR"
        }
        
        return $null
    }
}

function New-ChunkedPackage {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )
    
    try {
        Write-LogMessage "Attempting chunked compression method..." "INFO"
        
        # Create temporary directory for chunks
        $tempDir = Join-Path $env:TEMP "AutoDesk_Chunks_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Get all files and split into chunks
        $allFiles = Get-ChildItem -Path $SourcePath -Recurse -File
        $chunkSize = 1000  # Files per chunk
        $chunks = @()
        
        for ($i = 0; $i -lt $allFiles.Count; $i += $chunkSize) {
            $chunkFiles = $allFiles | Select-Object -Skip $i -First $chunkSize
            $chunkNumber = [math]::Floor($i / $chunkSize) + 1
            $chunkName = "chunk_$($chunkNumber.ToString('D3')).zip"
            $chunkPath = Join-Path $tempDir $chunkName
            
            Write-LogMessage "Creating chunk $chunkNumber with $($chunkFiles.Count) files..." "INFO"
            
            # Create individual chunk
            $chunkFiles | ForEach-Object {
                $relativePath = $_.FullName.Substring($SourcePath.Length + 1)
                Compress-Archive -Path $_.FullName -DestinationPath $chunkPath -Update
            }
            
            $chunks += $chunkPath
        }
        
        # Combine all chunks into final package
        Write-LogMessage "Combining $($chunks.Count) chunks into final package..." "INFO"
        
        if ($chunks.Count -eq 1) {
            # Only one chunk, just move it
            Move-Item $chunks[0] $DestinationPath -Force
        } else {
            # Multiple chunks - create container ZIP
            Compress-Archive -Path "$tempDir\*" -DestinationPath $DestinationPath -CompressionLevel Optimal
        }
        
        # Cleanup
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-LogMessage "Chunked compression completed successfully" "SUCCESS"
        return $true
        
    } catch {
        Write-LogMessage "Chunked compression failed: $($_.Exception.Message)" "ERROR"
        
        # Cleanup on failure
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        return $false
    }
}

function New-InstallationScript {
    param(
        [object]$SelectedFolder,
        [hashtable]$SummaryInfo,
        [string]$BasePath
    )
    
    Write-LogMessage "=== GENERATING INSTALLATION SCRIPT ===" "INFO"
    
    $zipName = "$($SelectedFolder.Name).zip"
    $scriptName = "$($SelectedFolder.Name)_Install.ps1"
    $scriptPath = Join-Path "$BasePath\Source" $scriptName
    
    $installScriptContent = @"
# AutoDesk Installation Script
# Generated by AutoDesk Package Tool v2.0
# Application: $($SummaryInfo.ProgramName)
# Build: $($SummaryInfo.BuildNumber)
# Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

param(
    [Parameter(Mandatory=`$true)]
    [ValidateSet("Install", "Uninstall")]
    [string]`$Action
)

# Configuration
`$programName = "$($SummaryInfo.ProgramName)"
`$buildNumber = "$($SummaryInfo.BuildNumber)"
`$productCode = "$($SummaryInfo.ProductCode)"
`$installerVersion = "$($SummaryInfo.InstallerVersion)"

`$zipPath = "`$PSScriptRoot\$zipName"
`$extractRoot = Join-Path `$env:windir "Temp"
`$extractPath = Join-Path `$extractRoot "$($SelectedFolder.Name)"
`$installerPath = Join-Path `$extractPath "image\Installer.exe"
`$collectionXmlPath = Join-Path `$extractPath "image\Collection.xml"

# Logging function
function Write-InstallLog {
    param([string]`$Message, [string]`$Level = "INFO")
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    `$logEntry = "[`$timestamp] `$Level`: `$Message"
    Write-Host `$logEntry
    # Add to Windows Event Log if possible
    try {
        Write-EventLog -LogName Application -Source "AutoDesk Installer" -EventId 1000 -Message `$logEntry -ErrorAction SilentlyContinue
    } catch { }
}

function Test-AdminRights {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

function Get-InstalledVersion {
    `$registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach (`$path in `$registryPaths) {
        try {
            `$apps = Get-ItemProperty -Path `$path -ErrorAction SilentlyContinue
            foreach (`$app in `$apps) {
                if (`$app.DisplayName -like "*`$programName*") {
                    Write-InstallLog "Found installed version: `$(`$app.DisplayVersion)" "INFO"
                    return `$app.DisplayVersion
                }
            }
        } catch {
            # Continue searching
        }
    }
    Write-InstallLog "No existing installation found" "INFO"
    return `$null
}

function Install-Application {
    Write-InstallLog "Starting installation of `$programName" "INFO"
    
    if (!(Test-AdminRights)) {
        Write-InstallLog "Administrator rights required for installation" "ERROR"
        return 1
    }
    
    # Check current installation
    `$installedVersion = Get-InstalledVersion
    if (`$installedVersion) {
        try {
            if ([version]`$installedVersion -ge [version]`$buildNumber) {
                Write-InstallLog "`$programName is already up to date (installed: `$installedVersion, package: `$buildNumber)" "INFO"
                return 0
            } else {
                Write-InstallLog "Upgrading from version `$installedVersion to `$buildNumber" "INFO"
            }
        } catch {
            Write-InstallLog "Version comparison failed, proceeding with installation" "WARNING"
        }
    }
    
    # Validate package file
    if (!(Test-Path `$zipPath)) {
        Write-InstallLog "Package file not found: `$zipPath" "ERROR"
        return 1
    }
    
    `$packageSize = [math]::Round((Get-Item `$zipPath).Length / 1MB, 2)
    Write-InstallLog "Package file found (`$packageSize MB)" "INFO"
    
    try {
        # Clean up any existing extraction
        if (Test-Path `$extractPath) {
            Write-InstallLog "Cleaning up previous extraction..." "INFO"
            Remove-Item `$extractPath -Recurse -Force
        }
        
        # Extract package
        Write-InstallLog "Extracting package to `$extractPath..." "INFO"
        Expand-Archive -Path `$zipPath -DestinationPath `$extractPath -Force
        
        # Validate extracted files
        if (!(Test-Path `$installerPath)) {
            throw "Installer.exe not found after extraction"
        }
        if (!(Test-Path `$collectionXmlPath)) {
            throw "Collection.xml not found after extraction"
        }
        
        Write-InstallLog "Package extracted successfully" "SUCCESS"
        
        # Run installation
        `$arguments = "-i deploy --offline_mode -q -o `"`$collectionXmlPath`" --installer_version `"`$installerVersion`""
        Write-InstallLog "Running installer with arguments: `$arguments" "INFO"
        
        `$process = Start-Process -FilePath `$installerPath -ArgumentList `$arguments -Wait -PassThru -WindowStyle Hidden
        
        if (`$process.ExitCode -eq 0) {
            Write-InstallLog "Installation completed successfully" "SUCCESS"
            
            # Verify installation
            `$newVersion = Get-InstalledVersion
            if (`$newVersion) {
                Write-InstallLog "Installation verified - version `$newVersion is now installed" "SUCCESS"
            }
            
            return 0
        } else {
            Write-InstallLog "Installation failed with exit code: `$(`$process.ExitCode)" "ERROR"
            return `$process.ExitCode
        }
        
    } catch {
        Write-InstallLog "Installation failed: `$(`$_.Exception.Message)" "ERROR"
        return 1
    } finally {
        # Clean up extraction folder
        try {
            if (Test-Path `$extractPath) {
                Remove-Item `$extractPath -Recurse -Force
                Write-InstallLog "Cleanup completed" "INFO"
            }
        } catch {
            Write-InstallLog "Cleanup failed: `$(`$_.Exception.Message)" "WARNING"
        }
    }
}

function Uninstall-Application {
    Write-InstallLog "Starting uninstallation of `$programName" "INFO"
    
    if (!(Test-AdminRights)) {
        Write-InstallLog "Administrator rights required for uninstallation" "ERROR"
        return 1
    }
    
    # Check if application is installed
    `$installedVersion = Get-InstalledVersion
    if (!`$installedVersion) {
        Write-InstallLog "`$programName is not installed" "INFO"
        return 0
    }
    
    try {
        Write-InstallLog "Uninstalling `$programName (Product Code: `$productCode)..." "INFO"
        `$process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x", `$productCode, "/quiet", "/norestart" -Wait -PassThru -WindowStyle Hidden
        
        if (`$process.ExitCode -eq 0) {
            Write-InstallLog "Uninstallation completed successfully" "SUCCESS"
            return 0
        } else {
            Write-InstallLog "Uninstallation failed with exit code: `$(`$process.ExitCode)" "ERROR"
            return `$process.ExitCode
        }
    } catch {
        Write-InstallLog "Uninstallation failed: `$(`$_.Exception.Message)" "ERROR"
        return 1
    }
}

# Main execution
try {
    Write-InstallLog "AutoDesk Installation Script starting - Action: `$Action" "INFO"
    Write-InstallLog "Target application: `$programName v`$buildNumber" "INFO"
    
    `$exitCode = switch (`$Action) {
        "Install"   { Install-Application }
        "Uninstall" { Uninstall-Application }
        default     { 
            Write-InstallLog "Invalid action: `$Action" "ERROR"
            1
        }
    }
    
    Write-InstallLog "Script completed with exit code: `$exitCode" "INFO"
    exit `$exitCode
    
} catch {
    Write-InstallLog "Script execution failed: `$(`$_.Exception.Message)" "ERROR"
    exit 1
}
"@

    try {
        $installScriptContent | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
        Write-LogMessage "Installation script saved: $scriptPath" "SUCCESS"
        return $scriptPath
    } catch {
        Write-LogMessage "Failed to create installation script: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function New-IntuneInfo {
    param(
        [object]$SelectedFolder,
        [hashtable]$SummaryInfo,
        [string]$BasePath
    )
    
    Write-LogMessage "=== GENERATING INTUNE DEPLOYMENT INFO ===" "INFO"
    
    $installCmd = "powershell.exe -ExecutionPolicy Bypass -File .\$($SelectedFolder.Name)_Install.ps1 -Action Install"
    $uninstallCmd = "powershell.exe -ExecutionPolicy Bypass -File .\$($SelectedFolder.Name)_Install.ps1 -Action Uninstall"
    
    $outputTextFile = Join-Path "$BasePath\Output" "$($SelectedFolder.Name)_IntuneInfo.txt"
    
    $intuneInfo = @"
=============================================================================
                        AUTODESK INTUNE DEPLOYMENT INFO
=============================================================================
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Package Tool Version: 2.0

APPLICATION INFORMATION:
-----------------------
Program Name        : $($SummaryInfo.ProgramName)
Build Number        : $($SummaryInfo.BuildNumber)
Product Code        : $($SummaryInfo.ProductCode)
Installer Version   : $($SummaryInfo.InstallerVersion)

INTUNE CONFIGURATION:
--------------------
Install Command     : $installCmd
Uninstall Command   : $uninstallCmd
Install Behavior    : System
Restart Behavior    : Determine behavior based on return codes
Return Codes        : 0=Success, 3010=Success (restart required)

DETECTION RULE (RECOMMENDED):
----------------------------
Detection Type      : MSI
Product Code        : $($SummaryInfo.ProductCode)
Product Version     : $($SummaryInfo.BuildNumber)

ALTERNATIVE DETECTION RULE:
---------------------------
Detection Type      : Registry
Registry Key        : HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$($SummaryInfo.ProductCode)
Registry Value      : DisplayVersion
Detection Method    : Version comparison
Operator           : Greater than or equal to
Value              : $($SummaryInfo.BuildNumber)

REQUIREMENTS:
------------
- Windows 10/11 (64-bit)
- Administrator privileges required
- Minimum 2GB free disk space
- .NET Framework 4.7.2 or higher

DEPLOYMENT NOTES:
----------------
1. This package uses PowerShell execution policy bypass
2. Installation runs in system context
3. Package includes automatic version checking
4. Supports upgrade scenarios
5. Includes comprehensive logging

INTUNE SETUP STEPS:
------------------
1. Upload the .intunewin file to Intune
2. Configure detection rule using MSI Product Code (recommended)
3. Set install/uninstall commands as shown above
4. Configure requirements as listed
5. Set install behavior to "System"
6. Configure assignments and deployment

TROUBLESHOOTING:
---------------
- Check Windows Event Log (Application) for detailed installation logs
- Verify PowerShell execution policy allows script execution
- Ensure target devices have sufficient disk space
- Validate network connectivity for license activation (if required)
- Use MSI Product Code detection for most reliable application detection

=============================================================================
"@

    try {
        $intuneInfo | Out-File -FilePath $outputTextFile -Encoding UTF8 -Force
        Write-LogMessage "Intune deployment info saved: $outputTextFile" "SUCCESS"
        return $outputTextFile
    } catch {
        Write-LogMessage "Failed to create Intune info file: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Invoke-IntunePackaging {
    param(
        [string]$IntuneToolPath,
        [string]$PackageZipPath,
        [string]$InstallScriptPath,
        [string]$ApplicationName,
        [string]$OutputFolder
    )
    
    Write-LogMessage "=== CREATING INTUNE WIN32 PACKAGE ===" "INFO"
    
    try {
        # Validate inputs
        if (!(Test-Path $IntuneToolPath)) {
            throw "IntuneWinAppUtil.exe not found: $IntuneToolPath"
        }
        if (!(Test-Path $PackageZipPath)) {
            throw "Package ZIP not found: $PackageZipPath"
        }
        if (!(Test-Path $InstallScriptPath)) {
            throw "Installation script not found: $InstallScriptPath"
        }
        
        # Create clean temporary folder for this specific application
        $tempPackageFolder = Join-Path $env:TEMP "IntunePackage_$ApplicationName_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        Write-LogMessage "Creating clean package folder: $tempPackageFolder" "INFO"
        
        if (Test-Path $tempPackageFolder) {
            Remove-Item $tempPackageFolder -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempPackageFolder -Force | Out-Null
        
        # Copy only the files needed for this specific application
        $zipName = Split-Path $PackageZipPath -Leaf
        $scriptName = Split-Path $InstallScriptPath -Leaf
        
        $tempZipPath = Join-Path $tempPackageFolder $zipName
        $tempScriptPath = Join-Path $tempPackageFolder $scriptName
        
        Write-LogMessage "Copying application-specific files to clean folder..." "INFO"
        Write-LogMessage "  ZIP package: $zipName" "INFO"
        Write-LogMessage "  Install script: $scriptName" "INFO"
        
        Copy-Item -Path $PackageZipPath -Destination $tempZipPath -Force
        Copy-Item -Path $InstallScriptPath -Destination $tempScriptPath -Force
        
        # Verify files were copied
        if (!(Test-Path $tempZipPath) -or !(Test-Path $tempScriptPath)) {
            throw "Failed to copy files to temporary package folder"
        }
        
        # Show what's being packaged
        $packageContents = Get-ChildItem -Path $tempPackageFolder
        Write-LogMessage "Package contents:" "INFO"
        foreach ($file in $packageContents) {
            $fileSize = [math]::Round($file.Length / 1MB, 2)
            Write-LogMessage "  - $($file.Name) ($fileSize MB)" "INFO"
        }
        
        # Build command arguments for IntuneWinAppUtil
        $arguments = @(
            "-c", "`"$tempPackageFolder`"",
            "-s", "`"$scriptName`"",
            "-o", "`"$OutputFolder`"",
            "-q"
        )
        
        $argumentString = $arguments -join " "
        Write-LogMessage "Executing IntuneWinAppUtil.exe..." "INFO"
        Write-LogMessage "Command: `"$IntuneToolPath`" $argumentString" "INFO"
        
        # Execute IntuneWinAppUtil
        $process = Start-Process -FilePath $IntuneToolPath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
        
        if ($process.ExitCode -eq 0) {
            # Find the generated .intunewin file
            $intunewinFiles = Get-ChildItem -Path $OutputFolder -Filter "*.intunewin" -ErrorAction SilentlyContinue | 
                             Where-Object { $_.Name -like "*$ApplicationName*" -or $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
            
            if ($intunewinFiles.Count -eq 0) {
                # Fallback: get the most recent .intunewin file
                $intunewinFiles = Get-ChildItem -Path $OutputFolder -Filter "*.intunewin" -ErrorAction SilentlyContinue | 
                                 Sort-Object LastWriteTime -Descending
            }
            
            if ($intunewinFiles.Count -gt 0) {
                $latestFile = $intunewinFiles[0]
                $fileSize = [math]::Round($latestFile.Length / 1MB, 2)
                Write-LogMessage "Intune package created successfully!" "SUCCESS"
                Write-LogMessage "Package name: $($latestFile.Name)" "SUCCESS"
                Write-LogMessage "Package size: $fileSize MB" "SUCCESS"
                Write-LogMessage "Package location: $($latestFile.FullName)" "INFO"
                
                # Show package summary
                Write-LogMessage "Package Summary:" "INFO"
                Write-LogMessage "  Source files: $($packageContents.Count) files" "INFO"
                Write-LogMessage "  Setup file: $scriptName" "INFO"
                Write-LogMessage "  Final package: $($latestFile.Name) ($fileSize MB)" "INFO"
                
                # Cleanup temporary folder
                try {
                    Remove-Item $tempPackageFolder -Recurse -Force
                    Write-LogMessage "Cleaned up temporary package folder" "INFO"
                } catch {
                    Write-LogMessage "Could not clean up temporary folder: $($_.Exception.Message)" "WARNING"
                }
                
                return $latestFile.FullName
            } else {
                throw "IntuneWinAppUtil completed but no .intunewin file was found in output folder"
            }
        } else {
            throw "IntuneWinAppUtil failed with exit code: $($process.ExitCode)"
        }
    } catch {
        Write-LogMessage "Intune packaging failed: $($_.Exception.Message)" "ERROR"
        
        # Cleanup on failure
        if (Test-Path $tempPackageFolder) {
            try {
                Remove-Item $tempPackageFolder -Recurse -Force
                Write-LogMessage "Cleaned up temporary folder after error" "INFO"
            } catch {
                Write-LogMessage "Could not clean up temporary folder: $($_.Exception.Message)" "WARNING"
            }
        }
        
        return $null
    }
}

#endregion

#region Main Script Execution

Write-Host ""
Write-Host "===================================" -ForegroundColor Cyan
Write-Host "   AutoDesk Package Tool v2.0      " -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

Write-LogMessage "AutoDesk Package Tool v2.0 starting..." "INFO"
Write-LogMessage "Log file: $Script:LogFile" "INFO"
Write-LogMessage "Parameters:" "INFO"
Write-LogMessage "  Base path: $BasePath" "INFO"
Write-LogMessage "  AutoDesk temp path: $AutoDeskTempPath" "INFO"
Write-LogMessage "  Force download: $ForceDownload" "INFO"
Write-LogMessage "  Max file age: $MaxFileAgeDays days" "INFO"

try {
    # Step 1: Check prerequisites
    if (!(Test-Prerequisites)) {
        Write-Host ""
        Write-Host "Prerequisites check failed. Please resolve the issues above and try again." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 2: Create required folders
    if (!(New-RequiredFolders -BasePath $BasePath)) {
        Write-Host ""
        Write-Host "Failed to create required folders. Please check permissions and try again." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 3: Get IntuneWinAppUtil
    $intuneToolPath = Get-IntuneWinAppUtil -BasePath $BasePath -ForceDownload:$ForceDownload -MaxFileAgeDays $MaxFileAgeDays
    if (!$intuneToolPath) {
        Write-Host ""
        Write-Host "Failed to acquire IntuneWinAppUtil.exe. Please check your internet connection and try again." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 4: Find AutoDesk folders
    $autoDeskFolders = Get-AutoDeskFolders -AutoDeskTempPath $AutoDeskTempPath
    if (!$autoDeskFolders) {
        Write-Host ""
        Write-Host "No AutoDesk application folders found. Please ensure your deployment completed successfully." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 5: Select application
    $selectedFolder = Select-AutoDeskFolder -Folders $autoDeskFolders
    
    # Step 6: Validate the package
    if (!(Test-AutoDeskPackage -SourcePath $selectedFolder.FullName)) {
        Write-Host ""
        Write-Host "Package validation failed. The selected folder does not contain a valid AutoDesk package." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 7: Parse application information
    $summaryInfo = Get-SummaryInfo -SourcePath $selectedFolder.FullName
    if (!$summaryInfo) {
        Write-Host ""
        Write-Host "Failed to parse application information from Summary.txt." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 8: Create application package
    $packagePath = New-ApplicationPackage -SelectedFolder $selectedFolder -BasePath $BasePath
    if (!$packagePath) {
        Write-Host ""
        Write-Host "Failed to create application package." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 9: Generate installation script
    $scriptPath = New-InstallationScript -SelectedFolder $selectedFolder -SummaryInfo $summaryInfo -BasePath $BasePath
    if (!$scriptPath) {
        Write-Host ""
        Write-Host "Failed to generate installation script." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    
    # Step 10: Generate Intune deployment info
    $intuneInfoPath = New-IntuneInfo -SelectedFolder $selectedFolder -SummaryInfo $summaryInfo -BasePath $BasePath
    if (!$intuneInfoPath) {
        Write-LogMessage "Failed to generate Intune info file (non-critical)" "WARNING"
    }
    
    # Step 11: Create Intune Win32 package
    Write-LogMessage "Preparing for Intune packaging..." "INFO"
    Write-LogMessage "Files to be included in Intune package:" "INFO"
    Write-LogMessage "  1. $(Split-Path $packagePath -Leaf) - AutoDesk application package" "INFO"
    Write-LogMessage "  2. $(Split-Path $scriptPath -Leaf) - PowerShell installation script" "INFO"
    Write-LogMessage "No other files from previous runs will be included" "SUCCESS"
    
    $outputFolder = Join-Path $BasePath "Output"
    
    $intunePackagePath = Invoke-IntunePackaging -IntuneToolPath $intuneToolPath -PackageZipPath $packagePath -InstallScriptPath $scriptPath -ApplicationName $selectedFolder.Name -OutputFolder $outputFolder
    
    if ($intunePackagePath) {
        Write-LogMessage "=== PACKAGE CREATION COMPLETED SUCCESSFULLY ===" "SUCCESS"
        Write-Host ""
        Write-Host "Package Creation Summary:" -ForegroundColor Green
        Write-Host "========================" -ForegroundColor Green
        Write-Host "Application: $($summaryInfo.ProgramName) v$($summaryInfo.BuildNumber)" -ForegroundColor White
        Write-Host "Intune Package: $(Split-Path $intunePackagePath -Leaf)" -ForegroundColor White
        Write-Host "Package Location: $intunePackagePath" -ForegroundColor White
        Write-Host "Deployment Info: $intuneInfoPath" -ForegroundColor White
        Write-Host "Log File: $Script:LogFile" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Next Steps:" -ForegroundColor Cyan
        Write-Host "1. Upload the .intunewin file to Microsoft Intune" -ForegroundColor White
        Write-Host "2. Use the deployment info file for configuration guidance" -ForegroundColor White
        Write-Host "3. Test deployment on a pilot group first" -ForegroundColor White
        Write-Host ""
    } else {
        Write-LogMessage "Intune packaging failed, but other components were created successfully" "WARNING"
        Write-Host ""
        Write-Host "Partial Success:" -ForegroundColor Yellow
        Write-Host "===============" -ForegroundColor Yellow
        Write-Host "The installation script and package were created, but Intune packaging failed." -ForegroundColor White
        Write-Host "You can manually run IntuneWinAppUtil.exe on the generated files." -ForegroundColor White
    }
    
} catch {
    Write-LogMessage "Script execution failed with unexpected error: $($_.Exception.Message)" "ERROR"
    Write-LogMessage "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    Write-Host ""
    Write-Host "An unexpected error occurred. Please check the log file for details:" -ForegroundColor Red
    Write-Host "$Script:LogFile" -ForegroundColor Gray
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Read-Host "Press Enter to exit"

#endregion
