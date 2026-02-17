# AutoDesk Package Tool for Microsoft Intune

A comprehensive PowerShell tool for automatically packaging AutoDesk applications (Revit, AutoCAD, Maya, 3ds Max, Inventor, Fusion 360) for deployment through Microsoft Intune Win32 apps.

## Features

- **Automatic Discovery** - Scans and detects AutoDesk deployment packages
- **Smart Validation** - Validates package integrity and critical files
- **Multiple Compression Methods** - Handles large packages (up to 16GB+) with fallback compression
- **Enhanced Error Handling** - Comprehensive error detection and recovery
- **Clean Packaging** - Isolated package creation preventing cross-contamination
- **Detailed Logging** - Complete audit trail with timestamped logs
- **Intune-Ready Output** - Generates complete .intunewin packages with deployment guides
- **Version Management** - Automatic version detection and upgrade handling

## Prerequisites

### System Requirements
- **Windows 10/11** (PowerShell 5.1 or higher)
- **Administrator privileges** required
- **Internet connection** (for downloading IntuneWinAppUtil.exe)

### AutoDesk Requirements
- AutoDesk deployment packages must be deployed to `C:\Windows\Temp\AutoDesk\`. You specify this part on AutoDesk deploy package download page.

## Installation

- View installation guide under Wiki page or klick link:
https://github.com/3TeRswe/AutoDesk_Intune_Packaging_Tool/wiki/Usage-and-Deployment-Guide

### Method 1: Direct Download
```powershell
# Download the script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/3TeRswe/autodesk-package-tool/main/AutoDesk_Package_Tool.ps1" -OutFile "AutoDesk_Package_Tool.ps1"

# Run as Administrator
.\AutoDesk_Package_Tool.ps1
```

### Method 2: Clone Repository
```bash
git clone https://github.com/3TeRswe/autodesk-package-tool.git
cd autodesk-package-tool
```

## Quick Start

### Basic Usage
```
powershell
# Run with default settings
.\AutoDesk_Package_Tool.ps1
#or to bypass execution policy, if scripts is disabled on your system
powershell.exe  -ExecutionPolicy Bypass -File '.\AutoDesk_Package_Tool.ps1'

# Force fresh download of IntuneWinAppUtil.exe
.\AutoDesk_Package_Tool.ps1 -ForceDownload

# Use custom paths
.\AutoDesk_Package_Tool.ps1 -BasePath "D:\Intune" -AutoDeskTempPath "C:\Temp\AutoDesk"
```

### Step-by-Step Process
1. **Deploy AutoDesk Package** - Use AutoDesk deployment tool to deploy to `C:\Windows\Temp\AutoDesk\`

2. **Run Script as Administrator** - Execute the PowerShell script with elevated privileges
3. **Select Application** - Choose from detected AutoDesk applications
4. **Review Package Info** - Verify application details and package size
5. **Wait for Completion** - Script creates compressed package and Intune deployment files
6. **Upload to Intune** - Use generated `.intunewin` file in Microsoft Intune

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `BasePath` | String | `C:\Intune` | Base directory for all Intune packaging files |
| `AutoDeskTempPath` | String | `C:\Windows\Temp\AutoDesk` | Location where AutoDesk packages are deployed |
| `ForceDownload` | Switch | `$false` | Force download of IntuneWinAppUtil.exe even if exists |
| `MaxFileAgeDays` | Integer | `30` | Maximum age in days before redownloading IntuneWinAppUtil.exe |

### Usage Examples

```powershell
# Basic execution
.\AutoDesk_Package_Tool.ps1

# Force download and use custom base path
.\AutoDesk_Package_Tool.ps1 -BasePath "D:\IntunePackages" -ForceDownload

# Set shorter cache period for IntuneWinAppUtil.exe
.\AutoDesk_Package_Tool.ps1 -MaxFileAgeDays 7

# Custom AutoDesk deployment location
.\AutoDesk_Package_Tool.ps1 -AutoDeskTempPath "C:\CustomPath\AutoDesk"
```

## Output Structure

The script creates the following directory structure:

```
C:\Intune\
├── App\
│   └── IntuneWinAppUtil.exe                    # Microsoft packaging tool
├── Source\
│   ├── Revit_2023.1.7.zip                     # Compressed AutoDesk package
│   └── Revit_2023.1.7_Install.ps1             # Installation script
├── Output\
│   ├── Revit_2023.1.7.intunewin               # Final Intune package
│   └── Revit_2023.1.7_IntuneInfo.txt          # Deployment configuration guide, includes all information you need to create Intune Application.
└── SourceFiles\                               # (Reserved for future use)
```

## Generated Files

### Installation Script (`*_Install.ps1`)
- **Automated installation/uninstallation** with comprehensive error handling
- **Version checking** to prevent downgrades
- **Administrator privilege validation**
- **Detailed logging** to Windows Event Log
- **Cleanup** of temporary files

### Intune Configuration File (`*_IntuneInfo.txt`)
Contains ready-to-use Intune deployment settings:
- Install/uninstall commands
- Detection rules (registry-based)
- System requirements
- Return codes
- Troubleshooting guidance

### Intune Win32 Package (`*.intunewin`)
- Microsoft Intune-compatible package file
- Contains compressed AutoDesk application and installation script
- Ready for upload to Intune portal

### Log Files
Detailed logs are saved to: `%TEMP%\AutoDesk_Package_Tool_YYYYMMDD_HHMMSS.log`

## Best Practices

### Intune Deployment
1. **Test deployments** on pilot groups first
2. **Monitor installation logs** in Windows Event Viewer
3. **Set appropriate detection rules** using provided configuration
4. **Configure restart behavior** based on application requirements

### Network Considerations
- Large packages may take significant time to upload to Intune
- Consider using Content Gateway for branch offices
- Schedule deployments during maintenance windows

## Security Considerations

- Script requires Administrator privileges for system access
- Installation scripts include privilege validation
- All operations are logged for audit purposes
- No sensitive data is stored or transmitted
- PowerShell execution policy bypass is used for deployment flexibility

