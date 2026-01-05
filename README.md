# Multi-ADR Patch Strategy Builder (ConfigMgr)

This repository contains a single PowerShell script, `BuildADRFramework.ps1`, used to create a simple phased patching strategy in Microsoft Configuration Manager (ConfigMgr / SCCM).

Key points
- Single-script repo — most edits happen in `BuildADRFramework.ps1`.
- Script is intended to be run from a machine with the ConfigMgr Console installed.

Requirements
- Windows machine with the SCCM Console installed (so `$ENV:SMS_ADMIN_UI_PATH` is defined).
- Elevated PowerShell session (run as Administrator).
- Access to an SMB UNC share for update package sources.

Quick run
1. Open an elevated PowerShell prompt on a machine with the SCCM Console.
2. Run:

```powershell
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\BuildADRFramework.ps1
```

Configuration (edit near top of `BuildADRFramework.ps1`)
- `$SiteCode` — User Defined
- `$SiteServer` — User Defined
- `$CollectionFolderName` — folder name used under Device Collections (default: `Patch Strategy`).
- `$Collections`, `$DeploymentPackages`, `$WindowsProducts`, `$OfficeProducts`, `$DefenderProducts` — arrays used to control what is created and which products the ADRs target.

What the script creates (behavior observed in the current code)
- Device collections (created if not present):
  - `01 - Test - All Devices`
  - `02 - Broad - All Devices`
  - `03 - Production - All Devices`
- Deployment packages (created if not present):
  - `Microsoft Updates` -> subfolder `Microsoft Updates`
  - `Office Updates` -> subfolder `Office Updates`
  - `Defender Updates` -> subfolder `Defender Updates`
  - `Third Party Updates` -> subfolder `Third Party Updates`
- Automatic Deployment Rules (ADRs) created (if not present):
  1. `Windows OS Updates` — runs daily at 01:00; phased deployments: Test (1 day), Broad (3 days), Production (7 days).
  2. `Office Updates` — runs daily at 02:00; phased deployments: Test (1 day), Broad (3 days), Production (7 days).
  3. `Defender Updates` — runs every 8 hours; Test (immediate), Broad (2 hours) Production (4 hours); configured to ignore maintenance windows.
  4. `Third Party Updates` — runs daily at 03:00; phased deployments: Test (immediate), Broad (3 days), Production (5 days); vendor set to "Patch My PC" and this ADR creates new Software Update Groups per run.


A PowerShell automation script that creates a complete phased patching strategy in Microsoft Configuration Manager (SCCM/MECM). This script sets up device collections, deployment packages, and Automatic Deployment Rules (ADRs) with multiple phased deployments for Windows, Office, Defender, and third-party updates.

## Features

- **Automated Infrastructure Setup**: Creates all necessary collections, packages, and ADRs with a single command
- **Phased Deployment Strategy**: Test → Broad → Production deployment phases with configurable deadlines
- **Multiple Update Categories**: Separate ADRs for Windows OS, Office, Defender, and third-party updates
- **Idempotent Operations**: Safe to run multiple times - checks for existing objects before creating
- **Easy Cleanup**: Built-in `-Uninstall` switch to remove all created components

## What It Creates

### Device Collections (3)
- **01 - Test - All Devices** - Early adopters/test systems
- **02 - Broad - All Devices** - Broader deployment group
- **03 - Production - All Devices** - Production systems

### Deployment Packages (4)
- **Microsoft Updates** - Windows OS updates
- **Office Updates** - Microsoft 365/Office updates
- **Defender Updates** - Defender AV and Endpoint updates
- **Third Party Updates** - Third-party application updates

### Automatic Deployment Rules (4)

#### 1. Windows OS Updates
- **Products**: Windows 10, Windows 11, Windows Server 2016/2019, .NET 8.0/9.0/10.0
- **Schedule**: Daily at 1:00 AM
- **Deployments**:
  - Test: 1 day deadline
  - Broad: 3 day deadline
  - Production: 7 day deadline

#### 2. Office Updates
- **Products**: Microsoft 365 Apps/Office 2019/Office LTSC
- **Schedule**: Daily at 2:00 AM
- **Deployments**:
  - Test: 1 day deadline
  - Broad: 3 day deadline
  - Production: 7 day deadline

#### 3. Defender Updates
- **Products**: Microsoft Defender Antivirus, Microsoft Defender for Endpoint
- **Schedule**: Every 8 hours
- **Deployments**:
  - Test: Immediate deadline
  - Broad: 2 hour deadline
  - Production: 4 hour deadline
- **Special**: Ignores maintenance windows, suppresses restarts

#### 4. Third Party Updates
- **Vendor**: Patch My PC
- **Schedule**: Daily at 3:00 AM
- **Deployments**:
  - Test: Immediate deadline
  - Broad: 3 day deadline
  - Production: 5 day deadline
- **Special**: Ignores maintenance windows, suppresses restarts

## Prerequisites

- **Configuration Manager Console** installed on the machine where you run the script
- **Elevated PowerShell session** (Run as Administrator)
- **Access to a UNC path** for storing deployment package sources (e.g., `\\CM1\Sources\Updates`)
- **Appropriate SCCM permissions** to create collections, packages, and ADRs

## Usage

### Install Mode (Create Components)

Run the script in an elevated PowerShell session:

```powershell
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\BuildADRFramework.ps1
```


### Uninstall Mode (Remove Components)

To remove all created components:

```powershell
PowerShell -NoProfile -ExecutionPolicy Bypass -File .\BuildADRFramework.ps1 -Uninstall
```

You'll be prompted to type `YES` to confirm the removal.

**Note**: The uninstall process:
- Removes ADRs and their deployments
- Removes Software Update Groups created by the ADRs
- Removes deployment packages
- Removes collections (only if they have no active deployments)
- Does NOT delete the physical UNC path folders

## Configuration

### Product Categories

Customize which products are included in each ADR:

```powershell
$WindowsProducts = @(
    "Windows 10",
    "Windows 11",
    "Windows Server 2016",
    "Windows Server 2019",
    "Microsoft Server operating system-21H2",
    "Microsoft Server Operating System-22H2",
    "Microsoft Server Operating System-23H2",
    "Microsoft Server Operating System-24H2",
    ".NET 9.0",
    ".NET 8.0",
    ".NET 10.0"
)

$OfficeProducts = @("Microsoft 365 Apps/Office 2019/Office LTSC")

$DefenderProducts = @(
    "Microsoft Defender Antivirus",
    "Microsoft Defender for Endpoint"
)
```

### Deployment Timelines

Modify the deadline parameters in each ADR creation block:

```powershell
-DeadlineTime 7      # Number of days/hours
-DeadlineTimeUnit Days  # Or "Hours"
```

## Post-Deployment Steps

After running the script:

1. **Populate Collections**: Add devices to the Test, Broad, and Production collections using membership rules
2. **Verify ADR Configuration**: Open the ConfigMgr Console and review the ADR settings
3. **Test First Sync**: Manually run one ADR to verify it creates update groups and deployments correctly
4. **Monitor Deployments**: Watch the first few deployment cycles before rolling out broadly
5. **Adjust as Needed**: Fine-tune deadlines, schedules, and collection membership based on results

## Troubleshooting

### Common Issues

**"Configuration Manager module not found"**
- Ensure you're running the script on a machine with the SCCM Console installed
- The module path is: `$ENV:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1`

**"Path must be a UNC path"**
- The deployment package path must start with `\\` (e.g., `\\server\share\path`)
- Do not use mapped drives

**"Failed to create collection/package/ADR"**
- Check that you have appropriate permissions in SCCM
- Verify the limiting collection exists
- Review the error message for specific issues

**Collections won't delete during uninstall**
- Collections with active deployments cannot be deleted
- The script will warn you which collections have active deployments
- Remove the deployments manually in the console, then re-run the uninstall

## Best Practices

- **Start Small**: Begin with a small test collection before adding more devices
- **Monitor Compliance**: Use ConfigMgr reports to track deployment success rates
- **Maintenance Windows**: Configure maintenance windows on collections to control when updates install
- **Backup**: Test the uninstall process in a lab before using in production
- **Documentation**: Document any customizations you make to the script

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built for Microsoft Configuration Manager (SCCM/MECM)
- Designed to work with Patch My PC for third-party patching

## Support

For issues, questions, or suggestions:
- Open an issue on GitHub

---

**Warning**: This script creates active deployment rules that will automatically deploy updates. Always test in a lab environment first and ensure you understand the implications of each ADR configuration before deploying to production.