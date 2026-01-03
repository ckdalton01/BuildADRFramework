# ============================================================================
# Multi-ADR Patch Strategy Builder for ConfigMgr
# ============================================================================
# This script creates:
# - 3 Device Collections (Test, Broad, Production)
# - 4 Deployment Packages (Microsoft, Office, Defender, Third-Party)
# - 4 ADRs with multiple phased deployments
#
# Run with -Uninstall to remove all created items
# ============================================================================

[CmdletBinding()]
param(
    [switch]$Uninstall
)

# --- CONFIGURATION VARIABLES ---
$SiteCode = ""          # Will be prompted if empty
$SiteServer = ""        # Will be prompted if empty
$CollectionFolderName = "Patch Strategy"

# Collection Definitions
$Collections = @(
    @{ Name = "01 - Test - All Devices"; LimitingCollection = "All Systems" }
    @{ Name = "02 - Broad - All Devices"; LimitingCollection = "All Systems" }
    @{ Name = "03 - Production - All Devices"; LimitingCollection = "All Systems" }
)

# Deployment Package Definitions
$DeploymentPackages = @(
    @{ Name = "Microsoft Updates"; SubFolder = "Microsoft Updates" }
    @{ Name = "Office Updates"; SubFolder = "Office Updates" }
    @{ Name = "Defender Updates"; SubFolder = "Defender Updates" }
    @{ Name = "Third Party Updates"; SubFolder = "Third Party Updates" }
)

# Products Configuration
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

# ADR Names
$ADRNames = @(
    "Windows OS Updates",
    "Office Updates",
    "Defender Updates",
    "Third Party Updates"
)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Test-UNCPathAccess {
    param(
        [string]$Path
    )
    
    # Save current location and switch to a filesystem drive
    $CurrentLocation = Get-Location
    try {
        Set-Location $env:SystemDrive\ -ErrorAction Stop
    } catch {
        Write-Warning "  -> Could not switch to filesystem drive"
        return $false
    }
    
    try {
        # Test if path exists
        if (-not (Test-Path -Path $Path -PathType Container)) {
            Write-Warning "Path does not exist: $Path"
            Write-Host "Attempting to create path..." -ForegroundColor Yellow
            
            try {
                New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "  -> Successfully created path" -ForegroundColor Green
            } catch {
                Write-Warning "  -> Failed to create path: $($_.Exception.Message)"
                Set-Location $CurrentLocation
                return $false
            }
        }
        
        # Test write access by creating a temporary file
        $TestFile = Join-Path -Path $Path -ChildPath "~test_write_access_$(Get-Random).tmp"
        try {
            New-Item -Path $TestFile -ItemType File -Force -ErrorAction Stop | Out-Null
            Remove-Item -Path $TestFile -Force -ErrorAction SilentlyContinue
            Set-Location $CurrentLocation
            return $true
        } catch {
            Write-Warning "  -> No write access to path: $($_.Exception.Message)"
            Set-Location $CurrentLocation
            return $false
        }
    } catch {
        Write-Warning "  -> Error accessing path: $($_.Exception.Message)"
        Set-Location $CurrentLocation
        return $false
    }
}

function Test-SCCMConnection {
    param(
        [string]$SiteCode,
        [string]$SiteServer
    )
    
    try {
        # Save current location
        $CurrentLocation = Get-Location
        
        # Test if we can connect to the site
        $TestDrive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue
        
        if ($TestDrive) {
            # Switch to the CM drive to test
            Set-Location "$($SiteCode):" -ErrorAction Stop
            
            # Test if we can query the site
            $SiteInfo = Get-CMSite -SiteCode $SiteCode -ErrorAction SilentlyContinue
            
            # Return to original location
            Set-Location $CurrentLocation
            
            if ($SiteInfo) {
                return $true
            } else {
                Write-Warning "Connected to drive but cannot query site: $SiteCode"
                return $false
            }
        } else {
            Write-Warning "Cannot connect to site: $SiteCode"
            return $false
        }
    } catch {
        # Attempt to return to original location
        try { Set-Location $CurrentLocation } catch { }
        Write-Warning "Error testing SCCM connection: $($_.Exception.Message)"
        return $false
    }
}
function Set-CMSoftwareUpdateAutoDeploymentRuleIsDeployed {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteCode,
        [Parameter(Mandatory = $true)]
        [string]$ADRName,
        [Parameter(Mandatory = $true)]
        [bool]$IsDeployed
    )

    # Get the ADR
    try {
        $Namespace = "root/sms/site_" + $SiteCode
        [wmi]$ADR = (Get-WmiObject -Class SMS_AutoDeployment -Namespace $Namespace | Where-Object -FilterScript { $_.Name -eq $ADRName }).__PATH
    }
    catch {
        throw 'Failed to Set ADR IsDeployed'
    }

    try {
        # Load the Current ADR UpdateRuleXML
        $UpdateRuleXML = New-Object -TypeName XML
        $UpdateRuleXML.LoadXML($ADR.UpdateRuleXML)
        $UpdateRuleXML.UpdateXML.UpdateXMLDescriptionItems.OutterXml

        # Create the UpdateXMLDescriptionItem node if it doesn't exist
        $IsDeployedNode = $UpdateRuleXML.SelectSingleNode("/UpdateXML/UpdateXMLDescriptionItems/UpdateXMLDescriptionItem[@PropertyName='IsDeployed']")
        if ($null -eq $IsDeployedNode) {
            $IsDeployedNode = $UpdateRuleXML.CreateElement("UpdateXMLDescriptionItem")
            $IsDeployedNode.SetAttribute("PropertyName", "IsDeployed")
            $UpdateRuleXML.SelectSingleNode("/UpdateXML/UpdateXMLDescriptionItems").AppendChild($IsDeployedNode) | Out-Null
        }

        # Create the MatchRules node if it doesn't exist
        $MatchRulesNode = $IsDeployedNode.SelectSingleNode("MatchRules")
        if ($null -eq $MatchRulesNode) {
            $MatchRulesNode = $UpdateRuleXML.CreateElement("MatchRules")
            $IsDeployedNode.AppendChild($MatchRulesNode) | Out-Null
        }
        else {
            # Replace the existing MatchRules node
            $IsDeployedNode.RemoveChild($MatchRulesNode) | Out-Null
            $MatchRulesNode = $UpdateRuleXML.CreateElement("MatchRules")
            $IsDeployedNode.AppendChild($MatchRulesNode) | Out-Null
        }

        # Create a new string element with the innertext of IsDeployed
        $CreateXMLElement = $UpdateRuleXML.CreateElement("string")
        $CreateXMLElement.InnerText = "$($IsDeployed.ToString().ToLower())"

        # Append the new string element to the MatchRules node
        $MatchRulesNode.AppendChild($CreateXMLElement) | Out-Null

        # Get the OuterXML property
        $OuterXML = $UpdateRuleXML.UpdateXML.OuterXml

        # Put the OuterXML on the ADR object
        $ADR.UpdateRuleXML = $OuterXML
        $ADR.Put() | Out-Null
    }
    catch {
        throw 'Failed to update ADR UpdateRuleXML'
    }
}

# ============================================================================
# UNINSTALL MODE
# ============================================================================

if ($Uninstall) {
    Write-Host "============================================================================" -ForegroundColor Red
    Write-Host " UNINSTALL MODE - Removing Patch Strategy Components" -ForegroundColor Red
    Write-Host "============================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Warning "This will remove all ADRs, packages, and collections created by this script."
    $Confirm = Read-Host "Type 'YES' to continue"
    
    if ($Confirm -ne "YES") {
        Write-Host "Uninstall cancelled." -ForegroundColor Yellow
        exit
    }
    
    Write-Host ""
    Write-Host "Connecting to ConfigMgr..." -ForegroundColor Yellow
    
    # Prompt for Site Code if not configured
    if ([string]::IsNullOrWhiteSpace($SiteCode)) {
        do {
            $SiteCode = Read-Host "Enter your SCCM Site Code (e.g., CHQ, PS1)"
            if ([string]::IsNullOrWhiteSpace($SiteCode)) {
                Write-Warning "Site Code cannot be empty."
            }
        } while ([string]::IsNullOrWhiteSpace($SiteCode))
        $SiteCode = $SiteCode.Trim().ToUpper()
    }
    
    # Prompt for Site Server if not configured
    if ([string]::IsNullOrWhiteSpace($SiteServer)) {
        do {
            $SiteServer = Read-Host "Enter your SCCM Site Server (e.g., PrimaryServer, hostname)"
            if ([string]::IsNullOrWhiteSpace($SiteServer)) {
                Write-Warning "Site Server cannot be empty."
            }
        } while ([string]::IsNullOrWhiteSpace($SiteServer))
        $SiteServer = $SiteServer.Trim()
    }
    
    try {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    } catch {
        Write-Error "Configuration Manager module not found. Run this from a machine with the SCCM Console."
        exit
    }
    
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -Description "SCCM Site" | Out-Null
    }
    
    Set-Location "$($SiteCode):"
    Write-Host "Connected to site: $SiteCode" -ForegroundColor Green
    Write-Host ""
    
    # Suppress the Fast parameter warning
    $CMPSSuppressFastNotUsedCheck = $true
    
    # --- Step 1: Remove ADRs (this also removes their deployments) ---
    Write-Host "[1/4] Removing Automatic Deployment Rules..." -ForegroundColor Yellow
    foreach ($ADRName in $ADRNames) {
        $ADR = Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName -ErrorAction SilentlyContinue
        if ($ADR) {
            try {
                Remove-CMSoftwareUpdateAutoDeploymentRule -InputObject $ADR -Force -ErrorAction Stop
                Write-Host "  -> Removed ADR: $ADRName" -ForegroundColor Green
            } catch {
                Write-Warning "  -> Failed to remove ADR: $ADRName - $($_.Exception.Message)"
            }
        } else {
            Write-Host "  -> ADR not found: $ADRName" -ForegroundColor Gray
        }
    }
    Write-Host ""
    
    # --- Step 2: Remove Software Update Groups created by ADRs ---
    Write-Host "[2/4] Removing Software Update Groups..." -ForegroundColor Yellow
    foreach ($ADRName in $ADRNames) {
        # ADRs typically create SUGs with the same name or similar pattern
        $SUG = Get-CMSoftwareUpdateGroup -Name $ADRName -ErrorAction SilentlyContinue
        if ($SUG) {
            try {
                Remove-CMSoftwareUpdateGroup -InputObject $SUG -Force -ErrorAction Stop
                Write-Host "  -> Removed SUG: $ADRName" -ForegroundColor Green
            } catch {
                Write-Warning "  -> Failed to remove SUG: $ADRName - $($_.Exception.Message)"
            }
        } else {
            Write-Host "  -> SUG not found: $ADRName" -ForegroundColor Gray
        }
    }
    Write-Host ""
    
    # --- Step 3: Remove Deployment Packages ---
    Write-Host "[3/4] Removing Deployment Packages..." -ForegroundColor Yellow
    foreach ($Package in $DeploymentPackages) {
        $PackageName = $Package.Name
        $Pkg = Get-CMSoftwareUpdateDeploymentPackage -Name $PackageName -ErrorAction SilentlyContinue
        if ($Pkg) {
            try {
                Remove-CMSoftwareUpdateDeploymentPackage -InputObject $Pkg -Force -ErrorAction Stop
                Write-Host "  -> Removed package: $PackageName" -ForegroundColor Green
            } catch {
                Write-Warning "  -> Failed to remove package: $PackageName - $($_.Exception.Message)"
            }
        } else {
            Write-Host "  -> Package not found: $PackageName" -ForegroundColor Gray
        }
    }
    Write-Host ""
    
    # --- Step 4: Remove Collections ---
    Write-Host "[4/4] Removing Device Collections..." -ForegroundColor Yellow
    foreach ($Collection in $Collections) {
        $CollectionName = $Collection.Name
        $Coll = Get-CMDeviceCollection -Name $CollectionName -ErrorAction SilentlyContinue
        if ($Coll) {
            # Check for active deployments
            $Deployments = Get-CMDeployment -CollectionName $CollectionName -ErrorAction SilentlyContinue
            if ($Deployments) {
                Write-Warning "  -> Collection has active deployments: $CollectionName"
                Write-Warning "     Cannot remove collection with active deployments. Remove deployments first."
            } else {
                try {
                    Remove-CMCollection -InputObject $Coll -Force -ErrorAction Stop
                    Write-Host "  -> Removed collection: $CollectionName" -ForegroundColor Green
                } catch {
                    Write-Warning "  -> Failed to remove collection: $CollectionName - $($_.Exception.Message)"
                }
            }
        } else {
            Write-Host "  -> Collection not found: $CollectionName" -ForegroundColor Gray
        }
    }
    
    # Try to remove the collection folder if empty
    try {
        $FolderPath = "$($SiteCode):\DeviceCollection\$CollectionFolderName"
        if (Test-Path $FolderPath) {
            $FolderContents = Get-ChildItem -Path $FolderPath -ErrorAction SilentlyContinue
            if ($FolderContents.Count -eq 0) {
                Remove-Item -Path $FolderPath -Force -ErrorAction Stop
                Write-Host "  -> Removed folder: $CollectionFolderName" -ForegroundColor Green
            } else {
                Write-Host "  -> Folder not empty, skipping removal: $CollectionFolderName" -ForegroundColor Gray
            }
        }
    } catch {
        Write-Warning "  -> Could not remove folder: $CollectionFolderName - $($_.Exception.Message)"
    }
    
    Write-Host ""
    Write-Host "============================================================================" -ForegroundColor Red
    Write-Host " Uninstall Complete!" -ForegroundColor Green
    Write-Host "============================================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Note: Package source folders on the file system were NOT removed." -ForegroundColor Yellow
    Write-Host "You may need to manually delete folders under your UNC path." -ForegroundColor Yellow
    Write-Host ""
    
    exit
}

# ============================================================================
# SCRIPT EXECUTION (Install Mode)
# ============================================================================

Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host " Multi-ADR Patch Strategy Builder" -ForegroundColor Cyan
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""

# --- 1. Import Module & Connect ---
Write-Host "[1/6] Connecting to ConfigMgr..." -ForegroundColor Yellow

# Import ConfigMgr module first
try {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    Write-Host "  -> Module imported successfully" -ForegroundColor Green
} catch {
    Write-Error "Configuration Manager module not found. Run this from a machine with the SCCM Console."
    Break
}

# Prompt for and validate Site Code
$SiteCodeValid = $false
do {
    if ([string]::IsNullOrWhiteSpace($SiteCode)) {
        $SiteCode = Read-Host "Enter your SCCM Site Code (e.g., CHQ, PS1)"
        if ([string]::IsNullOrWhiteSpace($SiteCode)) {
            Write-Warning "Site Code cannot be empty."
            continue
        }
        $SiteCode = $SiteCode.Trim().ToUpper()
    }
    
    # Prompt for Site Server if needed
    if ([string]::IsNullOrWhiteSpace($SiteServer)) {
        $SiteServer = Read-Host "Enter your SCCM Site Server (e.g., PrimarySiteServer, hostname)"
        if ([string]::IsNullOrWhiteSpace($SiteServer)) {
            Write-Warning "Site Server cannot be empty."
            $SiteCode = ""
            continue
        }
        $SiteServer = $SiteServer.Trim()
    }
    
    # Try to connect to the site
    Write-Host "  -> Validating connection to site: $SiteCode on server: $SiteServer" -ForegroundColor Yellow
    
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -Description "SCCM Site" -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "  -> Failed to connect: $($_.Exception.Message)"
            Write-Warning "  -> Please verify Site Code and Site Server are correct."
            $SiteCode = ""
            $SiteServer = ""
            continue
        }
    }
    
    # Validate the connection
    if (Test-SCCMConnection -SiteCode $SiteCode -SiteServer $SiteServer) {
        $SiteCodeValid = $true
        Write-Host "  -> Successfully connected to site: $SiteCode" -ForegroundColor Green
    } else {
        Write-Warning "  -> Could not validate connection to site: $SiteCode"
        Write-Warning "  -> Please verify Site Code and Site Server are correct."
        $SiteCode = ""
        $SiteServer = ""
    }
    
} while (-not $SiteCodeValid)

Set-Location "$($SiteCode):"
Write-Host ""

# --- 2. Get Source Path from User ---
Write-Host "[2/6] Deployment Package Configuration" -ForegroundColor Yellow

$SourcePathValid = $false
do {
    $SourceBasePath = Read-Host "Enter the UNC path for deployment packages (e.g., \\CM1\Sources\Updates)"
    
    # Validate UNC path format
    if ($SourceBasePath -notmatch "^\\\\") {
        Write-Warning "Path must be a UNC path (starting with \\)."
        continue
    }
    
    # Remove trailing backslash if present
    $SourceBasePath = $SourceBasePath.TrimEnd('\')
    
    # Test path access
    Write-Host "  -> Validating path access..." -ForegroundColor Yellow
    if (Test-UNCPathAccess -Path $SourceBasePath) {
        $SourcePathValid = $true
        Write-Host "  -> Base path validated: $SourceBasePath" -ForegroundColor Green
        
        # Test access to subfolders
        Write-Host "  -> Testing subfolder access..." -ForegroundColor Yellow
        $AllSubfoldersValid = $true
        foreach ($Package in $DeploymentPackages) {
            $SubPath = "$SourceBasePath\$($Package.SubFolder)"
            if (-not (Test-UNCPathAccess -Path $SubPath)) {
                $AllSubfoldersValid = $false
                Write-Warning "  -> Could not access/create: $SubPath"
            }
        }
        
        if (-not $AllSubfoldersValid) {
            Write-Warning "One or more subfolders could not be validated."
            $Retry = Read-Host "Do you want to try a different path? (Y/N)"
            if ($Retry -eq "Y" -or $Retry -eq "y") {
                $SourcePathValid = $false
                continue
            } else {
                Write-Host "  -> Continuing with current path..." -ForegroundColor Yellow
            }
        } else {
            Write-Host "  -> All subfolders validated successfully" -ForegroundColor Green
        }
    } else {
        Write-Warning "Cannot access the specified path. Please verify:"
        Write-Warning "  1. The path exists or you have permissions to create it"
        Write-Warning "  2. You have read/write permissions to the path"
        Write-Warning "  3. The path is accessible from this machine"
    }
    
} while (-not $SourcePathValid)

Write-Host ""

# --- 3. Create Device Collections ---
Write-Host "[3/6] Creating Device Collections..." -ForegroundColor Yellow

# Create folder for collections
try {
    $FolderPath = "$($SiteCode):\DeviceCollection\$CollectionFolderName"
    if (-not (Test-Path $FolderPath)) {
        New-Item -Path "$($SiteCode):\DeviceCollection" -Name $CollectionFolderName -ErrorAction Stop | Out-Null
        Write-Host "  -> Created folder: $CollectionFolderName" -ForegroundColor Green
    } else {
        Write-Host "  -> Folder already exists: $CollectionFolderName" -ForegroundColor Gray
    }
} catch {
    Write-Warning "  -> Could not create folder: $($_.Exception.Message)"
    Write-Warning "  -> Collections will be created in root."
}

foreach ($Collection in $Collections) {
    $CollectionName = $Collection.Name
    $LimitingCollection = $Collection.LimitingCollection
    
    if (Get-CMDeviceCollection -Name $CollectionName -ErrorAction SilentlyContinue) {
        Write-Host "  -> Collection already exists: $CollectionName" -ForegroundColor Gray
    } else {
        try {
            $Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7
            $NewCollection = New-CMDeviceCollection `
                -Name $CollectionName `
                -LimitingCollectionName $LimitingCollection `
                -RefreshType Periodic `
                -RefreshSchedule $Schedule `
                -ErrorAction Stop
            
            # Move to folder if it exists
            if (Test-Path $FolderPath) {
                Move-CMObject -FolderPath $FolderPath -InputObject $NewCollection -ErrorAction SilentlyContinue | Out-Null
            }
            
            Write-Host "  -> Created collection: $CollectionName" -ForegroundColor Green
        } catch {
            Write-Warning "  -> Failed to create collection: $CollectionName - $($_.Exception.Message)"
        }
    }
}
Write-Host ""

# --- 4. Create Deployment Packages ---
Write-Host "[4/6] Creating Deployment Packages..." -ForegroundColor Yellow

foreach ($Package in $DeploymentPackages) {
    $PackageName = $Package.Name
    $PackagePath = "$SourceBasePath\$($Package.SubFolder)"
    
    if (Get-CMSoftwareUpdateDeploymentPackage -Name $PackageName -ErrorAction SilentlyContinue) {
        Write-Host "  -> Package already exists: $PackageName" -ForegroundColor Gray
    } else {
        try {
            New-CMSoftwareUpdateDeploymentPackage -Name $PackageName -Path $PackagePath -ErrorAction Stop | Out-Null
            Write-Host "  -> Created package: $PackageName at $PackagePath" -ForegroundColor Green
        } catch {
            Write-Warning "  -> Failed to create package: $PackageName - $($_.Exception.Message)"
        }
    }
}
Write-Host ""

# --- 5. Create ADRs with Multiple Deployments ---
Write-Host "[5/6] Creating Automatic Deployment Rules..." -ForegroundColor Yellow
Write-Host ""

# Suppress the Fast parameter warning
$CMPSSuppressFastNotUsedCheck = $true

# ============================================================================
# WINDOWS OS UPDATES ADR
# ============================================================================
Write-Host "  Creating Windows OS Updates ADR..." -ForegroundColor Cyan
$ADRName = "Windows OS Updates"

if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName -ErrorAction SilentlyContinue) {
    Write-Host "    -> ADR already exists: $ADRName" -ForegroundColor Gray
} else {
    try {
        # Create schedule for daily runs at 1 AM
        $Schedule = New-CMSchedule -Start (Get-Date "01:00") -RecurInterval Days -RecurCount 1
        
        # Create the ADR with Test collection (1 day deadline)
        $ADR = New-CMSoftwareUpdateAutoDeploymentRule `
            -Name $ADRName `
            -Description "Windows OS updates with phased deployment to Test/Broad/Production - CDalton Template" `
            -Collection (Get-CMDeviceCollection -Name "01 - Test - All Devices") `
            -Product $WindowsProducts `
            -Superseded $false `
            -DateReleasedOrRevised Last1Month `
            -Required ">=1" `
            -DeploymentPackageName "Microsoft Updates" `
            -AddToExistingSoftwareUpdateGroup $true `
            -EnabledAfterCreate $true `
            -RunType RunTheRuleOnSchedule `
            -Schedule $Schedule `
            -AvailableImmediately $true `
            -UserNotification DisplayAll `
            -AllowSoftwareInstallationOutsideMaintenanceWindow $false `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -WriteFilterHandling $false `
            -GenerateSuccessAlert $false `
            -DeadlineTime 1 `
            -DeadlineTimeUnit Days `
            -RequirePostRebootFullScan $true `
            -ErrorAction Stop
        
        Write-Host "    -> Created ADR: $ADRName with Test deployment (1 day)" -ForegroundColor Green
        
        # Add Broad deployment (3 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "02 - Broad - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 3 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Broad deployment (3 days)" -ForegroundColor Green
        
        # Add Production deployment (7 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "03 - Production - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 7 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Production deployment (7 days)" -ForegroundColor Green
    } catch {
        Write-Warning "    -> Failed to create Windows OS Updates ADR: $($_.Exception.Message)"
    }
}
Write-Host ""

# ============================================================================
# OFFICE UPDATES ADR
# ============================================================================
Write-Host "  Creating Office Updates ADR..." -ForegroundColor Cyan
$ADRName = "Office Updates"

if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName -ErrorAction SilentlyContinue) {
    Write-Host "    -> ADR already exists: $ADRName" -ForegroundColor Gray
} else {
    try {
        # Create schedule for daily runs at 2 AM
        $Schedule = New-CMSchedule -Start (Get-Date "02:00") -RecurInterval Days -RecurCount 1
        
        # Create the ADR with Test collection (1 day deadline)
        $ADR = New-CMSoftwareUpdateAutoDeploymentRule `
            -Name $ADRName `
            -Description "Office updates with phased deployment to Test/Broad/Production - CDalton Template" `
            -Collection (Get-CMDeviceCollection -Name "01 - Test - All Devices") `
            -Product $OfficeProducts `
            -Superseded $false `
            -DateReleasedOrRevised Last1Month `
            -Required ">=1" `
            -DeploymentPackageName "Office Updates" `
            -AddToExistingSoftwareUpdateGroup $true `
            -EnabledAfterCreate $true `
            -RunType RunTheRuleOnSchedule `
            -Schedule $Schedule `
            -AvailableImmediately $true `
            -UserNotification DisplayAll `
            -AllowSoftwareInstallationOutsideMaintenanceWindow $false `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -WriteFilterHandling $false `
            -GenerateSuccessAlert $false `
            -DeadlineTime 1 `
            -DeadlineTimeUnit Days `
            -ErrorAction Stop
        
        Write-Host "    -> Created ADR: $ADRName with Test deployment (1 day)" -ForegroundColor Green
        
        # Add Broad deployment (3 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "02 - Broad - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 3 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Broad deployment (3 days)" -ForegroundColor Green
        
        # Add Production deployment (7 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "03 - Production - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 7 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $false `
            -SuppressRestartWorkstation $false `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Production deployment (7 days)" -ForegroundColor Green
    } catch {
        Write-Warning "    -> Failed to create Office Updates ADR: $($_.Exception.Message)"
    }
}
Write-Host ""

# ============================================================================
# DEFENDER UPDATES ADR
# ============================================================================
Write-Host "  Creating Defender Updates ADR..." -ForegroundColor Cyan
$ADRName = "Defender Updates"

if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName -ErrorAction SilentlyContinue) {
    Write-Host "    -> ADR already exists: $ADRName" -ForegroundColor Gray
} else {
    try {
        # Create schedule for every 8 hours
        $Schedule = New-CMSchedule -Start (Get-Date "00:00") -RecurInterval Hours -RecurCount 8
        
        # Create the ADR with Test collection (immediate deadline)
        $ADR = New-CMSoftwareUpdateAutoDeploymentRule `
            -Name $ADRName `
            -Description "Defender updates with phased deployment to Test/Production - CDalton Template" `
            -Collection (Get-CMDeviceCollection -Name "01 - Test - All Devices") `
            -Product $DefenderProducts `
            -DateReleasedOrRevised Last1Month `
            -Superseded $false `
            -DeploymentPackageName "Defender Updates" `
            -AddToExistingSoftwareUpdateGroup $true `
            -EnabledAfterCreate $true `
            -RunType RunTheRuleOnSchedule `
            -Schedule $Schedule `
            -AvailableImmediately $true `
            -DeadlineImmediately $true `
            -UserNotification DisplayAll `
            -AllowSoftwareInstallationOutsideMaintenanceWindow $true `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -WriteFilterHandling $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop
        
        Write-Host "    -> Created ADR: $ADRName with Test deployment (immediate)" -ForegroundColor Green
        
        # Add Broad deployment (2 hour deadline)        
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "02 - Broad - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 2 `
            -DeadlineTimeUnit Hours `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Broad deployment (2 hours)" -ForegroundColor Green

        # Add Production deployment (4 hour deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "03 - Production - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 4 `
            -DeadlineTimeUnit Hours `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Production deployment (4 hours)" -ForegroundColor Green
    } catch {
        Write-Warning "    -> Failed to create Defender Updates ADR: $($_.Exception.Message)"
    }
}
Write-Host ""

# ============================================================================
# THIRD PARTY UPDATES ADR
# ============================================================================
Write-Host "  Creating Third Party Updates ADR..." -ForegroundColor Cyan
$ADRName = "Third Party Updates"

if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName -ErrorAction SilentlyContinue) {
    Write-Host "    -> ADR already exists: $ADRName" -ForegroundColor Gray
} else {
    try {
        # Create schedule for daily runs at 3 AM
        $Schedule = New-CMSchedule -Start (Get-Date "03:00") -RecurInterval Days -RecurCount 1
        
        # Create the ADR with Test collection (immediate deadline)
        $ADR = New-CMSoftwareUpdateAutoDeploymentRule `
            -Name $ADRName `
            -Description "Third-party updates with phased deployment to Test/Broad/Production - CDalton Template" `
            -Collection (Get-CMDeviceCollection -Name "01 - Test - All Devices") `
            -Vendor "Patch My PC" `
            -Superseded $false `
            -DeploymentPackageName "Third Party Updates" `
            -AddToExistingSoftwareUpdateGroup $false `
            -EnabledAfterCreate $true `
            -RunType RunTheRuleOnSchedule `
            -Schedule $Schedule `
            -AvailableImmediately $true `
            -DeadlineImmediately $true `
            -UserNotification DisplayAll `
            -AllowSoftwareInstallationOutsideMaintenanceWindow $true `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -WriteFilterHandling $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop
        
        Write-Host "    -> Created ADR: $ADRName with Test deployment (immediate)" -ForegroundColor Green

        # --- SetISDeployed Flag for PMPC Updates ---
        Set-CMSoftwareUpdateAutoDeploymentRuleIsDeployed `
            -SiteCode $SiteCode `
            -ADRName $ADRName `
            -IsDeployed $false

        # Add Broad deployment (3 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "02 - Broad - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 3 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Broad deployment (3 days)" -ForegroundColor Green
        
        # Add Production deployment (5 day deadline)
        New-CMAutoDeploymentRuleDeployment `
            -InputObject $ADR `
            -Collection (Get-CMDeviceCollection -Name "03 - Production - All Devices") `
            -EnableDeployment $true `
            -SendWakeUpPacket $false `
            -UseUtc $false `
            -AvailableImmediately $true `
            -DeadlineTime 5 `
            -DeadlineTimeUnit Days `
            -UserNotification DisplayAll `
            -AllowRestart $false `
            -SuppressRestartServer $true `
            -SuppressRestartWorkstation $true `
            -RequirePostRebootFullScan $false `
            -DisableOperationsManager $false `
            -GenerateOperationsManagerAlert $false `
            -GenerateSuccessAlert $false `
            -ErrorAction Stop | Out-Null
        
        Write-Host "    -> Added Production deployment (5 days)" -ForegroundColor Green
    } catch {
        Write-Warning "    -> Failed to create Third Party Updates ADR: $($_.Exception.Message)"
    }
}

Write-Host ""

# --- 6. Summary ---
Write-Host ""
Write-Host "[6/6] Deployment Summary" -ForegroundColor Yellow
Write-Host "  Collections Created: 3" -ForegroundColor Green
Write-Host "    - 01 - Test - All Devices"
Write-Host "    - 02 - Broad - All Devices"
Write-Host "    - 03 - Production - All Devices"
Write-Host ""
Write-Host "  Deployment Packages Created: 4" -ForegroundColor Green
Write-Host "    - Microsoft Updates -> $SourceBasePath\Microsoft Updates"
Write-Host "    - Office Updates -> $SourceBasePath\Office Updates"
Write-Host "    - Defender Updates -> $SourceBasePath\Defender Updates"
Write-Host "    - Third Party Updates -> $SourceBasePath\Third Party Updates"
Write-Host ""
Write-Host "  ADRs Created: 4 (with multiple deployments each)" -ForegroundColor Green
Write-Host "    1. Windows OS Updates" -ForegroundColor Cyan
Write-Host "       - Test (1 day), Broad (3 days), Production (7 days)"
Write-Host "       - Runs: Daily at 1 AM"
Write-Host ""
Write-Host "    2. Office Updates" -ForegroundColor Cyan
Write-Host "       - Test (1 day), Broad (3 days), Production (7 days)"
Write-Host "       - Runs: Daily at 2 AM"
Write-Host ""
Write-Host "    3. Defender Updates" -ForegroundColor Cyan
Write-Host "       - Test (immediate), Production (4 hours)"
Write-Host "       - Runs: Every 8 hours"
Write-Host "       - Ignores maintenance windows"
Write-Host ""
Write-Host "    4. Third Party Updates" -ForegroundColor Cyan
Write-Host "       - Test (immediate), Broad (3 days), Production (5 days)"
Write-Host "       - Runs: Daily at 3 AM"
Write-Host "       - Creates new Software Update Group each run"
Write-Host "       - Ignores maintenance windows"
Write-Host ""
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host " Patch Strategy Deployment Complete!" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Populate device collections with appropriate members"
Write-Host "  2. Verify ADR schedules and deployments in the ConfigMgr Console"
Write-Host "  3. Monitor first sync and deployment creation"
Write-Host "  4. Test with a small group before full rollout"
Write-Host ""