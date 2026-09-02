<#
    .SYNOPSIS
    Creates or incrementally updates a "Take Over" package in Application Workspace.

    .DESCRIPTION
    Rerunnable version of Create-TakeOverPackage.ps1. Uses the same proven Liquit
    SDK calls as the original, but instead of always creating a new package it will:
    find or create the package, snapshot, and install action set, read the actions
    that already exist, skip any package whose take-over action already exists, and
    create take-over actions only for packages that are new since the last run.

    .EXAMPLE
    .\Create-TakeOverPackagev2.ps1
    .\Create-TakeOverPackagev2.ps1 -AddUnmanaged

    .EXAMPLE
    # Interactively pick (single- or multi-select) which packages get the
    # process-not-running filter via a grid; unselected packages get only the
    # file-exists detection.
    .\Create-TakeOverPackagev2.ps1 -SelectProcessFilterPackages

    .EXAMPLE
    # Interactively pick which packages get a take-over action at all; the grid
    # lets you single- or multi-select, and unselected candidates are skipped.
    .\Create-TakeOverPackagev2.ps1 -SelectPackages

    .NOTES
    Based on Create-TakeOverPackage.ps1 (John Yoakum, Recast Software).
    Rerun/idempotency support added.
#>
param(
    [switch]$AddUnmanaged,
    [switch]$SelectPackages = $true,
    [switch]$AddProcessNotRunningFilter = $true,
    [switch]$SelectProcessFilterPackages = $true,
    [switch]$CreateCollections = $true,
    [switch]$CreateEntitlements = $true,
    [switch]$CreateDesktopIcons = $true,
    [switch]$CreateStartMenuIcons = $true,
    [string]$LiquitURL = 'https://chrisa.recastsoftware.cloud',
    [string]$Username = 'LOCAL\aw.api',
    [pscredential]$Credential,
    [string]$TakeOverPackageName = 'Take Over Applicationsv2',
    [string]$SnapshotName = 'Take Over',
    [string]$ActionSetName = 'Take Over Install',
    [string]$LiquitModulePath = 'C:\Program Files (x86)\Liquit Workspace\PowerShell\Liquit.Server.PowerShell.dll',
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

# ---- Parameter validation (same rules as the original) ------------------
if ($CreateEntitlements -and -not $CreateCollections) {
    throw 'The -CreateCollections parameter must be specified when -CreateEntitlements is specified.'
}
if (($CreateDesktopIcons -or $CreateStartMenuIcons) -and -not $CreateEntitlements) {
    throw 'The -CreateEntitlements parameter must be specified when icon creation is enabled.'
}

# ---- Logging ------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $scriptDirectory = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDirectory) -and $MyInvocation.MyCommand.Path) {
        $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    if ([string]::IsNullOrWhiteSpace($scriptDirectory)) {
        $scriptDirectory = (Get-Location).Path
    }
    $LogPath = Join-Path -Path $scriptDirectory -ChildPath ("Create-TakeOverPackage-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
}

$script:Summary = [ordered]@{
    PackagesScanned = 0
    PackagesSkipped = 0
    SkippedUnmanaged      = 0
    SkippedNoProduction   = 0
    SkippedNoLaunchAction = 0
    CandidatesFound = 0
    ActionsCreated  = 0
    ActionsSkipped  = 0
    CollectionsCreated  = 0
    EntitlementsCreated = 0
    Warnings        = 0
    Errors          = 0
}

function Write-Log {
    param(
        [Parameter(Mandatory)][ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level,
        [Parameter(Mandatory)][string]$Message
    )
    $line = '{0:u} [{1}] {2}' -f (Get-Date), $Level, $Message
    $line | Tee-Object -FilePath $LogPath -Append | Write-Host
    if ($Level -eq 'WARN') { $script:Summary.Warnings++ }
    if ($Level -eq 'ERROR') { $script:Summary.Errors++ }
}

function New-TakeOverEntitlementIcons {
    param([bool]$Desktop, [bool]$StartMenu)
    if (-not $Desktop -and -not $StartMenu) { return $null }
    $command = Get-Command New-LiquitPackageEntitlement -ErrorAction Stop
    $iconsParameter = $command.Parameters['Icons']
    if (-not $iconsParameter -or -not $iconsParameter.ParameterType) {
        throw 'New-LiquitPackageEntitlement does not expose a usable -Icons parameter.'
    }
    $icons = [Activator]::CreateInstance($iconsParameter.ParameterType)
    $icons.Desktop = $Desktop
    $icons.StartMenu = $StartMenu
    return $icons
}

# ---- Load module (same as original) -------------------------------------
if (-not (Test-Path $LiquitModulePath)) {
    if ($MyInvocation.MyCommand.Source) {
        $LiquitModulePath = (Split-Path -Parent $MyInvocation.MyCommand.Definition) + '\Liquit.Server.PowerShell.dll'
    }
    if (-not (Test-Path $LiquitModulePath)) {
        throw 'Unable to find Liquit.Server.PowerShell.dll'
    }
}
Import-Module $LiquitModulePath -Global

[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null

# ---- Credentials --------------------------------------------------------
if (-not $Credential) {
    $Credential = Get-Credential -UserName $Username -Message "Enter credentials for $Username"
}
if (-not $Credential) { throw 'Credential prompt was canceled.' }

Connect-LiquitWorkspace -URI $LiquitURL -Credential $Credential
Write-Log INFO "Connected to $LiquitURL."

# ---- Discover one eligible processstart launch action per package -------
$TakeOverCommands = [System.Collections.ArrayList]::new()
$AllPackages = Get-LiquitPackage

foreach ($Package in $AllPackages) {
    $script:Summary.PackagesScanned++
    try {
        $Attribute = Get-LiquitAttribute -Entity $Package
        if ($Attribute) { $Managed = $true } else { $Managed = $false }

        if (-not $Managed -and -not $AddUnmanaged) {
            $script:Summary.PackagesSkipped++
            $script:Summary.SkippedUnmanaged++
            Write-Log INFO "Skipped package '$($Package.Name)' because it is unmanaged (use -AddUnmanaged to include)."
            continue
        }

        $Snapshot = Get-LiquitPackageSnapshot -Package $Package | Where-Object { $_.Type -eq 'Production' } | Select-Object -First 1
        if (-not $Snapshot) {
            $script:Summary.PackagesSkipped++
            $script:Summary.SkippedNoProduction++
            Write-Log INFO "Skipped package '$($Package.Name)' because it has no Production snapshot."
            continue
        }

        $ActionSets = Get-LiquitActionSet -Snapshot $Snapshot | Where-Object { $_.Type -eq 'Launch' }
        $validActionFound = $false

        foreach ($Actionset in $ActionSets) {
            $Actions = Get-LiquitAction -ActionSet $Actionset | Where-Object { $_.Type -eq 'processstart' }
            foreach ($Action in $Actions) {
                $directory = [string]$Action.Settings.directory
                $name = [string]$Action.Settings.name
                if ([string]::IsNullOrWhiteSpace($name)) { continue }

                # The full path may live entirely in 'name' (directory empty), or be
                # split across 'directory' + 'name'. Join only when a directory exists
                # so Join-Path does not throw on an empty -Path.
                if ([string]::IsNullOrWhiteSpace($directory)) {
                    $filePath = $name
                }
                else {
                    $filePath = Join-Path -Path $directory -ChildPath $name
                }

                # The process-not-running filter needs just the executable name.
                $processName = Split-Path -Path $filePath -Leaf

                $systemPaths = @(
                    "$env:SystemRoot\System32",
                    "$env:SystemRoot\SysWOW64",
                    "$env:SystemRoot"
                )
                $isSystemFile = $false
                foreach ($path in $systemPaths) {
                    if ($filePath.ToLower().StartsWith($path.ToLower())) {
                        $isSystemFile = $true
                        break
                    }
                }
                if ($isSystemFile) {
                    Write-Log INFO "Skipped a launch action in '$($Package.Name)' because it points to a system executable: $filePath"
                    continue
                }

                $displayName = if ($Package.DisplayName) { [string]$Package.DisplayName } else { [string]$Package.Name }

                $NewPackage = [PSCustomObject]@{
                    PackageID   = [string]$Package.ID
                    PackageName = [string]$Package.Name
                    DisplayName = $displayName
                    PathToFile  = $directory
                    FileName    = $processName
                    FullPath    = [string]$filePath
                    Managed     = $Managed
                }
                [void]$TakeOverCommands.Add($NewPackage)
                $validActionFound = $true
                break
            }
            if ($validActionFound) { break }
        }

        if (-not $validActionFound) { 
            $script:Summary.PackagesSkipped++
            $script:Summary.SkippedNoLaunchAction++
            Write-Log INFO "Skipped package '$($Package.Name)' because it has no eligible (non-system) processstart launch action."
        }
    }
    catch {
        Write-Log ERROR "Discovery failed for package '$([string]$Package.Name)': $($_.Exception.Message)"
    }
}

$script:Summary.CandidatesFound = $TakeOverCommands.Count
Write-Log INFO "Found $($TakeOverCommands.Count) takeover candidates."

# ---- Optionally let the admin choose which packages get a take-over -----
#      action at all (single- or multi-select grid). Unselected candidates
#      are dropped before any create activity.
if ($SelectPackages -and $TakeOverCommands.Count -gt 0) {
    $pickInput = $TakeOverCommands | ForEach-Object {
        [pscustomobject]@{
            DisplayName = $_.DisplayName
            FileName    = $_.FileName
            FullPath    = $_.FullPath
            PackageName = $_.PackageName
            PackageID   = $_.PackageID
        }
    }

    $picked = @($pickInput | Out-GridView -Title 'Select which packages to create take-over actions for (Ctrl/Shift for multi-select, then OK). Cancel = none.' -OutputMode Multiple)
    $pickedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $picked) { [void]$pickedIds.Add([string]$p.PackageID) }

    $filtered = [System.Collections.ArrayList]::new()
    foreach ($cmd in $TakeOverCommands) {
        if ($pickedIds.Contains([string]$cmd.PackageID)) { [void]$filtered.Add($cmd) }
    }
    $TakeOverCommands = $filtered

    $script:Summary.CandidatesFound = $TakeOverCommands.Count
    Write-Log INFO "Admin selected $($TakeOverCommands.Count) package(s) to create take-over actions for."
}

# ---- Optionally let the admin choose which packages get the -------------
#      process-not-running filter (single- or multi-select grid).
# Every candidate defaults to the value of -AddProcessNotRunningFilter, then
# the grid selection (if used) overrides which ones actually receive it.
$processFilterIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

if ($AddProcessNotRunningFilter -and $SelectProcessFilterPackages -and $TakeOverCommands.Count -gt 0) {
    $gridInput = $TakeOverCommands | ForEach-Object {
        [pscustomobject]@{
            DisplayName = $_.DisplayName
            FileName    = $_.FileName
            FullPath    = $_.FullPath
            PackageName = $_.PackageName
            PackageID   = $_.PackageID
        }
    }

    $selection = @($gridInput | Out-GridView -Title 'Select packages to receive the "process not running" filter (Ctrl/Shift for multi-select, then OK). Cancel = none.' -OutputMode Multiple)

    foreach ($sel in $selection) { [void]$processFilterIds.Add([string]$sel.PackageID) }
    Write-Log INFO "Admin selected $($processFilterIds.Count) package(s) for the process-not-running filter."
}
elseif ($AddProcessNotRunningFilter) {
    # No interactive selection: apply the filter to every candidate (original behavior).
    foreach ($cmd in $TakeOverCommands) { [void]$processFilterIds.Add([string]$cmd.PackageID) }
}

# ---- Find or create package / snapshot / install action set -------------
$targetPackageName = "TO - $TakeOverPackageName"

$AWPackage = Get-LiquitPackage -Name $targetPackageName -ErrorAction SilentlyContinue |
    Where-Object { [string]$_.Name -eq $targetPackageName } |
    Select-Object -First 1

if (-not $AWPackage) {
    $AWPackage = New-LiquitPackage -Name $targetPackageName -Type 'Custom' -DisplayName $TakeOverPackageName -Priority 100 -Enabled $true -Web $false
    Write-Log INFO "Created package '$targetPackageName'."
}
else {
    Write-Log INFO "Found existing package '$targetPackageName'."
}

$AWSnapshot = Get-LiquitPackageSnapshot -Package $AWPackage | Where-Object { $_.Name -eq $SnapshotName } | Select-Object -First 1
if (-not $AWSnapshot) {
    $AWSnapshot = New-LiquitPackageSnapshot -Package $AWPackage -Name $SnapshotName
    Write-Log INFO "Created snapshot '$SnapshotName'."
}
else {
    Write-Log INFO "Found existing snapshot '$SnapshotName'."
}

$ActionSet = Get-LiquitActionSet -Snapshot $AWSnapshot |
    Where-Object { $_.Type -eq 'Install' -and $_.Name -eq $ActionSetName } |
    Select-Object -First 1

if (-not $ActionSet) {
    $ActionSet = New-LiquitActionSet -Snapshot $AWSnapshot -Type Install -Name $ActionSetName -Enabled $true -Frequency OncePerDevice -Process Sequential
    Write-Log INFO "Created action set '$ActionSetName'."
}
else {
    Write-Log INFO "Found existing action set '$ActionSetName'."
}

# ---- Read existing action names so we can skip work already done --------
$existingActionNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($existing in @(Get-LiquitAction -ActionSet $ActionSet)) {
    $name = [string]$existing.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$existingActionNames.Add($name) }
}
Write-Log INFO "Loaded $($existingActionNames.Count) existing action names."

# ---- Prepare entitlement icons once (if requested) ----------------------
$icons = $null
if ($CreateEntitlements -and ($CreateDesktopIcons -or $CreateStartMenuIcons)) {
    $icons = New-TakeOverEntitlementIcons -Desktop ([bool]$CreateDesktopIcons) -StartMenu ([bool]$CreateStartMenuIcons)
}

# ---- Create only the missing takeover actions ---------------------------
foreach ($Command in $TakeOverCommands) {
    if (-not $Command.Managed) { continue }

    $ActionName = "Take Over $($Command.DisplayName)"

    try {
        $CurrentPackage = Get-LiquitPackage -ID $Command.PackageID

        # --- Take-over install action (skip if it already exists) ---
        if ($existingActionNames.Contains($ActionName)) {
            $script:Summary.ActionsSkipped++
            Write-Log INFO "Existing action found; skipped '$ActionName'."
        }
        else {
            $Action = New-LiquitAction -ActionSet $ActionSet -Name $ActionName -Type 'installpackage' -Enabled $true -Settings @{ title = "$($Command.DisplayName)"; value = $CurrentPackage.ID }
            $null = New-LiquitAttribute -Entity $Action -Link $CurrentPackage -ID 'package'
            $FilterSet = New-LiquitFilterSet -Action $Action
            $null = New-LiquitFilter -FilterSet $FilterSet -Type fileexists -Settings @{ path = "$($Command.FullPath)" } -Value 'true'

            if ($AddProcessNotRunningFilter -and $processFilterIds.Contains([string]$Command.PackageID)) {
                # Filters within one filter set are OR'd; separate filter sets are AND'd.
                # Put the process-not-running filter in its OWN filter set so the file
                # must exist AND the process must not be running.
                $ProcessFilterSet = New-LiquitFilterSet -Action $Action
                $null = New-LiquitFilter -FilterSet $ProcessFilterSet -Type processexists -Settings @{ name = "$($Command.FileName)" } -Value 'false'
                Write-Log INFO "Added process-not-running filter for '$($Command.FileName)' to '$ActionName'."
            }

            [void]$existingActionNames.Add($ActionName)
            $script:Summary.ActionsCreated++
            Write-Log INFO "Created '$ActionName'."
        }

        # --- User collection (find or create) ---
        $collection = $null
        if ($CreateCollections) {
            $collection = Get-LiquitUserCollection -Name $Command.DisplayName -ErrorAction SilentlyContinue |
                Where-Object { [string]$_.Name -eq $Command.DisplayName } |
                Select-Object -First 1
            if (-not $collection) {
                $collection = New-LiquitUserCollection -Name $Command.DisplayName
                $script:Summary.CollectionsCreated++
                Write-Log INFO "Created user collection '$($Command.DisplayName)'."
            }
        }

        # --- Entitlement + move-user action ---
        if ($CreateEntitlements -and $collection) {
            # A brand-new collection's identity can take a moment to become
            # queryable. Retry a few times before giving up.
            $identity = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                $identity = Get-LiquitIdentity -Name $collection.Name -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($identity) { break }
                Start-Sleep -Milliseconds 750
            }
            if (-not $identity) {
                Write-Log WARN "No identity found for collection '$($collection.Name)'; skipped entitlement for '$($Command.DisplayName)'."
            }
            else {
                # Create entitlement; treat an explicit duplicate response as an
                # idempotent skip, and retry the transient "resource does not exist"
                # error that can occur while a new identity is still committing.
                $entitlementParams = @{
                    Package  = $CurrentPackage
                    Publish  = 'Workspace'
                    Stage    = 'Production'
                    Identity = $identity
                }
                if ($icons) { $entitlementParams.Icons = $icons }

                $entitlementDone = $false
                for ($attempt = 1; $attempt -le 5 -and -not $entitlementDone; $attempt++) {
                    try {
                        $null = New-LiquitPackageEntitlement @entitlementParams
                        $script:Summary.EntitlementsCreated++
                        Write-Log INFO "Created entitlement for '$($Command.DisplayName)'."
                        $entitlementDone = $true
                    }
                    catch {
                        $msg = $_.Exception.Message
                        if ($msg -match '(?i)already exists|duplicate|conflict') {
                            Write-Log INFO "Entitlement already exists for '$($Command.DisplayName)'; skipped."
                            $entitlementDone = $true
                        }
                        elseif ($msg -match '(?i)does not exist or one of its queried reference' -and $attempt -lt 5) {
                            Write-Log INFO "Entitlement identity not ready for '$($Command.DisplayName)'; retry $attempt."
                            Start-Sleep -Milliseconds 1000
                        }
                        else {
                            Write-Log ERROR "Entitlement failed for '$($Command.DisplayName)': $msg"
                            $entitlementDone = $true
                        }
                    }
                }

                # Move-user action (skip if it already exists).
                $memberActionName = "Move User to $($Command.DisplayName)"
                if ($existingActionNames.Contains($memberActionName)) {
                    Write-Log INFO "Existing move-user action found; skipped '$memberActionName'."
                }
                else {
                    $AddUserAction = New-LiquitAction -ActionSet $ActionSet -Type 'identitymember' -Name $memberActionName -Settings @{
                        add   = $true
                        title = "$($Command.DisplayName)"
                        group = $identity.ID
                    } -Context Server
                    $UserFilterSet = New-LiquitFilterSet -Action $AddUserAction
                    $UserFilter = New-LiquitFilter -FilterSet $UserFilterSet -Type 'packageinstalled' -Settings @{ title = "$($CurrentPackage.Name)"; value = "$($CurrentPackage.ID)" } -Value '0'
                    $null = New-LiquitAttribute -Entity $UserFilter -Id 'package' -Link $CurrentPackage
                    [void]$existingActionNames.Add($memberActionName)
                    Write-Log INFO "Created '$memberActionName'."
                }
            }
        }
    }
    catch {
        Write-Log ERROR "Failed to process '$($Command.DisplayName)': $($_.Exception.Message)"
    }
}

# ---- Summary ------------------------------------------------------------
Write-Host ''
Write-Host 'Run summary'
[pscustomobject]$script:Summary | Format-List | Out-String | Tee-Object -FilePath $LogPath -Append | Write-Host

if ($script:Summary.Errors -gt 0) { exit 1 }
exit 0
