<#
.SYNOPSIS
    Synchronizes device members of selected Microsoft Entra ID groups
    to Application Workspace device collections.

.DESCRIPTION
    The script enumerates Entra ID groups and presents a UI for selecting
    groups to synchronize.

    Selected groups are saved to selected-groups.json. Subsequent runs can
    use -UseSaved to bypass the UI.

    Performance optimizations:
      - Uses Get-MgGroupMemberAsDevice so only device members are returned.
      - Loads Application Workspace devices once.
      - Creates case-insensitive lookup tables for Application Workspace devices.
      - Uses ArrayList objects to store mutable collections.
      - Uses HashSet objects for fast membership comparisons.
      - Does not query every group's members while building the UI.
      - Avoids repeated Get-LiquitDevice -Search calls.
      - Requests only required properties from Microsoft Graph.

    Required Microsoft Graph application permissions:
      - GroupMember.Read.All
      - Group.Read.All
      - Device.Read.All

    Directory.Read.All can also provide the needed directory read access,
    but use the least-privileged permissions appropriate for your environment.

.PARAMETER UseSaved
    Skips the UI and processes groups stored in selected-groups.json.

.EXAMPLE
    .\Sync-EntraDeviceGroups.ps1

.EXAMPLE
    .\Sync-EntraDeviceGroups.ps1 -UseSaved
#>

[CmdletBinding()]
param(
    [switch]$UseSaved
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Microsoft Entra ID application registration
$TenantId     = 'tenantID'
$ClientId     = 'ApplicationID'
$ClientSecret = 'CLIENTSecret'

# Application Workspace access
$LiquitURI      = 'https://zone.fqdn.com'
$LiquitUsername = 'local\SERVICEACCOUNT'
$LiquitPassword = 'SERVICEACCOUNTPASSWORD'

# Saved group selection
$SavedGroupsFile = Join-Path -Path $PSScriptRoot -ChildPath 'selected-groups.json'

# ---------------------------------------------------------------------------
# Credentials
# ---------------------------------------------------------------------------

$SecureClientSecret = ConvertTo-SecureString `
    -String $ClientSecret `
    -AsPlainText `
    -Force

$GraphCredential = [System.Management.Automation.PSCredential\]::new(
    $ClientId,
    $SecureClientSecret
)

$SecureLiquitPassword = ConvertTo-SecureString `
    -String $LiquitPassword `
    -AsPlainText `
    -Force

$LiquitCredential = [System.Management.Automation.PSCredential\]::new(
    $LiquitUsername,
    $SecureLiquitPassword
)

# ---------------------------------------------------------------------------
# Script-level collections and lookup tables
# ---------------------------------------------------------------------------

# Mutable list containing all Application Workspace devices.
$script:AllDevices = [System.Collections.ArrayList\]::new()

# Case-insensitive device-name lookup.
$script:DevicesByName =
    [System.Collections.Generic.Dictionary[string, object]\]::new(
        [System.StringComparer\]::OrdinalIgnoreCase
    )

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

function Get-ObjectDisplayName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject
    )

    if ($InputObject.PSObject.Properties.Name -contains 'DisplayName') {
        if (-not [string\]::IsNullOrWhiteSpace([string]$InputObject.DisplayName)) {
            return [string]$InputObject.DisplayName
        }
    }

    if ($InputObject.PSObject.Properties.Name -contains 'Name') {
        if (-not [string\]::IsNullOrWhiteSpace([string]$InputObject.Name)) {
            return [string]$InputObject.Name
        }
    }

    if ($InputObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $AdditionalProperties = $InputObject.AdditionalProperties

        if ($null -ne $AdditionalProperties) {
            if ($AdditionalProperties.ContainsKey('displayName')) {
                return [string]$AdditionalProperties['displayName']
            }
        }
    }

    return $null
}

function Initialize-AWDeviceCache {
    [CmdletBinding()]
    param()

    Write-Host 'Loading Application Workspace devices...' -ForegroundColor Cyan

    $script:AllDevices.Clear()
    $script:DevicesByName.Clear()

    $LiquitDevices = @(Get-LiquitDevice)

    foreach ($Device in $LiquitDevices) {
        if ($null -eq $Device) {
            continue
        }

        [void]$script:AllDevices.Add($Device)

        $DeviceName = Get-ObjectDisplayName -InputObject $Device

        if ([string\]::IsNullOrWhiteSpace($DeviceName)) {
            continue
        }

        # If duplicate names exist, retain the first device returned.
        if (-not $script:DevicesByName.ContainsKey($DeviceName)) {
            $script:DevicesByName.Add($DeviceName, $Device)
        }
    }

    Write-Host (
        'Loaded {0:N0} Application Workspace devices.' -f
        $script:AllDevices.Count
    ) -ForegroundColor Green
}

function Get-EntraDeviceMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    $DeviceMembers = [System.Collections.ArrayList\]::new()

    # This asks Graph for device objects only.
    $GraphDevices = @(
        Get-MgGroupMemberAsDevice `
            -GroupId $GroupId `
            -Property 'id,displayName' `
            -All
    )

    foreach ($GraphDevice in $GraphDevices) {
        if ($null -eq $GraphDevice) {
            continue
        }

        $DeviceName = Get-ObjectDisplayName -InputObject $GraphDevice

        if ([string\]::IsNullOrWhiteSpace($DeviceName)) {
            continue
        }

        $DeviceRecord = [PSCustomObject]@{
            Id          = [string]$GraphDevice.Id
            DisplayName = $DeviceName
        }

        [void]$DeviceMembers.Add($DeviceRecord)
    }

    return $DeviceMembers
}

function Get-OrCreateAWDeviceCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CollectionName
    )

    $Matches = @(Get-LiquitDeviceCollection -Search $CollectionName)

    # Prefer an exact, case-insensitive name match.
    $CurrentCollection = $Matches |
        Where-Object {
            [string\]::Equals(
                [string]$_.Name,
                $CollectionName,
                [System.StringComparison\]::OrdinalIgnoreCase
            )
        } |
        Select-Object -First 1

    if ($null -eq $CurrentCollection) {
        Write-Host "Creating collection: $CollectionName" -ForegroundColor Yellow

        $CreatedCollection = New-LiquitDeviceCollection -Name $CollectionName

        if ($null -ne $CreatedCollection) {
            $CurrentCollection = $CreatedCollection
        }
        else {
            $Matches = @(Get-LiquitDeviceCollection -Search $CollectionName)

            $CurrentCollection = $Matches |
                Where-Object {
                    [string\]::Equals(
                        [string]$_.Name,
                        $CollectionName,
                        [System.StringComparison\]::OrdinalIgnoreCase
                    )
                } |
                Select-Object -First 1
        }
    }

    if ($null -eq $CurrentCollection) {
        throw "Unable to find or create collection '$CollectionName'."
    }

    return $CurrentCollection
}

function Update-AWCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Group
    )

    $GroupId = [string]$Group.Id
    $GroupDisplayName = [string]$Group.DisplayName

    if ([string\]::IsNullOrWhiteSpace($GroupId)) {
        throw 'The supplied group does not contain an Id.'
    }

    if ([string\]::IsNullOrWhiteSpace($GroupDisplayName)) {
        throw "The group '$GroupId' does not contain a DisplayName."
    }

    Write-Host ''
    Write-Host "Processing group: $GroupDisplayName" -ForegroundColor Cyan
    Write-Host "Group ID: $GroupId" -ForegroundColor DarkGray

    # -----------------------------------------------------------------------
    # Get Entra device members
    # -----------------------------------------------------------------------

    $EntraDevices = [System.Collections.ArrayList\]::new()

    $RetrievedDevices = Get-EntraDeviceMembers -GroupId $GroupId

    foreach ($RetrievedDevice in $RetrievedDevices) {
        [void]$EntraDevices.Add($RetrievedDevice)
    }

    # HashSet provides fast case-insensitive membership checks.
    $EntraDeviceNames =
        [System.Collections.Generic.HashSet[string]\]::new(
            [System.StringComparer\]::OrdinalIgnoreCase
        )

    foreach ($EntraDevice in $EntraDevices) {
        if (-not [string\]::IsNullOrWhiteSpace($EntraDevice.DisplayName)) {
            [void]$EntraDeviceNames.Add([string]$EntraDevice.DisplayName)
        }
    }

    Write-Host (
        'Entra device members found: {0:N0}' -f $EntraDeviceNames.Count
    )

    # -----------------------------------------------------------------------
    # Get or create Application Workspace collection
    # -----------------------------------------------------------------------

    $CurrentCollection = Get-OrCreateAWDeviceCollection `
        -CollectionName $GroupDisplayName

    # -----------------------------------------------------------------------
    # Get current Application Workspace collection membership
    # -----------------------------------------------------------------------

    $CurrentCollectionMembers = [System.Collections.ArrayList\]::new()

    $RetrievedCollectionMembers = @(
        Get-LiquitDeviceCollectionMember `
            -DeviceCollection $CurrentCollection
    )

    foreach ($CollectionMember in $RetrievedCollectionMembers) {
        if ($null -eq $CollectionMember) {
            continue
        }

        [void]$CurrentCollectionMembers.Add($CollectionMember)
    }

    $CurrentCollectionMemberNames =
        [System.Collections.Generic.HashSet[string]\]::new(
            [System.StringComparer\]::OrdinalIgnoreCase
        )

    $CurrentCollectionMemberByName =
        [System.Collections.Generic.Dictionary[string, object]\]::new(
            [System.StringComparer\]::OrdinalIgnoreCase
        )

    foreach ($CollectionMember in $CurrentCollectionMembers) {
        $MemberName = Get-ObjectDisplayName -InputObject $CollectionMember

        if ([string\]::IsNullOrWhiteSpace($MemberName)) {
            continue
        }

        [void]$CurrentCollectionMemberNames.Add($MemberName)

        if (-not $CurrentCollectionMemberByName.ContainsKey($MemberName)) {
            $CurrentCollectionMemberByName.Add(
                $MemberName,
                $CollectionMember
            )
        }
    }

    Write-Host (
        'Current collection members: {0:N0}' -f
        $CurrentCollectionMemberNames.Count
    )

    # -----------------------------------------------------------------------
    # Calculate differences
    # -----------------------------------------------------------------------

    $DevicesToAdd    = [System.Collections.ArrayList\]::new()
    $DevicesToRemove = [System.Collections.ArrayList\]::new()
    $MissingAWDevices = [System.Collections.ArrayList\]::new()

    foreach ($ExistingName in $CurrentCollectionMemberNames) {
        if (-not $EntraDeviceNames.Contains($ExistingName)) {
            [void]$DevicesToRemove.Add($ExistingName)
        }
    }

    foreach ($EntraName in $EntraDeviceNames) {
        if ($CurrentCollectionMemberNames.Contains($EntraName)) {
            continue
        }

        if ($script:DevicesByName.ContainsKey($EntraName)) {
            [void]$DevicesToAdd.Add($EntraName)
        }
        else {
            [void]$MissingAWDevices.Add($EntraName)
        }
    }

    Write-Host (
        'Changes: {0:N0} add, {1:N0} remove, {2:N0} not registered in Application Workspace.' -f
        $DevicesToAdd.Count,
        $DevicesToRemove.Count,
        $MissingAWDevices.Count
    )

    # -----------------------------------------------------------------------
    # Remove devices no longer in Entra group
    # -----------------------------------------------------------------------

    foreach ($DeviceName in $DevicesToRemove) {
        $CurrentDevice = $null

        # Prefer the global Application Workspace cache.
        if ($script:DevicesByName.ContainsKey($DeviceName)) {
            $CurrentDevice = $script:DevicesByName[$DeviceName]
        }
        elseif ($CurrentCollectionMemberByName.ContainsKey($DeviceName)) {
            # The collection-member object might itself be accepted by the cmdlet.
            $CurrentDevice = $CurrentCollectionMemberByName[$DeviceName]
        }
        else {
            # Last-resort search for unusual/stale collection entries.
            $SearchResults = @(Get-LiquitDevice -Search $DeviceName)

            $CurrentDevice = $SearchResults |
                Where-Object {
                    [string\]::Equals(
                        [string]$_.Name,
                        [string]$DeviceName,
                        [System.StringComparison\]::OrdinalIgnoreCase
                    )
                } |
                Select-Object -First 1
        }

        if ($null -eq $CurrentDevice) {
            Write-Warning (
                "Unable to resolve Application Workspace device '$DeviceName' " +
                "for removal from '$GroupDisplayName'."
            )
            continue
        }

        Remove-LiquitDeviceCollectionMember `
            -DeviceCollection $CurrentCollection `
            -Device $CurrentDevice

        Write-Host (
            "Removed $DeviceName from $GroupDisplayName"
        ) -ForegroundColor Yellow
    }

    # -----------------------------------------------------------------------
    # Add newly assigned devices
    # -----------------------------------------------------------------------

    foreach ($DeviceName in $DevicesToAdd) {
        $CurrentDevice = $script:DevicesByName[$DeviceName]

        Add-LiquitDeviceCollectionMember `
            -DeviceCollection $CurrentCollection `
            -Device $CurrentDevice

        Write-Host (
            "Added $DeviceName to $GroupDisplayName"
        ) -ForegroundColor Green
    }

    if ($MissingAWDevices.Count -gt 0) {
        Write-Warning (
            '{0:N0} Entra devices were not found in Application Workspace ' +
            'and could not be added to collection "{1}".'
        ) -f $MissingAWDevices.Count, $GroupDisplayName
    }

    return [PSCustomObject]@{
        GroupId               = $GroupId
        GroupName             = $GroupDisplayName
        EntraDeviceCount      = $EntraDeviceNames.Count
        ExistingMemberCount   = $CurrentCollectionMemberNames.Count
        AddedCount            = $DevicesToAdd.Count
        RemovedCount          = $DevicesToRemove.Count
        MissingAWDeviceCount  = $MissingAWDevices.Count
    }
}

function Get-AllEntraGroups {
    [CmdletBinding()]
    param()

    Write-Host 'Enumerating Entra ID groups...' -ForegroundColor Cyan

    $Groups = [System.Collections.ArrayList\]::new()

    $GraphGroups = @(
        Get-MgGroup `
            -All `
            -Property 'id,displayName,groupTypes,securityEnabled'
    )

    foreach ($GraphGroup in $GraphGroups) {
        if ($null -eq $GraphGroup) {
            continue
        }

        if ([string\]::IsNullOrWhiteSpace([string]$GraphGroup.DisplayName)) {
            continue
        }

        $GroupRecord = [PSCustomObject]@{
            Id              = [string]$GraphGroup.Id
            DisplayName     = [string]$GraphGroup.DisplayName
            GroupTypes      = $GraphGroup.GroupTypes
            SecurityEnabled = $GraphGroup.SecurityEnabled
        }

        [void]$Groups.Add($GroupRecord)
    }

    Write-Host (
        'Loaded {0:N0} Entra ID groups.' -f $Groups.Count
    ) -ForegroundColor Green

    return $Groups
}

function Save-SelectedGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Groups,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $GroupsToSave = [System.Collections.ArrayList\]::new()

    foreach ($Group in $Groups) {
        $SavedGroup = [PSCustomObject]@{
            Id          = [string]$Group.Id
            DisplayName = [string]$Group.DisplayName
        }

        [void]$GroupsToSave.Add($SavedGroup)
    }

    # -InputObject prevents a one-item ArrayList from being written as a
    # single JSON object instead of an array.
    ConvertTo-Json `
        -InputObject $GroupsToSave `
        -Depth 4 |
        Set-Content `
            -Path $Path `
            -Encoding UTF8
}

function Get-SavedGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Saved groups file not found: $Path"
    }

    $RawJson = Get-Content `
        -LiteralPath $Path `
        -Raw

    if ([string\]::IsNullOrWhiteSpace($RawJson)) {
        return [System.Collections.ArrayList\]::new()
    }

    $ParsedGroups = @($RawJson | ConvertFrom-Json)
    $SavedGroups = [System.Collections.ArrayList\]::new()

    foreach ($Group in $ParsedGroups) {
        if ($null -eq $Group) {
            continue
        }

        if ([string\]::IsNullOrWhiteSpace([string]$Group.Id)) {
            Write-Warning 'A saved group entry was skipped because it has no Id.'
            continue
        }

        $SavedGroup = [PSCustomObject]@{
            Id          = [string]$Group.Id
            DisplayName = [string]$Group.DisplayName
        }

        [void]$SavedGroups.Add($SavedGroup)
    }

    return $SavedGroups
}

function Invoke-GroupSynchronization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Groups
    )

    $Results = [System.Collections.ArrayList\]::new()

    foreach ($Group in $Groups) {
        try {
            $Result = Update-AWCollection -Group $Group
            [void]$Results.Add($Result)
        }
        catch {
            Write-Error (
                "Failed to synchronize group '$($Group.DisplayName)' " +
                "($($Group.Id)): $($_.Exception.Message)"
            )
        }
    }

    Write-Host ''
    Write-Host 'Synchronization summary' -ForegroundColor Cyan
    Write-Host '-----------------------' -ForegroundColor Cyan

    foreach ($Result in $Results) {
        Write-Host (
            '{0}: Entra={1:N0}, Added={2:N0}, Removed={3:N0}, Missing={4:N0}' -f
            $Result.GroupName,
            $Result.EntraDeviceCount,
            $Result.AddedCount,
            $Result.RemovedCount,
            $Result.MissingAWDeviceCount
        )
    }

    return $Results
}

# ---------------------------------------------------------------------------
# Connect
# ---------------------------------------------------------------------------

Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor -ForegroundColor Cyan

Connect-LiquitWorkspace `
    -URI $LiquitURI `
    -Credential $LiquitCredential `
    -ErrorAction Stop

# Load this once for the entire run.
Initialize-AWDeviceCache

# ---------------------------------------------------------------------------
# Saved mode
# ---------------------------------------------------------------------------

if ($UseSaved) {
    $SavedGroups = Get-SavedGroups -Path $SavedGroupsFile

    if ($SavedGroups.Count -eq 0) {
        Write-Warning "Saved groups file is empty: $SavedGroupsFile"
        return
    }

    Write-Host ''
    Write-Host (
        'Running in -UseSaved mode. Loaded {0:N0} groups from {1}' -f
        $SavedGroups.Count,
        $SavedGroupsFile
    ) -ForegroundColor Cyan

    foreach ($Group in $SavedGroups) {
        Write-Host (
            ' - {0} ({1})' -f $Group.DisplayName, $Group.Id
        )
    }

    Invoke-GroupSynchronization -Groups $SavedGroups

    Write-Host ''
    Write-Host 'Completed processing saved groups.' -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# Interactive mode
# ---------------------------------------------------------------------------

$DeviceGroups = Get-AllEntraGroups

if ($DeviceGroups.Count -eq 0) {
    Write-Warning 'No Entra ID groups were returned.'
    return
}

Add-Type -AssemblyName PresentationFramework

[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Select Device-based Entra Groups"
    Height="700"
    Width="900"
    MinHeight="500"
    MinWidth="700"
    ResizeMode="CanResize"
    Background="#FF2D2D30"
    WindowStartupLocation="CenterScreen">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock
            Grid.Row="0"
            Text="Select the Entra groups to synchronize with Application Workspace device collections:"
            Foreground="White"
            FontSize="14"
            Margin="6"/>

        <TextBlock
            Grid.Row="1"
            Text="Member counts are not loaded here because retrieving them for every group significantly slows the initial group list."
            Foreground="#FFBEBEBE"
            FontSize="12"
            TextWrapping="Wrap"
            Margin="6,0,6,6"/>

        <ListBox
            Grid.Row="2"
            Name="GroupList"
            Margin="6"
            Background="#FF252526"
            BorderBrush="#FF555555"
            ScrollViewer.VerticalScrollBarVisibility="Auto"
            VirtualizingStackPanel.IsVirtualizing="True"
            VirtualizingStackPanel.VirtualizationMode="Recycling"/>

        <StackPanel
            Grid.Row="3"
            Orientation="Horizontal"
            HorizontalAlignment="Right"
            Margin="6">

            <Button
                Name="RefreshButton"
                Width="100"
                Height="30"
                Margin="4"
                Content="Refresh"/>

            <Button
                Name="SaveSelection"
                Width="140"
                Height="30"
                Margin="4"
                Content="Save &amp; Continue"/>

            <Button
                Name="CancelButton"
                Width="100"
                Height="30"
                Margin="4"
                Content="Cancel"/>
        </StackPanel>
    </Grid>
</Window>
"@

$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$GroupList  = $Window.FindName('GroupList')
$RefreshBtn = $Window.FindName('RefreshButton')
$SaveBtn    = $Window.FindName('SaveSelection')
$CancelBtn  = $Window.FindName('CancelButton')

function Populate-GroupList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.ListBox]$TargetList,

        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Groups
    )

    $TargetList.Items.Clear()

    foreach ($Group in ($Groups | Sort-Object -Property DisplayName)) {
        $CheckBox = [System.Windows.Controls.CheckBox]::new()
        $CheckBox.Content = $Group.DisplayName
        $CheckBox.Tag = $Group
        $CheckBox.Foreground = [System.Windows.Media.Brushes]::White
        $CheckBox.Margin = [System.Windows.Thickness]::new(2, 2, 2, 2)

        [void]$TargetList.Items.Add($CheckBox)
    }
}

Populate-GroupList `
    -TargetList $GroupList `
    -Groups $DeviceGroups

# ---------------------------------------------------------------------------
# UI events
# ---------------------------------------------------------------------------

$RefreshBtn.Add_Click({
    $RefreshBtn.IsEnabled = $false
    $SaveBtn.IsEnabled = $false

    try {
        $RefreshedGroups = Get-AllEntraGroups

        if ($RefreshedGroups.Count -eq 0) {
            [void][System.Windows.MessageBox]::Show(
                'No Entra ID groups were returned.',
                'Information',
                'OK',
                'Information'
            )

            return
        }

        Populate-GroupList `
            -TargetList $GroupList `
            -Groups $RefreshedGroups
    }
    catch {
        [void][System.Windows.MessageBox]::Show(
            "Refresh failed: $($_.Exception.Message)",
            'Error',
            'OK',
            'Error'
        )
    }
    finally {
        $RefreshBtn.IsEnabled = $true
        $SaveBtn.IsEnabled = $true
    }
})

$SaveBtn.Add_Click({
    $SelectedGroups = [System.Collections.ArrayList]::new()

    foreach ($Item in $GroupList.Items) {
        if (
            $Item -is [System.Windows.Controls.CheckBox] -and
            $Item.IsChecked -eq $true
        ) {
            [void]$SelectedGroups.Add($Item.Tag)
        }
    }

    if ($SelectedGroups.Count -eq 0) {
        [void][System.Windows.MessageBox]::Show(
            'No groups selected. Select at least one group or select Cancel.',
            'No selection',
            'OK',
            'Warning'
        )

        return
    }

    try {
        Save-SelectedGroups `
            -Groups $SelectedGroups `
            -Path $SavedGroupsFile
    }
    catch {
        [void][System.Windows.MessageBox]::Show(
            "Failed to save the selected groups: $($_.Exception.Message)",
            'Save failed',
            'OK',
            'Error'
        )

        return
    }

    $Window.DialogResult = $true
    $Window.Close()

    try {
        Invoke-GroupSynchronization -Groups $SelectedGroups

        [void][System.Windows.MessageBox]::Show(
            "Processing complete.`n`nSelection saved to:`n$SavedGroupsFile",
            'Done',
            'OK',
            'Information'
        )
    }
    catch {
        [void][System.Windows.MessageBox]::Show(
            "Processing failed: $($_.Exception.Message)",
            'Error',
            'OK',
            'Error'
        )
    }
})

$CancelBtn.Add_Click({
    $Window.DialogResult = $false
    $Window.Close()
})

[void]$Window.ShowDialog()