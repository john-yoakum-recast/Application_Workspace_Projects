<#
.SYNOPSIS
    Synchronizes device members of selected Microsoft Entra ID groups to
    Application Workspace device collections by using direct Microsoft Graph REST calls.

.DESCRIPTION
    Interactive mode displays all Entra groups in a WPF selection window. Selected groups
    are saved to selected-groups.json. Use -UseSaved to skip the UI on later runs.

    This script does not require the Microsoft.Graph PowerShell module.

    App Registration Permissions
    | Permission             | Why it's needed                                                             |
| ---------------------- | --------------------------------------------------------------------------- |
| `Group.Read.All`       | Read all Entra groups so the UI can enumerate groups.                       |
| `GroupMember.Read.All` | Read members of the selected groups.                                        |
| `Device.Read.All`      | Read device objects and retrieve properties such as `id` and `displayName`. |

.PARAMETER UseSaved
    Processes groups stored in selected-groups.json and skips the UI.

.PARAMETER UseTransitiveMembers
    Includes device members found through nested groups. The default is direct membership.
#>

[CmdletBinding()]
param(
    [switch]$UseSaved,
    [switch]$UseTransitiveMembers
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

$TenantId     = 'TENANT-ID'
$ClientId     = 'APPLICATION-ID'
$ClientSecret = 'CLIENT-SECRET'

$LiquitURI      = 'https://zone.fqdn.com'
$LiquitUsername = 'local\SERVICEACCOUNT'
$LiquitPassword = 'SERVICE-ACCOUNT-PASSWORD'

$SavedGroupsFile = Join-Path -Path $PSScriptRoot -ChildPath 'selected-groups.json'
$GraphBaseUri = 'https://graph.microsoft.com/v1.0'
$TokenBaseUri = 'https://login.microsoftonline.com'
$GraphMaximumRetryCount = 5
$GraphDefaultRetryDelaySeconds = 5
$GraphTokenRefreshBufferMinutes = 5

# ---------------------------------------------------------------------------
# Credentials and script-level state
# ---------------------------------------------------------------------------

$SecureLiquitPassword = ConvertTo-SecureString -String $LiquitPassword -AsPlainText -Force
$LiquitCredential = [System.Management.Automation.PSCredential]::new(
    $LiquitUsername,
    $SecureLiquitPassword
)

$script:GraphAccessToken = $null
$script:GraphTokenExpiresUtc = [datetime]::MinValue
$script:AllAWDevices = [System.Collections.ArrayList]::new()
$script:AWDevicesByName = [System.Collections.Generic.Dictionary[string, object]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

function Get-DisplayName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject
    )

    foreach ($PropertyName in @('DisplayName', 'displayName', 'Name')) {
        if ($InputObject.PSObject.Properties.Name -contains $PropertyName) {
            $Value = [string]$InputObject.$PropertyName
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
        }
    }

    if ($InputObject.PSObject.Properties.Name -contains 'AdditionalProperties') {
        $AdditionalProperties = $InputObject.AdditionalProperties
        if (
            $null -ne $AdditionalProperties -and
            $AdditionalProperties.ContainsKey('displayName')
        ) {
            $Value = [string]$AdditionalProperties['displayName']
            if (-not [string]::IsNullOrWhiteSpace($Value)) {
                return $Value
            }
        }
    }

    return $null
}

function Test-Configuration {
    [CmdletBinding()]
    param()

    $ConfigurationErrors = [System.Collections.ArrayList]::new()

    $RequiredValues = @(
        @{ Name = 'TenantId'; Value = $TenantId; Placeholder = 'TENANT-ID' },
        @{ Name = 'ClientId'; Value = $ClientId; Placeholder = 'APPLICATION-ID' },
        @{ Name = 'ClientSecret'; Value = $ClientSecret; Placeholder = 'CLIENT-SECRET' },
        @{ Name = 'LiquitURI'; Value = $LiquitURI; Placeholder = 'https://zone.fqdn.com' },
        @{ Name = 'LiquitUsername'; Value = $LiquitUsername; Placeholder = 'local\SERVICEACCOUNT' },
        @{ Name = 'LiquitPassword'; Value = $LiquitPassword; Placeholder = 'SERVICE-ACCOUNT-PASSWORD' }
    )

    foreach ($RequiredValue in $RequiredValues) {
        if (
            [string]::IsNullOrWhiteSpace([string]$RequiredValue.Value) -or
            [string]$RequiredValue.Value -eq [string]$RequiredValue.Placeholder
        ) {
            [void]$ConfigurationErrors.Add(
                "Set `$${($RequiredValue.Name)} to the correct value."
            )
        }
    }

    if ($ConfigurationErrors.Count -gt 0) {
        throw ($ConfigurationErrors -join [Environment]::NewLine)
    }
}

# ---------------------------------------------------------------------------
# Microsoft Graph authentication and requests
# ---------------------------------------------------------------------------

function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [switch]$ForceRefresh
    )

    $CurrentUtc = [datetime]::UtcNow
    $RefreshThreshold = $script:GraphTokenExpiresUtc.AddMinutes(
        -$GraphTokenRefreshBufferMinutes
    )

    if (
        -not $ForceRefresh -and
        -not [string]::IsNullOrWhiteSpace($script:GraphAccessToken) -and
        $CurrentUtc -lt $RefreshThreshold
    ) {
        return $script:GraphAccessToken
    }

    $TokenUri = '{0}/{1}/oauth2/v2.0/token' -f $TokenBaseUri.TrimEnd('/'), $TenantId
    $TokenBody = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }

    try {
        $TokenResponse = Invoke-RestMethod `
            -Method Post `
            -Uri $TokenUri `
            -Body $TokenBody `
            -ContentType 'application/x-www-form-urlencoded' `
            -ErrorAction Stop
    }
    catch {
        throw "Microsoft Graph authentication failed: $($_.Exception.Message)"
    }

    if ([string]::IsNullOrWhiteSpace([string]$TokenResponse.access_token)) {
        throw 'Microsoft Graph authentication returned no access token.'
    }

    $ExpiresInSeconds = 3599
    if ($null -ne $TokenResponse.expires_in) {
        $ExpiresInSeconds = [int]$TokenResponse.expires_in
    }

    $script:GraphAccessToken = [string]$TokenResponse.access_token
    $script:GraphTokenExpiresUtc = [datetime]::UtcNow.AddSeconds($ExpiresInSeconds)
    return $script:GraphAccessToken
}

function Get-GraphRequestHeaders {
    [CmdletBinding()]
    param(
        [switch]$ForceTokenRefresh
    )

    $AccessToken = Get-GraphAccessToken -ForceRefresh:$ForceTokenRefresh
    return @{
        Authorization = "Bearer $AccessToken"
        Accept        = 'application/json'
    }
}

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [object]$Body,
        [hashtable]$AdditionalHeaders
    )

    $Attempt = 0
    $TokenRefreshAttempted = $false

    while ($Attempt -lt $GraphMaximumRetryCount) {
        $Attempt++
        $Headers = Get-GraphRequestHeaders

        if ($null -ne $AdditionalHeaders) {
            foreach ($HeaderName in $AdditionalHeaders.Keys) {
                $Headers[$HeaderName] = $AdditionalHeaders[$HeaderName]
            }
        }

        $RequestParameters = @{
            Method      = $Method
            Uri         = $Uri
            Headers     = $Headers
            ErrorAction = 'Stop'
        }

        if ($null -ne $Body) {
            $RequestParameters.Body = $Body | ConvertTo-Json -Depth 10 -Compress
            $RequestParameters.ContentType = 'application/json'
        }

        try {
            return Invoke-RestMethod @RequestParameters
        }
        catch {
            $StatusCode = $null
            $RetryAfterSeconds = $GraphDefaultRetryDelaySeconds

            if ($null -ne $_.Exception.Response) {
                try {
                    $StatusCode = [int]$_.Exception.Response.StatusCode
                }
                catch {
                    $StatusCode = $null
                }

                try {
                    $RetryAfterValue = $_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds
                    if ($null -ne $RetryAfterValue) {
                        $RetryAfterSeconds = [int][math]::Ceiling($RetryAfterValue)
                    }
                }
                catch {
                    try {
                        $RetryAfterHeader = $_.Exception.Response.Headers.GetValues('Retry-After') |
                            Select-Object -First 1
                        if ($RetryAfterHeader -match '^\d+$') {
                            $RetryAfterSeconds = [int]$RetryAfterHeader
                        }
                    }
                    catch {
                        $RetryAfterSeconds = $GraphDefaultRetryDelaySeconds
                    }
                }
            }

            if ($StatusCode -eq 401 -and -not $TokenRefreshAttempted) {
                Write-Warning 'Microsoft Graph returned HTTP 401. Refreshing the access token.'
                [void](Get-GraphAccessToken -ForceRefresh)
                $TokenRefreshAttempted = $true
                continue
            }

            if (
                $StatusCode -in @(429, 500, 502, 503, 504) -and
                $Attempt -lt $GraphMaximumRetryCount
            ) {
                if ($RetryAfterSeconds -lt 1) {
                    $RetryAfterSeconds = $GraphDefaultRetryDelaySeconds
                }

                Write-Warning (
                    'Microsoft Graph returned HTTP {0}. Retrying in {1} seconds. Attempt {2} of {3}.' -f
                    $StatusCode,
                    $RetryAfterSeconds,
                    $Attempt,
                    $GraphMaximumRetryCount
                )
                Start-Sleep -Seconds $RetryAfterSeconds
                continue
            }

            throw (
                'Microsoft Graph request failed. Method={0}; URI={1}; HTTP status={2}; Error={3}' -f
                $Method,
                $Uri,
                $StatusCode,
                $_.Exception.Message
            )
        }
    }

    throw "Microsoft Graph request exceeded the maximum retry count. URI: $Uri"
}

function Invoke-GraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [hashtable]$AdditionalHeaders
    )

    $Items = [System.Collections.ArrayList]::new()
    $NextLink = $Uri

    while (-not [string]::IsNullOrWhiteSpace($NextLink)) {
        $Response = Invoke-GraphRequest `
            -Method GET `
            -Uri $NextLink `
            -AdditionalHeaders $AdditionalHeaders

        if (
            $Response.PSObject.Properties.Name -contains 'value' -and
            $null -ne $Response.value
        ) {
            foreach ($Item in $Response.value) {
                [void]$Items.Add($Item)
            }
        }

        $NextLink = $null
        if ($Response.PSObject.Properties.Name -contains '@odata.nextLink') {
            $NextLink = [string]$Response.'@odata.nextLink'
        }
    }

    return $Items
}

function Get-AllEntraGroups {
    [CmdletBinding()]
    param()

    Write-Host 'Enumerating Microsoft Entra ID groups...' -ForegroundColor Cyan
    $Groups = [System.Collections.ArrayList]::new()
    $InitialUri = "$GraphBaseUri/groups?`$select=id,displayName&`$top=999"
    $GraphGroups = Invoke-GraphPagedRequest -Uri $InitialUri

    foreach ($GraphGroup in $GraphGroups) {
        if ($null -eq $GraphGroup) {
            continue
        }

        $GroupId = [string]$GraphGroup.id
        $GroupDisplayName = [string]$GraphGroup.displayName
        if (
            [string]::IsNullOrWhiteSpace($GroupId) -or
            [string]::IsNullOrWhiteSpace($GroupDisplayName)
        ) {
            continue
        }

        [void]$Groups.Add([PSCustomObject]@{
            Id          = $GroupId
            DisplayName = $GroupDisplayName
        })
    }

    Write-Host ('Loaded {0:N0} Microsoft Entra ID groups.' -f $Groups.Count) `
        -ForegroundColor Green
    return $Groups
}

function Get-EntraDeviceMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )

    $DeviceMembers = [System.Collections.ArrayList]::new()
    $EncodedGroupId = [uri]::EscapeDataString($GroupId)
    $MembershipPath = if ($UseTransitiveMembers) { 'transitiveMembers' } else { 'members' }
    $InitialUri = (
        "$GraphBaseUri/groups/$EncodedGroupId/$MembershipPath/" +
        "microsoft.graph.device?`$select=id,displayName&`$top=999"
    )

    $GraphDevices = Invoke-GraphPagedRequest -Uri $InitialUri
    foreach ($GraphDevice in $GraphDevices) {
        if ($null -eq $GraphDevice) {
            continue
        }

        $DeviceDisplayName = [string]$GraphDevice.displayName
        if ([string]::IsNullOrWhiteSpace($DeviceDisplayName)) {
            continue
        }

        [void]$DeviceMembers.Add([PSCustomObject]@{
            Id          = [string]$GraphDevice.id
            DisplayName = $DeviceDisplayName
        })
    }

    return $DeviceMembers
}

# ---------------------------------------------------------------------------
# Application Workspace cache and collection helpers
# ---------------------------------------------------------------------------

function Initialize-AWDeviceCache {
    [CmdletBinding()]
    param()

    Write-Host 'Loading Application Workspace devices...' -ForegroundColor Cyan
    $script:AllAWDevices.Clear()
    $script:AWDevicesByName.Clear()

    foreach ($Device in @(Get-LiquitDevice)) {
        if ($null -eq $Device) {
            continue
        }

        [void]$script:AllAWDevices.Add($Device)
        $DeviceName = Get-DisplayName -InputObject $Device
        if ([string]::IsNullOrWhiteSpace($DeviceName)) {
            continue
        }

        if (-not $script:AWDevicesByName.ContainsKey($DeviceName)) {
            $script:AWDevicesByName.Add($DeviceName, $Device)
        }
    }

    Write-Host ('Loaded {0:N0} Application Workspace devices.' -f $script:AllAWDevices.Count) `
        -ForegroundColor Green
}

function Confirm-AWDeviceCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CollectionName
    )

    $CollectionSearchResults = @(Get-LiquitDeviceCollection -Search $CollectionName)
    $CurrentCollection = $CollectionSearchResults |
        Where-Object {
            [string]::Equals(
                [string]$_.Name,
                $CollectionName,
                [System.StringComparison]::OrdinalIgnoreCase
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
            $CollectionSearchResults = @(Get-LiquitDeviceCollection -Search $CollectionName)
            $CurrentCollection = $CollectionSearchResults |
                Where-Object {
                    [string]::Equals(
                        [string]$_.Name,
                        $CollectionName,
                        [System.StringComparison]::OrdinalIgnoreCase
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

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw 'The supplied group does not contain an ID.'
    }
    if ([string]::IsNullOrWhiteSpace($GroupDisplayName)) {
        throw "The group '$GroupId' does not contain a display name."
    }

    Write-Host ''
    Write-Host "Processing group: $GroupDisplayName" -ForegroundColor Cyan
    Write-Host "Group ID: $GroupId" -ForegroundColor DarkGray
    Write-Host (if ($UseTransitiveMembers) { 'Membership scope: transitive' } else { 'Membership scope: direct' }) `
        -ForegroundColor DarkGray

    $EntraDevices = Get-EntraDeviceMembers -GroupId $GroupId
    $EntraDeviceNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($EntraDevice in $EntraDevices) {
        $DeviceName = [string]$EntraDevice.DisplayName
        if (-not [string]::IsNullOrWhiteSpace($DeviceName)) {
            [void]$EntraDeviceNames.Add($DeviceName)
        }
    }

    Write-Host ('Microsoft Entra device members found: {0:N0}' -f $EntraDeviceNames.Count)
    $CurrentCollection = Confirm-AWDeviceCollection -CollectionName $GroupDisplayName
    $CurrentCollectionMembers = [System.Collections.ArrayList]::new()

    foreach ($CollectionMember in @(
        Get-LiquitDeviceCollectionMember -DeviceCollection $CurrentCollection
    )) {
        if ($null -ne $CollectionMember) {
            [void]$CurrentCollectionMembers.Add($CollectionMember)
        }
    }

    $CurrentCollectionMemberNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $CurrentCollectionMemberByName = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($CollectionMember in $CurrentCollectionMembers) {
        $MemberName = Get-DisplayName -InputObject $CollectionMember
        if ([string]::IsNullOrWhiteSpace($MemberName)) {
            continue
        }

        [void]$CurrentCollectionMemberNames.Add($MemberName)
        if (-not $CurrentCollectionMemberByName.ContainsKey($MemberName)) {
            $CurrentCollectionMemberByName.Add($MemberName, $CollectionMember)
        }
    }

    Write-Host ('Current collection members: {0:N0}' -f $CurrentCollectionMemberNames.Count)

    $DevicesToAdd = [System.Collections.ArrayList]::new()
    $DevicesToRemove = [System.Collections.ArrayList]::new()
    $MissingAWDevices = [System.Collections.ArrayList]::new()

    foreach ($ExistingName in $CurrentCollectionMemberNames) {
        if (-not $EntraDeviceNames.Contains($ExistingName)) {
            [void]$DevicesToRemove.Add($ExistingName)
        }
    }

    foreach ($EntraName in $EntraDeviceNames) {
        if ($CurrentCollectionMemberNames.Contains($EntraName)) {
            continue
        }

        if ($script:AWDevicesByName.ContainsKey($EntraName)) {
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

    foreach ($DeviceName in $DevicesToRemove) {
        $CurrentDevice = $null

        if ($script:AWDevicesByName.ContainsKey($DeviceName)) {
            $CurrentDevice = $script:AWDevicesByName[$DeviceName]
        }
        elseif ($CurrentCollectionMemberByName.ContainsKey($DeviceName)) {
            $CurrentDevice = $CurrentCollectionMemberByName[$DeviceName]
        }
        else {
            $DeviceSearchResults = @(Get-LiquitDevice -Search $DeviceName)
            $CurrentDevice = $DeviceSearchResults |
                Where-Object {
                    [string]::Equals(
                        [string]$_.Name,
                        [string]$DeviceName,
                        [System.StringComparison]::OrdinalIgnoreCase
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
        Write-Host "Removed $DeviceName from $GroupDisplayName" -ForegroundColor Yellow
    }

    foreach ($DeviceName in $DevicesToAdd) {
        $CurrentDevice = $script:AWDevicesByName[$DeviceName]
        Add-LiquitDeviceCollectionMember `
            -DeviceCollection $CurrentCollection `
            -Device $CurrentDevice
        Write-Host "Added $DeviceName to $GroupDisplayName" -ForegroundColor Green
    }

    if ($MissingAWDevices.Count -gt 0) {
        Write-Warning (
            '{0:N0} Entra devices were not found in Application Workspace and could not be added to collection "{1}".' -f
            $MissingAWDevices.Count,
            $GroupDisplayName
        )
    }

    return [PSCustomObject]@{
        GroupId              = $GroupId
        GroupName            = $GroupDisplayName
        EntraDeviceCount     = $EntraDeviceNames.Count
        ExistingMemberCount  = $CurrentCollectionMemberNames.Count
        AddedCount           = $DevicesToAdd.Count
        RemovedCount         = $DevicesToRemove.Count
        MissingAWDeviceCount = $MissingAWDevices.Count
    }
}

function Invoke-GroupSynchronization {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Groups
    )

    $Results = [System.Collections.ArrayList]::new()
    $Failures = [System.Collections.ArrayList]::new()

    foreach ($Group in $Groups) {
        try {
            [void]$Results.Add((Update-AWCollection -Group $Group))
        }
        catch {
            [void]$Failures.Add([PSCustomObject]@{
                GroupId   = [string]$Group.Id
                GroupName = [string]$Group.DisplayName
                Error     = $_.Exception.Message
            })
            Write-Warning (
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

    if ($Failures.Count -gt 0) {
        Write-Warning ('{0:N0} group synchronizations failed.' -f $Failures.Count)
    }

    return [PSCustomObject]@{
        Successful = $Results
        Failed     = $Failures
    }
}

# ---------------------------------------------------------------------------
# Saved group functions
# ---------------------------------------------------------------------------

function Save-SelectedGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Groups,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $GroupsToSave = [System.Collections.ArrayList]::new()
    foreach ($Group in $Groups) {
        [void]$GroupsToSave.Add([PSCustomObject]@{
            Id          = [string]$Group.Id
            DisplayName = [string]$Group.DisplayName
        })
    }

    ConvertTo-Json -InputObject $GroupsToSave -Depth 4 |
        Set-Content -LiteralPath $Path -Encoding UTF8
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

    $SavedGroups = [System.Collections.ArrayList]::new()
    $RawJson = Get-Content -LiteralPath $Path -Raw
    if ([string]::IsNullOrWhiteSpace($RawJson)) {
        return $SavedGroups
    }

    foreach ($Group in @($RawJson | ConvertFrom-Json)) {
        if ($null -eq $Group) {
            continue
        }

        $GroupId = [string]$Group.Id
        if ([string]::IsNullOrWhiteSpace($GroupId)) {
            Write-Warning 'A saved group entry was skipped because it has no ID.'
            continue
        }

        $GroupDisplayName = [string]$Group.DisplayName
        if ([string]::IsNullOrWhiteSpace($GroupDisplayName)) {
            $GroupDisplayName = $GroupId
        }

        [void]$SavedGroups.Add([PSCustomObject]@{
            Id          = $GroupId
            DisplayName = $GroupDisplayName
        })
    }

    return $SavedGroups
}

# ---------------------------------------------------------------------------
# Validate, authenticate, and connect
# ---------------------------------------------------------------------------

Test-Configuration
Write-Host 'Authenticating to Microsoft Graph...' -ForegroundColor Cyan
[void](Get-GraphAccessToken)
Write-Host 'Microsoft Graph authentication succeeded.' -ForegroundColor Green

Write-Host 'Connecting to Application Workspace...' -ForegroundColor Cyan
Connect-LiquitWorkspace `
    -URI $LiquitURI `
    -Credential $LiquitCredential `
    -ErrorAction Stop

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

    Write-Host (
        'Running in saved-group mode. Loaded {0:N0} groups from {1}' -f
        $SavedGroups.Count,
        $SavedGroupsFile
    ) -ForegroundColor Cyan

    [void](Invoke-GroupSynchronization -Groups $SavedGroups)
    Write-Host 'Completed processing saved groups.' -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# Interactive WPF mode
# ---------------------------------------------------------------------------

$EntraGroups = Get-AllEntraGroups
if ($EntraGroups.Count -eq 0) {
    Write-Warning 'No Microsoft Entra ID groups were returned.'
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
            Text="Select Entra groups to synchronize with Application Workspace device collections:"
            Foreground="White"
            FontSize="14"
            Margin="6"/>
        <TextBlock
            Grid.Row="1"
            Text="Member counts are not loaded because querying every group would delay the initial list."
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
            <Button Name="RefreshButton" Width="100" Height="30" Margin="4" Content="Refresh"/>
            <Button Name="SaveSelection" Width="140" Height="30" Margin="4" Content="Save &amp; Continue"/>
            <Button Name="CancelButton" Width="100" Height="30" Margin="4" Content="Cancel"/>
        </StackPanel>
    </Grid>
</Window>
"@

$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)
$GroupList = $Window.FindName('GroupList')
$RefreshButton = $Window.FindName('RefreshButton')
$SaveButton = $Window.FindName('SaveSelection')
$CancelButton = $Window.FindName('CancelButton')

function Update-GroupList {
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

Update-GroupList -TargetList $GroupList -Groups $EntraGroups

$RefreshButton.Add_Click({
    $RefreshButton.IsEnabled = $false
    $SaveButton.IsEnabled = $false

    try {
        $RefreshedGroups = Get-AllEntraGroups
        if ($RefreshedGroups.Count -eq 0) {
            [void][System.Windows.MessageBox]::Show(
                'No Microsoft Entra ID groups were returned.',
                'Information',
                'OK',
                'Information'
            )
            return
        }

        Update-GroupList -TargetList $GroupList -Groups $RefreshedGroups
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
        $RefreshButton.IsEnabled = $true
        $SaveButton.IsEnabled = $true
    }
})

$SaveButton.Add_Click({
    $SelectedGroups = [System.Collections.ArrayList]::new()
    foreach ($Item in $GroupList.Items) {
        if (
            $Item -is [System.Windows.Controls.CheckBox] -and
            $Item.IsChecked
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
        Save-SelectedGroups -Groups $SelectedGroups -Path $SavedGroupsFile
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
        $SynchronizationResult = Invoke-GroupSynchronization -Groups $SelectedGroups
        $CompletionMessage = (
            "Processing complete.`n`n" +
            "Successful groups: $($SynchronizationResult.Successful.Count)`n" +
            "Failed groups: $($SynchronizationResult.Failed.Count)`n`n" +
            "Selection saved to:`n$SavedGroupsFile"
        )
        [void][System.Windows.MessageBox]::Show(
            $CompletionMessage,
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

$CancelButton.Add_Click({
    $Window.DialogResult = $false
    $Window.Close()
})

[void]$Window.ShowDialog()
