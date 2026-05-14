<#
    .SYNOPSIS
    Script to query ConfigMgr Environment and then automatically create a Software Update Process in RCT Patching

    .DESCRIPTION
    This script will query your configMgr environment for all installed applications and then will attempt
    to match to the closest match within the RCT Patching Catalog. Once it has the matches, then it will
    create a Software Update Process within RCT Patching for all those matched apps. This assumes that no previous
    Software Update Processes exist as you can't have the same application in more than one Software Update
    Process. This needs to be ran by an account that has read access to the ConfigMgr Database and 
    Write Access to the RMS database.

    You can hard code the parameters, or you can pass them in the command line


    .EXAMPLE
    .\Create-RCTSoftwareUpdateProcess.ps1

    .NOTES
    Version:       1.1
    Author:        John Yoakum, Recast Software
    Creation Date: 05/12/2026
    Purpose/Change: Updated code to be more dynamic in finding the right Service Connection, also added the ability to add the option to patch only minor versions
#>
param(
    $CMSQLServer = 'demo-mecm.demo.recastsoftware.com', # Enter your ConfigMgr SQL Server FQDN
    $CMDB = 'cm_dm1', # Enter your ConfigMgr Database Name
    $RMSServer = "https://demo-rms-dev.demo.recastsoftware.com:444", # Enter your FQDN of your RMS Server
    $RMSSQLServer = 'demo-rms-dev.demo.recastsoftware.com', # Enter the fqdn of your RMS Database Host
    $RMSDB = 'RecastManagementServer', # Enter the name of your RMS Database
    $patchingProcessName = "Synced Matching Updates", # Enter what you would like to name the Software Update Process
    [switch]$AllowMinorUpgrades = $true
)

Import-Module SQLServer

$ErrorActionPreference = 'Stop'

function Escape-SqlString {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return $Value.Replace("'", "''")
}

function Normalize-AppName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }

    $n = $Name.ToLowerInvariant()
    $n = $n -replace '\(x64\)|\(x86\)', ''
    $n = $n -replace '\bx64\b|\bx86\b|\b64-bit\b|\b32-bit\b', ''
    $n = $n -replace '[®™©]', ''
    $n = $n -replace '\bcorporation\b|\bcorp\b|\bincorporated\b|\binc\b|\bltd\b|\bllc\b|\bcompany\b|\bco\b', ''
    $n = $n -replace '[^a-z0-9]+', ' '
    $n = $n -replace '\s+', ' '

    return $n.Trim()
}

function Get-Tokens {
    param([string]$Text)

    @(Normalize-AppName $Text -split ' ' |
        Where-Object {
            $_.Length -gt 2 -and
            $_ -notin @(
                'the','and','for','with','app','application',
                'setup','installer','runtime','update'
            )
        } |
        Sort-Object -Unique)
}

function Normalize-Version {
    param([string]$Version)

    if ([string]::IsNullOrWhiteSpace($Version)) { return "" }

    return ($Version -replace '[^\d\.]', '').Trim('.')
}

function Get-VersionScore {
    param(
        [string]$SourceVersion,
        [string]$TargetVersion
    )

    $source = Normalize-Version $SourceVersion
    $target = Normalize-Version $TargetVersion

    if (-not $source -or -not $target) { return 50 }
    if ($source -eq $target) { return 100 }

    $s = $source.Split('.')
    $t = $target.Split('.')

    if ($s[0] -eq $t[0]) { return 70 }

    return 0
}

function Get-TokenScore {
    param(
        [string[]]$SourceTokens,
        [string[]]$TargetTokens
    )

    if (-not $SourceTokens -or -not $TargetTokens) { return 0 }

    $sourceSet = @{}

    foreach ($token in $SourceTokens) {
        $sourceSet[$token] = $true
    }

    $common = 0

    foreach ($token in $TargetTokens) {
        if ($sourceSet.ContainsKey($token)) {
            $common++
        }
    }

    $totalUnique = @($SourceTokens + $TargetTokens | Sort-Object -Unique).Count

    if ($totalUnique -eq 0) { return 0 }

    return [math]::Round(($common / $totalUnique) * 100, 2)
}

function Find-AppCatalogMatchVeryFast {
    param(
        [Parameter(Mandatory)]
        [array]$SourceApps,

        [Parameter(Mandatory)]
        [array]$TargetApps,

        [int]$MinimumScore = 85,

        [int]$MaxCandidates = 50
    )

    $processedTargets = @()
    $exactNameIndex = @{}
    $tokenIndex = @{}

    foreach ($target in $TargetApps) {
        $targetDisplayName = "$($target.Publisher) $($target.Name)"
        $normalizedName = Normalize-AppName $targetDisplayName
        $tokens = Get-Tokens $targetDisplayName

        $targetRecord = [pscustomobject]@{
            Id                = $target.Id
            TargetDisplayName = $targetDisplayName
            NormalizedName    = $normalizedName
            Tokens            = $tokens
            LatestVersion     = $target.LatestVersion
            OriginalTarget    = $target
        }

        $processedTargets += $targetRecord

        if (-not $exactNameIndex.ContainsKey($normalizedName)) {
            $exactNameIndex[$normalizedName] = @()
        }

        $exactNameIndex[$normalizedName] += $targetRecord

        foreach ($token in $tokens) {
            if (-not $tokenIndex.ContainsKey($token)) {
                $tokenIndex[$token] = @()
            }

            $tokenIndex[$token] += $targetRecord
        }
    }

    foreach ($source in $SourceApps) {
        $sourceDisplayName = "$($source.Publisher) $($source.DisplayName)"
        $sourceNormalizedName = Normalize-AppName $sourceDisplayName
        $sourceTokens = Get-Tokens $sourceDisplayName

        $candidateCounts = @{}

        if ($exactNameIndex.ContainsKey($sourceNormalizedName)) {
            $candidateTargets = $exactNameIndex[$sourceNormalizedName]
        }
        else {
            foreach ($token in $sourceTokens) {
                if ($tokenIndex.ContainsKey($token)) {
                    foreach ($target in $tokenIndex[$token]) {
                        if (-not $candidateCounts.ContainsKey($target.Id)) {
                            $candidateCounts[$target.Id] = [pscustomobject]@{
                                Target = $target
                                Count  = 0
                            }
                        }

                        $candidateCounts[$target.Id].Count++
                    }
                }
            }

            $candidateTargets = $candidateCounts.Values |
                Sort-Object Count -Descending |
                Select-Object -First $MaxCandidates |
                ForEach-Object { $_.Target }
        }

        if (-not $candidateTargets) {
            continue
        }

        $bestMatch = $null
        $bestScore = 0

        foreach ($target in $candidateTargets) {
            $nameScore = Get-TokenScore `
                -SourceTokens $sourceTokens `
                -TargetTokens $target.Tokens

            if ($nameScore -lt 60) {
                continue
            }

            $versionScore = Get-VersionScore `
                -SourceVersion $source.InstalledVersion `
                -TargetVersion $target.LatestVersion

            $score = [math]::Round(
                ($nameScore * 0.85) + ($versionScore * 0.15),
                2
            )

            if ($score -gt $bestScore) {
                $bestScore = $score

                $bestMatch = [pscustomobject]@{
                    Matched           = $score -ge $MinimumScore
                    Id                = $target.Id
                    SourcePublisher   = $source.Publisher
                    SourceDisplayName = $source.DisplayName
                    SourceVersion     = $source.InstalledVersion
                    TargetDisplayName = $target.TargetDisplayName
                    TargetVersion     = $target.LatestVersion
                    Score             = $score
                    NameScore         = $nameScore
                    VersionScore      = $versionScore
                    TargetObject      = $target.OriginalTarget
                }
            }
        }

        if ($bestMatch -and $bestMatch.Matched) {
            $bestMatch
        }
    }
}

# Get Installed Applications from ConfigMgr

$DeviceInstalledSoftwareQuery = @"
SELECT DISTINCT
    arp.publisher0,
    arp.displayname0,
    arp.version0
FROM v_gs_Add_Remove_programs arp
WHERE arp.displayname0 IS NOT NULL

UNION

SELECT DISTINCT
    arp2.publisher0,
    arp2.displayname0,
    arp2.version0
FROM v_GS_ADD_REMOVE_PROGRAMS_64 arp2
WHERE arp2.displayname0 IS NOT NULL
"@

$UserInstalledSoftwareQuery = @"
SELECT
    Publisher0,
    DisplayName0,
    ProductId0 AS ProdId0,
    Version0,
    InstanceCount
FROM
(
    SELECT
        Publisher0,
        DisplayName0,
        ProductId0,
        Version0,
        COUNT(ResourceID) AS InstanceCount
    FROM v_GS_RecastSoftware_EI_UserArp0
    WHERE Version0 <> '[Not Available]'
      AND DisplayName0 IS NOT NULL
    GROUP BY
        Publisher0,
        DisplayName0,
        ProductId0,
        Version0
) AS UserInstalledApps
"@

$InstalledApps = [System.Collections.ArrayList]::new()

try {
    $IADevices = Invoke-Sqlcmd `
        -ServerInstance $CMSQLServer `
        -Database $CMDB `
        -Query $DeviceInstalledSoftwareQuery `
        -TrustServerCertificate
}
catch {
    Write-Warning "No device apps found or ConfigMgr device query failed: $($_.Exception.Message)"
    $IADevices = @()
}

try {
    $IAUsers = Invoke-Sqlcmd `
        -ServerInstance $CMSQLServer `
        -Database $CMDB `
        -Query $UserInstalledSoftwareQuery `
        -TrustServerCertificate
}
catch {
    Write-Warning "No user apps found or ConfigMgr user query failed: $($_.Exception.Message)"
    $IAUsers = @()
}

foreach ($app in $IADevices) {
    [void]$InstalledApps.Add([pscustomobject]@{
        Publisher        = $app.Publisher0
        DisplayName      = $app.DisplayName0
        InstalledVersion = $app.Version0
    })
}

foreach ($app in $IAUsers) {
    [void]$InstalledApps.Add([pscustomobject]@{
        Publisher        = $app.Publisher0
        DisplayName      = $app.DisplayName0
        InstalledVersion = $app.Version0
    })
}

# Get Available Applications from RMS

$Results = Invoke-RestMethod `
    -Method Post `
    -Uri "$RMSServer/api/application-catalog" `
    -UseDefaultCredentials

$output = $Results.value

# Match ConfigMgr installed apps to RMS application catalog

$matches = Find-AppCatalogMatchVeryFast `
    -SourceApps $InstalledApps `
    -TargetApps $output `
    -MinimumScore 85 `
    -MaxCandidates 50

# Keep one best match per target application Id.
# Important: do NOT remove TargetObject or Matched here.

$AllMatches = $matches |
    Where-Object { $_.Matched -and $_.Id -and $_.TargetObject } |
    Group-Object Id |
    ForEach-Object {
        $_.Group |
            Sort-Object Score -Descending |
            Select-Object -First 1
    }

Write-Host "Matched unique applications: $($AllMatches.Count)"

# Create the Software Update Process

$patchingProcessId = [guid]::NewGuid()

$escapedPatchingProcessName = Escape-SqlString $patchingProcessName

$processSql = @"
INSERT INTO [AM].[PatchingProcesses]
(
    [Id],
    [Name],
    [Settings]
)
VALUES
(
    '$patchingProcessId',
    '$escapedPatchingProcessName',
    '{}'
)
"@

Invoke-Sqlcmd `
    -ServerInstance $RMSSQLServer `
    -Database $RMSDB `
    -Query $processSql `
    -TrustServerCertificate

Write-Host "Created patching process: $patchingProcessName [$patchingProcessId]"

# Get Integration Id for Type 2
$integration = Invoke-Sqlcmd `
    -ServerInstance $RMSSQLServer `
    -Database $RMSDB `
    -Query "SELECT TOP 1 [Id] FROM [AM].[Integrations] WHERE [Type] = 2" `
    -TrustServerCertificate

if (-not $integration -or -not $integration.Id) {
    throw "No integration found in [AM].[Integrations] where Type = 2."
}

$integrationId = $integration.Id

Write-Host "Using IntegrationId: $integrationId"

# Add settings for minor version if chosen

If ($AllowMinorUpgrades) {
    $Settings = '{"AllowOnlyMinorVersionUpgrades":true}'
} else {
    $Settings = $null
}
# Insert the connection to the Service Connection

$integrationSql = @"
INSERT INTO [AM].[IntegrationPatchingProcesses]
           ([IntegrationId]
           ,[PatchingProcessId]
           ,[Status]
           ,[LastRun]
           ,[StatusSet]
           ,[StatusMessage]
           ,[Settings])
     VALUES
           ($integrationId
           ,'$patchingProcessId'
           ,'Paused'
           ,null
           ,null
           ,null
           ,'$Settings')
"@

Invoke-Sqlcmd `
    -ServerInstance $RMSSQLServer `
    -Database $RMSDB `
    -Query $integrationSql `
    -TrustServerCertificate

# Insert matched apps into AM.Applications and associate them to the process

foreach ($match in $AllMatches) {

    $applicationId = $match.Id
    $target = $match.TargetObject

    $publisher = Escape-SqlString $target.Publisher
    $appName = Escape-SqlString $target.Name
    $architecture = Escape-SqlString $target.Architecture
    $language = Escape-SqlString $target.Language
    $iconUrl = Escape-SqlString $target.IconUrl

    $created = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fffffff")
    $latestVersionCheck = $created

    $applicationSql = @"
IF NOT EXISTS (
    SELECT 1
    FROM [AM].[Applications]
    WHERE [Id] = '$applicationId'
)
BEGIN
    INSERT INTO [AM].[Applications]
    (
        [Id],
        [Publisher],
        [Name],
        [Architecture],
        [Language],
        [IconUrl],
        [LatestVersionCheck],
        [Created]
    )
    VALUES
    (
        '$applicationId',
        '$publisher',
        '$appName',
        '$architecture',
        '$language',
        '$iconUrl',
        '$latestVersionCheck',
        '$created'
    )
END
"@

    Invoke-Sqlcmd `
        -ServerInstance $RMSSQLServer `
        -Database $RMSDB `
        -Query $applicationSql `
        -TrustServerCertificate

    $processAppSql = @"
IF NOT EXISTS (
    SELECT 1
    FROM [AM].[PatchingProcessApplications]
    WHERE [PatchingProcessId] = '$patchingProcessId'
      AND [ApplicationId] = '$applicationId'
)
BEGIN
    INSERT INTO [AM].[PatchingProcessApplications]
    (
        [PatchingProcessId],
        [ApplicationId]
    )
    VALUES
    (
        '$patchingProcessId',
        '$applicationId'
    )
END
"@

    Invoke-Sqlcmd `
        -ServerInstance $RMSSQLServer `
        -Database $RMSDB `
        -Query $processAppSql `
        -TrustServerCertificate

    Write-Host "Added application: $publisher - $appName [$applicationId]"
}

Write-Host "Complete. Added $($AllMatches.Count) applications to patching process [$patchingProcessId]."
