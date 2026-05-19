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
    Version:       1.0
    Author:        John Yoakum, Recast Software
    Creation Date: 05/12/2026
    Purpose/Change: Initial script development
#>
param(
    $CMSQLServer = 'demo-mecm.demo.recastsoftware.com', # Enter your ConfigMgr SQL Server FQDN
    $CMDB = 'cm_dm1', # Enter your ConfigMgr Database Name
    $RMSServer = "https://demo-rms-dev.demo.recastsoftware.com:444", # Enter your FQDN of your RMS Server
    $RMSSQLServer = 'demo-rms-dev.demo.recastsoftware.com', # Enter the fqdn of your RMS Database Host
    $RMSDB = 'RecastManagementServer', # Enter the name of your RMS Database
    $patchingProcessName = "Synced Matching Updates", # Enter what you would like to name the Software Update Process
    [switch]$AllowMinorUpgrades = $true,
    $csvFile = $false,
    [int]$MaxCandidates = 200,
    [double]$AutoMatchThreshold = 95.0,
    [double]$StrongReviewThreshold = 85.0,
    [double]$PossibleReviewThreshold = 75.0
)

Import-Module SQLServer

$ErrorActionPreference = 'Stop'

$LegalVendorWords = @{}
"inc incorporated corporation corp company co ltd limited llc gmbh ag sa srl bv plc software technologies technology systems system solutions solution group groups international intl" -split " " | ForEach-Object { $LegalVendorWords[$_] = $true }

$NoiseTokens = @{}
"x86 x64 x32 x86_64 amd64 32bit 64bit 32 64 bit bits edition editions version ver v en us neutral" -split " " | ForEach-Object { $NoiseTokens[$_] = $true }

$GenericTokens = @{}
"suite desktop installer software setup tool tools client server manager management viewer driver application app service services update updater pro standard professional enterprise runtime platform console agent assistant utility utilities extension component components pack language office support" -split " " | ForEach-Object { $GenericTokens[$_] = $true }

function Normalize-Text {
    param(
        [AllowNull()][string]$Text,
        [bool]$IsVendor = $false,
        [bool]$RemoveVersions = $true
    )
    if ($null -eq $Text) { return "" }
    $s = $Text.ToLowerInvariant()
    $s = $s -replace "&", " and "
    $s = $s -replace "@", " at "
    $s = $s -replace "\.net", " dotnet "
    $s = $s -replace "c\+\+", " cpp "
    $s = $s -replace "c#", " csharp "
    $s = $s -replace "\bms\b", " microsoft "
    $s = $s -replace "\bwin\b", " windows "
    $s = $s -replace "\bconfigmgr\b|\bsccm\b", " configuration manager "
    $s = $s -replace "\((x86|x64|32-bit|64-bit|32 bit|64 bit|amd64)\)", " "
    $s = $s -replace "\b(x86|x64|32-bit|64-bit|32 bit|64 bit|amd64)\b", " "
    if ($RemoveVersions) {
        $s = $s -replace "\b(v|version|ver)?\s*\d+([\._-]\d+){1,}[a-z0-9\.-]*\b", " "
        $s = $s -replace "\b20\d{2}\b", " "
    } else {
        $s = $s -replace "\b(v|version|ver)\s*", " "
    }
    $s = $s -replace "[^a-z0-9]+", " "
    $parts = @()
    foreach ($t in ($s -split "\s+")) {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if ($t.Length -le 1 -and $t -notmatch "^\d$") { continue }
        if ($NoiseTokens.ContainsKey($t)) { continue }
        if ($IsVendor -and $LegalVendorWords.ContainsKey($t)) { continue }
        $parts += $t
    }
    return ($parts -join " ")
}

function Get-TokenSet {
    param([string]$Text)
    $h = @{}
    if ([string]::IsNullOrWhiteSpace($Text)) { return $h }
    foreach ($t in ($Text -split "\s+")) {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if ($t -eq "for" -or $t -eq "the" -or $t -eq "and" -or $t -eq "with" -or $t -eq "from") { continue }
        $h[$t] = $true
    }
    return $h
}

function Copy-WithoutTokens {
    param([hashtable]$Source, [hashtable]$Remove)
    $h = @{}
    foreach ($k in $Source.Keys) {
        if (-not $Remove.ContainsKey($k)) { $h[$k] = $true }
    }
    return $h
}

function Remove-GenericTokens {
    param([hashtable]$Source)
    $h = @{}
    foreach ($k in $Source.Keys) {
        if (-not $GenericTokens.ContainsKey($k)) { $h[$k] = $true }
    }
    return $h
}

function Get-TokenArray {
    param([hashtable]$Set)
    return @($Set.Keys)
}

function Get-WeightedCoverage {
    param([hashtable]$A, [hashtable]$B, [hashtable]$Idf)
    $wa = 0.0; $wb = 0.0; $wi = 0.0; $overlap = 0
    foreach ($k in $A.Keys) {
        $v = 1.0
        if ($Idf.ContainsKey($k)) { $v = [double]$Idf[$k] }
        $wa += $v
        if ($B.ContainsKey($k)) {
            $wi += $v
            $overlap++
        }
    }
    foreach ($k in $B.Keys) {
        $v = 1.0
        if ($Idf.ContainsKey($k)) { $v = [double]$Idf[$k] }
        $wb += $v
    }
    if ($wa -le 0 -or $wb -le 0) {
        return @{ Coverage = 0.0; Overlap = 0 }
    }
    $minw = [Math]::Min($wa, $wb)
    return @{ Coverage = [Math]::Round((100.0 * $wi / $minw), 1); Overlap = $overlap }
}

function Get-Bigrams {
    param([string]$s)
    $s = ($s -replace "\s+", "")
    $grams = @{}
    if ($s.Length -lt 2) {
        if ($s.Length -eq 1) { $grams[$s] = 1 }
        return $grams
    }
    for ($i = 0; $i -lt ($s.Length - 1); $i++) {
        $g = $s.Substring($i, 2)
        if ($grams.ContainsKey($g)) { $grams[$g]++ } else { $grams[$g] = 1 }
    }
    return $grams
}

function Get-StringSimilarity {
    param([string]$A, [string]$B)
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return 0.0 }
    if ($A -eq $B) { return 100.0 }
    $ga = Get-Bigrams $A
    $gb = Get-Bigrams $B
    $total = 0; $inter = 0
    foreach ($k in $ga.Keys) {
        $total += [int]$ga[$k]
        if ($gb.ContainsKey($k)) { $inter += [Math]::Min([int]$ga[$k], [int]$gb[$k]) }
    }
    foreach ($k in $gb.Keys) { $total += [int]$gb[$k] }
    if ($total -eq 0) { return 0.0 }
    return [Math]::Round((200.0 * $inter / $total), 1)
}

function Get-TokenSetSimilarity {
    param([hashtable]$A, [hashtable]$B)
    if ($A.Count -eq 0 -or $B.Count -eq 0) { return 0.0 }
    $inter = 0
    foreach ($k in $A.Keys) {
        if ($B.ContainsKey($k)) { $inter++ }
    }
    return [Math]::Round((200.0 * $inter / ($A.Count + $B.Count)), 1)
}

function Get-Confidence {
    param([double]$Score)
    if ($Score -ge $AutoMatchThreshold) { return "AutoMatch" }
    if ($Score -ge $StrongReviewThreshold) { return "Review-Strong" }
    if ($Score -ge $PossibleReviewThreshold) { return "Review-Possible" }
    return "NoAutoMatch"
}

function Score-Pair {
    param($Arp, $Cat, [hashtable]$Idf)
    $arpVendorTokens = Get-TokenSet $Arp.VendorNorm
    $catPublisherTokens = Get-TokenSet $Cat.PubNorm

    $ta = Remove-GenericTokens (Copy-WithoutTokens $Arp.Tokens $arpVendorTokens)
    if ($ta.Count -eq 0) { $ta = Copy-WithoutTokens $Arp.Tokens $arpVendorTokens }
    if ($ta.Count -eq 0) { $ta = $Arp.Tokens }

    $tc = Remove-GenericTokens (Copy-WithoutTokens $Cat.Tokens $catPublisherTokens)
    if ($tc.Count -eq 0) { $tc = Copy-WithoutTokens $Cat.Tokens $catPublisherTokens }
    if ($tc.Count -eq 0) { $tc = $Cat.Tokens }

    $covObj = Get-WeightedCoverage $ta $tc $Idf
    $coverage = [double]$covObj.Coverage
    $overlap = [int]$covObj.Overlap

    $nameTokenScore = Get-TokenSetSimilarity $ta $tc
    $nameStringScore = Get-StringSimilarity $Arp.ProductNorm $Cat.NameNorm
    $nameScore = [Math]::Max($nameTokenScore, $nameStringScore)

    $pubScore = Get-StringSimilarity $Arp.VendorNorm $Cat.PubNorm
    $fullScore = Get-StringSimilarity $Arp.ProductFullNorm $Cat.NameFullNorm

    $exact = $false
    if ($Arp.ProductNorm -eq $Cat.NameNorm -or $Arp.ProductFullNorm -eq $Cat.NameFullNorm) { $exact = $true }

    $score = ($nameScore * 0.45) + ($coverage * 0.30) + ($pubScore * 0.20) + ($fullScore * 0.05)

    # Architecture preference:
    # - If the ARP/source app explicitly says x64, prefer x64 and penalize x86.
    # - If the ARP/source app explicitly says x86, prefer x86 and penalize x64.
    # - If the ARP/source app does not specify architecture, mildly prefer x64 over x86.
    if ($Arp.ArchitectureHint -eq "x64") {
        if ($Cat.ArchitectureHint -eq "x64") { $score += 4.0 }
        elseif ($Cat.ArchitectureHint -eq "x86") { $score -= 8.0 }
    }
    elseif ($Arp.ArchitectureHint -eq "x86") {
        if ($Cat.ArchitectureHint -eq "x86") { $score += 4.0 }
        elseif ($Cat.ArchitectureHint -eq "x64") { $score -= 8.0 }
    }
    else {
        if ($Cat.ArchitectureHint -eq "x64") { $score += 3.0 }
        elseif ($Cat.ArchitectureHint -eq "x86") { $score -= 3.0 }
    }

    if ($exact) { $score = [Math]::Max($score, 98.0) }
    if ($pubScore -ge 90 -and $coverage -ge 85 -and $overlap -ge 2) {
        $score = [Math]::Max($score, 92.0 + [Math]::Min(7.0, (($coverage - 85.0) / 15.0 * 7.0)))
    }
    if ($pubScore -ge 90 -and $nameScore -ge 92 -and $overlap -ge 2) {
        $score = [Math]::Max($score, 90.0)
    }

    if ($overlap -eq 0 -and -not $exact) {
        if ($pubScore -ge 90) { $score = [Math]::Min($score, 60.0) } else { $score = [Math]::Min($score, 45.0) }
    }
    if ($overlap -eq 1 -and -not $exact) {
        if ($pubScore -ge 90 -and $coverage -ge 90) { $score = [Math]::Min($score, 88.0) }
        elseif ($pubScore -ge 90) { $score = [Math]::Min($score, 78.0) }
        else { $score = [Math]::Min($score, 70.0) }
    }
    if ($pubScore -lt 70 -and $overlap -lt 2 -and -not $exact) {
        $score = [Math]::Min($score, 68.0)
    }

    return @{
        Score = [Math]::Round([Math]::Min(100.0, $score), 1)
        NameScore = [Math]::Round($nameScore, 1)
        PublisherScore = [Math]::Round($pubScore, 1)
        TokenCoverage = [Math]::Round($coverage, 1)
        TokenOverlap = $overlap
    }
}

function Escape-SqlString {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "" }
    return $Value.Replace("'", "''")
}

function Get-ArchitectureHint {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    if ($Text -match '(?i)\b(x64|64-bit|64 bit|amd64)\b') { return "x64" }
    if ($Text -match '(?i)\b(x86|32-bit|32 bit|i386)\b') { return "x86" }

    return ""
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

If ($csvFile) {
    $csvData = Import-Csv -Path $csvFile
    ForEach ($app in $csvData) {
        [void]$InstalledApps.Add([pscustomobject]@{
            Publisher        = $app.Vendor
            DisplayName      = $app.Product
            InstalledVersion = 0
        })
    }

} else {
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
}
# Get Available Applications from RMS

$Results = Invoke-RestMethod `
    -Method Post `
    -Uri "$RMSServer/api/application-catalog" `
    -UseDefaultCredentials

$output = $Results.value

# Exclude catalog entries that reference User Installer packages before building the match catalog.
# This prevents User Installer entries from being indexed, scored, matched, or inserted.
$output = @(
    $output | Where-Object {
        $catalogJson = ($_ | ConvertTo-Json -Depth 20 -Compress)
        $catalogJson -notmatch '(?i)\buser\s+installer\b'
    }
)

# *********************************************************

$catalog = New-Object System.Collections.ArrayList
$tokenDocFreq = @{}

for ($i = 0; $i -lt $output.Count; $i++) {
    $r = $output[$i]
    $nameNorm = Normalize-Text $r.Name $false $true
    $pubNorm = Normalize-Text $r.Publisher $true $true
    $tokens = Get-TokenSet $nameNorm
    foreach ($t in $tokens.Keys) {
        if ($tokenDocFreq.ContainsKey($t)) { $tokenDocFreq[$t]++ } else { $tokenDocFreq[$t] = 1 }
    }
    [void]$catalog.Add([pscustomobject]@{
        Index = $i
        Row = $r
        NameNorm = $nameNorm
        NameFullNorm = (Normalize-Text $r.Name $false $false)
        PubNorm = $pubNorm
        Tokens = $tokens
        PubTokens = (Get-TokenSet $pubNorm)
        ArchitectureHint = Get-ArchitectureHint (($r.Name, $r.Architecture, $r.References) -join ' ')
    })
}

$idf = @{}
foreach ($t in $tokenDocFreq.Keys) {
    $idf[$t] = [Math]::Log(($catalog.Count + 1.0) / ([double]$tokenDocFreq[$t] + 1.0)) + 1.0
}

$tokenIndex = @{}
for ($i = 0; $i -lt $catalog.Count; $i++) {
    foreach ($t in $catalog[$i].Tokens.Keys) {
        if (-not $tokenIndex.ContainsKey($t)) {
            $tokenIndex[$t] = New-Object System.Collections.ArrayList
        }
        [void]$tokenIndex[$t].Add($i)
    }
}

$FinalResults = New-Object System.Collections.ArrayList
$processed = 0

foreach ($r in $InstalledApps) {
    $processed++
    if (($processed % 1000) -eq 0) { Write-Host "Processed $processed / $($InstalledApps.Count)" }

    $prodNorm = Normalize-Text $r.DisplayName $false $true
    $prodFullNorm = Normalize-Text $r.DisplayName $false $false
    $vendorNorm = Normalize-Text $r.Publisher $true $true
    $prodTokens = Get-TokenSet $prodNorm

    $arpObj = [pscustomobject]@{
        Row = $r
        ProductNorm = $prodNorm
        ProductFullNorm = $prodFullNorm
        VendorNorm = $vendorNorm
        Tokens = $prodTokens
        ArchitectureHint = Get-ArchitectureHint $r.DisplayName
    }

    $candidateCounts = @{}
    foreach ($t in $prodTokens.Keys) {
        if ($tokenIndex.ContainsKey($t)) {
            $list = $tokenIndex[$t]
            if ($list.Count -lt 1000) {
                foreach ($idx in $list) {
                    $key = [string]$idx
                    if ($candidateCounts.ContainsKey($key)) { $candidateCounts[$key] += 3 } else { $candidateCounts[$key] = 3 }
                }
            }
        }
    }

    $candidateIds = @($candidateCounts.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $MaxCandidates | ForEach-Object { [int]$_.Key })

    $bestCat = $null
    $best = @{ Score = 0.0; NameScore = 0.0; PublisherScore = 0.0; TokenCoverage = 0.0; TokenOverlap = 0 }

    foreach ($idx in $candidateIds) {
        $catObj = $catalog[$idx]
        $sc = Score-Pair $arpObj $catObj $idf
        if ([double]$sc.Score -gt [double]$best.Score) {
            $best = $sc
            $bestCat = $catObj
        }
    }

    $catalogId = ""; $catalogPublisher = ""; $catalogName = ""; $catalogLatestVersion = ""
    if ($null -ne $bestCat) {
        $catalogId = $bestCat.Row.Id
        $catalogPublisher = $bestCat.Row.Publisher
        $catalogName = $bestCat.Row.Name
        $catalogLatestVersion = $bestCat.Row.LatestVersion
        #$catalog
    }

    [void]$FinalResults.Add([pscustomobject]@{
        ARPVendor = $r.Publisher
        ARPProduct = $r.DisplayName
        CatalogId = $catalogId
        CatalogPublisher = $catalogPublisher
        CatalogName = $catalogName
        CatalogLatestVersion = $catalogLatestVersion
        ARPArchitectureHint = $arpObj.ArchitectureHint
        CatalogArchitectureHint = if ($null -ne $bestCat) { $bestCat.ArchitectureHint } else { "" }
        Score = $best.Score
        Confidence = (Get-Confidence ([double]$best.Score))
        NameScore = $best.NameScore
        PublisherScore = $best.PublisherScore
        TokenCoverage = $best.TokenCoverage
        TokenOverlap = $best.TokenOverlap
    })
}


# Keep one best match per catalog application Id.

$AllMatches = $FinalResults |
    Where-Object { $_.Confidence -ne 'NoAutoMatch' -and -not [string]::IsNullOrWhiteSpace($_.CatalogId) } |
    Group-Object CatalogId |
    ForEach-Object {
        $_.Group |
            Sort-Object Score -Descending |
            Select-Object -First 1
    }


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

    $applicationId = $match.CatalogId
    $currentTarget = $output | Where-Object { $_.id -eq $applicationId } | Select-Object -First 1

    if ($null -eq $currentTarget) {
        Write-Warning "Skipping CatalogId $applicationId because it was filtered from the catalog or could not be found."
        continue
    }

    # Safety check in case User Installer appears in nested reference metadata.
    $referenceText = ($currentTarget | ConvertTo-Json -Depth 20 -Compress)
    if ($referenceText -match '(?i)\buser\s+installer\b') {
        Write-Warning "Skipping user installer match: $($currentTarget.Publisher) - $($currentTarget.Name) [$applicationId]"
        continue
    }

    $publisher = Escape-SqlString $currentTarget.Publisher
    $appName = Escape-SqlString $currentTarget.Name
    $architecture = Escape-SqlString $currentTarget.Architecture
    $language = Escape-SqlString $currentTarget.Language
    $iconUrl = Escape-SqlString $currentTarget.IconUrl

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
