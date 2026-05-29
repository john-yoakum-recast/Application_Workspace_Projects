<#
    .SYNOPSIS
    Script To automate the process of creating Windows Update Packages

    .DESCRIPTION
    This script can be used to automate the process of creating Windows update Packages in
    Application Workspace. It uses the powershell module MSCatalogLTS
    (https://powershellisfun.com/2025/09/18/search-and-download-microsoft-updates-using-the-mscataloglts-powershell-module/)
    I set up this script to choose the correct windows versions you would want to create
    update packages for, however, you can modify this script for whatever criteria you
    would like based on the different options within the module. If you want to run all 
    Windows updates through Application Workspace, you can by setting this script up on a scheduled task on a machine.
    This way you can have it always and automatically create those update packages for you.

    You need to have the following modules installed for this to work correctly:
        MSCatalogLTS
        Liquit.Server.Powershell

    .EXAMPLE
    .\Build-WindowsUpdates.ps1

    .NOTES
    Version:       1.0
    Author:        John Yoakum, Recast Software
    Creation Date: 05/29/2026
    Purpose/Change: Initial script development
#>

Import-Module MSCatalogLTS

# Number of days ago you want to pull updates for
$numberOfDays = 30

# Specify the path you want to download files to
$downloadedFilePath = $null
$zipFilePath = $null
$downloadPath = "C:\Temp\Downloads"
$zipFolder = "C:\Temp\Zips"
$tempFolder = "C:\Temp"

# Specify Application Workspace Credentials for creating the packages
$awURL = "https://john.recastsoftware.cloud"
$awUsername = "LOCAL\admin"
$awPassword = "IsaiahMaddux@2014"
$awCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $awUsername, (ConvertTo-SecureString -String $awPassword -AsPlainText -Force)
$awIcon = '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'
# Search Criteria for which platforms and versions to capture
$searchCriteria = @(
    "Windows 11 24H2 x64"
    "Windows 11 25H2 x64"
    "Windows 11 23H2 x64"
)
<#
$searchCriteria = @(
    "Windows 11 24H2 x64"
    "Windows 11 25H2 x64"
    "Windows 11 23H2 x64"
    "Windows Server 21H2 x64"
    "Windows Server 24H2 x64"
    "Windows Server 2022"
    "Windows Server 2025"
)
#>

function Get-7ZipPath {
    $paths = @(
        "C:\Program Files\7-Zip\7z.exe"
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }

    Write-Warning "7-Zip is not installed. Please install it from https://www.7-zip.org/ and re-run the script."
    exit 1
}

function Get-UpdateMetadata {
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    $kbArticle = $null
    $osBuild   = $null
    $release   = $null
    $type      = $null
    $fullBuild = $null
    $packageName = $null

    if ($title -match '\((KB\d{7})\)') {
        $kbArticle = $matches[1]
    }

    if ($title -match '\((\d{5})\.\d+\)') {
        $osBuild = $matches[1]
    }

    if ($title -match '(?i)\b(2[0-9]H[12])\b') {
        $release = $matches[1].ToUpper()
    }

    if ($title -match '(?i)server') {
        $type = 'Server'
    }
    elseif ($title -match '(?i)Windows 11') {
        $type = 'Workstation'
    }
    if ($title -match '\((\d{5}\.\d+)\)') {
        $fullBuild = $matches[1]
    }
    if ($title -match '^(?<Date>\d{4}-\d{2}).*?\((?<KB>KB\d{7})\)') {
        $dateValue = $matches['Date']
        $kbValue   = $matches['KB']

        $platform = if ($title -match '(?i)server') {
            'Windows Server'
        }
        elseif ($title -match '(?i)Windows 11') {
            'Windows 11'
        }
        else {
            'Unknown'
        }

        $packageName = "$dateValue $kbValue for $platform $release"
    }
    [pscustomobject]@{
        KBArticle = $kbArticle
        OSBuild   = $osBuild
        Release   = $release
        Type      = $type
        FullBuild = $fullBuild
        PackageName = $packageName
    }
}

function New-ZipWith7Zip {
    param(
        [Parameter(Mandatory)]
        [string]$SourceFilePath,

        [Parameter(Mandatory)]
        [string]$ZipFilePath,

        [string]$SevenZipExe = "C:\Program Files\7-Zip\7z.exe"
    )

    if (-not (Test-Path $SourceFilePath)) {
        throw "Source file not found: $SourceFilePath"
    }

    if (-not (Test-Path $SevenZipExe)) {
        throw "7-Zip executable not found: $SevenZipExe"
    }

    if (Test-Path $ZipFilePath) {
        Remove-Item -Path $ZipFilePath -Force
    }

    $null = & $SevenZipExe a -tzip $ZipFilePath $SourceFilePath
    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip failed creating archive: $ZipFilePath"
    }

    return $ZipFilePath
}

function Create-UpdatePackage {
param(
        [Parameter(Mandatory)]
        $Update
)

    try {
        
        if (-not (Test-Path $downloadPath)) {
            New-Item -Path $downloadPath -ItemType Directory -Force | Out-Null
        }

        if (-not (Test-Path $zipFolder)) {
            New-Item -Path $zipFolder -ItemType Directory -Force | Out-Null
        }

        # Download Update
        Save-MSCatalogUpdate -Guid $Update.Guid -Destination $downloadPath -DownloadAll | Out-Null

        $newFiles = Get-ChildItem -Path $downloadPath -File

        if (-not $newFiles) {
            throw "No new downloaded file found."
        }

        # Download and process each Windows Update
        Connect-LiquitWorkspace -URI $awURL -Credential $awCredentials

        # Create Icon
        $bytes = [System.Convert]::FromBase64String($awIcon)
        [System.IO.File]::WriteAllBytes("$tempFolder\icon.jpg", $bytes) | Out-Null
        $iconPath = "$tempFolder\icon.jpg"
        $iconContent = New-LiquitContent -Path $iconPath -FileName "icon.jpg"

        # Set a counter for number of Update
        $packageNumber = 1

        ForEach ($newFile in $newFiles) {
            $downloadedFilePath = $newFile.FullName
            $downloadedFileName = $newFile.Name
            
            Write-Host "MSU Full Path: $downloadedFilePath"
            Write-Host "MSU FileName: $downloadedFileName"

            $zipFilePath = Join-Path $zipFolder ("{0}.zip" -f [System.IO.Path]::GetFileNameWithoutExtension($downloadedFileName))

            Write-Host "Creating the zip file for $($Update.PackageName) Update: $packageNumber"
            New-ZipWith7Zip -SourceFilePath $downloadedFilePath -ZipFilePath $zipFilePath

            Write-Host "Zip File Path 2: $zipFilePath"
            $zipFileName = Split-Path -Path $zipFilePath -Leaf

            Write-Host "ZipFileName: $zipFileName"
        
            # Create the new package
            Write-Host "Creating the Package Shell for $($Update.PackageName) Update: $packageNumber"
            $newPackage = New-LiquitPackage -Name "$($Update.PackageName) Update: $packageNumber" -Type "Update" -DisplayName "$($Update.PackageName) Update: $packageNumber" -Priority 100 -Enabled $true -Offline $true -Web $false -Icon $iconContent
        
            # Create the snapshot
            Write-Host "Creating the Package Snapshot for $($Update.PackageName) Update: $packageNumber"
            $newSnapshot = New-LiquitPackageSnapshot -Package $newPackage -Name $Update.KBArticle

            # Create the filter objects for this update package
            Write-Host "Creating the Snapshot Filter for $($Update.PackageName) Update: $packageNumber"
            $awFilterSet = New-LiquitFilterSet -Snapshot $newSnapshot

            # Create the filters for this snapshot
            If ($Update.Type -eq 'Server') {
                    switch ($Update.Release) {
                    "21H2" {
                        $awFilter1 = New-LiquitFilter -FilterSet $awFilterSet -Type PlatformVersion -Operator Equal -Value '10.0.20348' -Settings @{type = "winnt"; platformType = $Update.Type}
                    }
                    "24H2" {
                        $awFilter1 = New-LiquitFilter -FilterSet $awFilterSet -Type PlatformVersion -Operator Equal -Value '10.0.26100' -Settings @{type = "winnt"; platformType = $Update.Type}
                    }
                }
            } else {
                $awFilter1 = New-LiquitFilter -FilterSet $awFilterSet -Type PlatformVersion -Operator Equal -Value $('10.0.' + $Update.OSBuild) -Settings @{type = "winnt"; platformType = $Update.Type}
            }
        
            # Create the Action Set
            Write-Host "Creating the Install Action Set for $($Update.PackageName) Update: $packageNumber"
            $awActionSet = New-LiquitActionSet -Snapshot $newSnapshot -Name 'Install Update' -Type Install -Frequency OncePerDevice

            # Create the Action to copy content locally
            Write-Host "Uploading content for $($Update.PackageName) Update: $packageNumber"
            $awContent = New-LiquitContent -path $zipFilePath
            Try {
            Write-Host "Creating Content Copy for $($Update.PackageName) Update: $packageNumber"
            $awAction1 = New-LiquitAction -ActionSet $awActionSet -Name 'Copy MSU to local machine' -Type 'contentextract' -Enabled $true -IgnoreErrors $true -Settings @{content = "$zipFileName"; destination = '${PackageTempDir}'}
            } catch {}
            Write-Host "Linking Content to Action for $($Update.PackageName) Update: $packageNumber"
            $awAttribute = New-LiquitAttribute -Entity $awAction1 -Link $awContent -ID 'content' -Settings @{filename = "$zipFileName"}

            # Create the Action to install windows update
            Write-Host "Adding installation step to $($Update.PackageName) Update: $packageNumber"
            $awAction2 = New-LiquitAction -ActionSet $awActionSet -Name 'Install Update' -Type pathmsuinstall -Enabled $true -IgnoreErrors $true -Settings @{arguments = '/quiet /norestart'; path = '${PackageTempDir}' + "\$downloadedFileName"}

            # Create Action to delete temporary files
            Write-host "Adding deleting temporary files for $($Update.PackageName) Update: $packageNumber"
            $awAction3 = New-LiquitAction -ActionSet $awActionSet -Name 'Remove Temporary Directory' -type 'dirdelete' -Enabled $true -IgnoreErrors $false -Settings @{path = '${PackageTempDir}'}

        # *******************************************************
            $packageNumber++
        }
        foreach ($file in @($downloadPath, $zipFilePath)) {
            if ($file -and (Test-Path $file)) {
                Remove-Item -Path $file -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
    } catch {}
}

$UpdatesAvailable = [System.Collections.ArrayList]::new()

$7ZipPath = Get-7ZipPath

# Get all the Windows update that match Search Criteria
foreach ($item in $searchCriteria) {
    try {
        $results = Get-MSCatalogUpdate -Search $item -LastDays $numberOfDays -AllPages -ErrorAction Stop 
        foreach ($result in $results) {
            $meta = Get-UpdateMetadata -Title $result.Title

            
            $update = New-Object PSObject -Property @{
                Title = $result.Title
                Products = $result.Products
                Classification = $result.Classification
                LastUpdated = $result.LastUpdated
                Size = $result.Size
                FileNames = $result.FileNames
                Guid = $result.Guid
                Version = $result.Version
                KBArticle = $meta.kbArticle
                OSBuild = $meta.osBuild
                Release = $meta.release
                Type = $meta.type
                FullBuild = $meta.fullBuild
                PackageName = $meta.PackageName
            }
            [void]$UpdatesAvailable.Add($update)
        }
    }
    catch {
        Write-Warning "Failed to query '$item': $($_.Exception.Message)"
    }
}

ForEach ($updateAvailable in $UpdatesAvailable) {
    Create-UpdatePackage -Update $updateAvailable
}


#$UpdatesAvailable
