<#
    .SYNOPSIS
    Script Based installation that can support multiple deployments based on group tag.

    .DESCRIPTION
    This script can be used as a dynamic installation script for the Agent Bootstrapper. 
    This embeds the device registration certificate in the script so that you can do 
    everything with this single script, no need to store files anywhere. This will generate the agent.json 
    file on demand although some modifications may be needed to reach your desired state.
    See the sections below to configure what group tags are associated with which deployment.
    This also needs an app registratiod and client secret to access the device information
    direct from Intune. This app registration needs the following permissions:

        Device.Read.All
        DeviceManagementManagedDevices.Read.All

    You can hard code the parameters, or you can pass them in the command line


    .EXAMPLE
    .\Install-AWDynamicFromIntuneScript.ps1

    .NOTES
    Version:       1.3
    Author:        John Yoakum, Recast Software
    Creation Date: 05/24/2026
    Purpose/Change: Tweaked the script to put all variables necessary at the top and available as parameters.
#>

param (
  [String]$url = "https://download.liquit.com/extra/Bootstrapper/AgentBootstrapper-Win-4.4.4130.3708.exe",
  [switch]$StartDeployment = $true,
  [string]$deployment = "Autopilot - ThinKiosk (Testing)", # name of default deployment to run
  [string]$logPath = "C:\Windows\Temp",
  [switch]$UseDeviceTags = $false, # if using device tags, you will modify the section below to have a more dynamic selection of deployments
  [switch]$UseCertificate = $true,
  [String]$AgentURL = "https://download.liquit.com/release/4.4/4225/Liquit-Universal-Agent-Win-4.4.4225.7279.exe",
  [string]$ZoneURL = "https://john.liquit.com", # Enter your zoneURL here.
  [string]$identitySource = "AzureAD", # Enter the name of your Identity Source for SSO.
  # Define the Azure App Registration details
  [string]$clientId = "2e4e0a31-3cd0-4c3b-b418-b2eca4a9b7e9",  #Client ID
  [string]$tenantId = "d37cd50c-80c6-4fd2-9be2-6b24ff526332",  #Directory ID
  [string]$clientSecret = "6Bf8Q~hqoywoDvC8uZwp2IpNRq0WF1c_Mb7aodjp" #Client Secret
)

######################
#   Configuration    #
######################
$InstallerArguments = ''

# Files to download
$DestinationPath = "C:\InstallFiles"
$InstallerPath = "C:\InstallFiles\AgentBootstrapper.exe"
$AgentPath = "C:\InstallFiles\Agent.exe"

# Create destination directory
If (!(Test-Path $DestinationPath)) {  
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
}

If ($StartDeployment) {$InstallerArguments += " --startDeployment --wait"}
If ($logPath) {$InstallerArguments += " --logPath $($logPath)"}
If ($UseCertificate) {
    $InstallerArguments += " --certificate C:\InstallFiles\AgentRegistration.cer"
    # Replace the below with your certificate for Device Registration
    $Certificate = @"
-----BEGIN CERTIFICATE-----
MIIDPjCCAiagAwIBAgIQMpop5F2aMqRCubsj16subTANBgkqhkiG9w0BAQsFADAkMSIwIAYDVQQD
DBlMaXF1aXQgQWdlbnQgUmVnaXN0cmF0aW9uMB4XDTI0MDUyMDIxNTYyNloXDTM0MDUxODIxNTYy
NlowJDEiMCAGA1UEAwwZTGlxdWl0IEFnZW50IFJlZ2lzdHJhdGlvbjCCASIwDQYJKoZIhvcNAQEB
BQADggEPADCCAQoCggEBAJroIsoLX9iUGpnaeWOaHPKc3pjm8XxSJGTqkZU9U0Y3EsnAiX8drLuf
yL+gARuZ3ZItayJxlHw9IjqLDASXBJHE+oDoojo7dc+7y8nq2DUUAiJT6LiRLGaUUQSrOWpQAbbI
EnIIU29hc7+NolapXHqbNJzvlJFil9FuU2NAyw0OMJ1CUbIWGLvw0adhCFrt/EYNXBbNUJmReT/c
gkFNW9gmb0GIQDXN+xYkcGa/qICqjQfmFGCeRzjXxB0xFJc6j3h3YNhBjQuutXbbfz3uyB91OW+1
/fZPH0u3mOXue8QwjKOJEJ2rMSKLyOgDnvuzp/5ayrZvbwtCD4Vj56vNXVUCAwEAAaNsMGowDgYD
VR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMCMCQGA1UdEQQdMBuCGUxpcXVpdCBBZ2Vu
dCBSZWdpc3RyYXRpb24wHQYDVR0OBBYEFILhJ05HlHEvzfbt2YxZQGrRkptFMA0GCSqGSIb3DQEB
CwUAA4IBAQBucR8gOPE9mJKkuFeFu+d0nFu66EPBsCqlYU9T6vyvik8tRMas+4i4E6ZV4+sjCK86
AtEqa0AXDhdTOql+lyuOogYwrDJQ6pxj34x8bQoTLRgj7kNbG6BasV+WE1VcO3GxwI/Bxf2ByaY8
zWwbAxz//C7SjzEU69dLMRtARb38SuGI6/YyaoSVVtUHGImS+lUCzh1bdQIpGmsrYVYXgSt7rikL
ei45VTaOrbp/pS8ddLxncu+5xsFlnkCaODJ89dgmA8Iwndmt9FtN5ulP0lNUwEwVMfdMMr5TIopU
5yNS9TB0MvMwe1hNy2aFWXTSo9yV6eX9xC92LgpM1OPxqGcJ
-----END CERTIFICATE-----
"@
    New-Item -Path "$DestinationPath\AgentRegistration.cer" -ItemType File -Value $Certificate -Force

}

If ($UseDeviceTags) {
    # Define the device name to search for
    $deviceName = "$env:COMPUTERNAME"

    # Get an access token using client credentials flow
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body
    $accessToken = $response.access_token
    $deviceSearchResults = $null

    $graphApiUrl = "https://graph.microsoft.com/v1.0/devices?`$filter=startswith(displayName,'$deviceName')"
    $headers = @{
        Authorization = "Bearer $accessToken"
    }
    Try {
        $deviceSearchResults = Invoke-RestMethod -Method Get -Uri $graphApiUrl -Headers $headers
    } catch {
    
    }
    if ($deviceSearchResults.value.Length -eq 0 -or $null -eq $deviceSearchResults) {
        Write-Output "No devices found with the name: $deviceName"
        return
    } elseif ($deviceSearchResults.value.Length -gt 1) {
        $deviceSearchResults.value
        return
    }

    $device = $deviceSearchResults.value[0]

    $physicalIds = $device.physicalIds

    if ($physicalIds) {
        $orderId = $physicalIds -match "^\[OrderId\]:.*"

        if ($orderId) {
            $cleanOrderId = $orderId -replace "^\[OrderId\]:", ""
        }
    }
    ###### Set your different Tags here with associated Deployments
    switch ($cleanOrderId) {
        'Finance' {
            $deployment = 'Finance'
        }
        'IT' {
            $deployment = 'IT'
        }
        default {
            $deployment = 'Autopilot'
        }
    }
}


#######################################
#    DOWNLOAD FILES TO DESTINATION    #
#######################################


# Download Agent Bootstrapper direct from Internet
Invoke-WebRequest -Uri $url -OutFile $InstallerPath -UseBasicParsing

# Download Agent Installer
Try {
    # Try to download installer from zone
    Invoke-WebRequest -Uri "$ZoneURL/api/agent/installers/118A90AD-C2EB-4AE3-A69E-B1154CF46962" -TimeoutSec 60 -OutFile $AgentPath -UseBasicParsing
} catch {
    # Download Agent Bootstrapper direct from Internet
    Invoke-WebRequest -Uri $AgentURL -OutFile $AgentPath -UseBasicParsing
}

Write-Output "All downloads completed."

############################################################
#    CREATE AGENT JSON
###########################################################



##### Tweak the next section of code for the agent.json settings you would like.
$jsonData = @{
    zone = $ZoneURL
    promptZone = "Disabled"
    login = @{
        enabled = $true
        sso = $true
        identitySource = $identitySource
        timeout = 4
    }
    log = @{
        level = "Debug"
        agentPath = "Agent.log"
        userHostPath = "UserHost.log"
        rotateCount = 5
        rotateSize = 1048576
    }
    registration = @{
        type = "Certificate"
    }
    nativeIcons = @{
        enabled = $true
        primary = $true
        startMenuPath = '${Programs}'
    }
    launcher = @{
        enabled = $true
        state = "Default"
        start = "Auto"
        tiles = $false
        minimal = $false
        contextMenu = $true
        sideMenu = "Tags"
        close = $true
    }
    deployment = @{
        zoneTimeout = 60
        enabled = $false
        start = $false
        context = "Device"
        cancel = $false
        triggers = $true
        autoStart = @{
            enabled = [bool]$StartDeployment
            deployment = $deployment
            timer = 0
        }

    }
}

$jsonString = $jsonData | ConvertTo-Json -Depth 10
$jsonFilePath = "$DestinationPath\Agent.json"
$jsonString | Set-Content -Path $jsonFilePath

#####################################################
#       Initiate Install Process
#####################################################

set-location $DestinationPath

# Start the install process

Write-Host "Starting the installation process..."
Write-Host "Command line being ran:  $InstallerPath $InstallerArguments"

if (Test-Path -Path $InstallerPath) {

   try {

       Start-Process -FilePath $InstallerPath -ArgumentList $InstallerArguments -Wait

       Exit 0

   } catch {

       Write-Error "Error starting the installer '$InstallerPath': $($_.Exception.Message)"

       exit 1

   }

} else {

   Write-Warning "Installer executable not found: '$InstallerPath'"


}














