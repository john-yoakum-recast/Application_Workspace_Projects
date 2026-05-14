#!/bin/zsh

# =================================================================
# .SYNOPSIS
# Script Based installation that can support multiple deployments based on group tag.
#
# .DESCRIPTION
# This script is a macOS version rewritten in ZSH intended for dynamic installation processes.
# Adjust parameters as necessary for your specific deployment.
#
# .NOTES
# Version:       1.0
# Author:        John Yoakum, Recast Software
# Creation Date: 01/29/2026
# Purpose/Change: Rewriting the script to support macOS in ZSH
# =================================================================

# Set default parameters
url="https://download.liquit.com/extra/Bootstrapper/AgentBootstrapper-Mac-4.4.4130.3708"
StartDeployment=true
deployment="Macs"
logPath="/tmp"
UseCertificate=true
AgentURL="https://download.liquit.com/release/4.4/4225/Liquit-Universal-Agent-Mac-4.4.4225.7279.pkg"
DestinationPath="$HOME/InstallFiles"
InstallerPath="$DestinationPath/AgentBootstrapper"
AgentPath="$DestinationPath/Agent.pkg"
ZoneURL="https://john.liquit.com"
appName="Liquit.app"
appPath="/Applications/$appName"

InstallerArguments=("--startDeployment" "--wait")
if [ -n "$logPath" ]; then
  InstallerArguments+=("--logPath" "$logPath")
fi

# Ensure the destination directory exists
if [ ! -d "$DestinationPath" ]; then
  echo "Creating directory: $DestinationPath"
  mkdir -p "$DestinationPath"
fi

# Check for already installed and if so, exit
echo "Checking if $appName is already installed..."

# Check if the application exists at the default path
if [ -d "$appPath" ]; then
  echo "$appName is already installed at $appPath. Exiting script..."
  exit 0  # Exit successfully since the application is already installed
else
  echo "$appName is not installed. Proceeding with the installation..."
fi

# Handle Certificates if enabled
if $UseCertificate; then
  cat <<EOF >"$DestinationPath/AgentRegistration.cer"
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
EOF
  InstallerArguments+=("--certificate" "$DestinationPath/AgentRegistration.cer")
fi

# Download files
echo "Downloading files..."
curl -L "$url" -o "$InstallerPath"
if curl -L "$ZoneURL/api/agent/installers/F84543F0-F440-4200-9A2B-E13FC30C71BB" --connect-timeout 60 -o "$AgentPath"; then
  echo "Agent successfully downloaded from zone URL."
else
  echo "Failed to download from zone, fallback to Agent URL..."
  curl -L "$AgentURL" -o "$AgentPath"
fi
echo "All downloads completed."

# Create JSON configuration
jsonFilePath="$DestinationPath/Agent.json"
cat <<EOF >"$jsonFilePath"
{
  "zone": "$ZoneURL",
  "promptZone": "Disabled",
  "login": {
    "enabled": true,
    "sso": true,
    "identitySource": "AzureAD",
    "timeout": 4
  },
  "log": {
    "level": "Debug",
    "agentPath": "Agent.log",
    "userHostPath": "UserHost.log",
    "rotateCount": 5,
    "rotateSize": 1048576
  },
  "registration": {
    "type": "Certificate"
  },
  "nativeIcons": {
    "enabled": true,
    "primary": true,
    "startMenuPath": "\${Programs}"
  },
  "launcher": {
    "enabled": true,
    "state": "Default",
    "tiles": false,
    "minimal": false,
    "contextMenu": true,
    "sideMenu": "Tags",
    "close": true
  },
  "deployment": {
    "zoneTimeout": 60,
    "enabled": false,
    "start": false,
    "context": "Device",
    "cancel": false,
    "triggers": true,
    "autoStart": {
      "enabled": true,
      "deployment": "$deployment",
      "timer": 0
    }
  }
}
EOF

# Initiate installation
echo "Initiating installation process..."

# Run the bootstrapper if the installer exists
if [ -f "$InstallerPath" ]; then
  echo "Making the bootstrapper executable..."
  chmod +x "$InstallerPath" # Make the file executable

  echo "Starting the installation process..."
  echo "Command: ./$InstallerPath $InstallerArguments"

  # Execute the bootstrapper with the provided arguments
  sudo "$InstallerPath" "${InstallerArguments[@]}"

  echo "Installation process completed."
else
  echo "Installer file not found: $InstallerPath"
fi

# Get logged-in user and UID 
loggedInUser=$(stat -f%Su /dev/console) 
userUID=$(id -u "$loggedInUser") 

# Kill UserHost as the user 
echo "Running 'killall UserHost' as $loggedInUser..." 
sudo -u "$loggedInUser" killall UserHost 2>/dev/null 
echo "Waiting 5 seconds after killall..." sleep 5 

# Restart the Liquit agent service 
if [ -f "/Library/LaunchDaemons/com.liquit.Agent.plist" ]; then 
	echo "Reloading Liquit LaunchDaemon..." 
	sudo launchctl bootout system /Library/LaunchDaemons/com.liquit.Agent.plist 
	sleep 2 
	sudo launchctl bootstrap system /Library/LaunchDaemons/com.liquit.Agent.plist 
else 
	echo "LaunchDaemon not found." 
fi 


echo "Installation script completed."
