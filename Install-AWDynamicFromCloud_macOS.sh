#!/bin/zsh
set -euo pipefail

# =================================================================
# .SYNOPSIS
# Script-based installation that can support multiple deployments based on group tag.
#
# .DESCRIPTION
# macOS ZSH version intended for dynamic installation of the Liquit / Application Workspace Agent.
# Adjust parameters as necessary for your specific deployment.
#
# .NOTES
# Version:       1.1
# Author:        John Yoakum, Recast Software
# Revised By:    Christopher Antoku / Copilot
# Creation Date: 01/29/2026
# Purpose/Change:
#   - Fixed ZSH syntax
#   - Removed HTML encoding from copied URLs and heredocs
#   - Added safer argument handling
#   - Added JSON validation
#   - Added improved download and launch daemon handling
# =================================================================

# -------------------------------
# Default Parameters
# -------------------------------

BootstrapperURL="https://download.liquit.com/extra/Bootstrapper/AgentBootstrapper-Mac-4.4.4130.3708"

StartDeployment=true
deployment="Mac Base Apps"
logPath="/tmp"
UseCertificate=true

AgentURL="https://download.liquit.com/release/4.4/4225/Liquit-Universal-Agent-Mac-4.4.4225.7279.pkg"

DestinationPath="$HOME/InstallFiles"
InstallerPath="$DestinationPath/AgentBootstrapper"
AgentPath="$DestinationPath/Agent.pkg"
CertificatePath="$DestinationPath/AgentRegistration.cer"
jsonFilePath="$DestinationPath/Agent.json"

ZoneURL="https://AppWorkspace-lmigov.msappproxy.us"
identitySource="Provisioning"

appName="Liquit.app"
appPath="/Applications/$appName"

# -------------------------------
# Build Installer Arguments
# -------------------------------

InstallerArguments=()

if [ "$StartDeployment" = true ]; then
  InstallerArguments+=("--startDeployment" "--wait")
fi

if [ -n "$logPath" ]; then
  InstallerArguments+=("--logPath" "$logPath")
fi

# -------------------------------
# Ensure Destination Directory Exists
# -------------------------------

if [ ! -d "$DestinationPath" ]; then
  echo "Creating directory: $DestinationPath"
  mkdir -p "$DestinationPath"
fi

# -------------------------------
# Check for Existing Installation
# -------------------------------

echo "Checking if $appName is already installed..."

if [ -d "$appPath" ]; then
  echo "$appName is already installed at $appPath. Exiting script."
  exit 0
else
  echo "$appName is not installed. Proceeding with installation."
fi

# -------------------------------
# Certificate Handling
# -------------------------------

if [ "$UseCertificate" = true ]; then
  echo "Creating agent registration certificate at: $CertificatePath"

  cat <<EOF >"$CertificatePath"
-----BEGIN CERTIFICATE-----
MIIDazCCAlOgAwIBAgIQETKA7ne6Ep9C8QTPNAFZFTANBgkqhkiG9w0BAQsFADAzMTEwLwYDVQQD
DChBcHBsaWNhdGlvbiBXb3Jrc3BhY2UgQWdlbnQgUmVnaXN0cmF0aW9uMB4XDTI2MDYzMDE1MTU1
MFoXDTM2MDYyNzE1MTU1MFowMzExMC8GA1UEAwwoQXBwbGljYXRpb24gV29ya3NwYWNlIEFnZW50
IFJlZ2lzdHJhdGlvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJpcIhi+V6eHmWRi
g2TIXStD84nRtmGECsn4ZZUr4N88FFTG2bWzSC3dTViHwZClhaVyoUo3wEBPM9/TZBEKTcBhh5Nf
x4Sj5GKJpCG0aV/UcqRIMOM4I+PXUD9f+Fadfx02oGLqlpX7QSv45/G3zb3gGk/D0rms570ixw1O
Dm+P2r4/6KLGVgBdQpirOlUN6jfRUZfx5r8qIJV24GB9vujLuXJxK16W7x8nuPfL7riqglHixX5k
cI587vYiOl13VBHDkqHtt8olUztEyDZNXhe6KifwpMt+VZ0n54ExWvpTvDByssECAjd/Q8zfemHB
xLRJTOZA7ep03d7TH2Yk1pUCAwEAAaN7MHkwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsG
AQUFBwMCMDMGA1UdEQQsMCqCKEFwcGxpY2F0aW9uIFdvcmtzcGFjZSBBZ2VudCBSZWdpc3RyYXRp
b24wHQYDVR0OBBYEFKVXI3hy1jN78zDvKXcLk23teSqqMA0GCSqGSIb3DQEBCwUAA4IBAQBP3Uge
qZatc68zyHWuW2Lyw/HUC0lTg5LERmn7oY88yy0BBRmSEfHOnsAXJ1b/XdpVUldETsRs6Q2FfdMX
5/PFCYPE+2y816xSY2nAl4UtJYA8SO3jiXsLhlnpXeWTXvmLixtQkOIcaSuQ+XGTLjb45jhV2zuV
jvUlzOLi9Og26IFuQvwOVvX5+ruv2rN+P3g4mXNbjOjjvx6dc5J3hLyInondHjJMRoCWE+cRyvOk
PZcsSJH63H7qrSyVDtC1VCXa2m3tIyaYfoTAHd8wTj8J/WYTkqMvCP8ulZ7H2i0J+gamVFi/n5UA
E8oZpyuS3CpHdJwKNE4393C6xWY74TP7
-----END CERTIFICATE-----
EOF

  InstallerArguments+=("--certificate" "$CertificatePath")
fi

# -------------------------------
# Download Files
# -------------------------------

echo "Downloading bootstrapper..."

if ! curl -fL "$BootstrapperURL" -o "$InstallerPath"; then
  echo "Failed to download bootstrapper from: $BootstrapperURL"
  exit 1
fi

echo "Bootstrapper downloaded to: $InstallerPath"

echo "Downloading agent package from zone URL..."

if curl -fL "$ZoneURL/api/agent/installers/F84543F0-F440-4200-9A2B-E13FC30C71BB" --connect-timeout 60 -o "$AgentPath"; then
  echo "Agent successfully downloaded from zone URL."
else
  echo "Failed to download agent from zone URL. Falling back to public Agent URL..."

  if ! curl -fL "$AgentURL" -o "$AgentPath"; then
    echo "Failed to download agent package from fallback URL."
    exit 1
  fi

  echo "Agent successfully downloaded from fallback URL."
fi

echo "All downloads completed."

# -------------------------------
# Build Optional Deployment JSON Block
# -------------------------------

deploymentBlock=""

if [ "$StartDeployment" = true ]; then
deploymentBlock=$(cat <<EOF
,
  "deployment": {
    "zoneTimeout": 60,
    "enabled": true,
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
EOF
)
fi

# -------------------------------
# Create Agent JSON Configuration
# -------------------------------

echo "Creating agent configuration file at: $jsonFilePath"

cat <<EOF >"$jsonFilePath"
{
  "zone": "$ZoneURL",
  "promptZone": "Disabled",
  "login": {
    "enabled": true,
    "sso": true,
    "identitySource": "$identitySource",
    "timeout": 15
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
  "launcher": {
    "enabled": true,
    "state": "default",
    "start": "Disabled",
    "tiles": false,
    "minimal": false,
    "contextMenu": true,
    "sideMenu": "Tags",
    "close": true
  },
  "userHostStartupOperatingMode": "Background"$deploymentBlock
}
EOF

# -------------------------------
# Validate JSON Configuration
# -------------------------------

echo "Validating generated JSON configuration..."

if ! python3 -m json.tool "$jsonFilePath" >/dev/null 2>&1; then
  echo "Generated JSON is invalid. Outputting file contents for troubleshooting:"
  cat "$jsonFilePath"
  exit 1
fi

echo "JSON validated successfully."

# -------------------------------
# Initiate Installation
# -------------------------------

echo "Initiating installation process..."

if [ ! -f "$InstallerPath" ]; then
  echo "Installer file not found: $InstallerPath"
  exit 1
fi

echo "Downloaded bootstrapper file information:"
file "$InstallerPath"

echo "Making bootstrapper executable..."
chmod +x "$InstallerPath"

if [ ! -x "$InstallerPath" ]; then
  echo "Bootstrapper is not executable after chmod."
  ls -l "$InstallerPath"
  exit 1
fi

echo "Installer arguments:"
printf '  %s\n' "${InstallerArguments[@]}"

echo "Starting bootstrapper installation..."

sudo "$InstallerPath" "${InstallerArguments[@]}"

installExitCode=$?

if [ $installExitCode -ne 0 ]; then
  echo "Bootstrapper returned exit code: $installExitCode"
  exit $installExitCode
fi

# Get logged-in user
loggedInUser=$(stat -f%Su /dev/console)
loggedInUID=$(id -u "$loggedInUser")

LaunchAgentPath="/Library/LaunchAgents/com.liquit.UserHost.plist"

if [ -f "$LaunchAgentPath" ]; then
    echo "Restarting UserHost LaunchAgent for $loggedInUser..."

    # Unload if already loaded
    launchctl bootout "gui/$loggedInUID" "$LaunchAgentPath" 2>/dev/null || true

    sleep 2

    # Reload
    launchctl bootstrap "gui/$loggedInUID" "$LaunchAgentPath"

    sleep 2

    # Force restart if already running
    launchctl kickstart -k "gui/$loggedInUID/com.liquit.UserHost" 2>/dev/null || true

    echo "UserHost LaunchAgent restarted."
else
    echo "LaunchAgent not found: $LaunchAgentPath"
fi

echo "Installation script completed."
