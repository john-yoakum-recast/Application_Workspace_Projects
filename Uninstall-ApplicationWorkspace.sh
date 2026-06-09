#!/bin/zsh

# =================================================================
# .SYNOPSIS
# Script Based Uninstall for macOS to remove Application Workspace from mac Devices
#
# .DESCRIPTION
# This script is designed as an uninstall script for macOS to completely remove Application Workspace
# including any additional files, launch daemon and certificates.
#
# .NOTES
# Version:       1.0
# Author:        John Yoakum, Recast Software
# Creation Date: 06/09/2026
# Purpose/Change: Initial Script Development
# =================================================================

# Must be run as root
if [[ "$EUID" -ne 0 ]]; then
  echo "Please run as root: sudo $0"
  exit 1
fi

# Set default parameters
logPath="/tmp"
appName="Liquit.app"
appPath="/Applications/$appName"

DIRECTORY_TO_DELETE="/Library/Application Support/Liquit"

CERTIFICATE_NAMES=(
  "Liquit Agent Authentication"
  "Liquit Root CA"
)

# Get logged-in user and UID 
loggedInUser=$(stat -f%Su /dev/console) 
userUID=$(id -u "$loggedInUser") 

# Kill UserHost as the user 
echo "Running 'killall UserHost' as $loggedInUser..." 
sudo -u "$loggedInUser" killall UserHost 2>/dev/null 
echo "Waiting 5 seconds after killall..." sleep 5 

if [ -f "/Library/LaunchDaemons/com.liquit.Agent.plist" ]; then 
	echo "Reloading Liquit LaunchDaemon..." 
	sudo launchctl bootout system /Library/LaunchDaemons/com.liquit.Agent.plist 
else 
	echo "LaunchDaemon not found." 
fi 

sudo rm -f /Library/LaunchDaemons/com.liquit.Agent.plist

###############################################################################
# Delete directory
###############################################################################

if [[ -d "$appPath" ]]; then
    rm -rf "$appPath"
fi

if [[ -d "$DIRECTORY_TO_DELETE" ]]; then
  rm -rf "$DIRECTORY_TO_DELETE"
fi

###############################################################################
# Remove certificates by subject/common name
###############################################################################

for CERT_NAME in "${CERTIFICATE_NAMES[@]}"; do
  security find-certificate -a -c "$CERT_NAME" -Z /Library/Keychains/System.keychain 2>/dev/null |
    awk '/SHA-1 hash:/ {print $3}' |
    while read -r SHA1; do
      echo "Removing certificate: $CERT_NAME ($SHA1)"
      security delete-certificate -Z "$SHA1" /Library/Keychains/System.keychain
    done
done

echo "Done."