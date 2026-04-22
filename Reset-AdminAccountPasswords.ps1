<#
.SYNOPSIS
    Script to Dynamically reset admin accounts in Application Workspace

.DESCRIPTION
    This script will delete and recreate admin accounts in Application Workspace. Unfortunately, there is no direct module to just reset a password, so this is the workaround.
    It does need 2 admin accounts in the LOCAL database to perform these operations. And we use each other to recreate themselves.

.EXAMPLE
    .\Get-InstalledApplications.ps1 -UserAccountA "AdminA' -UserAccountB "AdminB" -UserAccountAPreviousPassword "A PreviousPassword" -UserAccountANewePassword "UserAccountA New Password" -UserAccountBNewPassword "UserAccount B New Password" -LiquitURI "FQDn of Your Zone"

#>
param (
    [string]$UserAccountA,
    [string]$UserAccountB,
    [string]$UserAccountAPreviousPassword,
    [string]$UserAccountANewPassword,
    [string]$UserAccountBNewPassword,
    [string]$LiquitURI = 'https://liquit.corp.viamonstra.com'
)

if (-not (Get-Module -ListAvailable -Name Liquit.Server.PowerShell)) {
    Install-Module -Name Liquit.Server.PowerShell -Scope CurrentUser -Force
}

Import-Module Liquit.Server.PowerShell

$username = "local\$UserAcountA" 
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (ConvertTo-SecureString -String "$UserAccountAPreviousPassword" -AsPlainText -Force)

# Connect to Application Workspace and delete UserAccountB
$ServiceRoot = New-Object Liquit.API.Server.V3.ServiceRoot([uri]"$LiquitURI", $credentials)
$ServiceRoot.Authenticate()
$AdminUser = $ServiceRoot.Users.List() | Where-Object {$_.Name -eq "$UserAccountB"}
$AdminUser.Delete()

# Refresh Connection to Application Workspace
Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials

# Recreate UserAccountB with new password
$IdentitySource = Get-LiquitIdentitySource -Name 'LOCAL' -Verbose
$SecurePassword = ConvertTo-SecureString -String "$UserAccountBNewPassword" -AsPlainText -Force
$NewAdminUser = New-LiquitUser -Name "$UserAccountB" -SecurePassword $SecurePassword -DisplayName "Admin Account B" -Source $IdentitySource -Verbose

# Refresh Connection to see new Account Created and assign Access Policy
Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials
$accessPolicy = Get-LiquitAccessPolicy -Name 'Administrator' -Verbose
$identity = Get-LiquitIdentity -Name "$UserAccountB"
New-LiquitPermission -Identity $identity -AccessPolicy $accessPolicy

# Congratulations, you have just recreated UserAccountB with the new password, time to repeat the process for UserAccountA



$username = "local\$UserAcountB" 
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (ConvertTo-SecureString -String "$UserAccountBNewPassword" -AsPlainText -Force)

# Connect to Application Workspace and delete UserAccountB
$ServiceRoot = New-Object Liquit.API.Server.V3.ServiceRoot([uri]"$LiquitURI", $credentials)
$ServiceRoot.Authenticate()
$AdminUser = $ServiceRoot.Users.List() | Where-Object {$_.Name -eq "$UserAccountA"}
$AdminUser.Delete()

# Refresh Connection to Application Workspace
Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials

# Recreate UserAccountB with new password
$IdentitySource = Get-LiquitIdentitySource -Name 'LOCAL' -Verbose
$SecurePassword = ConvertTo-SecureString -String "$UserAccountANewPassword" -AsPlainText -Force
$NewAdminUser = New-LiquitUser -Name "$UserAccountA" -SecurePassword $SecurePassword -DisplayName "Admin Account A" -Source $IdentitySource -Verbose

# Refresh Connection to see new Account Created and assign Access Policy
Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials
$accessPolicy = Get-LiquitAccessPolicy -Name 'Administrator' -Verbose
$identity = Get-LiquitIdentity -Name "$UserAccountA"
New-LiquitPermission -Identity $identity -AccessPolicy $accessPolicy

#Congratulations, you have now recreated both Admin accounts with new passwords
