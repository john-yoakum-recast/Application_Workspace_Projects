if (-not (Get-Module -ListAvailable -Name Liquit.Server.PowerShell)) {
    Install-Module -Name Liquit.Server.PowerShell -Scope CurrentUser -Force
}

$LiquitURI = 'https://zone.liquit.com' # Replace this with your zone
$username = 'local\admin' # Replace this with a service account you have created for creating and accessing this information
$password = 'PASSWORD' # Enter the password for that service Account
$credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (ConvertTo-SecureString -String $password -AsPlainText -Force)
$ScheduledTasks = "Synchronize Windows Apps","Synchronize Mac Apps"

Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials

ForEach ($ScheduledTask in $ScheduledTasks) {
    $CurrentTask = Get-LiquitScheduledTask -Name "$ScheduledTask"
    Start-LiquitScheduledTask -ScheduledTask $CurrentTask
}
