<#
.SYNOPSIS
    Promotes Application Workspace package snapshots Test -> Acceptance -> Production
    after a configurable number of days, with optional HTML email summary.

.EXAMPLE
    .\Move-AWStages.ps1 -DaysBeforeAcceptance 3 -DaysBeforeProduction 7

.EXAMPLE
    $smtpCred = Get-Credential
    .\Move-AWStages.ps1 -DaysBeforeAcceptance 2 -DaysBeforeProduction 5 `
        -EnableEmail -SmtpServer 'smtp.office365.com' -SmtpPort 587 -UseSsl `
        -MailFrom 'aw-automation@contoso.com' -MailTo 'appteam@contoso.com','chris@contoso.com' `
        -SmtpCredential $smtpCred
#>
param(
    # Days a snapshot must sit in Test before moving to Acceptance
    [ValidateRange(0, 365)]
    [int]$DaysBeforeAcceptance = 3,

    # Days a snapshot must sit in Acceptance before moving to Production
    [ValidateRange(0, 365)]
    [int]$DaysBeforeProduction = 7,

    # Testing: evaluate eligibility as if the script were running on this date
    [datetime]$AsOfDate = (Get-Date),

    # --- Email notification options ---
    [switch]$EnableEmail,
    [string]$SmtpServer,
    [int]$SmtpPort = 587,
    [switch]$UseSsl,
    [string]$MailFrom,
    [string[]]$MailTo,
    [pscredential]$SmtpCredential,
    [string]$MailSubject = "Application Workspace Stage Promotions - $(Get-Date -Format 'yyyy-MM-dd')",

    # Send the email even when no packages were moved
    [switch]$SendEmailWhenNoChanges,
    # Show packages that will become eligible within the next N days
    [switch]$IncludeUpcoming,
    [ValidateRange(1, 365)]
    [int]$UpcomingDays = 7,

    # DPAPI-protected credential files (Export-Clixml)
    [string]$AWCredentialPath   = 'C:\ProgramData\AWStagePromote\aw.cred.xml',
    [string]$SmtpCredentialPath = 'C:\ProgramData\AWStagePromote\smtp.cred.xml'
)

# Check for powershell Module and install if necessary
if (-not (Get-Module -ListAvailable -Name Liquit.Server.PowerShell)) {
    Install-Module -Name Liquit.Server.PowerShell -Scope CurrentUser -Force
}

# --- Logging setup (continuous log) ---
$LogRoot = "C:\Temp\Application Workspace\Logs"
if (-not (Test-Path $LogRoot)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}

# ONE continuous log file
$LogFile = Join-Path $LogRoot "LiquitStagePromote.log"

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "{0} [{1}] {2}" -f $ts, $Level, $Message
    # Console output (captured by task history if configured)
    Write-Output $line
    # Append as UTF-8. Retry briefly if the file is temporarily locked (e.g.
    # open in an editor or log viewer) so a transient lock does not crash the
    # run. FileShare.ReadWrite also lets other apps keep it open while we write.
    $maxAttempts = 5
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        try {
            $stream = [System.IO.FileStream]::new(
                $LogFile,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            try {
                $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
                $writer.WriteLine($line)
                $writer.Flush()
                $writer.Dispose()
            }
            finally {
                $stream.Dispose()
            }
            break
        }
        catch [System.IO.IOException] {
            if ($attempt -eq $maxAttempts) {
                Write-Warning "Write-Log: could not write to '$LogFile' after $maxAttempts attempts: $($_.Exception.Message)"
            }
            else {
                Start-Sleep -Milliseconds 150
            }
        }
    }
}

# --- Email helpers ---
function ConvertTo-HtmlSafe {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-PromotionReportHtml {
    param(
        [object[]]$Results,
        [bool]$IsPreview,
        [int]$AcceptanceDays,
        [int]$ProductionDays,
        [object[]]$Upcoming = @()
    )

    $runTime   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $succeeded = @($Results | Where-Object { $_.Status -in 'Moved','Preview' }).Count
    $failed    = @($Results | Where-Object { $_.Status -eq 'Failed' }).Count

    $th = 'style="background:#1F3864;color:#FFFFFF;text-align:left;padding:8px 10px;font-size:13px;"'
    $td = 'padding:8px 10px;border-bottom:1px solid #E1E4E8;font-size:13px;'

    $rows = New-Object System.Text.StringBuilder
    $i = 0
    foreach ($r in ($Results | Sort-Object ToStage, PackageName)) {
        $bg = if ($i % 2 -eq 0) { '#FFFFFF' } else { '#F6F8FA' }
        $statusColor = switch ($r.Status) {
            'Moved'   { '#1E7B34' }
            'Preview' { '#9A6700' }
            'Failed'  { '#B42318' }
            default   { '#333333' }
        }
        $stageColor = if ($r.ToStage -eq 'Production') { '#5B2C83' } else { '#0B5CAD' }
        [void]$rows.Append(@"
<tr style="background:$bg;">
  <td style="$td">$(ConvertTo-HtmlSafe $r.PackageName)</td>
  <td style="$td">$(ConvertTo-HtmlSafe $r.PackageVersion)</td>
  <td style="$td">$(ConvertTo-HtmlSafe $r.FromStage)</td>
  <td style="$td"><span style="color:$stageColor;font-weight:bold;">$(ConvertTo-HtmlSafe $r.ToStage)</span></td>
  <td style="$td">$(ConvertTo-HtmlSafe $r.DaysInStage)</td>
  <td style="$td"><span style="color:$statusColor;font-weight:bold;">$(ConvertTo-HtmlSafe $r.Status)</span>$(if ($r.Error) { "<br/><span style='color:#B42318;font-size:12px;'>$(ConvertTo-HtmlSafe $r.Error)</span>" })</td>
</tr>
"@)
        $i++
    }

    $previewBanner = ''
    if ($IsPreview) {
        $previewBanner = '<div style="background:#FFF8C5;border:1px solid #D4A72C;padding:10px;margin-bottom:16px;font-size:13px;"><b>Preview mode:</b> no packages were moved. The items below are what <i>would</i> be promoted.</div>'
    }

    if ($Results.Count -gt 0) {
        $body = @"
<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;border:1px solid #E1E4E8;">
  <tr>
    <th $th>Package</th>
    <th $th>Version</th>
    <th $th>From</th>
    <th $th>To</th>
    <th $th>Days in Stage</th>
    <th $th>Status</th>
  </tr>
  $($rows.ToString())
</table>
"@
    }
    else {
        $body = '<p style="font-size:13px;">No packages met the criteria for promotion on this run.</p>'
    }
    $upcomingSection = ''
    if ($Upcoming -and $Upcoming.Count -gt 0) {
        $upRows = New-Object System.Text.StringBuilder
        $j = 0
        foreach ($u in ($Upcoming | Sort-Object EligibleOn, PackageName)) {
            $bg = if ($j % 2 -eq 0) { '#FFFFFF' } else { '#F6F8FA' }
            $stageColor = if ($u.ToStage -eq 'Production') { '#5B2C83' } else { '#0B5CAD' }
            [void]$upRows.Append(@"
<tr style="background:$bg;">
  <td style="$td">$(ConvertTo-HtmlSafe $u.PackageName)</td>
  <td style="$td">$(ConvertTo-HtmlSafe $u.PackageVersion)</td>
  <td style="$td">$(ConvertTo-HtmlSafe $u.FromStage)</td>
  <td style="$td"><span style="color:$stageColor;font-weight:bold;">$(ConvertTo-HtmlSafe $u.ToStage)</span></td>
  <td style="$td">$($u.EligibleOn.ToString('yyyy-MM-dd'))</td>
  <td style="$td">$($u.DaysRemaining)</td>
</tr>
"@)
            $j++
        }
        $upcomingSection = @"
<h3 style="margin:28px 0 4px 0;color:#1F3864;">Upcoming Promotions (next $UpcomingDays days)</h3>
<p style="margin:0 0 10px 0;font-size:12px;color:#57606A;">Projected from current snapshots. Publishing a new version to Test or Acceptance resets its timer.</p>
<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;border:1px solid #E1E4E8;">
  <tr>
    <th $th>Package</th>
    <th $th>Version</th>
    <th $th>From</th>
    <th $th>To</th>
    <th $th>Eligible On</th>
    <th $th>Days Remaining</th>
  </tr>
  $($upRows.ToString())
</table>
"@
    }
    return @"
<html>
<body style="font-family:Segoe UI,Arial,sans-serif;color:#24292F;margin:0;padding:20px;">
  <h2 style="margin:0 0 4px 0;color:#1F3864;">Application Workspace Stage Promotions</h2>
  <p style="margin:0 0 16px 0;font-size:12px;color:#57606A;">
    Run: $runTime &nbsp;|&nbsp; Host: $(ConvertTo-HtmlSafe $env:COMPUTERNAME) &nbsp;|&nbsp;
    Test &rarr; Acceptance after $AcceptanceDays day(s) &nbsp;|&nbsp; Acceptance &rarr; Production after $ProductionDays day(s)
  </p>
  $previewBanner
  <p style="font-size:13px;margin:0 0 12px 0;">
    <b>$succeeded</b> package(s) $(if ($IsPreview) { 'eligible' } else { 'promoted' }) &nbsp;|&nbsp; <b>$failed</b> failed
  </p>
  $body
  $upcomingSection
  <p style="font-size:11px;color:#8C959F;margin-top:20px;">Log file: $(ConvertTo-HtmlSafe $LogFile)</p>
</body>
</html>
"@
}

function Send-PromotionReport {
    param(
        [object[]]$Results,
        [object[]]$Upcoming = @(),
        [bool]$IsPreview
    )
    try {
        $html = New-PromotionReportHtml -Results $Results -Upcoming $Upcoming -IsPreview $IsPreview `
            -AcceptanceDays $DaysBeforeAcceptance -ProductionDays $DaysBeforeProduction

        $subject = if ($IsPreview) { "[PREVIEW] $MailSubject" } else { $MailSubject }

        $mailParams = @{
            SmtpServer  = $SmtpServer
            Port        = $SmtpPort
            From        = $MailFrom
            To          = $MailTo
            Subject     = $subject
            Body        = $html
            BodyAsHtml  = $true
            Encoding    = [System.Text.Encoding]::UTF8
            ErrorAction = 'Stop'
        }
        if ($UseSsl)         { $mailParams.UseSsl = $true }
        if ($SmtpCredential) { $mailParams.Credential = $SmtpCredential }

        Send-MailMessage @mailParams -WarningAction SilentlyContinue
        Write-Log -Message "Email summary sent to $($MailTo -join ', ')." -Level INFO
    }
    catch {
        Write-Log -Message "Failed to send email summary: $($_.Exception.Message)" -Level ERROR
    }
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($EnableEmail -and -not $SmtpCredential -and $SmtpCredentialPath) {
    if (-not (Test-Path $SmtpCredentialPath)) { throw "SMTP credential file not found: $SmtpCredentialPath" }
    $SmtpCredential = Import-Clixml -Path $SmtpCredentialPath
}
# Validate email parameters up front so a misconfiguration is caught before any work is done
if ($EnableEmail) {
    $missing = @()
    if (-not $SmtpServer) { $missing += 'SmtpServer' }
    if (-not $MailFrom)   { $missing += 'MailFrom' }
    if (-not $MailTo)     { $missing += 'MailTo' }
    if ($missing.Count -gt 0) {
        throw "-EnableEmail requires: $($missing -join ', ')"
    }
}

Write-Log "===== Run started. Host=$env:COMPUTERNAME User=$env:USERNAME ====="
Write-Log "DaysBeforeAcceptance=$DaysBeforeAcceptance DaysBeforeProduction=$DaysBeforeProduction EnableEmail=$EnableEmail"
if ($AsOfDate.Date -ne (Get-Date).Date) {
    Write-Log -Message "AsOfDate override active: evaluating as of $($AsOfDate.ToString('yyyy-MM-dd HH:mm'))." -Level WARN
}
$LiquitURI = 'https://chris.liquit.com' # Replace this with your zone
if (-not (Test-Path $AWCredentialPath)) {
    Write-Log -Message "AW credential file not found: $AWCredentialPath" -Level ERROR
    throw "AW credential file not found: $AWCredentialPath"
}
try {
    $credentials = Import-Clixml -Path $AWCredentialPath
}
catch {
    Write-Log -Message "Could not decrypt AW credential file. It must be created by $env:USERNAME on $env:COMPUTERNAME. $($_.Exception.Message)" -Level ERROR
    throw
}

$PackageEvents = [System.Collections.ArrayList]::new()
$PromotionResults = [System.Collections.ArrayList]::new()
$UpcomingResults = [System.Collections.ArrayList]::new()

Connect-LiquitWorkspace -URI $LiquitURI -Credential $credentials

#$ExcludedPackages = @('*firefox*', '*chrome*')
#$AllPackages = Get-LiquitPackage | Where-Object {($_.name -like "WIN -*" -or $_.name -like "MAC -*" -or $_.name -like "MS -*")}

function Get-LiquitPackageFiltered {
    [CmdletBinding()]
    param(
        # Wildcard patterns (recommended): '*chrome*', 'WIN -*', etc.
        [string[]]$IncludedPackages,
        # Wildcard patterns (recommended): '*firefox*', '*chrome*', etc.
        [string[]]$ExcludedPackages,
        # Keep your "WIN - / MAC - / MS -" scope as default behavior
        [switch]$LimitToPlatformPrefixes,
        # If you want to override prefixes, pass your own (wildcards)
        [string[]]$PlatformPrefixPatterns = @('WIN -*','MAC -*','MS -*')
    )
    # Normalize arrays (treat $null / empty as "not provided")
    $hasIncludes = $IncludedPackages -and $IncludedPackages.Count -gt 0
    $hasExcludes = $ExcludedPackages -and $ExcludedPackages.Count -gt 0
    # Pull packages once
    $pkgs = Get-LiquitPackage
    $pkgs | Where-Object {
        $name = $_.Name
        # 1) Optional platform prefix limiter
        $platformOk = $true
        if ($LimitToPlatformPrefixes) {
            $platformOk = [bool]($PlatformPrefixPatterns.Where({ $name -like $_ }, 'First').Count)
        }
        # 2) Include logic:
        #    - If includes provided: keep only if name matches ANY include pattern
        #    - If no includes provided: allow all
        $includeOk = $true
        if ($hasIncludes) {
            $includeOk = [bool]($IncludedPackages.Where({ $name -like $_ }, 'First').Count)
        }
        # 3) Exclude logic:
        #    - If excludes provided: reject if name matches ANY exclude pattern
        #    - If no excludes provided: reject none
        $excludeHit = $false
        if ($hasExcludes) {
            $excludeHit = [bool]($ExcludedPackages.Where({ $name -like $_ }, 'First').Count)
        }
        $platformOk -and $includeOk -and (-not $excludeHit)
    }
}

$IncludedPackages = @('WIN -*','MAC -*','MS -*')
#$IncludedPackages = @('MAC -*')
#$ExcludedPackages = @('*firefox*', '*chrome*')
$AllPackages = Get-LiquitPackageFiltered -IncludedPackages $IncludedPackages

ForEach ($Package in $AllPackages){
    Write-Log -Message "Processing $($Package.Name)." -Level INFO
    $TestPackageSnapshot = Get-LiquitPackageSnapshot -Package $Package -Type Test
    $AcceptancePackageSnapshot = Get-LiquitPackageSnapshot -Package $Package -Type Acceptance
    If ($TestPackageSnapshot) {
        Write-Log -Message "Found Test for $($Package.Name) with version $($TestPackageSnapshot.Name)." -Level INFO
        $EventDetails = New-Object PSObject -prop @{
            PackageID     = $Package.ID
            PackageName   = $Package.Name
            PackageVersion = $TestPackageSnapshot.Name
            SnapshotID    = $TestPackageSnapshot.ID.Guid
            DateOfStage   = $TestPackageSnapshot.CreatedAt   # or ModifiedAt if re-promotions matter
            SnapshotStage = 'Test'
        }
        [void]$PackageEvents.Add($EventDetails)
    }
    If ($AcceptancePackageSnapshot) {
        Write-Log -Message "Found Acceptance for $($Package.Name) with version $($AcceptancePackageSnapshot.Name)." -Level INFO
        $EventDetails = New-Object PSObject -prop @{
            PackageID     = $Package.ID
            PackageName   = $Package.Name
            PackageVersion = $AcceptancePackageSnapshot.Name
            SnapshotID    = $AcceptancePackageSnapshot.ID.Guid
            DateOfStage   = $AcceptancePackageSnapshot.CreatedAt   # or ModifiedAt if re-promotions matter
            SnapshotStage = 'Acceptance'
        }
        [void]$PackageEvents.Add($EventDetails)
    }
}

$Debug = $true

# Now that I have the data, let's move it
# Eligibility is evaluated in both modes so Debug runs can still produce a preview email.
# Promote Acceptance -> Production before Test -> Acceptance so a Test promotion
# never replaces an Acceptance snapshot that is still waiting to be moved.
$OrderedEvents = $PackageEvents | Sort-Object `
    @{ Expression = { if ($_.SnapshotStage -eq 'Acceptance') { 0 } else { 1 } } }, `
    PackageName

ForEach ($StageEvent in $OrderedEvents) {
    $targetStage = $null
    Switch ($StageEvent.SnapshotStage) {
        'Test'       { $targetStage = 'Acceptance'; $daysRequired = $DaysBeforeAcceptance }
        'Acceptance' { $targetStage = 'Production'; $daysRequired = $DaysBeforeProduction }
    }
    If (-not $targetStage) { continue }

    $dateToCheck = $null
    If ($null -ne $StageEvent.DateOfStage) {
        $dateToCheck = Get-Date $($StageEvent.DateOfStage)
    }
    $cutoffDate = $AsOfDate.AddDays(-$daysRequired)

    If ($dateToCheck -lt $cutoffDate -or $null -eq $StageEvent.DateOfStage) {
        $daysInStage = if ($dateToCheck) { [math]::Floor(($AsOfDate - $dateToCheck).TotalDays) } else { 'Unknown' }

        $result = [PSCustomObject]@{
            PackageName    = $StageEvent.PackageName
            PackageVersion = $StageEvent.PackageVersion
            FromStage      = $StageEvent.SnapshotStage
            ToStage        = $targetStage
            DaysInStage    = $daysInStage
            Status         = $null
            Error          = $null
        }

        If ($Debug) {
            Write-Log -Message "[DEBUG] Would move $($StageEvent.PackageName) with version $($StageEvent.PackageVersion) to $targetStage stage." -Level DEBUG
            $result.Status = 'Preview'
        }
        Else {
            try {
                $CurrentPackage = Get-LiquitPackage -ID $StageEvent.PackageID
                $CurrentSnapshot = Get-LiquitPackageSnapshot -ID $StageEvent.SnapshotID -Package $CurrentPackage
                Write-Log -Message "Moving $($StageEvent.PackageName) with version $($StageEvent.PackageVersion) to $targetStage stage." -Level WARN
                Publish-LiquitPackageSnapshot -Snapshot $CurrentSnapshot -Stage $targetStage -Name $StageEvent.PackageVersion -Description $StageEvent.PackageName -ErrorAction Stop
                $result.Status = 'Moved'
            }
            catch {
                Write-Log -Message "Failed to move $($StageEvent.PackageName) $($StageEvent.PackageVersion) to ${targetStage}: $($_.Exception.Message)" -Level ERROR
                $result.Status = 'Failed'
                $result.Error  = $_.Exception.Message
            }
        }

        [void]$PromotionResults.Add($result)
    }
    ElseIf ($IncludeUpcoming) {
        $eligibleOn    = $dateToCheck.AddDays($daysRequired)
        $daysRemaining = ($eligibleOn.Date - $AsOfDate.Date).Days
        If ($eligibleOn -le $AsOfDate.AddDays($UpcomingDays)) {
            Write-Log -Message "Upcoming: $($StageEvent.PackageName) $($StageEvent.PackageVersion) eligible for $targetStage on $($eligibleOn.ToString('yyyy-MM-dd'))." -Level INFO
            [void]$UpcomingResults.Add([PSCustomObject]@{
                PackageName    = $StageEvent.PackageName
                PackageVersion = $StageEvent.PackageVersion
                FromStage      = $StageEvent.SnapshotStage
                ToStage        = $targetStage
                EligibleOn     = $eligibleOn
                DaysRemaining  = $daysRemaining
            })
        }
    }
}

# An Acceptance snapshot is replaced when its package's Test snapshot is promoted this run,
# so don't list that Acceptance snapshot as upcoming
$replacedPkgs = @($PromotionResults | Where-Object { $_.FromStage -eq 'Test' -and $_.Status -in 'Moved','Preview' } | ForEach-Object { $_.PackageName })
$UpcomingResults = [System.Collections.ArrayList]@($UpcomingResults | Where-Object {
    -not ($_.FromStage -eq 'Acceptance' -and $_.PackageName -in $replacedPkgs)
})

# --- Email summary ---
If ($EnableEmail) {
    If ($PromotionResults.Count -gt 0 -or $UpcomingResults.Count -gt 0 -or $SendEmailWhenNoChanges) {
        Send-PromotionReport -Results $PromotionResults.ToArray() -Upcoming $UpcomingResults.ToArray() -IsPreview $Debug
    }
    Else {
        Write-Log -Message "No packages promoted; email skipped (use -SendEmailWhenNoChanges to always send)." -Level INFO
    }
}

Write-Log "===== Run finished. Promoted/eligible=$($PromotionResults.Count) ====="
try { Stop-Transcript | Out-Null } catch {}
