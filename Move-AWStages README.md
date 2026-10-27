# Move-AWStages

Automatically promotes Application Workspace package snapshots through the release stages:

**Test → Acceptance → Production**

A snapshot is promoted once it has been in its current stage for a set number of days. The script can also send an HTML email that lists what was promoted, any failures, and (optionally) what is coming up in the next few days.

---

## Contents

- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Script settings](#script-settings)
- [Parameters](#parameters)
- [Credential storage (DPAPI)](#credential-storage-dpapi)
- [Scheduled task setup](#scheduled-task-setup)
- [Example commands](#example-commands)
- [Testing](#testing)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)

---

## How it works

1. Connects to the Application Workspace zone using the stored service account credential.
2. Collects every package that matches the include patterns (default: `WIN -*`, `MAC -*`, `MS -*`).
3. For each package, reads the current **Test** and **Acceptance** snapshots and their creation dates.
4. Decides which snapshots are eligible:
   - **Test → Acceptance** when the Test snapshot is older than `-DaysBeforeAcceptance`.
   - **Acceptance → Production** when the Acceptance snapshot is older than `-DaysBeforeProduction`.
5. Promotes eligible snapshots, **Acceptance → Production first**, then **Test → Acceptance**. This order matters: promoting Test first would replace the Acceptance snapshot before it could move to Production, and that move would fail with a "resource not found" error.
6. Each promotion runs on its own. If one fails, the error is logged and included in the email, and the run continues.
7. Optionally works out **upcoming** promotions: snapshots that aren't eligible yet but will be within `-UpcomingDays`.
8. Optionally sends an HTML summary email.

### Debug (preview) mode

The script contains a hard-coded switch:

```powershell
$Debug = $true
```

| Value | Behavior |
|---|---|
| `$true` | **Nothing is moved.** Eligible packages are logged and emailed with status **Preview**. The email subject starts with `[PREVIEW]` and includes a yellow banner. |
| `$false` | Eligible packages are promoted. Status is **Moved** or **Failed**. |

> Set `$Debug = $false` only after a preview run looks correct.

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows PowerShell 5.1 | Runs as `powershell.exe` |
| `Liquit.Server.PowerShell` module | Install for **all users** so the service account can use it (see below) |
| AW zone service account | A local AW account (for example `local\aw.api`) with rights to read and publish package snapshots |
| Windows service account | The account the scheduled task runs as (for example `CS\svc_awautomation`) |
| SMTP mailbox | Only needed for email. SMTP AUTH must be allowed for that mailbox in Microsoft 365. |
| Network | Outbound HTTPS to the AW zone; outbound TCP 587 to the SMTP server when email is enabled |

Install the module once from an **elevated** prompt:

```powershell
Install-Module -Name Liquit.Server.PowerShell -Scope AllUsers -Force
```

> The script's built-in install check uses `-Scope CurrentUser`. That doesn't reliably work in a scheduled task, so pre-install the module with `-Scope AllUsers`.

---

## Script settings

These are set inside the script rather than as parameters. Review them before first use.

| Setting | Default | Purpose |
|---|---|---|
| `$LiquitURI` | `https://chris.liquit.com` | Your Application Workspace zone URL |
| `$IncludedPackages` | `@('WIN -*','MAC -*','MS -*')` | Wildcard patterns for packages to process |
| `$ExcludedPackages` | Commented out | Optional wildcard patterns to skip, for example `@('*firefox*')` |
| `$Debug` | `$true` | Preview mode. See [Debug (preview) mode](#debug-preview-mode). |
| `$LogRoot` | `C:\Temp\Application Workspace\Logs` | Log folder |

---

## Parameters

### Promotion timing

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-DaysBeforeAcceptance` | int (0–365) | `3` | Days a snapshot must stay in Test before it moves to Acceptance |
| `-DaysBeforeProduction` | int (0–365) | `7` | Days a snapshot must stay in Acceptance before it moves to Production |
| `-AsOfDate` | datetime | Current date/time | **Testing only.** Evaluates eligibility as if the script were running on this date. Leave it out in production. |

### Credentials

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-AWCredentialPath` | string | `C:\ProgramData\AWStagePromote\aw.cred.xml` | DPAPI-protected credential file for the AW zone account |
| `-SmtpCredentialPath` | string | `C:\ProgramData\AWStagePromote\smtp.cred.xml` | DPAPI-protected credential file for the SMTP mailbox |
| `-SmtpCredential` | PSCredential | — | Pass a credential object directly. Interactive use only; takes priority over `-SmtpCredentialPath`. |

### Email

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-EnableEmail` | switch | Off | Turns on the summary email |
| `-SmtpServer` | string | — | SMTP host, for example `smtp.office365.com`. Required with `-EnableEmail`. |
| `-SmtpPort` | int | `587` | SMTP port |
| `-UseSsl` | switch | Off | Use TLS (STARTTLS). Needed for Microsoft 365. |
| `-MailFrom` | string | — | Sender address. Required with `-EnableEmail`. Must match the SMTP account or one it can send as. |
| `-MailTo` | string[] | — | One or more recipients. Required with `-EnableEmail`. |
| `-MailSubject` | string | `Application Workspace Stage Promotions - <date>` | Email subject |
| `-SendEmailWhenNoChanges` | switch | Off | Send the email even when nothing was promoted and nothing is upcoming |

### Upcoming promotions

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-IncludeUpcoming` | switch | Off | Adds an **Upcoming Promotions** table to the email |
| `-UpcomingDays` | int (1–365) | `7` | How many days ahead to look |

The upcoming table shows only each package's **next** move. Publishing a new version to Test or Acceptance resets that package's timer.

### When is the email sent?

With `-EnableEmail`, an email is sent if **any** of these is true:

- At least one package was promoted, previewed, or failed.
- `-IncludeUpcoming` is set and at least one package is upcoming.
- `-SendEmailWhenNoChanges` is set.

If an email can't be sent, the error is logged and the rest of the run is unaffected.

---

## Credential storage (DPAPI)

The script doesn't store passwords in plain text. Credentials are saved with `Export-Clixml`, which encrypts the password with Windows **DPAPI** (Data Protection API).

### How DPAPI protection works

- The password in the XML file is encrypted with a key tied to **the Windows account that created the file** and **the computer it was created on**.
- Only that same account, on that same computer, can decrypt it with `Import-Clixml`.
- If the file is copied to another server, or another account tries to read it (including an administrator or SYSTEM), decryption fails.
- The username is stored in readable text. Only the password is encrypted.

**This means the files must be created while signed in as the account the scheduled task runs as, on the server that runs the task.**

### Create the credential files

1. Open PowerShell **as the service account**:

   ```powershell
   runas /user:CS\svc_awautomation powershell.exe
   ```

2. In that window, confirm the account and create the files:

   ```powershell
   whoami   # must show cs\svc_awautomation

   $credDir = 'C:\ProgramData\AWStagePromote'
   New-Item -Path $credDir -ItemType Directory -Force | Out-Null

   # AW zone account - include the local\ prefix
   Get-Credential -UserName 'local\aw.api' -Message 'AW zone' |
       Export-Clixml "$credDir\aw.cred.xml"

   # SMTP mailbox account
   Get-Credential -UserName 'smtptestaccount@cs.recastsoftware.com' -Message 'SMTP' |
       Export-Clixml "$credDir\smtp.cred.xml"
   ```

3. Restrict the folder to the service account, administrators, and SYSTEM (run from an **elevated** prompt):

   ```powershell
   icacls 'C:\ProgramData\AWStagePromote' /inheritance:r /grant:r "CS\svc_awautomation:(OI)(CI)R" "Administrators:(OI)(CI)F" "SYSTEM:(OI)(CI)F"
   ```

4. Verify decryption, still as the service account:

   ```powershell
   Import-Clixml 'C:\ProgramData\AWStagePromote\aw.cred.xml'
   Import-Clixml 'C:\ProgramData\AWStagePromote\smtp.cred.xml'
   ```

   Each should return the username and `System.Security.SecureString`. An error means the file was created by a different account or on a different computer.

### When to recreate the files

Run the steps above again if any of these change:

- The scheduled task's run-as account
- The server the task runs on
- The AW account password or the SMTP mailbox password
- The service account's Windows password is reset by an administrator (rather than changed by the account itself), which can make existing DPAPI files unreadable

> A gMSA is not recommended with this method, because you can't sign in interactively as a gMSA to run `Get-Credential`.

---

## Scheduled task setup

### 1. Prepare the service account

- Create a domain account, for example `CS\svc_awautomation`.
- Give it the **Log on as a batch job** right on the server (Local Security Policy, or Group Policy).
- Give it **read** access to the script folder and **modify** access to the log folder:

  ```powershell
  icacls "C:\Temp\Application Workspace\Logs" /grant "CS\svc_awautomation:(OI)(CI)M"
  ```

### 2. Prepare the server

1. Install the Liquit module with `-Scope AllUsers` ([Requirements](#requirements)).
2. Copy the script to a fixed location, for example `C:\temp\Move-AWStages_Updated.ps1`.
3. Create the DPAPI credential files as the service account ([Credential storage](#credential-storage-dpapi)).
4. Review the [script settings](#script-settings) and run a preview with `$Debug = $true`.

### 3. Create the task

In **Task Scheduler**, choose **Create Task** (not Create Basic Task).

**General tab**

| Setting | Value |
|---|---|
| Name | `Application Workspace Move-AWStages` |
| When running the task, use the following user account | `CS\svc_awautomation` |
| Security option | **Run whether user is logged on or not** |
| Do not store password | **Unchecked** (the task needs the stored password to use DPAPI) |
| Run with highest privileges | Not required |

**Triggers tab**

- **New** → **Daily** → choose a time, for example 12:00 AM.

**Actions tab**

| Field | Value |
|---|---|
| Action | Start a program |
| Program/script | `powershell.exe` |
| Add arguments | See below |
| Start in | `C:\temp` (the script folder) |

**Add arguments:**

```text
-NoProfile -ExecutionPolicy Bypass -File "C:\temp\Move-AWStages_Updated.ps1" -DaysBeforeAcceptance 3 -DaysBeforeProduction 5 -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo chrisa@recastsoftware.com -SmtpCredentialPath "C:\ProgramData\AWStagePromote\smtp.cred.xml" -IncludeUpcoming -UpcomingDays 7
```

Click **OK** and enter the service account's password when prompted.

### Rules for task arguments

`-File` passes every argument as **plain text**. PowerShell expressions are not evaluated.

| Don't | Do |
|---|---|
| `-AsOfDate (Get-Date).AddDays(7)` | `-AsOfDate "2026-10-14"` (testing only) |
| `-SmtpCredential $smtpCred` | `-SmtpCredentialPath "C:\ProgramData\AWStagePromote\smtp.cred.xml"` |
| Curly quotes `“ ”` pasted from documents | Straight quotes `"` |

If an argument can't be converted, the script fails before logging starts and the task shows **Last Run Result `0x1`**.

### 4. Verify

1. Right-click the task → **Run**.
2. Press **F5**. **Last Run Result** should be `The operation completed successfully. (0x0)`.
3. Check `C:\Temp\Application Workspace\Logs\LiquitStagePromote.log` for `Run started` and `Run finished`.
4. When the preview looks right, set `$Debug = $false` in the script.

---

## Example commands

**Preview run with default timing, no email**

```powershell
.\Move-AWStages_Updated.ps1
```

**Custom timing**

```powershell
.\Move-AWStages_Updated.ps1 -DaysBeforeAcceptance 2 -DaysBeforeProduction 5
```

**Email using the stored SMTP credential**

```powershell
.\Move-AWStages_Updated.ps1 -DaysBeforeAcceptance 3 -DaysBeforeProduction 5 `
    -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl `
    -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo chrisa@recastsoftware.com `
    -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml'
```

**Email with an upcoming forecast for the next 7 days**

```powershell
.\Move-AWStages_Updated.ps1 -DaysBeforeAcceptance 3 -DaysBeforeProduction 5 `
    -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl `
    -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo chrisa@recastsoftware.com `
    -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml' `
    -IncludeUpcoming -UpcomingDays 7
```

**Multiple recipients**

```powershell
.\Move-AWStages_Updated.ps1 -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl `
    -MailFrom smtptestaccount@cs.recastsoftware.com `
    -MailTo 'chrisa@recastsoftware.com','appteam@recastsoftware.com' `
    -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml'
```

> In a scheduled task, multiple recipients are most reliable with the `-Command` form, because `-File` passes `-MailTo` as one string:
>
> ```text
> -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Move-AWStages_Updated.ps1' -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo 'chrisa@recastsoftware.com','appteam@recastsoftware.com' -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml' -IncludeUpcoming; exit $LASTEXITCODE"
> ```

**Always send an email, even when nothing changes**

```powershell
.\Move-AWStages_Updated.ps1 -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl `
    -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo chrisa@recastsoftware.com `
    -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml' -SendEmailWhenNoChanges
```

**Simulate a future date (testing)**

```powershell
.\Move-AWStages_Updated.ps1 -AsOfDate '2026-10-14' -EnableEmail -SmtpServer smtp-legacy.office365.com -UseSsl `
    -MailFrom smtptestaccount@cs.recastsoftware.com -MailTo chrisa@recastsoftware.com `
    -SmtpCredentialPath 'C:\ProgramData\AWStagePromote\smtp.cred.xml' -SendEmailWhenNoChanges
```

---

## Testing

### Test the SMTP connection

```powershell
Test-NetConnection smtp-legacy.office365.com -Port 587
```

`TcpTestSucceeded` must be `True`.

### Test sending an email without the script

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$cred = Import-Clixml 'C:\ProgramData\AWStagePromote\smtp.cred.xml'   # run as the service account
Send-MailMessage -SmtpServer smtp-legacy.office365.com -Port 587 -UseSsl -Credential $cred `
    -From smtptestaccount@cs.recastsoftware.com -To chrisa@recastsoftware.com `
    -Subject 'SMTP test' -Body '<b>It works</b>' -BodyAsHtml -ErrorAction Stop
```

| Error | Cause |
|---|---|
| 5.7.57 Client not authenticated | SMTP AUTH is disabled for the mailbox or tenant |
| 5.7.60 SendAsDenied | `-MailFrom` doesn't match the account, or the account lacks Send As rights |
| 5.7.139 Authentication unsuccessful | Wrong password, or MFA / Security Defaults block basic auth |

### Test promotions and the upcoming forecast

`-AsOfDate` changes only the date the script compares against. Nothing in Application Workspace is modified.

1. Publish a test package to **Test**.
2. With `$Debug = $true`, run without `-AsOfDate`. With the default 3 days, the package appears as **upcoming** (needs `-IncludeUpcoming`).
3. Run with `-AsOfDate` set 3 or more days after the snapshot was created. The package appears as **Preview** (eligible).
4. Set `$Debug = $false` and repeat to promote for real. Each run moves a snapshot one stage.

> A package that is already eligible appears in the main table, not the upcoming table. If `-AsOfDate` is far in the future, the upcoming table can be empty.

---

## Logging

| File | Contents |
|---|---|
| `C:\Temp\Application Workspace\Logs\LiquitStagePromote.log` | Continuous log. Each run starts with `===== Run started` and ends with `===== Run finished`. |

Useful log lines:

| Line | Meaning |
|---|---|
| `[DEBUG] Would move ...` | Preview mode: the package is eligible |
| `[WARN] Moving ...` | The package is being promoted |
| `[ERROR] Failed to move ...` | The promotion failed. The reason follows. |
| `[INFO] Upcoming: ...` | The package is upcoming |
| `[WARN] AsOfDate override active` | `-AsOfDate` is set. Shouldn't appear in production runs. |
| `[INFO] Email summary sent to ...` | Email handed off to the SMTP server |
| `[ERROR] Failed to send email summary: ...` | Email failed. See the SMTP error table. |

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Task result `0x1`, nothing in the log | Argument error, such as an expression passed with `-File` | Check the arguments against [Rules for task arguments](#rules-for-task-arguments) |
| `Could not decrypt AW credential file` | The file was created by a different account or on a different computer | Recreate it as the service account ([DPAPI](#credential-storage-dpapi)) |
| `AW credential file not found` | Wrong path, or the account can't read the folder | Check `-AWCredentialPath` and the folder permissions |
| Module not found / install fails | Module installed for another user only | `Install-Module Liquit.Server.PowerShell -Scope AllUsers -Force` |
| `Failed to move ... resource not found` | The snapshot was replaced in the Application Workspace console before the move | Normal if a newer version was published. The newer version follows its own timer. |
| Works in a PowerShell window but not as a task | Different account, `-File` argument handling, or empty **Start in** | Test with `runas` as the service account; set **Start in** to the script folder |
| No upcoming table in the email | Nothing falls within `-UpcomingDays`, everything is already eligible, or `-IncludeUpcoming` is missing | Check the log for `Upcoming:` lines |
| Event history shows `User: SYSTEM` | That's the Task Scheduler service logging the event, not the run-as account | Check the run-as account on the task's **General** tab |

To capture errors that happen before logging starts, temporarily change the task arguments to use `-Command` and send the output to a file:

```text
-NoProfile -ExecutionPolicy Bypass -Command "& 'C:\temp\Move-AWStages_Updated.ps1' <arguments> *> 'C:\Temp\Application Workspace\Logs\task-output.log'; exit $LASTEXITCODE"
```
