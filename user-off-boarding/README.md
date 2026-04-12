# AD & Microsoft 365 User Offboarding Script

A PowerShell script that automates the full offboarding workflow for Active Directory and Microsoft 365 users.

---

## Overview

This script streamlines the user offboarding process by performing the following tasks in sequence:

1. Converts the user's mailbox to a **shared mailbox**
2. **Disables** the on-premises Active Directory account
3. **Removes** the user from all AD groups (except Domain Users)
4. **Removes** the user from Azure AD cloud-only groups
5. **Hides** the user from the Global Address List (GAL)

All actions are logged to an output file for audit purposes.

---

## Prerequisites

### PowerShell Modules

The following modules must be installed and available before running the script:

| Module | Purpose |
|---|---|
| `ActiveDirectory` | Manage on-prem AD accounts and groups |
| `MSOnline` | Connect to Microsoft Online Services |
| `ExchangeOnlineManagement` | Manage Exchange Online mailboxes |
| `AzureAD` | Manage Azure AD users and cloud groups |

### Permissions

The operator must have sufficient administrative rights in:
- On-premises Active Directory
- Exchange Online
- Azure Active Directory / Microsoft 365

---

## Configuration

Before running the script, update the following placeholders:

### Organisational Unit Paths

Locate this section near the bottom of the script and replace both `<ORGANIZATIONAL UNIT PATH>` values with the Distinguished Names (DNs) of the OUs you want to search for users:

```powershell
$User_To_Disable = '<ORGANIZATIONAL UNIT PATH>', '<ORGANIZATIONAL UNIT PATH>' |
ForEach-Object {
    Get-ADUser -Filter {(enabled -eq $true) -and (displayName -like "*")} `
    -searchBase $_ ...
```

**Example:**
```powershell
$User_To_Disable = 'OU=Staff,DC=contoso,DC=com', 'OU=Contractors,DC=contoso,DC=com' |
```

### Log File Output Path

Locate the log redirection line and replace `<OUTPUT Path>` with the desired folder path:

```powershell
) *>> "<OUTPUT Path>\OffboardingLog.txt"
```

**Example:**
```powershell
) *>> "C:\Logs\Offboarding\OffboardingLog.txt"
```

---

## Usage

1. Open PowerShell as an administrator.
2. Run the script:
   ```powershell
   .\Offboarding.ps1
   ```
3. A grid view window will appear — select the user to offboard and click **OK**.
4. A confirmation dialog will prompt you to confirm the action.
5. If confirmed, you will be prompted to enter your Microsoft 365 credentials to connect to cloud services.
6. The script will execute each offboarding step and log all output to `OffboardingLog.txt`.
7. A completion popup will notify you when the process is finished.

> **Note:** Only one user can be offboarded per script execution.

---

## Logging

All console output during the offboarding process is captured and appended to:

```
<OUTPUT Path>\OffboardingLog.txt
```

Each run is logged with a timestamp and the operator's username (the account that ran the script). Log entries include confirmation of each completed step.

---

## Script Functions

| Function | Description |
|---|---|
| `Get-TimeStamp` | Returns a formatted timestamp for log entries |
| `ConnectToO365` | Prompts for credentials and connects to MSOnline, Exchange Online, and Azure AD |
| `ConvertToSharedMailbox` | Converts the specified mailbox to a shared mailbox |
| `DisableAccount` | Clears manager/phone attributes and disables the AD account |
| `RemoveFromGroups` | Removes the user from all on-prem AD groups except Domain Users |
| `HideUser` | Sets the `msExchHideFromAddressLists` attribute to hide the user from the GAL |
| `RemoveCloudGroups` | Removes the user from Azure AD cloud-only groups (not synced from on-prem) |

---

## Notes

- The script uses `wscript.shell` popups for operator interaction — ensure you are running in an interactive session.
- Azure AD group removal only targets **cloud-only groups** (where `DirSyncEnabled` is not set). On-prem synced groups are handled separately via the AD module.
- After mailbox conversion, the Exchange Online session is disconnected before proceeding with AD tasks.
- All variables are cleared at the end of the script for cleanup.

---

## Troubleshooting

**"No user selected to off board"** — You closed the grid view without selecting a user. Re-run the script and select a user.

**"You can only off board one user per time"** — Multiple users were selected in the grid view. Re-run and select only one.

**"[User] is not in AD, or script is not functioning properly"** — The `SamAccountName` could not be resolved when removing group memberships. Verify the account exists and the AD module is connected to the correct domain.

**"[UPN] doesn't have cloud assigned groups or the script encountered an error"** — The user may have no cloud-only group memberships, or the Azure AD connection failed. Check that `Connect-AzureAD` succeeded during the credential prompt.

---

## Author

| Field | Value |
|---|---|
| Author | `<Your Name>` |
| Date | `<Date>` |
| Version | 1.0 |
