<#
.SYNOPSIS
Automates the offboarding process for an Active Directory and Microsoft 365 user.

.DESCRIPTION
This script performs a complete offboarding workflow for a selected user, including:
- Converting the user's mailbox to a shared mailbox
- Disabling the on-prem Active Directory account
- Removing the user from all AD groups except "Domain Users"
- Removing the user from Azure AD cloud-only groups
- Hiding the user from the Global Address List (GAL)

The script prompts the operator to select a user from specified Organizational Units (OUs),
confirms the action, and logs all operations to a specified log file.

.REQUIREMENTS
- Active Directory PowerShell module
- MSOnline PowerShell module
- Exchange Online PowerShell module
- AzureAD PowerShell module
- Appropriate administrative permissions in AD and Microsoft 365

.REQUIRED MODULES
- ActiveDirectory
- MSOnline
- ExchangeOnlineManagement
- AzureAD

.PARAMETER None
This script does not take parameters. User selection is done via Out-GridView.

.NOTES
Author: <Your Name>
Date: <Date>
Version: 1.0

Ensure you update:
- <ORGANIZATIONAL UNIT PATH>
- <OUTPUT Path>

before running the script.

.LINK
https://learn.microsoft.com/powershell/

#>

<# 
    Off Boarding User Task List:
    1. Convert user's mailbox to a shared mailbox
    2. Disable the user account on-prem AD
    3. Remove the user from all groups except "Domain Users"
    4. Hide the user from GAL
#>

# Create a popup object for user interaction
$a = new-object -ComObject wscript.shell

# Initialise variables
$User_To_Disable = $null
$User_Name = $null
$User_UPN = $null
$User_Alias = $null
$ADgroups = $null

# Capture operator username (person running the script)
$operator = $env:username

# Function: Returns formatted timestamp for logging
function Get-TimeStamp {  
    return "[{0:dd/MM/yy} {0:HH:mm:ss}]" -f (Get-Date)    
}

# Function: Connect to Microsoft 365 services
function ConnectToO365
{
    # Prompt for credentials
    $cred = $host.ui.PromptForCredential("Need credentials", "Please enter your user name and password.", "", "NetBiosUserName")
    
    # Connect to required services
    Connect-MsolService
    Connect-ExchangeOnline
    Connect-AzureAD
}

# Function: Convert mailbox to shared mailbox
function ConvertToSharedMailbox($MailboxToConvert)
{  
    Set-Mailbox $MailboxToConvert -Type shared
}

# Function: Disable on-prem AD account and clear some attributes
function DisableAccount
{
    # Remove manager and telephone number
    Set-ADUser $User_Alias -Manager $null -Clear telephonenumber
    
    # Disable the account using display name lookup
    get-ADUser -filter "displayname -eq '$User_Name'" | Disable-ADAccount
}

# Function: Remove user from all AD groups except "Domain Users"
function RemoveFromGroups
{
    try
    {
        # Get all group memberships except "Domain Users"
        $ADgroups = Get-ADPrincipalGroupMembership -Identity $User_Alias -ResourceContextServer bc-gdc-dc2.rockdale.nsw.gov.au | where {$_.Name -ne "Domain Users"}

        if ($ADgroups -ne $null)
        {
            # Remove user from identified groups
            Remove-ADPrincipalGroupMembership -Identity $User_Alias -MemberOf $ADgroups  -Confirm:$false
        }

        # Output previous group memberships
        write-host "Previous Group Membership:"
        foreach ($g in $ADgroups) {
            write-host $g.name
        }
    }
    catch
    {
        Write-Host "$User_Name is not in AD, or script is not functioning properly"
    }
}

# Function: Hide user from Global Address List (GAL)
function HideUser
{
    Set-ADUser $User_Alias -replace @{msExchHideFromAddressLists="TRUE"}
}

# Function: Remove user from Azure AD cloud-only groups
function RemoveCloudGroups {
    try {
        # Get Azure AD Object ID for the user
        $AzureObject = Get-AzureADUser -SearchString $User_UPN | Select-Object -ExpandProperty ObjectID
        
        # Retrieve all Azure AD group memberships
        $AzureCloudGroups = Get-AzureADUserMembership -ObjectId $AzureObject 
        
        # Filter only cloud-only groups (not synced from on-prem)
        $CloudOnlyGroups = $AzureCloudGroups | Where-Object { -not $_.DirSyncEnabled }

        # Display cloud-only groups
        if ($CloudOnlyGroups) {
            "Previous Cloud-Only Groups:"
            $CloudOnlyGroups | Sort-Object | ForEach-Object {
                $_.DisplayName
            }
        }
        else {
            Write-Host "No cloud-only groups found for $User_UPN."
        }

        # Remove user from each cloud-only group
        $CloudOnlyGroups | ForEach-Object {
            Remove-AzureADGroupMember -ObjectId $_.ObjectId -MemberId $AzureObject
        }      
    }
    catch {
        Write-Host "$User_UPN doesn't have cloud assigned groups or the script encountered an error."
    }
}

# Prompt admin to select a user from specific OUs using Out-GridView
$User_To_Disable = '<ORGANIZATIONAL UNIT PATH>', '<ORGANIZATIONAL UNIT PATH>' | 
ForEach-Object { 
    Get-ADUser -Filter {(enabled -eq $true) -and (displayName -like "*")} `
    -searchBase $_ -Properties displayName,userPrincipalName, SamAccountName
} | 
select displayName,userPrincipalName, SamAccountName | 
Sort displayName | 
out-gridview -Title "Please Pick The User To Disable" -PassThru

# Assign selected user details
$User_Name = $User_To_Disable.displayName
$User_UPN = $User_To_Disable.userPrincipalName
$User_Alias = $User_To_Disable.SamAccountName

# Validation: Ensure only one user is selected
if($User_Name -eq $null)
{
    $a.Popup("No user selected to off board.")
    exit
}
elseif ($User_Name.count -gt 1)
{
    $a.Popup("You can only off board one user per time")
    exit
}
else
{
    # Confirmation popup
    $DialogMessage = "Are you going to off-boarding " + $User_Name +"?"
    $DialogTitle = "Confirmation"
    $PromoteType = 4
    $intAnswer = $a.popup($DialogMessage,50,$DialogTitle,$PromoteType)

    # If user confirms (Yes = 6)
    If($intAnswer -eq 6)
    {
        $(
            write-host "-----Offboarding Start-----"
            write-host "$(Get-TimeStamp) Operator: $($operator)"

            # Connect to cloud services
            ConnectToO365

            # Convert mailbox
            write-host "Converting $($User_Name) mailbox to a shared mailbox"
            ConvertToSharedMailbox -MailboxToConvert $User_UPN

            # Disconnect Exchange session (cleanup)
            write-host "Removing all O365 licenses from $($User_Name)"
            Disconnect-ExchangeOnline -Confirm:$false

            # Disable AD account
            write-host "Disabling $($User_Name) AD account"
            DisableAccount

            # Remove on-prem group memberships
            write-host "Removing $($User_Name) from all groups"
            RemoveFromGroups

            # Remove cloud group memberships
            Write-Host "Removing $($User_Name) from all cloud Groups"
            RemoveCloudGroups

            # Hide from GAL
            write-host "Hidding $($User_Name) from GAL"
            HideUser

            write-host "-----Offboarding Finish-----"
            write-host " "
        ) *>> "<OUTPUT Path>\OffboardingLog.txt"

        # Completion popup
        $a.Popup("Offboarding Process Finished. Please check OffboardingLog.txt for the details")
    }
    else
    {
        $a.popup("Operation Cancelled")
        exit
    }
}

# Cleanup variables
$User_To_Disable = $null
$User_Name = $null
$User_UPN = $null
$User_Alias = $null
$ADgroups = $null
