<#
.SYNOPSIS
    Generates a report of enabled Microsoft Entra ID Member users
    who have not signed in within the last 45 days.

.DESCRIPTION
    This script connects to Microsoft Graph and retrieves all enabled
    Member users from Microsoft Entra ID.

    It checks:
        - Interactive sign-ins
        - Non-interactive sign-ins
        - License assignment status

    Users are included in the report only if:
        - Account is enabled
        - UserType is Member
        - No interactive sign-in in the last 45 days
        - No non-interactive sign-in in the last 45 days

.PARAMETER None
    No parameters required.

.NOTES
    Required Microsoft Graph Permissions:
        - User.Read.All
        - AuditLog.Read.All

    Required Module:
        Microsoft.Graph

.EXAMPLE
    .\Get-InactiveUsers.ps1

    Runs the script and displays inactive enabled Member users.

#>

# ============================================
# Microsoft Graph - Inactive Users Report
# ============================================

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All"

# Set inactivity threshold
$thresholdDate = (Get-Date).AddDays(-45)

# Retrieve users with required properties
$users = Get-MgUser `
    -All `
    -Property `
        DisplayName,
        UserPrincipalName,
        AccountEnabled,
        UserType,
        SignInActivity,
        AssignedLicenses

# Process inactive users
$inactiveUsers = foreach ($user in $users) {

    # Only include enabled Member users
    if ($user.AccountEnabled -eq $true -and $user.UserType -eq "Member") {

        # Get sign-in activity
        $interactiveSignIn    = $user.SignInActivity.LastSignInDateTime
        $nonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime

        # Check license status
        $isLicensed = ($user.AssignedLicenses.Count -gt 0)

        # Check inactivity
        $interactiveInactive = (
            -not $interactiveSignIn -or
            $interactiveSignIn -lt $thresholdDate
        )

        $nonInteractiveInactive = (
            -not $nonInteractiveSignIn -or
            $nonInteractiveSignIn -lt $thresholdDate
        )

        # Include user if BOTH sign-ins are inactive
        if ($interactiveInactive -and $nonInteractiveInactive) {

            [PSCustomObject]@{
                DisplayName               = $user.DisplayName
                UserPrincipalName         = $user.UserPrincipalName
                UserType                  = $user.UserType
                AccountEnabled            = $user.AccountEnabled
                IsLicensed                = $isLicensed
                LastInteractiveSignIn     = $interactiveSignIn
                LastNonInteractiveSignIn  = $nonInteractiveSignIn
            }
        }
    }
}

# Display results
$inactiveUsers | Format-Table -AutoSize

# Export to CSV (Optional)
# $inactiveUsers | Export-Csv `
#     -Path "C:\Temp\Inactive_Member_Users_45Days.csv" `
#     -NoTypeInformation `
#     -Encoding UTF8
