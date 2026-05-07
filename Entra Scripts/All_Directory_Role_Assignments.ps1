<#
.SYNOPSIS
Exports all Entra ID directory role assignments (users and groups).

.DESCRIPTION
Retrieves all activated directory roles in Microsoft Entra ID and lists all assigned
users and groups. Resolves object names and exports results to a CSV file.

.REQUIREMENTS
- Microsoft Graph PowerShell module
- Permissions:
    - RoleManagement.Read.Directory
    - Directory.Read.All
    - User.Read.All
    - Group.Read.All
#>

# Connect to Microsoft Graph
Connect-MgGraph -Scopes `
    "RoleManagement.Read.Directory",
    "Directory.Read.All",
    "User.Read.All",
    "Group.Read.All"

# Get all activated directory roles
$roles = Get-MgDirectoryRole

# Collect results
$results = foreach ($role in $roles) {

    # Get members of each role
    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id

    foreach ($member in $members) {

        $principal = $null
        $type = "Unknown"

        # Try resolving as User
        try {
            $principal = Get-MgUser -UserId $member.Id -ErrorAction Stop
            $type = "User"
        }
        catch {
            # Try resolving as Group
            try {
                $principal = Get-MgGroup -GroupId $member.Id -ErrorAction Stop
                $type = "Group"
            }
            catch {
                $principal = $null
                $type = "Unknown"
            }
        }

        # Output object
        [PSCustomObject]@{
            RoleName      = $role.DisplayName
            PrincipalName = $principal.DisplayName
            PrincipalUPN  = $principal.UserPrincipalName
            ObjectType    = $type
            ObjectId      = $member.Id
        }
    }
}

# Export results
$results | Export-Csv "C:\temp\directory-role-assignments.csv" -NoTypeInformation