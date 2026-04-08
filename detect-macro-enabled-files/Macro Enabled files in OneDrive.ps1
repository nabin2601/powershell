#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Files

<#
.SYNOPSIS
Scans all OneDrive accounts for macro-enabled file types across the tenant.

.DESCRIPTION
This script connects to Microsoft Graph, enumerates all users,
and recursively scans each user's OneDrive for files with macro-enabled extensions.
Results are exported to a CSV file for reporting and analysis.
#>

# List of macro-enabled file extensions to search for
$Extensions = @("docm","xlsm","pptm","dotm","xltm")

# Initialize a list to store results
$Results = [System.Collections.Generic.List[object]]::new()

# Retrieve all users with required properties
$Users = Get-MgUser -All -Property Id,UserPrincipalName

# Track total users for progress reporting
$TotalUsers = $Users.Count
$UserCount = 0

# Recursive function to scan OneDrive folders
function Scan-OneDriveFolder {
    param (
        [string]$UserId,     # User ID (GUID)
        [string]$UserUPN,    # User Principal Name (email)
        [string]$ItemId,     # Current folder/item ID
        [string]$Path        # Current folder path
    )

    # Get all child items (files and folders) in the current directory
    $Items = Get-MgUserDriveItemChild -UserId $UserId -ItemId $ItemId -All

    foreach ($Item in $Items) {
        # Build full file/folder path
        $FullPath = "$Path/$($Item.Name)"

        # Check if the item is a file
        if ($Item.File) {
            # Extract file extension
            $Ext = $Item.Name.Split('.')[-1].ToLower()

            # Check if file matches macro-enabled extensions
            if ($Extensions -contains $Ext) {
                # Add matching file details to results list
                $Results.Add([PSCustomObject]@{
                    User          = $UserUPN
                    FileName      = $Item.Name
                    FullPath      = $FullPath
                    SizeKB        = [math]::Round($Item.Size / 1KB, 2)
                    LastModified  = $Item.LastModifiedDateTime
                })
            }
        }

        # If the item is a folder, recursively scan its contents
        if ($Item.Folder) {
            Scan-OneDriveFolder `
                -UserId $UserId `
                -UserUPN $UserUPN `
                -ItemId $Item.Id `
                -Path $FullPath
        }
    }
}

# Iterate through each user
foreach ($User in $Users) {
    $UserCount++

    # Display progress bar for tracking execution
    Write-Progress `
        -Activity "Scanning OneDrive (Online)" `
        -Status "User $UserCount of $TotalUsers : $($User.UserPrincipalName)" `
        -PercentComplete (($UserCount / $TotalUsers) * 100)

    try {
        # Attempt to retrieve the user's OneDrive
        $Drive = Get-MgUserDrive -UserId $User.Id -ErrorAction Stop

        # Start scanning from the root directory of OneDrive
        Scan-OneDriveFolder `
            -UserId $User.Id `
            -UserUPN $User.UserPrincipalName `
            -ItemId $Drive.Root.Id `
            -Path ""
    }
    catch {
        # Skip users without a provisioned OneDrive or inaccessible drives
        continue
    }
}

# Mark progress as complete
Write-Progress -Activity "Scanning OneDrive (Online)" -Completed

# Export results to CSV file
$Results | Export-Csv "C:\Temp\AllUsers_OneDrive_MacroFiles.csv" -NoTypeInformation
