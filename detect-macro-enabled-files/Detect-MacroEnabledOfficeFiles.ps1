# ===============================
# Detect Macro-Enabled Office Files
# ===============================

# Define macro-enabled Office file extensions to search for
$Extensions = @("*.docm","*.xlsm","*.pptm","*.dotm","*.xltm")

# Define target directories to scan
$Paths = @(
    "Enter the path in which you want to scan files"
)

# Initialize results array to store detected files
$Results = @()

# Loop through each defined path
foreach ($Path in $Paths) {

    # Check if the path exists before scanning
    if (Test-Path $Path) {

        # Loop through each file extension
        foreach ($Ext in $Extensions) {
            try {
                # Recursively search for files matching the extension
                # SilentlyContinue avoids interruption due to access issues
                $Files = Get-ChildItem -Path $Path -Recurse -Include $Ext -ErrorAction SilentlyContinue

                # Process each file found
                foreach ($File in $Files) {

                    # Create a custom object with relevant file details
                    $Results += [PSCustomObject]@{
                        FileName      = $File.Name
                        FullPath      = $File.FullName
                        SizeKB        = [math]::Round($File.Length / 1KB, 2)
                        LastModified  = $File.LastWriteTime
                    }
                }
            } catch {
                # Ignore errors such as access denied or locked directories
                # This ensures the script continues scanning other locations
            }
        }
    }
}

# ===============================
# Output Results
# ===============================

if ($Results.Count -gt 0) {

    # If macro-enabled files are found:
    Write-Output "Macro-enabled Office files detected."

    # Export results to CSV for reporting/auditing
    # -Append ensures results are added if the file already exists
    $Results | Export-Csv -Path "C:\Users\<UserName>\Documents\details.csv" -NoTypeInformation -Append

    # Exit with code 1 to indicate non-compliance (useful for Intune detection scripts)
    exit 1

} else {

    # If no macro-enabled files are found:
    Write-Output "No macro-enabled Office files found."

    # Exit with code 0 to indicate compliance
    exit 0
}
