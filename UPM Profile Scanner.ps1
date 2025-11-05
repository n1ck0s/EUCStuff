### Script Written by Nick Pylarinos ###
### This script will cycle through a profile store and determine the last time NTUSER.Dat was written ###
### The information extracted from the CSV file can be used to purge stale profiles on a profile store ###



# Data Array - use this to reference profile folder names
$fileListPath = "C:\Profile\output_file.txt"

# Profile Store location
$baseFolderPath = "E:\ProfileStore"

# NTUSER.DAT Variable
$targetFile = "NTUSER.dat"

# Output CSV file path
$outputCsvPath = "C:\output.csv"

# Ensure the directory for the output CSV file exists
$outputDirectory = Split-Path -Path $outputCsvPath -Parent
if (-not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory | Out-Null
}

# Read the list of folder names from the file
$folderList = Get-Content -Path $fileListPath

# Function to check the last written time of the target file in UPM_Profile subfolder
function Check-LastWrittenTime {
    param (
        [string]$folderName
    )
    try {
        $filePath = Join-Path -Path (Join-Path -Path $baseFolderPath -ChildPath $folderName) -ChildPath "UPM_Profile\$targetFile"
        if (Test-Path -Path $filePath) {
            $lastWrittenTime = (Get-Item -Path $filePath).LastWriteTime
            return [PSCustomObject]@{
                FolderName = $folderName
                LastWrittenTime = $lastWrittenTime
            }
        } else {
            return [PSCustomObject]@{
                FolderName = $folderName
                LastWrittenTime = "File does not exist"
            }
        }
    } catch {
        return [PSCustomObject]@{
            FolderName = $folderName
            LastWrittenTime = "Error: $_"
        }
    }
}

# List to store results
$results = @()

# Check the last written time for NTUSER.dat in each folder's UPM_Profile subfolder
foreach ($folderName in $folderList) {
    $result = Check-LastWrittenTime -folderName $folderName
    $results += $result
}

# Export results to a CSV file
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

Write-Output "Results have been written to $outputCsvPath"