# Path to your XML file
$xmlPath = "C:\Path\To\Your\File.xml"

# Load the XML
[xml]$xml = Get-Content -Path $xmlPath

# Find all <Section> nodes with 'Exclusion' in their name
$exclusionSections = $xml.SelectNodes("//Section[contains(translate(@name, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'exclusion')]")

# Prepare results
$results = foreach ($section in $exclusionSections) {
    $sectionName = $section.name
    foreach ($setting in $section.Setting) {
        if ($setting.name -like "ExcludedItem_*") {
            # Value format: "3|3|**\path\file.exe|"
            $rawValue = $setting.value
            # Extract the final pipe-delimited part that contains the path
            $parts = $rawValue -split "\|"
            $path = ($parts | Where-Object { $_ -match "[\\/]" })[-1]
            [PSCustomObject]@{
                Section = $sectionName
                Setting = $setting.name
                PathOrExecutable = $path
            }
        }
    }
}

# Display the table
$results | Format-Table -AutoSize

# Optional: export to CSV
$results | Export-Csv -Path ".\Exclusions.csv" -NoTypeInformation
