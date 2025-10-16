param(
    [Parameter(Mandatory)]
    [string]$XmlPath,

    [switch]$Unique,          # de-dupe paths
    [string]$CsvOut           # optional: write to CSV
)

function Get-ExcludedPathFromValue([string]$val) {
    # Values look like:  "3|3|**\path\thing.exe|"
    if (-not $val) { return $null }
    $parts = $val -split '\|'
    # take the last non-empty segment
    ($parts | Where-Object { $_ -ne '' })[-1]
}

$results = @()

# ---------- Attempt 1: strict XML ----------
try {
    [xml]$xml = Get-Content -Path $XmlPath -Raw
    # find any <Section> whose name contains "exclusion"
    $exSections = $xml.SelectNodes("//Section[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'exclusion')]")

    foreach ($section in $exSections) {
        foreach ($setting in @($section.Setting)) {
            if ($setting.name -like 'ExcludedItem_*') {
                $path = Get-ExcludedPathFromValue $setting.value
                if ($path) {
                    $results += [pscustomobject]@{
                        Section = $section.name
                        Setting = $setting.name
                        PathOrExecutable = $path
                    }
                }
            }
        }
    }
}
catch {
    Write-Warning "XML load failed (`$($_.Exception.Message)`). Falling back to text parse."

    # ---------- Attempt 2: tolerant text/regex scrape ----------
    $text = Get-Content -Path $XmlPath -Raw

    # Narrow to sections whose name contains "Exclusions" (case-insensitive)
    $sectionPattern = '<Section\s+name="([^"]*Exclusions[^"]*)"\s*>(.*?)</Section>'
    $settingPattern = '<Setting\s+name="(ExcludedItem_\d+)"\s+value="([^"]+)"\s*/?>'

    $secMatches = [regex]::Matches($text, $sectionPattern, 'Singleline, IgnoreCase')
    foreach ($sec in $secMatches) {
        $sectionName = $sec.Groups[1].Value
        $body        = $sec.Groups[2].Value
        $setMatches  = [regex]::Matches($body, $settingPattern, 'Singleline, IgnoreCase')

        foreach ($m in $setMatches) {
            $settingName = $m.Groups[1].Value
            $rawValue    = $m.Groups[2].Value
            $path        = Get-ExcludedPathFromValue $rawValue
            if ($path) {
                $results += [pscustomobject]@{
                    Section = $sectionName
                    Setting = $settingName
                    PathOrExecutable = $path
                }
            }
        }
    }
}

if ($Unique) {
    $results = $results | Sort-Object PathOrExecutable -Unique
}

# Show a nice table
$results | Format-Table -AutoSize

if ($CsvOut) {
    $results | Export-Csv -Path $CsvOut -NoTypeInformation -Encoding UTF8
    Write-Host "CSV written to: $CsvOut"
}