# Load the XML file
[xml]$xml = Get-Content -Path "YourPolicyFile.xml"  # Replace with your actual file path

# Add fuzzy matching capability
Add-Type -TypeDefinition @"
using System;
public class FuzzyMatcher
{
    public static int LevenshteinDistance(string s, string t)
    {
        int n = s.Length;
        int m = t.Length;
        int[,] d = new int[n + 1, m + 1];
        
        if (n == 0) return m;
        if (m == 0) return n;
        
        for (int i = 0; i <= n; i++)
            d[i, 0] = i;
        for (int j = 0; j <= m; j++)
            d[0, j] = j;
            
        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost);
            }
        }
        return d[n, m];
    }
    
    public static double Similarity(string s1, string s2)
    {
        if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2))
            return 0;
        
        int distance = LevenshteinDistance(s1.ToLower(), s2.ToLower());
        int maxLength = Math.Max(s1.Length, s2.Length);
        return 1.0 - (double)distance / maxLength;
    }
}
"@

# Function to extract exe information from text
function Extract-ExeInfo {
    param([string]$text)
    
    $exeInfo = @{
        FullPath = ""
        FileName = ""
        Directory = ""
        IsFullPath = $false
    }
    
    # Check if it's a full path
    if ($text -match '(.*\\)?([^\\]+\.exe)' -or $text -match '(.*\\)?([^\\]+)$') {
        $exeInfo.Directory = $matches[1]
        $exeInfo.FileName = $matches[2]
        $exeInfo.FullPath = $text
        $exeInfo.IsFullPath = ($text -match '\\' -or $text -match '^[A-Z]:')
    }
    else {
        $exeInfo.FileName = $text
        $exeInfo.FullPath = $text
    }
    
    # Clean up the filename (remove .exe extension for comparison)
    $exeInfo.CleanName = $exeInfo.FileName -replace '\.exe$', '' -replace '\.msc$', '' -replace '\*', ''
    
    return $exeInfo
}

# Get timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create output directory
$outputDir = "DefendPoint_ExeAnalysis_$timestamp"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Write-Host "`n========== Created output directory: $outputDir ==========" -ForegroundColor Green

# Collect all executables
Write-Host "`n========== Extracting all executables from policy ==========" -ForegroundColor Cyan

$allExecutables = @()
$applicationGroups = $xml.SelectNodes("//ApplicationGroup")

foreach ($group in $applicationGroups) {
    $applications = $group.SelectNodes("Application")
    
    foreach ($app in $applications) {
        # Only process exe, msc types (executable files)
        if ($app.Type -in @('exe', 'msc')) {
            
            # Extract from various attributes
            $searchAttributes = @{
                'Description' = $app.Description
                'FileName' = $app.FileName
                'ProductName' = $app.ProductName
                'FilePath' = $app.FilePath
                'CommandLine' = $app.CommandLine
            }
            
            foreach ($attr in $searchAttributes.GetEnumerator()) {
                if ($attr.Value) {
                    $exeInfo = Extract-ExeInfo -text $attr.Value
                    
                    $exeRecord = [PSCustomObject]@{
                        GroupName = $group.Name
                        GroupID = $group.ID
                        ApplicationID = $app.ID
                        ApplicationType = $app.Type
                        Description = $app.Description
                        SourceAttribute = $attr.Key
                        OriginalValue = $attr.Value
                        ExtractedFileName = $exeInfo.FileName
                        CleanName = $exeInfo.CleanName
                        Directory = $exeInfo.Directory
                        FullPath = $exeInfo.FullPath
                        IsFullPath = $exeInfo.IsFullPath
                        ChildrenInheritToken = $app.ChildrenInheritToken
                        OpenDLGDropRights = $app.OpenDLGDropRights
                        CheckFileName = $app.CheckFileName
                        CheckProductName = $app.CheckProductName
                    }
                    
                    $allExecutables += $exeRecord
                }
            }
        }
    }
}

Write-Host "Found $($allExecutables.Count) total executable references" -ForegroundColor Green

# Export raw list
$rawFile = Join-Path $outputDir "RAW_AllExecutables.csv"
$allExecutables | Export-Csv -Path $rawFile -NoTypeInformation
Write-Host "Exported raw list to: $rawFile" -ForegroundColor Green

# Perform fuzzy matching and grouping
Write-Host "`n========== Performing fuzzy matching and deduplication ==========" -ForegroundColor Cyan

# Group by exact CleanName first
$exactGroups = $allExecutables | Group-Object CleanName | Sort-Object Count -Descending

# Create similarity groups
$similarityThreshold = 0.75  # 75% similarity threshold
$processedNames = @{}
$similarityGroups = @()

foreach ($group in $exactGroups) {
    $cleanName = $group.Name
    
    if (-not $processedNames.ContainsKey($cleanName)) {
        # Find all similar names
        $similarNames = @($cleanName)
        $processedNames[$cleanName] = $true
        
        foreach ($otherGroup in $exactGroups) {
            if ($otherGroup.Name -ne $cleanName -and -not $processedNames.ContainsKey($otherGroup.Name)) {
                $similarity = [FuzzyMatcher]::Similarity($cleanName, $otherGroup.Name)
                if ($similarity -ge $similarityThreshold) {
                    $similarNames += $otherGroup.Name
                    $processedNames[$otherGroup.Name] = $true
                }
            }
        }
        
        # Create similarity group
        $groupMembers = $allExecutables | Where-Object { $_.CleanName -in $similarNames }
        
        $similarityGroups += [PSCustomObject]@{
            PrimaryName = $cleanName
            SimilarNames = ($similarNames -join "; ")
            TotalReferences = $groupMembers.Count
            UniqueFileNames = ($groupMembers | Select-Object -ExpandProperty ExtractedFileName -Unique).Count
            GroupCount = ($groupMembers | Select-Object -ExpandProperty GroupName -Unique).Count
            Groups = (($groupMembers | Select-Object -ExpandProperty GroupName -Unique) -join "; ")
            MostCommonPath = ($groupMembers | Group-Object FullPath | Sort-Object Count -Descending | Select-Object -First 1).Name
            AllPaths = (($groupMembers | Select-Object -ExpandProperty FullPath -Unique) -join "; ")
            HasChildInherit = ($groupMembers | Where-Object { $_.ChildrenInheritToken -eq "true" }).Count
            HasDropRights = ($groupMembers | Where-Object { $_.OpenDLGDropRights -eq "true" }).Count
            Members = $groupMembers
        }
    }
}

# Sort by importance (total references)
$similarityGroups = $similarityGroups | Sort-Object TotalReferences -Descending

# Export deduplicated summary
$deduplicatedSummary = $similarityGroups | Select-Object PrimaryName, SimilarNames, TotalReferences, UniqueFileNames, GroupCount, Groups, MostCommonPath, HasChildInherit, HasDropRights
$dedupFile = Join-Path $outputDir "DEDUPLICATED_ExecutableSummary.csv"
$deduplicatedSummary | Export-Csv -Path $dedupFile -NoTypeInformation
Write-Host "Exported deduplicated summary to: $dedupFile" -ForegroundColor Green

# Identify truly important executables (top tier)
Write-Host "`n========== Identifying Critical Executables ==========" -ForegroundColor Cyan

$criticalExecutables = @()

foreach ($simGroup in $similarityGroups) {
    $importance = 0
    $importanceFactors = @()
    
    # Factor 1: Frequency (appears multiple times)
    if ($simGroup.TotalReferences -ge 5) {
        $importance += 3
        $importanceFactors += "High frequency ($($simGroup.TotalReferences) refs)"
    } elseif ($simGroup.TotalReferences -ge 2) {
        $importance += 1
        $importanceFactors += "Medium frequency ($($simGroup.TotalReferences) refs)"
    }
    
    # Factor 2: Cross-group presence
    if ($simGroup.GroupCount -ge 3) {
        $importance += 3
        $importanceFactors += "Multiple groups ($($simGroup.GroupCount) groups)"
    } elseif ($simGroup.GroupCount -ge 2) {
        $importance += 2
        $importanceFactors += "Cross-group ($($simGroup.GroupCount) groups)"
    }
    
    # Factor 3: Security controls applied
    if ($simGroup.HasDropRights -gt 0 -and $simGroup.HasChildInherit -gt 0) {
        $importance += 2
        $importanceFactors += "Multiple security controls"
    }
    
    # Factor 4: System/Admin tools (pattern matching)
    $systemPatterns = @('powershell', 'cmd', 'mmc', 'regedit', 'wscript', 'cscript', 'msiexec', 'installer', 'setup', 'install', 'admin', 'system', 'config')
    foreach ($pattern in $systemPatterns) {
        if ($simGroup.PrimaryName -like "*$pattern*") {
            $importance += 2
            $importanceFactors += "System/Admin tool"
            break
        }
    }
    
    # Factor 5: Browser or common productivity app
    $commonApps = @('chrome', 'firefox', 'edge', 'iexplore', 'outlook', 'excel', 'word', 'powerpoint', 'teams', 'acrobat', 'reader')
    foreach ($pattern in $commonApps) {
        if ($simGroup.PrimaryName -like "*$pattern*") {
            $importance += 2
            $importanceFactors += "Common productivity app"
            break
        }
    }
    
    $criticalExecutables += [PSCustomObject]@{
        ExecutableName = $simGroup.PrimaryName
        ImportanceScore = $importance
        ImportanceFactors = ($importanceFactors -join "; ")
        TotalReferences = $simGroup.TotalReferences
        GroupCount = $simGroup.GroupCount
        Groups = $simGroup.Groups
        SecurityControls = "DropRights: $($simGroup.HasDropRights), ChildInherit: $($simGroup.HasChildInherit)"
        MostCommonPath = $simGroup.MostCommonPath
        Category = if ($importance -ge 5) { "CRITICAL" } elseif ($importance -ge 3) { "HIGH" } elseif ($importance -ge 1) { "MEDIUM" } else { "LOW" }
    }
}

# Sort by importance score
$criticalExecutables = $criticalExecutables | Sort-Object ImportanceScore -Descending, TotalReferences -Descending

# Export critical executables
$criticalFile = Join-Path $outputDir "CRITICAL_ImportantExecutables.csv"
$criticalExecutables | Export-Csv -Path $criticalFile -NoTypeInformation
Write-Host "Exported critical executables to: $criticalFile" -ForegroundColor Green

# Create executive summary
$executiveSummary = @()

# Top 10 most referenced
$executiveSummary += [PSCustomObject]@{
    Category = "TOP 10 MOST REFERENCED"
    Executables = (($criticalExecutables | Select-Object -First 10 | ForEach-Object { "$($_.ExecutableName) ($($_.TotalReferences) refs)" }) -join "; ")
}

# Critical importance executables
$criticalOnly = $criticalExecutables | Where-Object { $_.Category -eq "CRITICAL" }
$executiveSummary += [PSCustomObject]@{
    Category = "CRITICAL IMPORTANCE"
    Executables = (($criticalOnly | ForEach-Object { $_.ExecutableName }) -join "; ")
}

# System/Admin tools
$systemTools = $criticalExecutables | Where-Object { $_.ImportanceFactors -like "*System/Admin tool*" }
$executiveSummary += [PSCustomObject]@{
    Category = "SYSTEM/ADMIN TOOLS"
    Executables = (($systemTools | ForEach-Object { $_.ExecutableName }) -join "; ")
}

# Productivity apps
$productivityApps = $criticalExecutables | Where-Object { $_.ImportanceFactors -like "*Common productivity app*" }
$executiveSummary += [PSCustomObject]@{
    Category = "PRODUCTIVITY APPS"
    Executables = (($productivityApps | ForEach-Object { $_.ExecutableName }) -join "; ")
}

$execFile = Join-Path $outputDir "EXECUTIVE_Summary.csv"
$executiveSummary | Export-Csv -Path $execFile -NoTypeInformation
Write-Host "Exported executive summary to: $execFile" -ForegroundColor Green

# Display summary statistics
Write-Host "`n========== ANALYSIS COMPLETE ==========" -ForegroundColor Cyan
Write-Host "Total executable references: $($allExecutables.Count)" -ForegroundColor Green
Write-Host "Unique executable names (after dedup): $($similarityGroups.Count)" -ForegroundColor Green
Write-Host "Critical importance executables: $($criticalOnly.Count)" -ForegroundColor Red
Write-Host "High importance executables: $(($criticalExecutables | Where-Object { $_.Category -eq 'HIGH' }).Count)" -ForegroundColor Yellow
Write-Host "`nTop 5 Most Important Executables:" -ForegroundColor Cyan
$criticalExecutables | Select-Object -First 5 | ForEach-Object {
    Write-Host "  - $($_.ExecutableName) (Score: $($_.ImportanceScore), Refs: $($_.TotalReferences))" -ForegroundColor White
}

Write-Host "`nAll results saved to: $outputDir" -ForegroundColor Cyan
