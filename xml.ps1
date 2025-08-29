# Load the XML file
[xml]$xml = Get-Content -Path "YourPolicyFile.xml"  # Replace with your actual file path

# Create array to store results
$results = @()

# Find all ApplicationGroup elements
$applicationGroups = $xml.SelectNodes("//ApplicationGroup")

foreach ($group in $applicationGroups) {
    # Initialize counters
    $groupStats = [PSCustomObject]@{
        GroupID = $group.ID
        GroupName = $group.Name
        ParentGroup = $group.Parent
        TotalApplications = 0
        
        # Type counts
        Type_exe = 0
        Type_msc = 0
        Type_ps1 = 0
        Type_bat = 0
        Type_wsh = 0
        Type_msi = 0
        Type_dll = 0
        Type_reg = 0
        Type_other = 0
        
        # Attribute counts
        ChildrenInheritToken_True = 0
        ChildrenInheritToken_False = 0
        OpenDLGDropRights_True = 0
        OpenDLGDropRights_False = 0
        CheckFileName_True = 0
        CheckProductName_True = 0
        UseSourceFileName_True = 0
        
        # Unique attributes found
        UniqueAttributes = @()
    }
    
    # Get all Application elements within this group
    $applications = $group.SelectNodes("Application")
    $groupStats.TotalApplications = $applications.Count
    
    # Track unique attributes across all applications
    $allAttributes = @{}
    
    foreach ($app in $applications) {
        # Count types
        $type = $app.Type
        switch ($type) {
            "exe" { $groupStats.Type_exe++ }
            "msc" { $groupStats.Type_msc++ }
            "ps1" { $groupStats.Type_ps1++ }
            "bat" { $groupStats.Type_bat++ }
            "wsh" { $groupStats.Type_wsh++ }
            "msi" { $groupStats.Type_msi++ }
            "dll" { $groupStats.Type_dll++ }
            "reg" { $groupStats.Type_reg++ }
            default { $groupStats.Type_other++ }
        }
        
        # Count specific attributes
        if ($app.ChildrenInheritToken -eq "true") {
            $groupStats.ChildrenInheritToken_True++
        } elseif ($app.ChildrenInheritToken -eq "false") {
            $groupStats.ChildrenInheritToken_False++
        }
        
        if ($app.OpenDLGDropRights -eq "true") {
            $groupStats.OpenDLGDropRights_True++
        } elseif ($app.OpenDLGDropRights -eq "false") {
            $groupStats.OpenDLGDropRights_False++
        }
        
        if ($app.CheckFileName -eq "true") {
            $groupStats.CheckFileName_True++
        }
        
        if ($app.CheckProductName -eq "true") {
            $groupStats.CheckProductName_True++
        }
        
        if ($app.UseSourceFileName -eq "true") {
            $groupStats.UseSourceFileName_True++
        }
        
        # Collect all unique attributes for this application
        foreach ($attr in $app.Attributes) {
            $allAttributes[$attr.Name] = $true
        }
    }
    
    # Store unique attributes found
    $groupStats.UniqueAttributes = $allAttributes.Keys -join ", "
    
    $results += $groupStats
}

# Display and Export: APPLICATION GROUP SUMMARY
Write-Host "`n========== APPLICATION GROUP SUMMARY ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupID, GroupName, TotalApplications

$groupSummary = $results | Select-Object GroupID, GroupName, TotalApplications
$groupSummary | Export-Csv -Path "DefendPoint_GroupSummary.csv" -NoTypeInformation
Write-Host "Exported to DefendPoint_GroupSummary.csv" -ForegroundColor Green

# Display and Export: TYPE BREAKDOWN BY GROUP
Write-Host "`n========== TYPE BREAKDOWN BY GROUP ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupName, Type_exe, Type_msc, Type_ps1, Type_bat, Type_wsh, Type_msi, Type_dll, Type_reg

$typeBreakdown = $results | Select-Object GroupName, Type_exe, Type_msc, Type_ps1, Type_bat, Type_wsh, Type_msi, Type_dll, Type_reg, Type_other
$typeBreakdown | Export-Csv -Path "DefendPoint_TypeBreakdown.csv" -NoTypeInformation
Write-Host "Exported to DefendPoint_TypeBreakdown.csv" -ForegroundColor Green

# Display and Export: KEY ATTRIBUTES BY GROUP
Write-Host "`n========== KEY ATTRIBUTES BY GROUP ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupName, ChildrenInheritToken_True, OpenDLGDropRights_True, CheckFileName_True, CheckProductName_True

$keyAttributes = $results | Select-Object GroupName, ChildrenInheritToken_True, ChildrenInheritToken_False, OpenDLGDropRights_True, OpenDLGDropRights_False, CheckFileName_True, CheckProductName_True, UseSourceFileName_True
$keyAttributes | Export-Csv -Path "DefendPoint_KeyAttributes.csv" -NoTypeInformation
Write-Host "Exported to DefendPoint_KeyAttributes.csv" -ForegroundColor Green

# Display and Export: DETAILED GROUP ANALYSIS
Write-Host "`n========== DETAILED GROUP ANALYSIS ==========" -ForegroundColor Cyan
$detailedAnalysis = @()
foreach ($result in $results) {
    Write-Host "`nGroup: $($result.GroupName)" -ForegroundColor Yellow
    Write-Host "  ID: $($result.GroupID)"
    Write-Host "  Total Applications: $($result.TotalApplications)"
    Write-Host "  EXE Count: $($result.Type_exe)"
    Write-Host "  ChildrenInheritToken=True: $($result.ChildrenInheritToken_True)"
    Write-Host "  OpenDLGDropRights=True: $($result.OpenDLGDropRights_True)"
    Write-Host "  Unique Attributes: $($result.UniqueAttributes)"
    
    $detailedAnalysis += [PSCustomObject]@{
        GroupName = $result.GroupName
        GroupID = $result.GroupID
        TotalApplications = $result.TotalApplications
        EXE_Count = $result.Type_exe
        ChildrenInheritToken_True = $result.ChildrenInheritToken_True
        OpenDLGDropRights_True = $result.OpenDLGDropRights_True
        UniqueAttributes = $result.UniqueAttributes
    }
}
$detailedAnalysis | Export-Csv -Path "DefendPoint_DetailedAnalysis.csv" -NoTypeInformation
Write-Host "`nExported to DefendPoint_DetailedAnalysis.csv" -ForegroundColor Green

# Export COMPLETE data (all fields)
$results | Export-Csv -Path "DefendPoint_Complete.csv" -NoTypeInformation
Write-Host "Exported to DefendPoint_Complete.csv (all data)" -ForegroundColor Green

# Create and Export: OVERALL POLICY SUMMARY
$totalSummary = [PSCustomObject]@{
    TotalGroups = $results.Count
    TotalApplications = ($results | Measure-Object -Property TotalApplications -Sum).Sum
    TotalEXEs = ($results | Measure-Object -Property Type_exe -Sum).Sum
    TotalWithChildInherit = ($results | Measure-Object -Property ChildrenInheritToken_True -Sum).Sum
    TotalWithDropRights = ($results | Measure-Object -Property OpenDLGDropRights_True -Sum).Sum
    TotalMSCs = ($results | Measure-Object -Property Type_msc -Sum).Sum
    TotalPS1s = ($results | Measure-Object -Property Type_ps1 -Sum).Sum
    TotalBATs = ($results | Measure-Object -Property Type_bat -Sum).Sum
    TotalWSHs = ($results | Measure-Object -Property Type_wsh -Sum).Sum
    TotalMSIs = ($results | Measure-Object -Property Type_msi -Sum).Sum
}

Write-Host "`n========== OVERALL POLICY SUMMARY ==========" -ForegroundColor Magenta
$totalSummary | Format-List

# Export overall summary
$totalSummary | Export-Csv -Path "DefendPoint_OverallSummary.csv" -NoTypeInformation
Write-Host "Exported to DefendPoint_OverallSummary.csv" -ForegroundColor Green

# Summary of all exports
Write-Host "`n========== ALL EXPORTS COMPLETED ==========" -ForegroundColor Cyan
Write-Host "The following CSV files have been created:" -ForegroundColor Yellow
Write-Host "  1. DefendPoint_GroupSummary.csv     - Basic group information"
Write-Host "  2. DefendPoint_TypeBreakdown.csv    - Application types per group"
Write-Host "  3. DefendPoint_KeyAttributes.csv    - Security attributes per group"
Write-Host "  4. DefendPoint_DetailedAnalysis.csv - Focused analysis per group"
Write-Host "  5. DefendPoint_Complete.csv         - Complete data (all fields)"
Write-Host "  6. DefendPoint_OverallSummary.csv   - Policy-wide statistics"
