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

# Display results in a nice table format
Write-Host "`n========== APPLICATION GROUP SUMMARY ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupID, GroupName, TotalApplications

Write-Host "`n========== TYPE BREAKDOWN BY GROUP ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupName, Type_exe, Type_msc, Type_ps1, Type_bat, Type_wsh, Type_msi, Type_dll, Type_reg

Write-Host "`n========== KEY ATTRIBUTES BY GROUP ==========" -ForegroundColor Cyan
$results | Format-Table -AutoSize -Property GroupName, ChildrenInheritToken_True, OpenDLGDropRights_True, CheckFileName_True, CheckProductName_True

Write-Host "`n========== DETAILED GROUP ANALYSIS ==========" -ForegroundColor Cyan
foreach ($result in $results) {
    Write-Host "`nGroup: $($result.GroupName)" -ForegroundColor Yellow
    Write-Host "  ID: $($result.GroupID)"
    Write-Host "  Total Applications: $($result.TotalApplications)"
    Write-Host "  EXE Count: $($result.Type_exe)"
    Write-Host "  ChildrenInheritToken=True: $($result.ChildrenInheritToken_True)"
    Write-Host "  OpenDLGDropRights=True: $($result.OpenDLGDropRights_True)"
    Write-Host "  Unique Attributes: $($result.UniqueAttributes)"
}

# Export to CSV for further analysis
$results | Export-Csv -Path "DefendPoint_Analysis.csv" -NoTypeInformation
Write-Host "`n========== Results exported to DefendPoint_Analysis.csv ==========" -ForegroundColor Green

# Create a summary object for all groups combined
$totalSummary = [PSCustomObject]@{
    TotalGroups = $results.Count
    TotalApplications = ($results | Measure-Object -Property TotalApplications -Sum).Sum
    TotalEXEs = ($results | Measure-Object -Property Type_exe -Sum).Sum
    TotalWithChildInherit = ($results | Measure-Object -Property ChildrenInheritToken_True -Sum).Sum
    TotalWithDropRights = ($results | Measure-Object -Property OpenDLGDropRights_True -Sum).Sum
}

Write-Host "`n========== OVERALL POLICY SUMMARY ==========" -ForegroundColor Magenta
$totalSummary | Format-List
