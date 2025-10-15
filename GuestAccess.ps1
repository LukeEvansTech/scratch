# Connect to Exchange Online (works with Global Reader / Security Admin)
Connect-ExchangeOnline

# Get all Microsoft 365 Groups that have guest members
$groups = Get-UnifiedGroup -ResultSize Unlimited

$report = @()
foreach ($g in $groups) {
    $members = Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Members
    $guests = $members | Where-Object {$_.RecipientTypeDetails -eq "GuestMailUser"}
    if ($guests) {
        $report += [PSCustomObject]@{
            GroupName   = $g.DisplayName
            GroupEmail  = $g.PrimarySmtpAddress
            TeamEnabled = $g.ResourceProvisioningOptions -contains "Team"
            GuestCount  = $guests.Count
            GuestEmails = ($guests.PrimarySmtpAddress -join "; ")
        }
    }
}

# Output all M365 Groups that have guests
$report | Where-Object {$_.TeamEnabled -eq $true} | Format-Table -AutoSize

# Optional: Export to CSV for auditing
$report | Where-Object {$_.TeamEnabled -eq $true} | Export-Csv ".\TeamsWithGuests.csv" -NoTypeInformation
