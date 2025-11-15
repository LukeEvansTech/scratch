# Microsoft 365 Guidance Polling Scripts

Automated monitoring scripts for tracking changes to the UK Government's Microsoft 365 guidance published at https://www.gov.uk/guidance/microsoft-365-guidance-for-uk-government

## Overview

These scripts help you monitor updates to the official Microsoft 365 guidance for UK Government, including:
- Secure Configuration Blueprint
- Information Protection guidance
- External Collaboration guidance
- BYOD (Bring Your Own Device) guidance

## Scripts

### Poll-M365Guidance.ps1

Main polling script that monitors the guidance page for changes.

**Features:**
- Single check or continuous monitoring mode
- Tracks document additions, removals, and modifications
- Detects page content changes
- Maintains historical data in JSON format
- Optional email notifications
- Configurable polling intervals

**Basic Usage:**

```powershell
# Single check
.\Poll-M365Guidance.ps1

# Continuous monitoring (polls every 60 minutes by default)
.\Poll-M365Guidance.ps1 -ContinuousMode

# Custom polling interval (30 minutes)
.\Poll-M365Guidance.ps1 -ContinuousMode -PollIntervalMinutes 30

# With email notifications
.\Poll-M365Guidance.ps1 `
    -NotificationEmail "admin@example.gov.uk" `
    -SmtpServer "smtp.example.gov.uk"

# Custom history file location
.\Poll-M365Guidance.ps1 -HistoryFile "C:\Monitoring\M365History.json"
```

### Get-M365GuidanceHistory.ps1

Companion script for viewing and analyzing polling history.

**Features:**
- View all historical checks
- Display only the latest check
- Show changes between checks
- Export history to CSV for analysis

**Usage:**

```powershell
# Show latest check
.\Get-M365GuidanceHistory.ps1 -ShowLatest

# Show all historical data
.\Get-M365GuidanceHistory.ps1 -ShowAll

# Show only checks where changes were detected
.\Get-M365GuidanceHistory.ps1 -ShowChanges

# Export history to CSV
.\Get-M365GuidanceHistory.ps1 -ExportCsv

# Custom file locations
.\Get-M365GuidanceHistory.ps1 `
    -HistoryFile "C:\Monitoring\M365History.json" `
    -ShowAll
```

## Setting Up Scheduled Polling

### Using Windows Task Scheduler

1. Open Task Scheduler
2. Create a new task:
   - **Trigger**: Daily at your preferred time (or multiple times per day)
   - **Action**: Start a program
     - Program: `powershell.exe`
     - Arguments: `-ExecutionPolicy Bypass -File "C:\Path\To\Poll-M365Guidance.ps1"`
3. Configure additional settings as needed

### Using PowerShell Scheduled Job

```powershell
# Create a scheduled job that runs every 4 hours
$trigger = New-JobTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 4) -RepetitionDuration ([TimeSpan]::MaxValue)

Register-ScheduledJob -Name "M365GuidanceMonitor" `
    -FilePath "C:\Path\To\Poll-M365Guidance.ps1" `
    -Trigger $trigger
```

## Output Files

### M365GuidanceHistory.json

JSON file containing historical polling data:
- Check timestamps
- Page content hashes
- Last updated dates
- Document lists with URLs
- Maintains last 100 checks

### M365GuidanceHistory.csv (optional)

CSV export of history data for analysis in Excel or other tools.

## What Gets Tracked

The scripts monitor:

1. **Page Updates**: Detects any changes to the guidance page content
2. **Last Updated Date**: Tracks when gov.uk last updated the guidance
3. **Documents**: Monitors all linked documents including:
   - PDF guidance documents
   - Supporting materials
   - Blueprint documents
4. **Document Changes**: Detects when documents are added, removed, or modified

## Example Output

```
Polling Microsoft 365 Guidance page...
URL: https://www.gov.uk/guidance/microsoft-365-guidance-for-uk-government

Current Status:
  Last Updated: 15 November 2024
  Documents Found: 8

Documents:
  - Microsoft 365 Secure Configuration Blueprint
  - Information Protection Guidance
  - External Collaboration Guidance
  - BYOD Guidance

*** CHANGES DETECTED ***
  - Last updated date changed

  New Documents:
    + Microsoft 365 Security Update Q4 2024
```

## Email Notifications

When configured with `-NotificationEmail` and `-SmtpServer`, the script sends detailed email notifications including:
- What changed (page content, documents, dates)
- List of new documents with URLs
- List of removed documents
- Timestamp of detection

## Best Practices

1. **Polling Frequency**:
   - For production: Every 6-12 hours is reasonable
   - For critical monitoring: Every 1-4 hours
   - Avoid polling more frequently than every 30 minutes

2. **History Management**:
   - Script automatically maintains last 100 checks
   - Archive history file periodically if needed for long-term records

3. **Error Handling**:
   - Script includes retry logic and error handling
   - Check history file if polling seems to have stopped

4. **Security**:
   - Store scripts in a secure location
   - Limit access to history files (may contain sensitive info)
   - Use secure SMTP for email notifications

## Troubleshooting

### "Failed to fetch webpage" Error

- Check internet connectivity
- Verify the URL is accessible from your network
- Some networks may block automated requests - consider using a proxy

### No Documents Found

- The page structure may have changed
- Script may need updating to match new HTML structure
- Check manually that documents are visible on the page

### Email Notifications Not Sending

- Verify SMTP server is accessible
- Check firewall rules
- Ensure SMTP authentication is configured if required
- Test with a simple `Send-MailMessage` command first

## Requirements

- PowerShell 5.1 or later
- Internet access to gov.uk
- For email notifications: Access to an SMTP server

## Future Enhancements

Potential improvements:
- Download and hash-check actual PDF documents
- Microsoft Teams/Slack webhook notifications
- Integration with monitoring systems (Prometheus, etc.)
- Differential analysis of PDF content changes
- Support for monitoring multiple guidance pages

## License

These scripts are provided as-is for monitoring UK Government guidance publications.
