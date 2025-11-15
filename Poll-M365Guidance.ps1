<#
.SYNOPSIS
    Polls gov.uk for changes to Microsoft 365 guidance documents.

.DESCRIPTION
    This script monitors the UK Government's Microsoft 365 guidance page for updates.
    It tracks document versions, update dates, and can notify when changes are detected.
    Results are stored in a JSON file for historical tracking.

.PARAMETER PollIntervalMinutes
    How often to poll for changes (in minutes). Default is 60.

.PARAMETER ContinuousMode
    If specified, the script runs continuously polling at the specified interval.
    Otherwise, it performs a single check.

.PARAMETER NotificationEmail
    Email address to send notifications to when changes are detected.

.PARAMETER SmtpServer
    SMTP server to use for email notifications.

.PARAMETER HistoryFile
    Path to the JSON file storing historical data. Default is ".\M365GuidanceHistory.json"

.EXAMPLE
    .\Poll-M365Guidance.ps1
    Performs a single check for changes.

.EXAMPLE
    .\Poll-M365Guidance.ps1 -ContinuousMode -PollIntervalMinutes 30
    Continuously polls every 30 minutes.

.EXAMPLE
    .\Poll-M365Guidance.ps1 -NotificationEmail "admin@example.com" -SmtpServer "smtp.example.com"
    Checks once and sends email if changes detected.
#>

[CmdletBinding()]
param(
    [int]$PollIntervalMinutes = 60,
    [switch]$ContinuousMode,
    [string]$NotificationEmail,
    [string]$SmtpServer,
    [string]$HistoryFile = ".\M365GuidanceHistory.json"
)

# Configuration
$GuidanceUrl = "https://www.gov.uk/guidance/microsoft-365-guidance-for-uk-government"
$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"

function Get-WebPageContent {
    [CmdletBinding()]
    param(
        [string]$Url
    )

    try {
        $webRequest = [System.Net.WebRequest]::Create($Url)
        $webRequest.UserAgent = $UserAgent
        $webRequest.Method = "GET"
        $webRequest.Timeout = 30000

        $response = $webRequest.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()

        $reader.Close()
        $stream.Close()
        $response.Close()

        return $content
    }
    catch {
        Write-Error "Failed to fetch webpage: $_"
        return $null
    }
}

function Parse-M365GuidancePage {
    [CmdletBinding()]
    param(
        [string]$HtmlContent
    )

    $results = @{
        CheckDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PageHash = (Get-FileHash -InputStream ([System.IO.MemoryStream]::new([System.Text.Encoding]::UTF8.GetBytes($HtmlContent)))).Hash
        Documents = @()
        LastUpdated = $null
    }

    # Extract last updated date
    if ($HtmlContent -match 'Last updated[:\s]+(\d{1,2}\s+\w+\s+\d{4})') {
        $results.LastUpdated = $matches[1]
    }

    # Extract document links (PDFs and other downloadable content)
    $pdfPattern = 'href="([^"]*\.pdf[^"]*)"[^>]*>([^<]+)<'
    $matches = [regex]::Matches($HtmlContent, $pdfPattern)

    foreach ($match in $matches) {
        $docUrl = $match.Groups[1].Value
        $docTitle = $match.Groups[2].Value.Trim()

        # Make URL absolute if relative
        if ($docUrl -notmatch '^https?://') {
            $docUrl = "https://www.gov.uk" + $docUrl
        }

        $results.Documents += [PSCustomObject]@{
            Title = $docTitle
            Url = $docUrl
            ExtractedDate = $results.CheckDate
        }
    }

    # Also look for attachment links
    $attachmentPattern = 'class="[^"]*attachment[^"]*"[^>]*href="([^"]+)"[^>]*>([^<]+)<|href="([^"]+)"[^>]*class="[^"]*attachment[^"]*"[^>]*>([^<]+)<'
    $attachMatches = [regex]::Matches($HtmlContent, $attachmentPattern)

    foreach ($match in $attachMatches) {
        $docUrl = if ($match.Groups[1].Value) { $match.Groups[1].Value } else { $match.Groups[3].Value }
        $docTitle = if ($match.Groups[2].Value) { $match.Groups[2].Value } else { $match.Groups[4].Value }

        $docTitle = $docTitle.Trim()

        # Make URL absolute if relative
        if ($docUrl -notmatch '^https?://') {
            $docUrl = "https://www.gov.uk" + $docUrl
        }

        # Check if not already added
        if (-not ($results.Documents | Where-Object { $_.Url -eq $docUrl })) {
            $results.Documents += [PSCustomObject]@{
                Title = $docTitle
                Url = $docUrl
                ExtractedDate = $results.CheckDate
            }
        }
    }

    return $results
}

function Compare-GuidanceVersions {
    [CmdletBinding()]
    param(
        [object]$Previous,
        [object]$Current
    )

    $changes = @{
        HasChanges = $false
        PageModified = $false
        NewDocuments = @()
        RemovedDocuments = @()
        ModifiedDocuments = @()
        LastUpdatedChanged = $false
    }

    if (-not $Previous) {
        $changes.HasChanges = $true
        $changes.NewDocuments = $Current.Documents
        return $changes
    }

    # Check if page hash changed
    if ($Previous.PageHash -ne $Current.PageHash) {
        $changes.PageModified = $true
        $changes.HasChanges = $true
    }

    # Check if last updated date changed
    if ($Previous.LastUpdated -ne $Current.LastUpdated) {
        $changes.LastUpdatedChanged = $true
        $changes.HasChanges = $true
    }

    # Find new documents
    foreach ($doc in $Current.Documents) {
        if (-not ($Previous.Documents | Where-Object { $_.Url -eq $doc.Url })) {
            $changes.NewDocuments += $doc
            $changes.HasChanges = $true
        }
    }

    # Find removed documents
    foreach ($doc in $Previous.Documents) {
        if (-not ($Current.Documents | Where-Object { $_.Url -eq $doc.Url })) {
            $changes.RemovedDocuments += $doc
            $changes.HasChanges = $true
        }
    }

    return $changes
}

function Send-ChangeNotification {
    [CmdletBinding()]
    param(
        [object]$Changes,
        [object]$CurrentData,
        [string]$EmailTo,
        [string]$SmtpServer
    )

    if (-not $EmailTo -or -not $SmtpServer) {
        Write-Warning "Email notification requested but email or SMTP server not configured"
        return
    }

    $subject = "Microsoft 365 Guidance Update Detected"

    $body = @"
Microsoft 365 Guidance for UK Government has been updated.

Check Date: $($CurrentData.CheckDate)
Last Updated: $($CurrentData.LastUpdated)
Page URL: $GuidanceUrl

Changes Detected:
"@

    if ($Changes.LastUpdatedChanged) {
        $body += "`n- Last updated date changed"
    }

    if ($Changes.PageModified) {
        $body += "`n- Page content modified"
    }

    if ($Changes.NewDocuments.Count -gt 0) {
        $body += "`n`nNew Documents ($($Changes.NewDocuments.Count)):"
        foreach ($doc in $Changes.NewDocuments) {
            $body += "`n  - $($doc.Title)"
            $body += "`n    $($doc.Url)"
        }
    }

    if ($Changes.RemovedDocuments.Count -gt 0) {
        $body += "`n`nRemoved Documents ($($Changes.RemovedDocuments.Count)):"
        foreach ($doc in $Changes.RemovedDocuments) {
            $body += "`n  - $($doc.Title)"
        }
    }

    try {
        Send-MailMessage -To $EmailTo `
                        -From "M365GuidanceMonitor@gov.uk" `
                        -Subject $subject `
                        -Body $body `
                        -SmtpServer $SmtpServer `
                        -ErrorAction Stop

        Write-Host "Email notification sent to $EmailTo" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to send email notification: $_"
    }
}

function Export-HistoryData {
    [CmdletBinding()]
    param(
        [object]$Data,
        [string]$FilePath
    )

    try {
        # Load existing history
        $history = @()
        if (Test-Path $FilePath) {
            $history = Get-Content $FilePath -Raw | ConvertFrom-Json
        }

        # Add current data
        $history += $Data

        # Keep last 100 entries
        if ($history.Count -gt 100) {
            $history = $history | Select-Object -Last 100
        }

        # Save
        $history | ConvertTo-Json -Depth 10 | Set-Content $FilePath
        Write-Verbose "History saved to $FilePath"
    }
    catch {
        Write-Error "Failed to save history: $_"
    }
}

function Get-PreviousCheck {
    [CmdletBinding()]
    param(
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        return $null
    }

    try {
        $history = Get-Content $FilePath -Raw | ConvertFrom-Json
        if ($history -and $history.Count -gt 0) {
            return $history[-1]
        }
    }
    catch {
        Write-Warning "Failed to load history: $_"
    }

    return $null
}

function Invoke-GuidanceCheck {
    Write-Host "Polling Microsoft 365 Guidance page..." -ForegroundColor Cyan
    Write-Host "URL: $GuidanceUrl" -ForegroundColor Gray

    # Fetch page content
    $htmlContent = Get-WebPageContent -Url $GuidanceUrl

    if (-not $htmlContent) {
        Write-Error "Failed to retrieve page content"
        return $null
    }

    # Parse page
    $currentData = Parse-M365GuidancePage -HtmlContent $htmlContent

    Write-Host "`nCurrent Status:" -ForegroundColor Yellow
    Write-Host "  Last Updated: $($currentData.LastUpdated)"
    Write-Host "  Documents Found: $($currentData.Documents.Count)"

    # Display documents
    if ($currentData.Documents.Count -gt 0) {
        Write-Host "`nDocuments:" -ForegroundColor Yellow
        foreach ($doc in $currentData.Documents) {
            Write-Host "  - $($doc.Title)" -ForegroundColor Gray
        }
    }

    # Get previous check
    $previousData = Get-PreviousCheck -FilePath $HistoryFile

    # Compare versions
    $changes = Compare-GuidanceVersions -Previous $previousData -Current $currentData

    if ($changes.HasChanges) {
        Write-Host "`n*** CHANGES DETECTED ***" -ForegroundColor Green

        if ($changes.LastUpdatedChanged) {
            Write-Host "  - Last updated date changed" -ForegroundColor Green
        }

        if ($changes.PageModified) {
            Write-Host "  - Page content modified" -ForegroundColor Green
        }

        if ($changes.NewDocuments.Count -gt 0) {
            Write-Host "`n  New Documents:" -ForegroundColor Green
            foreach ($doc in $changes.NewDocuments) {
                Write-Host "    + $($doc.Title)" -ForegroundColor Green
            }
        }

        if ($changes.RemovedDocuments.Count -gt 0) {
            Write-Host "`n  Removed Documents:" -ForegroundColor Red
            foreach ($doc in $changes.RemovedDocuments) {
                Write-Host "    - $($doc.Title)" -ForegroundColor Red
            }
        }

        # Send notification if configured
        if ($NotificationEmail -and $SmtpServer) {
            Send-ChangeNotification -Changes $changes `
                                   -CurrentData $currentData `
                                   -EmailTo $NotificationEmail `
                                   -SmtpServer $SmtpServer
        }
    }
    else {
        Write-Host "`nNo changes detected." -ForegroundColor Gray
    }

    # Save to history
    Export-HistoryData -Data $currentData -FilePath $HistoryFile

    return $currentData
}

# Main execution
try {
    if ($ContinuousMode) {
        Write-Host "Starting continuous monitoring mode" -ForegroundColor Cyan
        Write-Host "Poll interval: $PollIntervalMinutes minutes" -ForegroundColor Cyan
        Write-Host "Press Ctrl+C to stop" -ForegroundColor Yellow
        Write-Host ""

        while ($true) {
            Invoke-GuidanceCheck

            Write-Host "`nNext check in $PollIntervalMinutes minutes..." -ForegroundColor Gray
            Start-Sleep -Seconds ($PollIntervalMinutes * 60)
            Write-Host "`n$('-' * 80)" -ForegroundColor DarkGray
        }
    }
    else {
        Invoke-GuidanceCheck
    }
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
