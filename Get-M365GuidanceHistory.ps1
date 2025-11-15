<#
.SYNOPSIS
    Views and analyzes the Microsoft 365 guidance polling history.

.DESCRIPTION
    This script reads the history file created by Poll-M365Guidance.ps1 and provides
    various reporting and analysis options.

.PARAMETER HistoryFile
    Path to the JSON history file. Default is ".\M365GuidanceHistory.json"

.PARAMETER ShowAll
    Display all historical checks.

.PARAMETER ShowLatest
    Display only the most recent check.

.PARAMETER ShowChanges
    Display only checks where changes were detected.

.PARAMETER ExportCsv
    Export the history to a CSV file.

.PARAMETER CsvPath
    Path for the CSV export. Default is ".\M365GuidanceHistory.csv"

.EXAMPLE
    .\Get-M365GuidanceHistory.ps1 -ShowLatest
    Shows the most recent check.

.EXAMPLE
    .\Get-M365GuidanceHistory.ps1 -ShowAll
    Shows all historical checks.

.EXAMPLE
    .\Get-M365GuidanceHistory.ps1 -ExportCsv
    Exports history to CSV.
#>

[CmdletBinding(DefaultParameterSetName = 'ShowLatest')]
param(
    [string]$HistoryFile = ".\M365GuidanceHistory.json",

    [Parameter(ParameterSetName = 'ShowAll')]
    [switch]$ShowAll,

    [Parameter(ParameterSetName = 'ShowLatest')]
    [switch]$ShowLatest,

    [Parameter(ParameterSetName = 'ShowChanges')]
    [switch]$ShowChanges,

    [switch]$ExportCsv,
    [string]$CsvPath = ".\M365GuidanceHistory.csv"
)

function Get-HistoryData {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        Write-Error "History file not found: $FilePath"
        Write-Host "Run Poll-M365Guidance.ps1 first to create history data." -ForegroundColor Yellow
        return $null
    }

    try {
        $history = Get-Content $FilePath -Raw | ConvertFrom-Json
        return $history
    }
    catch {
        Write-Error "Failed to read history file: $_"
        return $null
    }
}

function Show-CheckData {
    param(
        [object]$Check,
        [bool]$Detailed = $false
    )

    Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
    Write-Host "Check Date: $($Check.CheckDate)" -ForegroundColor Yellow
    Write-Host "Last Updated: $($Check.LastUpdated)" -ForegroundColor Yellow
    Write-Host "Page Hash: $($Check.PageHash)" -ForegroundColor Gray
    Write-Host "Documents: $($Check.Documents.Count)" -ForegroundColor Yellow

    if ($Detailed -and $Check.Documents.Count -gt 0) {
        Write-Host "`nDocuments Found:" -ForegroundColor Cyan
        foreach ($doc in $Check.Documents) {
            Write-Host "  - $($doc.Title)" -ForegroundColor White
            Write-Host "    URL: $($doc.Url)" -ForegroundColor Gray
        }
    }
}

function Compare-Checks {
    param(
        [object]$Check1,
        [object]$Check2
    )

    $changes = @{
        PageHashChanged = $Check1.PageHash -ne $Check2.PageHash
        LastUpdatedChanged = $Check1.LastUpdated -ne $Check2.LastUpdated
        DocumentCountChanged = $Check1.Documents.Count -ne $Check2.Documents.Count
    }

    return $changes
}

# Main execution
$history = Get-HistoryData -FilePath $HistoryFile

if (-not $history) {
    exit 1
}

Write-Host "`nMicrosoft 365 Guidance Polling History" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Total checks: $($history.Count)" -ForegroundColor Yellow
Write-Host "History file: $HistoryFile" -ForegroundColor Gray

if ($ShowLatest -or $PSCmdlet.ParameterSetName -eq 'ShowLatest') {
    $latest = $history[-1]
    Show-CheckData -Check $latest -Detailed $true
}
elseif ($ShowAll) {
    foreach ($check in $history) {
        Show-CheckData -Check $check -Detailed $true
    }
}
elseif ($ShowChanges) {
    Write-Host "`nAnalyzing changes between checks..." -ForegroundColor Cyan

    for ($i = 1; $i -lt $history.Count; $i++) {
        $previous = $history[$i - 1]
        $current = $history[$i]

        $changes = Compare-Checks -Check1 $previous -Check2 $current

        if ($changes.PageHashChanged -or $changes.LastUpdatedChanged -or $changes.DocumentCountChanged) {
            Write-Host "`n$('=' * 80)" -ForegroundColor Green
            Write-Host "CHANGE DETECTED" -ForegroundColor Green
            Write-Host "From: $($previous.CheckDate)" -ForegroundColor Yellow
            Write-Host "To:   $($current.CheckDate)" -ForegroundColor Yellow

            if ($changes.LastUpdatedChanged) {
                Write-Host "  - Last Updated changed: $($previous.LastUpdated) -> $($current.LastUpdated)" -ForegroundColor Green
            }

            if ($changes.PageHashChanged) {
                Write-Host "  - Page content modified" -ForegroundColor Green
            }

            if ($changes.DocumentCountChanged) {
                Write-Host "  - Document count changed: $($previous.Documents.Count) -> $($current.Documents.Count)" -ForegroundColor Green
            }
        }
    }
}

# Export to CSV if requested
if ($ExportCsv) {
    try {
        $csvData = @()

        foreach ($check in $history) {
            $csvData += [PSCustomObject]@{
                CheckDate = $check.CheckDate
                LastUpdated = $check.LastUpdated
                PageHash = $check.PageHash
                DocumentCount = $check.Documents.Count
                Documents = ($check.Documents.Title -join "; ")
            }
        }

        $csvData | Export-Csv -Path $CsvPath -NoTypeInformation
        Write-Host "`nHistory exported to: $CsvPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to export CSV: $_"
    }
}

Write-Host ""
