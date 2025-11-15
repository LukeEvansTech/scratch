<#
.SYNOPSIS
    Helper script to set up automated Microsoft 365 guidance monitoring.

.DESCRIPTION
    This script helps configure scheduled monitoring of the UK Government's Microsoft 365 guidance.
    It can create scheduled tasks or scheduled jobs for automated polling.

.PARAMETER Method
    How to schedule monitoring: 'TaskScheduler' or 'ScheduledJob'

.PARAMETER IntervalHours
    How often to check for updates (in hours). Default is 6.

.PARAMETER NotificationEmail
    Email address to send notifications to when changes are detected.

.PARAMETER SmtpServer
    SMTP server to use for email notifications.

.PARAMETER ScriptPath
    Full path to the Poll-M365Guidance.ps1 script.

.PARAMETER TestRun
    If specified, performs a test run without creating the scheduled task/job.

.EXAMPLE
    .\Setup-M365Monitoring.ps1 -TestRun
    Performs a test run to verify everything works.

.EXAMPLE
    .\Setup-M365Monitoring.ps1 -Method ScheduledJob -IntervalHours 4
    Creates a scheduled job that runs every 4 hours.

.EXAMPLE
    .\Setup-M365Monitoring.ps1 -Method TaskScheduler -IntervalHours 6 `
        -NotificationEmail "admin@example.gov.uk" `
        -SmtpServer "smtp.example.gov.uk"
    Creates a scheduled task with email notifications.
#>

[CmdletBinding()]
param(
    [ValidateSet('TaskScheduler', 'ScheduledJob')]
    [string]$Method = 'ScheduledJob',

    [int]$IntervalHours = 6,

    [string]$NotificationEmail,

    [string]$SmtpServer,

    [string]$ScriptPath,

    [switch]$TestRun
)

# Auto-detect script path if not provided
if (-not $ScriptPath) {
    $ScriptPath = Join-Path $PSScriptRoot "Poll-M365Guidance.ps1"
}

# Verify script exists
if (-not (Test-Path $ScriptPath)) {
    Write-Error "Poll-M365Guidance.ps1 not found at: $ScriptPath"
    Write-Host "Please specify the correct path using -ScriptPath parameter" -ForegroundColor Yellow
    exit 1
}

Write-Host "Microsoft 365 Guidance Monitoring Setup" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  Script: $ScriptPath" -ForegroundColor Gray
Write-Host "  Check Interval: Every $IntervalHours hours" -ForegroundColor Gray
Write-Host "  Method: $Method" -ForegroundColor Gray

if ($NotificationEmail) {
    Write-Host "  Email: $NotificationEmail" -ForegroundColor Gray
    Write-Host "  SMTP: $SmtpServer" -ForegroundColor Gray
}

# Build script arguments
$scriptArgs = @()
if ($NotificationEmail) {
    $scriptArgs += "-NotificationEmail `"$NotificationEmail`""
}
if ($SmtpServer) {
    $scriptArgs += "-SmtpServer `"$SmtpServer`""
}

$argumentString = $scriptArgs -join " "

if ($TestRun) {
    Write-Host "`n=== TEST RUN ===" -ForegroundColor Yellow
    Write-Host "Executing a single check..." -ForegroundColor Cyan

    try {
        & $ScriptPath @scriptArgs
        Write-Host "`nTest run completed successfully!" -ForegroundColor Green
        Write-Host "You can now run this script without -TestRun to set up scheduling." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Test run failed: $_"
        exit 1
    }

    exit 0
}

# Create scheduled monitoring
Write-Host "`nSetting up scheduled monitoring..." -ForegroundColor Cyan

if ($Method -eq 'ScheduledJob') {
    try {
        # Remove existing job if it exists
        $existingJob = Get-ScheduledJob -Name "M365GuidanceMonitor" -ErrorAction SilentlyContinue
        if ($existingJob) {
            Write-Host "Removing existing scheduled job..." -ForegroundColor Yellow
            Unregister-ScheduledJob -Name "M365GuidanceMonitor" -Force
        }

        # Create new job trigger
        $trigger = New-JobTrigger `
            -Once `
            -At (Get-Date).AddMinutes(5) `
            -RepetitionInterval (New-TimeSpan -Hours $IntervalHours) `
            -RepetitionDuration ([TimeSpan]::MaxValue)

        # Build script block
        $scriptBlock = [ScriptBlock]::Create(@"
& '$ScriptPath' $argumentString
"@)

        # Register scheduled job
        Register-ScheduledJob `
            -Name "M365GuidanceMonitor" `
            -ScriptBlock $scriptBlock `
            -Trigger $trigger `
            -MaxResultCount 10

        Write-Host "`nScheduled job created successfully!" -ForegroundColor Green
        Write-Host "Name: M365GuidanceMonitor" -ForegroundColor Gray
        Write-Host "First run: in 5 minutes" -ForegroundColor Gray
        Write-Host "Interval: Every $IntervalHours hours" -ForegroundColor Gray
        Write-Host "`nManage the job with:" -ForegroundColor Yellow
        Write-Host "  Get-ScheduledJob -Name M365GuidanceMonitor" -ForegroundColor Gray
        Write-Host "  Unregister-ScheduledJob -Name M365GuidanceMonitor" -ForegroundColor Gray
    }
    catch {
        Write-Error "Failed to create scheduled job: $_"
        exit 1
    }
}
elseif ($Method -eq 'TaskScheduler') {
    try {
        # Check if running as administrator
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

        if (-not $isAdmin) {
            Write-Error "Creating a scheduled task requires administrator privileges."
            Write-Host "Please run this script as Administrator." -ForegroundColor Yellow
            exit 1
        }

        # Build task action
        $action = New-ScheduledTaskAction `
            -Execute "powershell.exe" `
            -Argument "-ExecutionPolicy Bypass -File `"$ScriptPath`" $argumentString"

        # Create trigger (runs every X hours)
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours $IntervalHours)

        # Task settings
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -RunOnlyIfNetworkAvailable

        # Register task
        $taskName = "M365GuidanceMonitor"
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

        if ($existingTask) {
            Write-Host "Removing existing scheduled task..." -ForegroundColor Yellow
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }

        Register-ScheduledTask `
            -TaskName $taskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Description "Monitors UK Government Microsoft 365 guidance for updates" `
            -RunLevel Highest

        Write-Host "`nScheduled task created successfully!" -ForegroundColor Green
        Write-Host "Name: $taskName" -ForegroundColor Gray
        Write-Host "First run: in 5 minutes" -ForegroundColor Gray
        Write-Host "Interval: Every $IntervalHours hours" -ForegroundColor Gray
        Write-Host "`nManage the task with:" -ForegroundColor Yellow
        Write-Host "  Get-ScheduledTask -TaskName $taskName" -ForegroundColor Gray
        Write-Host "  Start-ScheduledTask -TaskName $taskName" -ForegroundColor Gray
        Write-Host "  Unregister-ScheduledTask -TaskName $taskName" -ForegroundColor Gray
    }
    catch {
        Write-Error "Failed to create scheduled task: $_"
        exit 1
    }
}

Write-Host "`nSetup complete!" -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Monitor the history file: M365GuidanceHistory.json" -ForegroundColor Gray
Write-Host "2. View results: .\Get-M365GuidanceHistory.ps1 -ShowLatest" -ForegroundColor Gray
Write-Host "3. Check logs if using scheduled jobs: Get-Job" -ForegroundColor Gray
Write-Host ""
