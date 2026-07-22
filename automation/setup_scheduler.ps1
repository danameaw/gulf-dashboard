# setup_scheduler.ps1
# Run once to schedule the automation every Wednesday at 9:00 AM

$pythonPath = (Get-Command python).Source
$scriptPath = "$PSScriptRoot\run.py"
$logPath    = "$PSScriptRoot\run.log"

$action  = New-ScheduledTaskAction `
    -Execute $pythonPath `
    -Argument "`"$scriptPath`" >> `"$logPath`" 2>&1" `
    -WorkingDirectory $PSScriptRoot

$trigger = New-ScheduledTaskTrigger `
    -Weekly -DaysOfWeek Wednesday -At "09:00AM"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

Register-ScheduledTask `
    -TaskName   "GulfDashboard_WeeklyUpdate" `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -Description "Gulf Energy Dashboard - weekly PDF extraction and GitHub push" `
    -Force

Write-Host "Task Scheduler registered: GulfDashboard_WeeklyUpdate"
Write-Host "Runs every Wednesday at 09:00 AM"
Write-Host "Log file: $logPath"
