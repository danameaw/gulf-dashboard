# setup_scheduler.ps1
# Run once (as Administrator) to schedule the automation every Wednesday at 9:00 AM

# Task Scheduler launches the action with CreateProcess, not through a shell, so
# ">>" redirection in -Argument is passed straight to run.py as argv and argparse
# aborts with exit code 2. Go through cmd.exe /c and let run_weekly.bat do the
# redirection instead.
$batPath = "$PSScriptRoot\run_weekly.bat"

$action  = New-ScheduledTaskAction `
    -Execute "$env:SystemRoot\System32\cmd.exe" `
    -Argument "/c `"`"$batPath`"`"" `
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
Write-Host "Log file: $PSScriptRoot\run_log.txt"
