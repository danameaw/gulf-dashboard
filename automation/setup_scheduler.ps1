# setup_scheduler.ps1
# Run once to schedule the automation every Tuesday at 9:00 AM.
#
# Tuesday, not Wednesday: a week folder is created on Wednesday and the
# reports land in it over the days that follow, so on Wednesday morning the
# newest folder is the one that has just appeared and is still near-empty.
# find_week_folder() always takes the newest, so a Wednesday run reads that
# empty folder and reports almost every project missing. By Tuesday the
# newest folder is the previous week's, complete.

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
    -Weekly -DaysOfWeek Tuesday -At "09:00AM"

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
Write-Host "Runs every Tuesday at 09:00 AM"
Write-Host "Log file: $PSScriptRoot\run_log.txt"
