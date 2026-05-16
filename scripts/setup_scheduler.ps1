# setup_scheduler.ps1
# Run as Administrator in PowerShell:
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   .\scripts\setup_scheduler.ps1

$repoPath = Split-Path -Parent $PSScriptRoot
$batFile  = Join-Path $repoPath "scripts\run_realprice.bat"

$action  = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batFile`""
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "13:00"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 1) -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "RealPriceWeeklyUpdate" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force

Write-Host ""
Write-Host "Scheduled task created: RealPriceWeeklyUpdate"
Write-Host "Runs every Sunday at 13:00"
Write-Host "Script: $batFile"
