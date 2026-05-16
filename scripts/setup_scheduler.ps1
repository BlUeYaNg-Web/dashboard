# setup_scheduler.ps1
# 用 PowerShell 建立 Windows Task Scheduler 排程
# 避免 setup_scheduler.bat 的中文亂碼問題
#
# 執行方式（以系統管理員開啟 PowerShell）：
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#   cd C:\path\to\dashboard
#   .\scripts\setup_scheduler.ps1

$repoPath = Split-Path -Parent $PSScriptRoot
$batFile  = Join-Path $repoPath "scripts\run_realprice.bat"

if (-not (Test-Path $batFile)) {
    Write-Host "Error: $batFile not found" -ForegroundColor Red
    Write-Host "Please make sure you run this from inside the dashboard repo folder."
    exit 1
}

$action   = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batFile`""
$trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "13:00"
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName  "RealPriceWeeklyUpdate" `
    -Action    $action `
    -Trigger   $trigger `
    -Settings  $settings `
    -RunLevel  Highest `
    -Force

Write-Host ""
Write-Host "Scheduled task created successfully!" -ForegroundColor Green
Write-Host "  Task name : RealPriceWeeklyUpdate"
Write-Host "  Schedule  : Every Sunday at 13:00"
Write-Host "  Script    : $batFile"
Write-Host ""
Write-Host "To run manually: schtasks /run /tn RealPriceWeeklyUpdate"
Write-Host "To check status: schtasks /query /tn RealPriceWeeklyUpdate /fo list"
