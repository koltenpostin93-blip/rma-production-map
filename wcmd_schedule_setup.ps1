# Run this once (as Administrator) to register the monthly WCMD scraper task.
# After registering, test it with: Start-ScheduledTask -TaskName "WCMD-Warehouse-Scraper"

$ProjectDir = "C:\Users\KoltenPostin\OneDrive - John Stewart and Associates\Desktop\Claude Code\rma-production-map"
$ScriptPath = Join-Path $ProjectDir "wcmd_scraper.py"
$LogPath    = Join-Path $ProjectDir "data\wcmd_scraper.log"

$Action = New-ScheduledTaskAction `
    -Execute "python" `
    -Argument $ScriptPath `
    -WorkingDirectory $ProjectDir

# Run at 6 AM on the 1st of every month
$Trigger = New-ScheduledTaskTrigger `
    -Weekly -WeeksInterval 4 -DaysOfWeek Monday -At "06:00AM"

# Simpler: use monthly trigger via CIM (Task Scheduler doesn't expose Monthly natively in PS)
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "WCMD-Warehouse-Scraper" `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Monthly download of USDA FSA WCMD licensed grain warehouse data" `
    -RunLevel Highest `
    -Force

Write-Host "Task registered. To test: Start-ScheduledTask -TaskName 'WCMD-Warehouse-Scraper'"
Write-Host "To view log after run: Get-Content '$LogPath'"
