# action_cop was moved from D:\share_live\action_cop into the project itself.
$action = New-ScheduledTaskAction -Execute "D:\behaviour analysis\action_cop\run_daily_scraper.bat" -WorkingDirectory "D:\behaviour analysis\action_cop"
$t1 = New-ScheduledTaskTrigger -Daily -At "09:30AM"
$t2 = New-ScheduledTaskTrigger -Daily -At "09:00PM"
Register-ScheduledTask -TaskName "NSE_Daily_Corporate_Scraper" -Action $action -Trigger @($t1, $t2) -Force
Write-Host "Task NSE_Daily_Corporate_Scraper successfully registered in Task Scheduler!"

