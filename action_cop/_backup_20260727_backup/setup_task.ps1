$action = New-ScheduledTaskAction -Execute 'd:\action_coop\run_daily_scraper.bat' -WorkingDirectory 'd:\action_coop'
$t1 = New-ScheduledTaskTrigger -Daily -At "09:30AM"
$t2 = New-ScheduledTaskTrigger -Daily -At "05:30PM"
Register-ScheduledTask -TaskName "NSE_Daily_Corporate_Scraper" -Action $action -Trigger @($t1, $t2) -Force
Write-Host "Task NSE_Daily_Corporate_Scraper successfully registered in Task Scheduler!"
