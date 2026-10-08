@echo off
REM Windows Task Scheduler Runner for NSE Daily Scraper
REM action_cop was moved from D:\share_live\action_cop into the project itself.
cd /d "D:\behaviour analysis\action_cop"
echo [%DATE% %TIME%] Starting NSE Daily Scraper... >> scraper_log.txt
python "D:\behaviour analysis\action_cop\scrap.py" >> scraper_log.txt 2>&1
echo [%DATE% %TIME%] Syncing Live Corporate Feeds to Dashboard... >> scraper_log.txt
python "d:\behaviour analysis\sync_live_corporate_feeds.py" >> scraper_log.txt 2>&1
echo [%DATE% %TIME%] Finished NSE Daily Scraper and Dashboard Sync. >> scraper_log.txt

