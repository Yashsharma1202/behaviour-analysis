@echo off
REM Windows Task Scheduler Runner for NSE Daily Scraper
REM action_cop was moved from D:\share_live\action_cop into the project itself.
cd /d "D:\behaviour analysis\action_cop"
echo [%DATE% %TIME%] Starting NSE Daily Scraper... >> scraper_log.txt
python "D:\behaviour analysis\action_cop\scrap.py" >> scraper_log.txt 2>&1
echo [%DATE% %TIME%] Running full refresh + publish (RBI, quarterly, prices, futures, dashboard, git push)... >> scraper_log.txt
python "D:\behaviour analysis\daily_refresh_and_publish.py" >> scraper_log.txt 2>&1
echo [%DATE% %TIME%] Finished NSE Daily Scraper and full refresh + publish. >> scraper_log.txt

