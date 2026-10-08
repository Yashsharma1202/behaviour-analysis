@echo off
REM Windows Task Scheduler Runner for NSE Daily Scraper
cd /d "d:\action_coop"
echo [%DATE% %TIME%] Starting NSE Daily Scraper... >> scraper_log.txt
python "d:\action_coop\scrap.py" >> scraper_log.txt 2>&1
echo [%DATE% %TIME%] Finished NSE Daily Scraper. >> scraper_log.txt
