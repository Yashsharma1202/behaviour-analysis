@echo off
REM ============================================================================
REM  THE REFRESH BUTTON -- double-click this file any time to refresh every
REM  live data source and publish the update to the live dashboard right now,
REM  instead of waiting for the scheduled 9:30 AM / 9:00 PM run.
REM
REM  This cannot live on the public website itself -- that would mean
REM  embedding git push credentials in a page anyone can view. This file is
REM  the safe equivalent: a real one-click trigger that only works because
REM  it's running on YOUR machine, with YOUR Python environment and YOUR
REM  git credentials.
REM ============================================================================
cd /d "D:\behaviour analysis"
echo.
echo ==========================================================
echo   Refreshing all data and publishing to the live dashboard...
echo   (this can take 10-20 minutes -- progress is logged below
echo    and also saved to daily_refresh_and_publish.log)
echo ==========================================================
echo.
python "D:\behaviour analysis\daily_refresh_and_publish.py"
echo.
echo ==========================================================
echo   Done. See daily_refresh_and_publish.log for the full run.
echo ==========================================================
pause
