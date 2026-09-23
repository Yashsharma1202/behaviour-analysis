@echo off
REM Wrapper for the daily Nifty50 Q1 update — logs output so failures are visible.
cd /d "D:\behaviour analysis"
echo ================ %DATE% %TIME% ================ >> "D:\behaviour analysis\update_q1.log"
"C:\Python314\python.exe" "D:\behaviour analysis\update_q1.py" >> "D:\behaviour analysis\update_q1.log" 2>&1
echo exit code: %ERRORLEVEL% >> "D:\behaviour analysis\update_q1.log"
