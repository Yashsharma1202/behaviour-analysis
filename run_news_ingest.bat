@echo off
REM Daily Google-News ingestion — stamps ingested_at so the causality history builds up.
cd /d "D:\behaviour analysis"
echo ================ %DATE% %TIME% ================ >> "D:\behaviour analysis\news_ingest.log"
"C:\Python314\python.exe" "D:\behaviour analysis\news_ingest.py" >> "D:\behaviour analysis\news_ingest.log" 2>&1
"C:\Python314\python.exe" "D:\behaviour analysis\forward_track.py" >> "D:\behaviour analysis\news_ingest.log" 2>&1
echo exit code: %ERRORLEVEL% >> "D:\behaviour analysis\news_ingest.log"
