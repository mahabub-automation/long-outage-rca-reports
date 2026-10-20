@echo off
REM Windows Task Scheduler entry point. Schedule hourly, "Run only when user is logged on"
REM (Excel and WhatsApp Web need an interactive desktop). Set an execution time limit
REM of ~20 minutes so a stuck run cannot hang forever.
cd /d "%~dp0"
python run_all.py >> run.log 2>&1
