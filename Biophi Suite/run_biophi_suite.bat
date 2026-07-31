@echo off
REM Biophi Suite — Windows desktop launcher
REM
REM Launches the GUI analysis app (main.py) from the App directory.
REM API keys are read from Windows user environment variables.
REM Set them once with:
REM   setx ROBOFLOW_API_KEY  "your-key-here"
REM   setx AIRTABLE_PAT      "your-token-here"
REM Never store API keys in this file or in source code.

cd /d "%~dp0App"
python main.py
pause
