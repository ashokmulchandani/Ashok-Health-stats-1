@echo off
cd /d "%~dp0"
echo Starting Ashok Daily Entry...
start "" http://localhost:8765
python connector.py
pause
