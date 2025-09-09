@echo off
cd /d "%~dp0"

:: Activate virtual environment
call env\Scripts\activate.bat

:: Run your script
python test2.py

:: Pause so you can see output/errors
pause
