@echo off
cd /d %~dp0
call venv\Scripts\activate
set PYTHONUTF8=1
python main.py
pause
