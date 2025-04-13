@echo off
cd /d %~dp0
call venv\Scripts\activate
python paint_analytics.py
pause
