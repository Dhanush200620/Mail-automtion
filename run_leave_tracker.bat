@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python main.py --imap --loop --interval 60
pause

