@echo off
echo Starting Real Estate Mailer Generator (Tri-Fold Edition)...
echo.
cd /d "%~dp0"
call env\Scripts\activate
python mailer_app_trifold.py
pause
