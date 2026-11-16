@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 (
 echo Kutuphaneler yuklenemedi. Python 3.11 kurulumunu kontrol edin.
 pause
 exit /b 1
)
py -m streamlit run app.py
pause
