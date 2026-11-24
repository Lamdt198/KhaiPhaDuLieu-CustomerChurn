@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Khoi chay ung dung minh hoa Khai pha du lieu - Nhom 02...
python -m streamlit run app.py
pause
