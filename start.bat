@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
rem Если t.me блокируется провайдером, раскомментируйте и укажите прокси:
rem set TG_PROXY=socks5h://127.0.0.1:1080
echo Запуск сервера поиска заказов...
python app.py
pause
