@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 обновить_траекторию.py
    goto end
)
where python >nul 2>nul
if %errorlevel%==0 (
    python обновить_траекторию.py
    goto end
)
echo Python не найден.
echo Установите Python и запустите файл снова.
:end
pause
