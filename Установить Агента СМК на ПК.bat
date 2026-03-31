@echo off
chcp 65001 >nul
echo =======================================================
echo Установка и Обновление ИИ-Агента СМК (Для быстрого запуска)
echo =======================================================
echo.

REM --- ИМЯ ФАЙЛА ИКОНКИ ---
set "ICON_NAME=agent_icon_3.ico"

REM Получаем путь к текущей сетевой папке
set "SERVER_PATH=%~dp0"
set "LOCAL_DIR=%LOCALAPPDATA%\SMK_Agent_Client"

echo [1] Синхронизируем файлы программы с сервером...
echo Запущено скоростное многопоточное копирование. Пожалуйста, подождите...
echo.
if not exist "%LOCAL_DIR%" mkdir "%LOCAL_DIR%"

REM --- ИДЕАЛЬНАЯ ЧИСТОТА: Режим Зеркала (/MIR) ---
REM Удалили /E и поставили /MIR. Теперь папка клиента будет точной копией сервера.
robocopy "%SERVER_PATH%_internal" "%LOCAL_DIR%\_internal" /MIR /MT:8 /NP

REM Копируем exe и иконку
copy /Y "%SERVER_PATH%SMK_Agent.exe" "%LOCAL_DIR%\" >nul
if exist "%SERVER_PATH%%ICON_NAME%" copy /Y "%SERVER_PATH%%ICON_NAME%" "%LOCAL_DIR%\" >nul

echo.
echo [2] Создаем умный ярлык с иконкой на вашем Рабочем столе...
set "SHORTCUT=%USERPROFILE%\Desktop\ИИ-Агент СМК.lnk"
set "TARGET=%LOCAL_DIR%\SMK_Agent.exe"
set "CLEAN_SERVER=%SERVER_PATH:~0,-1%"
set "ARGS=--server ""%CLEAN_SERVER%"""

REM Магия PowerShell для ярлыка и привязки иконки
powershell -nologo -noprofile -Command "$wshell = New-Object -ComObject WScript.Shell; $s = $wshell.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%TARGET%'; $s.Arguments = '%ARGS%'; $s.WorkingDirectory = '%LOCAL_DIR%'; if (Test-Path '%LOCAL_DIR%\%ICON_NAME%') { $s.IconLocation = '%LOCAL_DIR%\%ICON_NAME%' }; $s.Save()"

echo.
echo =======================================================
echo ГОТОВО! Процесс завершен.
echo На вашем Рабочем столе находится актуальный "ИИ-Агент СМК".
echo =======================================================
pause