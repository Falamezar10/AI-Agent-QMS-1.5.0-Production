@echo off
chcp 65001 >nul
echo =======================================================
echo Установка и Обновление ИИ-Агента СМК (Для быстрого запуска)
echo =======================================================
echo.

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

REM Копируем главный exe файл Агента
copy /Y "%SERVER_PATH%SMK_Agent.exe" "%LOCAL_DIR%\" >nul

REM --- НОВОЕ: Копируем утилиты транскрибации (Аудио-движок) ---
if exist "%SERVER_PATH%ffmpeg.exe" copy /Y "%SERVER_PATH%ffmpeg.exe" "%LOCAL_DIR%\" >nul
if exist "%SERVER_PATH%ffprobe.exe" copy /Y "%SERVER_PATH%ffprobe.exe" "%LOCAL_DIR%\" >nul

REM --- НОВОЕ: Универсальное копирование иконок ---
REM Захватываем все .ico файлы (если они нужны для интерфейса программы)
if exist "%SERVER_PATH%*.ico" copy /Y "%SERVER_PATH%*.ico" "%LOCAL_DIR%\" >nul

echo.
echo [2] Создаем умный ярлык на вашем Рабочем столе...
set "SHORTCUT=%USERPROFILE%\Desktop\ИИ-Агент СМК.lnk"
set "TARGET=%LOCAL_DIR%\SMK_Agent.exe"
set "CLEAN_SERVER=%SERVER_PATH:~0,-1%"
set "ARGS=--server ""%CLEAN_SERVER%"""

REM Магия PowerShell для ярлыка
REM Теперь мы явно указываем ярлыку брать иконку прямо из самого SMK_Agent.exe ($s.IconLocation = '%TARGET%, 0')
powershell -nologo -noprofile -Command "$wshell = New-Object -ComObject WScript.Shell; $s = $wshell.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%TARGET%'; $s.Arguments = '%ARGS%'; $s.WorkingDirectory = '%LOCAL_DIR%'; $s.IconLocation = '%TARGET%, 0'; $s.Save()"

echo.
echo =======================================================
echo ГОТОВО! Процесс завершен.
echo На вашем Рабочем столе находится актуальный "ИИ-Агент СМК".
echo =======================================================
pause