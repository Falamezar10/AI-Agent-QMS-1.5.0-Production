@echo off
chcp 65001 >nul
color 0B
echo ===================================================
echo   СБОРКА ИИ-АГЕНТА СМК (ZIP РЕЛИЗ ДЛЯ СЕРВЕРА)
echo ===================================================
echo.

echo [Шаг 1] Проверка и установка PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 goto err_pyinstaller

echo.
echo [Шаг 2] Очистка промежуточных файлов сборки...
if exist "build" rmdir /s /q "build"
if exist "dist\SMK_Agent" rmdir /s /q "dist\SMK_Agent"
if exist "SMK_Agent.spec" del /q "SMK_Agent.spec"

echo.
echo [Шаг 3] Начинаем компиляцию (это займет 3-7 минут)...
echo Пожалуйста, не закрывайте это окно!
echo.

REM Если нарисуете новую иконку, поменяйте название в параметре --icon:
python -m PyInstaller --noconfirm --onedir --windowed --icon="agent_icon_3.ico" --name "SMK_Agent" --collect-all "chromadb" --collect-all "pydantic" --hidden-import="sqlite3" --hidden-import="cryptography" --hidden-import="openpyxl" --hidden-import="docx" --hidden-import="fitz" --hidden-import="win32timezone" main.py

if not exist "dist\SMK_Agent\SMK_Agent.exe" goto err_compile

echo.
echo [Шаг 4] Копирование зависимостей (FFmpeg)...
set "APP_DIR=dist\SMK_Agent"

if not exist "ffmpeg.exe" goto warn_ffmpeg
copy /Y "ffmpeg.exe" "%APP_DIR%\ffmpeg.exe" >nul
echo [OK] ffmpeg.exe успешно скопирован.
goto copy_ffprobe

:warn_ffmpeg
color 0E
echo [ВНИМАНИЕ] ffmpeg.exe НЕ НАЙДЕН рядом с этим скриптом! Аудио-резчик не попадет в сборку.
color 0B

:copy_ffprobe
if not exist "ffprobe.exe" goto skip_ffprobe
copy /Y "ffprobe.exe" "%APP_DIR%\ffprobe.exe" >nul
echo [OK] ffprobe.exe успешно скопирован.

:skip_ffprobe
echo.
echo [Шаг 5] Формирование имени релиза...
REM Надежное получение даты. Отсекаем все багованные скрытые символы
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set "dt=%dt:~0,14%"
set "TIMESTAMP=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%_%dt:~8,2%-%dt:~10,2%"
set "RELEASE_DIR=SMK_Agent_Build_%TIMESTAMP%"

echo Новое имя папки: %RELEASE_DIR%
ren "%APP_DIR%" "%RELEASE_DIR%"
if errorlevel 1 goto err_rename

echo.
echo [Шаг 6] Упаковка в единый ZIP-архив...
echo Запущен системный архиватор PowerShell, создаем dist\%RELEASE_DIR%.zip ...
powershell -nologo -noprofile -command "Compress-Archive -Path 'dist\%RELEASE_DIR%' -DestinationPath 'dist\%RELEASE_DIR%.zip' -Force"

echo.
color 0A
echo ===================================================
echo   СБОРКА УСПЕШНО ЗАВЕРШЕНА!
echo ===================================================
echo.
echo В папке dist появился готовый архив: %RELEASE_DIR%.zip
echo.
pause
exit /b

:err_pyinstaller
color 0C
echo [ОШИБКА] Проблема с установкой PyInstaller или Python!
pause
exit /b

:err_compile
color 0C
echo [КРИТИЧЕСКАЯ ОШИБКА] PyInstaller не смог собрать программу! Читайте логи выше.
pause
exit /b

:err_rename
color 0C
echo [ОШИБКА] Не удалось переименовать папку релиза!
pause
exit /b