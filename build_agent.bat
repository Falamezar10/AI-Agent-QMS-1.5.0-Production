@echo off
chcp 65001 >nul
echo ===================================================
echo   СБОРКА ИИ-АГЕНТА СМК (ZIP РЕЛИЗ ДЛЯ СЕРВЕРА)
echo ===================================================
echo.

echo [Шаг 1] Проверка и установка PyInstaller...
python -m pip install pyinstaller

echo.
echo [Шаг 2] Очистка промежуточных файлов сборки...
if exist "build" rmdir /s /q "build"
if exist "SMK_Agent.spec" del /q "SMK_Agent.spec"

echo.
echo [Шаг 3] Начинаем компиляцию (это займет 3-7 минут)...
echo Пожалуйста, не закрывайте это окно!
echo.

REM Собираем проект через --onedir для быстрого старта.
python -m PyInstaller --noconfirm --onedir --windowed --name "SMK_Agent" --collect-all "chromadb" --collect-all "pydantic" --hidden-import="sqlite3" --hidden-import="cryptography" --hidden-import="openpyxl" --hidden-import="docx" --hidden-import="fitz" --hidden-import="win32timezone" main.py

echo.
echo [Шаг 4] Формирование уникальной папки релиза...
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "TIMESTAMP=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%_%dt:~8,2%-%dt:~10,2%"
set "RELEASE_DIR=SMK_Agent_Build_%TIMESTAMP%"

if exist "dist\SMK_Agent" (
    ren "dist\SMK_Agent" "%RELEASE_DIR%"
)

echo.
echo [Шаг 5] Упаковка в единый ZIP-архив...
echo Запущен системный архиватор PowerShell, подождите...
REM Упаковываем саму папку релиза, чтобы при извлечении не было "каши" из файлов
powershell -nologo -noprofile -command "Compress-Archive -Path 'dist\%RELEASE_DIR%' -DestinationPath 'dist\%RELEASE_DIR%.zip' -Force"

echo.
echo ===================================================
echo   СБОРКА УСПЕШНО ЗАВЕРШЕНА!
echo ===================================================
echo.
echo В папке dist появился готовый архив: %RELEASE_DIR%.zip
echo.
echo [ПРАВИЛА ДЕПЛОЯ НА СЕРВЕР]:
echo 1. Скопируй ОДИН этот файл (.zip) на сетевой диск.
echo 2. Нажми по нему правой кнопкой мыши -^> "Извлечь все...".
echo 3. При первом запуске извлеченного .exe файла он сам развернет базы.
echo 4. ДЛЯ ОБНОВЛЕНИЯ: Точно так же извлекай новый zip с заменой файлов поверх старых. 
echo    Базы (smk_vector_db) и хранилище (secrets.vault) не пострадают!
echo.
pause
