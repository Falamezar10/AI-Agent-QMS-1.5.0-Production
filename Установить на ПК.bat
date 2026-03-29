@echo off
chcp 65001 >nul
echo ===================================================
echo   УСТАНОВКА ИИ-АГЕНТА СМК
echo ===================================================
echo.

echo [Шаг 1] Проверка Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден! Установите Python 3.10+
    pause
    exit /b 1
)
echo Python найден

echo.
echo [Шаг 2] Создание виртуального окружения...
if not exist ".venv" (
    python -m venv .venv
    echo Виртуальное окружение создано
) else (
    echo Виртуальное окружение уже существует
)

echo.
echo [Шаг 3] Активация виртуального окружения...
call .venv\Scripts\activate.bat

echo.
echo [Шаг 4] Установка зависимостей...
python -m pip install --upgrade pip

echo.
echo   Установка основных библиотек...
pip install chromadb
pip install python-docx
pip install python-dotenv
pip install openai
pip install customtkinter
pip install keyboard
pip install openpyxl
pip install pywin32
pip install PyMuPDF
pip install requests
pip install wikipedia
pip install cryptography

echo.
echo [Шаг 5] Проверка структуры папок...
if not exist "SMK_Docs" mkdir "SMK_Docs"
if not exist "SMK_Docs\.cache" mkdir "SMK_Docs\.cache"
if not exist "Memory" mkdir "Memory"
if not exist "Sessions" mkdir "Sessions"

echo.
echo [Шаг 6] Проверка файла настроек...
if not exist "global_settings.json" (
    echo ОШИБКА: global_settings.json не найден!
    echo Скопируйте файл global_settings.json в папку проекта
    pause
    exit /b 1
)
echo global_settings.json найден

echo.
echo ===================================================
echo   УСТАНОВКА УСПЕШНО ЗАВЕРШЕНА!
echo ===================================================
echo.
echo Для запуска используйте start.bat
echo.
pause
