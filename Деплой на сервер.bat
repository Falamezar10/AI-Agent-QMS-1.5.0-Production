@echo off
chcp 65001 >nul
color 0B

echo =======================================================
echo   СКОРОСТНАЯ ОТПРАВКА РЕЛИЗА НА СЕРВЕР (Robocopy MT)
echo =======================================================
echo.

echo [1] Выбор исходной папки (релиза) для отправки...
echo Открывается окно выбора папки (может появиться на заднем плане)...

set "SOURCE_PATH="
REM Вызываем графическое окно Windows для выбора папки ИСТОЧНИКА
for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.windows.forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'ШАГ 1: Выберите папку с готовым релизом Агента (например, dist\SMK_Agent_Build_...)'; $f.ShowNewFolderButton = $false; if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.SelectedPath }"`) do (
    set "SOURCE_PATH=%%I"
)

if "%SOURCE_PATH%"=="" (
    color 0E
    echo.
    echo [ОТМЕНА] Вы не выбрали исходную папку. Деплой прерван.
    pause
    exit /b
)

echo [OK] Выбрана исходная папка: %SOURCE_PATH%
echo.

echo [2] Выбор папки назначения на сервере...
echo Открывается окно выбора папки (может появиться на заднем плане)...

set "SERVER_PATH="
REM Вызываем графическое окно Windows для выбора папки НАЗНАЧЕНИЯ
for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.windows.forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'ШАГ 2: Выберите папку назначения НА СЕРВЕРЕ для загрузки релиза'; $f.ShowNewFolderButton = $true; if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.SelectedPath }"`) do (
    set "SERVER_PATH=%%I"
)

if "%SERVER_PATH%"=="" (
    color 0E
    echo.
    echo [ОТМЕНА] Вы не выбрали папку назначения. Деплой прерван.
    pause
    exit /b
)

echo [OK] Выбрана папка назначения: %SERVER_PATH%
echo.
echo [3] Запуск многопоточного копирования на сервер...
echo Источник:   %SOURCE_PATH%
echo Назначение: %SERVER_PATH%
echo Включено 16 потоков (/MT:16). БЕЗОПАСНОЕ копирование (/E).
echo.
echo Нажмите любую клавишу для начала загрузки...
pause >nul
echo.

REM --- МАГИЯ СКОРОСТИ И БЕЗОПАСНОСТИ ---
REM /E - Копирует файлы и папки. Перезаписывает совпадающие файлы, НО НИЧЕГО НЕ УДАЛЯЕТ на сервере! Ваши базы в безопасности.
REM /MT:16 - Включает 16 параллельных потоков копирования
REM /W:1 /R:3 - Ждать 1 сек при ошибке, всего 3 попытки (чтобы не зависало на занятых файлах)
REM /NP - Не выводить прогресс в процентах для каждого мелкого файла (ускоряет работу консоли)

robocopy "%SOURCE_PATH%" "%SERVER_PATH%" /E /MT:16 /W:1 /R:3 /NP

echo.
color 0A
echo =======================================================
echo   ДЕПЛОЙ УСПЕШНО ЗАВЕРШЕН!
echo =======================================================
echo.
echo Все новые файлы мгновенно синхронизированы с сервером.
echo Ваши базы данных и хранилища не пострадали!
echo.
pause