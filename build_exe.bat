@echo off
chcp 65001 >nul
echo ============================================
echo   Сборка "Генератор заявочников" в .exe
echo ============================================
echo.

echo [1/3] Устанавливаю/обновляю зависимости...
pip install --upgrade pyinstaller pandas openpyxl
if errorlevel 1 goto :error

echo.
echo [2/3] Собираю exe-файл (это может занять минуту-две)...
pyinstaller --noconfirm --onefile --windowed ^
    --name "Генератор заявочников" ^
    --icon "icon\zaevochnik.ico" ^
    --add-data "icon;icon" ^
    run_zaevochnik.py
if errorlevel 1 goto :error

echo.
echo [3/3] Готово!
echo Exe-файл находится здесь: dist\Генератор заявочников.exe
echo Можно скопировать его куда угодно (например, на Рабочий стол) -
echo он не зависит от установленного Python.
echo.
pause
exit /b 0

:error
echo.
echo ❌ Что-то пошло не так во время сборки. Смотрите сообщение об ошибке выше.
pause
exit /b 1
