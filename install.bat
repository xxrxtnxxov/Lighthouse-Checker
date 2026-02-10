@echo off
chcp 65001 >nul
echo ==========================================
echo 📦 Установка необходимых библиотек Python
echo ==========================================

:: Проверка наличия Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python не найден! Пожалуйста, установите Python с сайта python.org
    echo    Не забудьте поставить галочку "Add Python to PATH" при установке.
    pause
    exit /b
)

echo.
echo ⏳ Обновление pip...
python -m pip install --upgrade pip

echo.
echo ⏳ Установка зависимостей (pandas, aiohttp, openpyxl, etc)...
pip install aiohttp pandas openpyxl pydantic python-dotenv tqdm colorama

if %errorlevel% neq 0 (
    echo.
    echo ❌ Ошибка при установке библиотек. Проверьте интернет-соединение.
    pause
    exit /b
)

echo.
echo ==========================================
echo ✅ Все готово! Теперь можно запускать start.bat
echo ==========================================
pause