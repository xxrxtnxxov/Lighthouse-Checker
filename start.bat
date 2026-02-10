@echo off
chcp 65001 >nul
title Google PageSpeed Analyzer

echo Запуск скрипта...
echo.

:: Запуск Python скрипта (предполагается, что файл называется main.py)
python main.py

echo.
if %errorlevel% neq 0 (
    echo ❌ Скрипт завершился с ошибкой. См. сообщение выше.
) else (
    echo ✅ Работа завершена успешно.
)

echo.
echo Нажмите любую клавишу, чтобы закрыть это окно...
pause >nul