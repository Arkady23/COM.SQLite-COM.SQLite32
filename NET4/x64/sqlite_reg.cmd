@echo off

:: Проверка прав администратора
net session >nul 2>&1
if %errorlevel% neq 0 (
    rem Запрос прав администратора...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Возвращаемся в папку, где лежит сам .cmd файл (важно при повышении прав!)
cd /d "%~dp0"

echo ========================================
echo  REGISTRATION of library sqlite.net.dll
echo ========================================
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe sqlite.net.dll /codebase

pause
