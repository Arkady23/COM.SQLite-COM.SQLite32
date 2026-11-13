@echo off

:: Проверка прав администратора
net session >nul 2>&1
if %errorlevel% neq 0 (
    rem Запрос прав администратора...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"

echo ==========================================
echo  UNREGISTRATION of library sqlite.net.dll
echo ==========================================
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe sqlite.net.dll /unregister

pause
