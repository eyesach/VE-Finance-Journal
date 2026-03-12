@echo off
title VE Sales Dashboard
echo Starting VE Sales Dashboard...
echo.
cd /d "%~dp0UpdateAccountsFlow"
start /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:3000"
npm run serve
pause
