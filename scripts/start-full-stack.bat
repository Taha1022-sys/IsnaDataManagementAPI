@echo off
echo =======================================================
echo  Excel Data Management - Full Stack Starter
echo =======================================================
echo.
echo This script will start both Backend and Frontend
echo.
echo Backend will run on: http://localhost:5002
echo Frontend will run on: http://localhost:3000
echo.
echo Press any key to continue...
pause > nul

echo.
echo Starting Backend in new window...
start "Excel Backend" cmd /c "%~dp0start-backend.bat"

echo.
echo Waiting 10 seconds for backend to start...
timeout /t 10 /nobreak > nul

echo.
echo Starting Frontend in new window...
start "Excel Frontend" cmd /c "%~dp0start-frontend.bat"

echo.
echo =======================================================
echo  Both services are starting...
echo.
echo  Backend: http://localhost:5002
echo  Frontend: http://localhost:3000
echo  Swagger: http://localhost:5002
echo.
echo  Wait for both windows to finish loading,
echo  then open http://localhost:3000 in your browser
echo =======================================================
echo.

pause