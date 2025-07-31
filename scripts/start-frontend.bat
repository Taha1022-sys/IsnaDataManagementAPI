@echo off
echo =================================
echo  Excel Data Management Frontend
echo =================================
echo.

cd /d "%~dp0.."

echo Checking Node.js installation...
node --version
if %ERRORLEVEL% neq 0 (
    echo ERROR: Node.js not found!
    echo Please install Node.js 18+ from: https://nodejs.org/
    pause
    exit /b 1
)

echo.
echo Checking npm installation...
npm --version
if %ERRORLEVEL% neq 0 (
    echo ERROR: npm not found!
    pause
    exit /b 1
)

echo.
echo Installing dependencies...
npm install

if %ERRORLEVEL% neq 0 (
    echo ERROR: npm install failed!
    pause
    exit /b 1
)

echo.
echo =================================
echo  Starting Frontend Server
echo =================================
echo  React App: http://localhost:3000
echo =================================
echo.

npm run dev

pause