@echo off
echo =================================
echo  Excel Data Management Backend
echo =================================
echo.

cd /d "%~dp0..\ExcelDataManagementAPI"

echo Checking .NET installation...
dotnet --version
if %ERRORLEVEL% neq 0 (
    echo ERROR: .NET 9.0 SDK not found!
    echo Please install .NET 9.0 SDK from: https://dotnet.microsoft.com/download
    pause
    exit /b 1
)

echo.
echo Restoring NuGet packages...
dotnet restore

echo.
echo Building project...
dotnet build

if %ERRORLEVEL% neq 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo =================================
echo  Starting Backend Server
echo =================================
echo  API: http://localhost:5002/api
echo  Swagger: http://localhost:5002
echo =================================
echo.

dotnet run

pause