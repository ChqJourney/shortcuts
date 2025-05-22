@echo off
REM Build script for Folder Shortcut Creator

echo Building Folder Shortcut Creator...
echo ================================

REM Check if dotnet is installed
where dotnet >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: .NET SDK not found. Please install .NET 9.0 SDK.
    goto :end
)

REM Build and publish
echo Cleaning previous build...
dotnet clean -c Release

echo Building and publishing in AOT mode...
dotnet publish -c Release

if %ERRORLEVEL% neq 0 (
    echo ERROR: Build failed!
    goto :end
)

echo.
echo Build completed successfully!
echo Executable location: bin\Release\net9.0\win-x64\publish\scripts.exe

REM Copy batch files to publish directory
echo Copying batch files to publish directory...
copy run.bat bin\Release\net9.0\win-x64\publish\ >nul
copy dirs.txt bin\Release\net9.0\win-x64\publish\ >nul
copy README.md bin\Release\net9.0\win-x64\publish\ >nul

echo.
echo Setup completed! You can find all files in: bin\Release\net9.0\win-x64\publish\

:end
pause 