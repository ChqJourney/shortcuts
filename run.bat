@echo off
REM Set code page to UTF-8
chcp 65001 > nul
REM Enable delayed expansion for variables in loops
setlocal EnableDelayedExpansion
REM Read paths from dirs.txt file and execute the script

echo Windows Folder Shortcut Creator - Batch Processing Mode
echo ======================================================

REM Check if dirs.txt exists
if not exist dirs.txt (
    echo Error: dirs.txt file not found!
    echo Please create a dirs.txt file in the current directory, with target path on first line and source paths on subsequent lines.
    goto :end
)

REM Read first line as targetPath
set /p TARGET_PATH=<dirs.txt
echo Target Path: %TARGET_PATH%

REM Check if target path is empty
if "%TARGET_PATH%"=="" (
    echo Error: Target path cannot be empty!
    goto :end
)

echo.
echo Processing source paths...

REM Use a temporary file to process subsequent lines
set TEMP_FILE=%TEMP%\sources.txt
type dirs.txt > "%TEMP_FILE%"

REM Initialize counter
set COUNTER=0

REM Skip the first line
for /f "skip=1 tokens=*" %%a in (%TEMP_FILE%) do (
    if not "%%a"=="" (
        set /a COUNTER+=1
        echo [!COUNTER!] Processing source path: %%a
        scripts.exe "%TARGET_PATH%" "%%a"
        echo.
    )
)

REM Delete temporary file
del "%TEMP_FILE%" > nul 2>&1

echo All !COUNTER! operations completed!

:end
endlocal
pause 