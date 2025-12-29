@echo off
echo Searching for MuseScore 4 executable...
echo.

REM Try common installation paths
set MUSESCORE_PATH=

REM Check Program Files
if exist "C:\Program Files\MuseScore 4\bin\MuseScore4.exe" (
    set MUSESCORE_PATH=C:\Program Files\MuseScore 4\bin\MuseScore4.exe
    goto :found
)

REM Check Program Files (x86)
if exist "C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe" (
    set MUSESCORE_PATH=C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe
    goto :found
)

REM Check user's AppData Local
if exist "%LOCALAPPDATA%\Programs\MuseScore 4\MuseScore4.exe" (
    set MUSESCORE_PATH=%LOCALAPPDATA%\Programs\MuseScore 4\MuseScore4.exe
    goto :found
)

echo MuseScore 4 not found in common locations.
echo Please edit this script and set MUSESCORE_PATH to your MuseScore 4 installation path.
echo.
pause
exit /b 1

:found
echo Found MuseScore 4 at: %MUSESCORE_PATH%
echo.
echo Starting MuseScore 4 in debug mode...
echo Debug output will appear below:
echo.
"%MUSESCORE_PATH%" -d
pause

