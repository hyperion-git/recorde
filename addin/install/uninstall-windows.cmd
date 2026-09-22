@echo off
setlocal
rem Recorde - remove the Word add-in registration made by install-windows.cmd.
set "ID=9971ece4-30c4-421f-920e-8df1563b6bd5"
set "KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
set "DEST=%LOCALAPPDATA%\Recorde"

reg delete "%KEY%" /v "%ID%" /f >nul 2>&1
if exist "%DEST%\manifest.xml" del /Q "%DEST%\manifest.xml"
if exist "%DEST%" rmdir "%DEST%" 2>nul

echo Recorde is unregistered. Close Word completely and start it again.
echo Documents keep their equations (they are ordinary pictures with the LaTeX stored inside).
echo.
pause
endlocal
