@echo off
setlocal
rem Recorde - register the Word add-in for the current Windows user.
rem
rem Copies manifest.xml (next to this script) to %LOCALAPPDATA%\Recorde and
rem points Word's developer add-in registry at it - the same mechanism
rem Microsoft's office-addin-dev-settings tool uses. No admin rights, nothing
rem else is written. Undo with uninstall-windows.cmd.
rem
rem The value name is the add-in <Id> of the released manifest; the packaging
rem step checks that the two agree.
set "ID=9971ece4-30c4-421f-920e-8df1563b6bd5"
set "KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
set "SRC=%~dp0manifest.xml"
set "DEST=%LOCALAPPDATA%\Recorde"

if not exist "%SRC%" (
  echo manifest.xml was not found next to this script. Extract the whole zip first.
  goto :fail
)
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%SRC%" "%DEST%\manifest.xml" >nul || goto :fail
reg add "%KEY%" /v "%ID%" /t REG_SZ /d "%DEST%\manifest.xml" /f >nul || goto :fail

echo Recorde is registered: %DEST%\manifest.xml
echo Close Word completely and start it again. The Recorde button is on the Home tab.
echo (If the pane looks stale after an update, see INSTALL.md - "Updating".)
goto :end

:fail
echo Installation failed.
:end
echo.
pause
endlocal
