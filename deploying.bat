@echo off
set "startupFolder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "wpsFolder=%APPDATA%\kingsoft\wps\jsaddons\publish.xml"

echo "Creating Shortcut on %startupFolder%" 
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%startupFolder%\addonServer_shortcut.lnk'); $Shortcut.TargetPath = '%~dp0\addonServer.exe'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.Save()"

echo "Copying File to %wpsFolder%"
xcopy /s/y "jsplugins.xml" "%wpsFolder%"

pause