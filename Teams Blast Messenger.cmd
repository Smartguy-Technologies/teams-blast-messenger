@echo off
rem Starts Teams Blast Messenger without a console window.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Start-TeamsBlastMessenger.ps1"
