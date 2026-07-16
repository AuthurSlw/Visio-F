@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0new_visio_swimlane_diagram.ps1" %*
