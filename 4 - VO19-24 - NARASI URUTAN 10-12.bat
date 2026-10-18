@echo off
setlocal
set "INPUT=%~1"
if "%INPUT%"=="" set "INPUT=%~dp0INPUT PART"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0AutoPartKeVO_IsiNarasi_V14.ps1" -Batch 4 -InputPath "%INPUT%"
endlocal
