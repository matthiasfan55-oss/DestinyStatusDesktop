@echo off
if exist "%~dp0EditDestinyStatusConfig.ps1" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0EditDestinyStatusConfig.ps1"
) else (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0src\EditDestinyStatusConfig.ps1"
)
