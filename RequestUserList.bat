@ECHO OFF
setlocal
cd /d %~dp0

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {& '.\RequestUserList.ps1'}"
Timeout 3