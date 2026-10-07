@echo off
rem 더블클릭용: 실행 정책과 관계없이 build.ps1 실행
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1" %*
if errorlevel 1 pause
