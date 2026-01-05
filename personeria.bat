@echo off
if not "%1"=="am_admin" (
    powershell -WindowStyle Hidden -Command "Start-Process '%~f0' -ArgumentList 'am_admin' -WindowStyle Hidden"
    exit /b
)
npm start