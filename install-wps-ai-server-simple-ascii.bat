@echo off
setlocal

set "SERVER=http://192.168.9.226:8765"
set "ADDON_DIR=%APPDATA%\kingsoft\wps\jsaddons"
set "PUBLISH_FILE=%ADDON_DIR%\publish.xml"
set "TMP_FILE=%TEMP%\wps-ai-publish.xml"

echo ========================================
echo WPS AI add-in installer
echo version: 2026-07-02-ascii-v1
echo ========================================
echo.
echo Do NOT run as Administrator.
echo Close all WPS windows before running.
echo.
echo Server: %SERVER%
echo Config: %PUBLISH_FILE%
echo.

if not exist "%ADDON_DIR%" mkdir "%ADDON_DIR%"
if errorlevel 1 goto mkdir_failed

if exist "%PUBLISH_FILE%" attrib -R "%PUBLISH_FILE%" >nul 2>nul
if exist "%PUBLISH_FILE%" copy /Y "%PUBLISH_FILE%" "%PUBLISH_FILE%.bak" >nul 2>nul

> "%TMP_FILE%" echo ^<?xml version="1.0" encoding="UTF-8" standalone="yes"?^>
>>"%TMP_FILE%" echo ^<jsplugins^>
>>"%TMP_FILE%" echo   ^<jspluginonline name="wps-ai-qa" type="wps" url="%SERVER%/" debug="" enable="enable_dev" install="null"/^>
>>"%TMP_FILE%" echo ^</jsplugins^>
if errorlevel 1 goto tmp_failed

copy /Y "%TMP_FILE%" "%PUBLISH_FILE%" >nul
if errorlevel 1 goto copy_failed

echo.
echo [OK] WPS add-in config updated.
echo Open WPS and check add-in location:
echo %SERVER%
echo.
pause
exit /b 0

:mkdir_failed
echo.
echo [FAILED] Cannot create directory:
echo %ADDON_DIR%
pause
exit /b 1

:tmp_failed
echo.
echo [FAILED] Cannot write temp file:
echo %TMP_FILE%
pause
exit /b 1

:copy_failed
echo.
echo [FAILED] Cannot update:
echo %PUBLISH_FILE%
echo.
echo Fix:
echo 1. Close all WPS windows.
echo 2. End wps.exe, et.exe, wpp.exe, wpscloudsvr.exe in Task Manager.
echo 3. Run this file by normal double-click, not as Administrator.
echo.
pause
exit /b 1