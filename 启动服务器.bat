@echo off
chcp 65001 > nul
cd /d "%~dp0"
title WPS AI助手服务器版 - 192.168.9.226:8765
echo 正在启动 WPS AI助手服务器版...
echo 服务地址: http://192.168.9.226:8765
echo 加载项:   http://192.168.9.226:8765/manifest.xml
echo.
python proxy.py
pause