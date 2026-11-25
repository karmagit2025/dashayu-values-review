@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo   价值观考核 · 在线填写
echo   员工入口: http://127.0.0.1:8780/
echo   管理后台: http://127.0.0.1:8780/admin
echo   关闭本窗口即停止服务
echo.

python app.py || pause
