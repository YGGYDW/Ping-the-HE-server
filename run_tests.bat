@echo off
REM run_tests.bat — 在当前目录直接调用 run_ping.ps1 并打开结果文件
REM 注意：请确保 run_ping1.ps1 与本文件位于同一目录

REM 切换到脚本所在目录（处理双击或从别处调用的情况）
cd /d "%~dp0"

REM 可选：设置控制台为 UTF-8（有助于显示中文）
chcp 65001 >nul

REM 以 PowerShell 执行脚本（绕过执行策略）
REM 如果希望 PowerShell 窗口最小化执行，请将 "start /min" 前缀取消注释并注释掉直接执行行
REM start /min "" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_ping.ps1"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_ping.ps1"

REM 等待脚本结束后自动打开结果文件（优先 xlsx）
if exist "%~dp0results.xlsx" (
    start "" "%~dp0results.xlsx"
) else if exist "%~dp0results.csv" (
    start "" "%~dp0results.csv"
)

REM 保持窗口，便于查看输出；完成后任意按键关闭
pause
