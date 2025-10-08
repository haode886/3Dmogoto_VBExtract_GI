@echo off
chcp 65001 > nul
:: 切换到批处理文件所在目录
cd /d "%~dp0"

:: 启动Python程序
echo 正在启动3DMigoto VB数据提取工具...
python luncher.py

:: 检查程序是否启动成功
if %errorlevel% equ 0 (
    echo 程序已成功启动
) else (
    echo 程序启动失败，请检查Python环境是否正确安装
    pause
)

exit