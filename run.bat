@echo off
chcp 65001 >nul
echo ========================================
echo VoteSite 投票系统启动脚本
echo ========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.7+
    pause
    exit /b 1
)

REM 检查依赖是否安装
echo [信息] 检查依赖包...
python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo [警告] 检测到缺少依赖包，正在安装...
    pip install -r requirement.txt
    if errorlevel 1 (
        echo [错误] 依赖包安装失败，请手动执行: pip install -r requirement.txt
        pause
        exit /b 1
    )
)

echo [信息] 启动服务器...
echo.

REM 设置环境变量（可选）
REM set HOST=0.0.0.0
REM set PORT=5000
REM set DEBUG=False
REM set ADMIN_GATE_KEY=wzkjgz

REM 运行应用
python run.py

pause

