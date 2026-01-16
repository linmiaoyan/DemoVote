@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ============================================
echo Git 提交并推送
echo ============================================
echo.

cd /d "%~dp0"

REM 检查Git是否安装
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Git，请先安装 Git for Windows
    pause
    exit /b 1
)

REM 检查是否是Git仓库，如果不是则初始化
if not exist ".git" (
    echo [提示] 当前目录不是Git仓库，正在初始化...
    git init
    if %errorlevel% neq 0 (
        echo [错误] Git仓库初始化失败
        pause
        exit /b 1
    )
    echo [成功] Git仓库初始化完成
    echo.
    
    REM 添加远程仓库（如果还没有）
    git remote -v | findstr "origin" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [提示] 添加远程仓库...
        git remote add origin https://github.com/linmiaoyan/DemoVote.git
        if %errorlevel% neq 0 (
            echo [警告] 添加远程仓库失败，但可以继续
        ) else (
            echo [成功] 远程仓库已添加
        )
        echo.
    )
)

echo [步骤1] 添加所有文件
git add .
if %errorlevel% neq 0 (
    echo [错误] 添加文件失败
    pause
    exit /b 1
)
echo [成功] 文件已添加到暂存区
echo.

echo [步骤2] 提交更改
echo.
set /p commit_msg="请输入提交描述: "

if "!commit_msg!"=="" (
    echo [错误] 提交描述不能为空
    pause
    exit /b 1
)

git commit -m "!commit_msg!"
if %errorlevel% neq 0 (
    echo [错误] 提交失败
    pause
    exit /b 1
)
echo [成功] 代码已提交
echo.

echo [步骤3] 推送到GitHub
REM 检查是否有提交（如果是新仓库）
git rev-parse --verify HEAD >nul 2>&1
if %errorlevel% neq 0 (
    REM 新仓库，没有提交历史，需要设置上游分支
    git push -u origin main
) else (
    REM 已有提交历史
    git push origin main
)
if %errorlevel% equ 0 (
    echo.
    echo ============================================
    echo [成功] 代码已推送到GitHub
    echo ============================================
    echo.
    echo 仓库地址：https://github.com/linmiaoyan/DemoVote
    echo.
) else (
    echo.
    echo [错误] 推送失败
    echo.
    echo 可能的原因：
    echo 1. 网络连接问题
    echo 2. 认证失败（需要Personal Access Token）
    echo 3. 权限不足
    echo.
)

pause
