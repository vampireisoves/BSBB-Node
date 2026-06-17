@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
mode con cols=80 lines=20
color 0A
title 依赖安装中……

cls
echo.
echo ==============================
echo      正在安装项目依赖
echo ==============================
echo.

if not exist "requirements.txt" (
    echo  错误：未找到 requirements.txt
    pause >nul
    exit /b
)

set "progres=                                                 "
set /a max=50
set /a now=0

:progress
set /a now+=1
set /a per=now*100/max

set "bar="
for /l %%i in (1,1,!now!) do set "bar=!bar!█"
for /l %%i in (!now!,1,!max!) do set "bar=!bar!░"

echo  安装进度：!bar! !per!%%
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
goto complete

:complete
cls
echo.
echo ==============================
echo      ✅ 依赖安装完成
echo ==============================
echo.
echo  所有依赖包已成功安装。
echo.
pause >nul
exit /b