@echo off
setlocal EnableExtensions EnableDelayedExpansion
rem 说明：GBK 编码 + CRLF 换行，请勿用 UTF-8 另存。
rem 与 build.bat 同样的双写设计：日志进 run.log，进度实时打到屏幕。

title LubanEditor 启动器

if /i not "%~1"=="--inner" (
    cls
    echo.
    echo   ============================================
    echo     LubanEditor  源码启动
    echo   ============================================
    echo.
    echo     首次运行会安装依赖，约 1-3 分钟，请耐心等待。
    echo     进度实时显示在下面，完整日志写入 run.log
    echo.
    echo   ============================================
    echo.
    call "%~f0" --inner > run.log 2>&1
    set "RC=!errorlevel!"
    cls
    echo.
    echo   ---------- run.log 全文 ----------
    type run.log
    echo   ---------------------------------
    echo.
    echo   退出码 !RC!     0 表示正常
    echo.
    echo   按任意键关闭 ...
    pause >nul
    exit /b !RC!
)

cd /d "%~dp0"
title LubanEditor 运行中

call :say "[1/3] 探测 Python ..."
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY ( py -c "import sys" >nul 2>nul && set "PY=py" )
if not defined PY ( python -c "import sys" >nul 2>nul && set "PY=python" )
if not defined PY ( python3 -c "import sys" >nul 2>nul && set "PY=python3" )
if not defined PY (
    call :say ""
    call :say "[错误] 没找到 Python。请先安装 Python 3.9 到 3.12 并勾选 Add to PATH。"
    call :say "       https://www.python.org/downloads/windows/"
    exit /b 1
)
call :say "      [正常] 使用命令：%PY%"

call :say "[2/3] 准备环境与依赖 ..."
set "VPY=.venv\Scripts\python.exe"
if exist ".venv" if not exist "%VPY%" rmdir /s /q .venv
if not exist "%VPY%" %PY% -m venv .venv > CON 2>&1
if not exist "%VPY%" set "VPY=%PY%"

call :say "      ------ pip 输出，实时滚动 ------"
%VPY% -m pip install -q -r requirements.txt > CON 2>&1
if errorlevel 1 (
    call :say ""
    call :say "[错误] 依赖安装失败，请检查网络。"
    call :say "       可手动执行：%VPY% -m pip install -r requirements.txt"
    exit /b 1
)
call :say "      ------ pip 输出结束 ------"

call :say "[3/3] 启动界面，窗口马上弹出 ..."
%VPY% main.py > CON 2>&1
if errorlevel 1 (
    call :say "[错误] 程序异常退出，请看上面的输出。"
    exit /b 1
)
exit /b 0

:say
if "%~1"=="" (
    echo.
    echo. > CON
) else (
    echo %~1
    echo %~1 > CON
)
exit /b 0
