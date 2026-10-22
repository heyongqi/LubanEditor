@echo off
setlocal EnableExtensions EnableDelayedExpansion
rem 说明：本文件为 GBK 编码 + CRLF 换行，请勿用 UTF-8 另存，否则 cmd 解析会错乱。

title LubanEditor 打包工具

rem ============ 外层：调用内部 + 回显结果 ============
if /i not "%~1"=="--inner" (
    call "%~f0" --inner
    set "RC=!errorlevel!"
    echo.
    if "!RC!"=="0" (
        echo   ============================================
        echo     打包成功！exe 已生成在根目录：LubanEditor.exe
        echo   ============================================
        echo.
        echo   本窗口将在 3 秒后自动关闭 ...
        timeout /t 3 /nobreak >nul
    ) else (
        echo   ============================================
        echo     打包失败，请看上面的错误信息。
        echo   ============================================
        echo.
        echo   按任意键关闭 ...
        pause >nul
    )
    exit /b !RC!
)

rem ==================== 内部阶段 ====================
cd /d "%~dp0"
title LubanEditor 打包中 - 请勿关闭本窗口

call :say "============================================"
call :say "  LubanEditor 打包  开始时间 %date% %time%"
call :say "============================================"
echo.

rem ---------------- 1. 探测 Python ----------------
call :say "[1/4] 探测 Python ..."
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY ( py -c "import sys" >nul 2>nul && set "PY=py" )
if not defined PY ( python -c "import sys" >nul 2>nul && set "PY=python" )
if not defined PY ( python3 -c "import sys" >nul 2>nul && set "PY=python3" )

if not defined PY (
    echo.
    call :say "[错误] 本机没有找到可用的 Python。"
    call :say "   请安装 Python 3.9 到 3.12，安装时必须勾选 Add to PATH，"
    call :say "   然后重新双击本脚本。"
    call :say "   下载地址：https://www.python.org/downloads/windows/"
    echo.
    exit /b 1
)
call :say "      [正常] 使用命令：%PY%"

rem ---------------- 2. 虚拟环境 ----------------
call :say "[2/4] 准备虚拟环境 .venv ..."
set "VPY=.venv\Scripts\python.exe"
if exist ".venv" if not exist "%VPY%" (
    call :say "      发现损坏的 .venv，删除重建 ..."
    rmdir /s /q .venv
)
if not exist "%VPY%" (
    %PY% -m venv .venv >nul 2>&1
) else (
    call :say "      复用已有 .venv"
)
if not exist "%VPY%" (
    call :say "      [提示] .venv 创建失败，改为直接使用系统 Python 打包。"
    set "VPY=%PY%"
)
call :say "      解释器：%VPY%"

rem ---------------- 3. 依赖 ----------------
call :say "[3/4] 安装依赖，首次约 3-8 分钟，正在下载，请耐心等待 ..."
%VPY% -m pip install --upgrade pip >nul 2>&1
%VPY% -m pip install -r requirements.txt >nul 2>&1
if errorlevel 1 (
    echo.
    call :say "[错误] 依赖安装失败。可能原因："
    call :say "       1. 网络不通，或需要设置代理"
    call :say "       2. 上面提示的 venv 创建问题"
    echo.
    call :say "       可手动执行查看详细报错："
    call :say "       %VPY% -m pip install -r requirements.txt"
    echo.
    exit /b 1
)
%VPY% -m pip install pyinstaller >nul 2>&1
if errorlevel 1 (
    call :say "[错误] PyInstaller 安装失败。"
    exit /b 1
)
call :say "      依赖就绪"

rem ---------------- 4. 打包（exe 直接生成在根目录） ----------------
call :say "[4/4] PyInstaller 打包中，预计 2-5 分钟，请不要关闭窗口 ..."
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist LubanEditor.exe del /q LubanEditor.exe
%VPY% -m PyInstaller LubanEditor.spec --noconfirm --distpath .
if errorlevel 1 (
    call :say "[错误] PyInstaller 打包失败。"
    call :say "       可查看 build\LubanEditor\warn-LubanEditor.txt 了解缺失模块。"
    exit /b 1
)
if not exist "LubanEditor.exe" (
    call :say "[错误] 未找到打包产物 LubanEditor.exe"
    exit /b 1
)
for %%F in ("LubanEditor.exe") do call :say "      打包完成：%%~nxF    %%~zF 字节"

rem ---------------- 5. 自检 ----------------
call :say "[自检] 运行自检 ..."
"%~dp0LubanEditor.exe" --selftest
if errorlevel 1 (
    call :say "[警告] 自检未通过，请看上面的输出。"
    exit /b 1
)
call :say "      自检通过"

rem ---------------- 6. 清理临时日志 ----------------
call :say "[清理] 删除临时日志文件 ..."
if exist LubanEditor.log del /q LubanEditor.log
if exist build.log del /q build.log
if exist dep_ver.txt del /q dep_ver.txt

echo.
call :say "============================================"
call :say "  全部完成 %date% %time%"
call :say "  成品：LubanEditor.exe（根目录）"
call :say "============================================"
exit /b 0

rem ---------------- 子程序：输出一行进度 ----------------
:say
if "%~1"=="" (
    echo.
) else (
    echo %~1
)
exit /b 0
