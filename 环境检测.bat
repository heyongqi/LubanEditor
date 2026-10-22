@echo off
setlocal EnableExtensions EnableDelayedExpansion
rem 说明：GBK 编码 + CRLF 换行。用来快速判断为什么 build.bat 跑不起来。

cd /d "%~dp0"

if /i not "%~1"=="--inner" (
    call "%~f0" --inner > 环境检测结果.txt 2>&1
    cls
    type 环境检测结果.txt
    echo.
    echo 结果已保存：环境检测结果.txt
    pause
    exit /b 0
)

echo ============================================
echo   LubanEditor 环境检测
echo ============================================
echo 检测时间：%date% %time%
echo 脚本目录：%cd%
echo.

echo ---------- 1. Python 命令 ----------
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY ( py -c "import sys" >nul 2>nul && set "PY=py" )
if not defined PY ( python -c "import sys" >nul 2>nul && set "PY=python" )
if not defined PY ( python3 -c "import sys" >nul 2>nul && set "PY=python3" )

if defined PY (
    echo   [正常] 找到 Python，命令为：%PY%
    %PY% -c "import sys;print('         版本：'+sys.version.split()[0])"
    %PY% -c "import sys;print('         路径：'+sys.executable)"
) else (
    echo   [问题] 没有找到 Python！
    echo          请安装 Python 3.9~3.12 并勾选 Add to PATH
    echo          https://www.python.org/downloads/windows/
)

echo.
echo ---------- 2. 磁盘空间 ----------
for %%D in (%CD:~0,2%) do echo   当前盘 %%D 剩余空间：
dir %CD:~0,2%\ | findstr /C:"字节可用" /C:"bytes free"

echo.
echo ---------- 3. 关键文件 ----------
for %%F in (main.py LubanEditor.spec requirements.txt make_sample.py 使用说明.txt lubankit) do (
    if exist "%%F" ( echo   [有] %%F ) else ( echo   [缺失] %%F  <- 缺这个会打包失败 )
)

echo.
echo ---------- 4. 已有产物 ----------
if exist "dist\LubanEditor.exe" (
    echo   [已有] dist\LubanEditor.exe
    for %%F in ("dist\LubanEditor.exe") do echo          大小：%%~zF 字节
) else (
    if exist "dist\LubanEditor\LubanEditor.exe" (
        echo   [已有] dist\LubanEditor\LubanEditor.exe （单目录版）
    ) else (
        echo   还没有打包产物
    )
)

echo.
echo ---------- 5. 上次打包日志（末尾 20 行）----------
if exist build.log (
    echo   --- build.log 末尾 ---
    powershell -NoProfile -Command "Get-Content build.log -Tail 20" 2>nul || type build.log
) else (
    echo   没有 build.log，说明 build.bat 尚未成功运行过
)

echo.
echo ---------- 6. 结论 ----------
if defined PY (
    echo   环境看起来正常。若 build.bat 仍失败，请把
    echo   "环境检测结果.txt" 和 "build.log" 一起发给开发者。
) else (
    echo   先安装 Python，再重新双击 build.bat。
)
exit /b 0
