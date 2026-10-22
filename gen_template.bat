set WORKSPACE=../..
set LUBAN_DLL=%WORKSPACE%\Tools\Luban\Luban.dll
set CONF_ROOT=.
set OUTPUT_ROOT=%WORKSPACE%\Assets\Res\Config
set OUTPUT_CODE_ROOT=%WORKSPACE%\Assets\Scripts\Config

dotnet %LUBAN_DLL% ^
    -t all ^
    -d json ^
    -c cs-simple-json ^
    --conf %CONF_ROOT%\luban.conf ^
    -x outputDataDir=%OUTPUT_ROOT% ^
    -x outputCodeDir=%OUTPUT_CODE_ROOT%
pause