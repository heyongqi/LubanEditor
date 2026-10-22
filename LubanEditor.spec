# -*- mode: python ; coding: utf-8 -*-
"""LubanEditor 打包配置（PyInstaller）。

要点：
1. 砍掉用不到的 Qt 子模块 —— PySide6 全量能到 1GB+，砍完通常 60~90MB；
2. --noconsole：GUI 程序，不弹黑框；
3. 图标可选（有就放 app.ico，没有也不影响）。
"""
from PyInstaller.utils.hooks import collect_submodules
import os

APP_NAME = "LubanEditor"
ENTRY = "main.py"

# 只用 QtCore / QtGui / QtWidgets，其余全部排除
QT_EXCLUDES = [
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
    "PySide6.QtWebChannel", "PySide6.QtWebSockets", "PySide6.QtNetworkAuth",
    "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.Qt3DInput", "PySide6.Qt3DLogic",
    "PySide6.Qt3DAnimation", "PySide6.Qt3DExtras",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtGraphs",
    "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets", "PySide6.QtSpatialAudio",
    "PySide6.QtQuick", "PySide6.QtQuick3D", "PySide6.QtQuickWidgets",
    "PySide6.QtQuickControls2", "PySide6.QtQuickTest",
    "PySide6.QtQml", "PySide6.QtQmlCore", "PySide6.QtQmlNetwork", "PySide6.QtQmlModels",
    "PySide6.QtQmlWorkerScript", "PySide6.QtQmlXmlListModel", "PySide6.QtQmlMeta",
    "PySide6.QtQmlCompiler", "PySide6.QtQmlIntegration",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning",
    "PySide6.QtPositioningQuick", "PySide6.QtSensors", "PySide6.QtSensorsQuick",
    "PySide6.QtSerialPort", "PySide6.QtSerialBus", "PySide6.QtStateMachine",
    "PySide6.QtRemoteObjects", "PySide6.QtScxml", "PySide6.QtSvg", "PySide6.QtSvgWidgets",
    "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtDesigner", "PySide6.QtUiTools",
    "PySide6.QtHelp", "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets",
    "PySide6.QtPdf", "PySide6.QtPdfWidgets", "PySide6.QtTextToSpeech",
    "PySide6.QtHttpServer", "PySide6.QtAxContainer", "PySide6.QtDBus",
    "PySide6.QtConcurrent", "PySide6.QtNetwork", "PySide6.QtXml",
    "PySide6.QtPrintSupport",
]

OTHER_EXCLUDES = [
    "numpy", "pandas", "scipy", "matplotlib", "PIL", "tkinter",
    "pytest", "IPython", "notebook", "setuptools", "pip", "distutils",
    "PySide6.QtAsyncio", "PySide6.QtLite",
    # openpyxl 处理 xlsx 不需要 lxml；谁间接拉进来就砍谁，能省约 10MB
    "lxml", "lxml.etree", "lxml.objectify",
]

# 单文件（分发友好，启动慢几秒）/ 单目录（启动快，要带整个文件夹）
# 用法：build.bat            -> 单文件
#       set LUBANKIT_ONEDIR=1 && build.bat -> 单目录
ONEDIR = os.environ.get("LUBANKIT_ONEDIR", "") == "1"

hiddenimports = ["lubankit.ui", "lubankit.model", "lubankit.projectconf",
                 "lubankit.exceldef", "lubankit.typecatalog", "lubankit.template",
                 "lubankit.runner", "lubankit.knowledge", "lubankit.publish",
                 "lubankit.projectwizard", "lubankit.xmldef"]

icon = "app.ico" if os.path.exists("app.ico") else None

a = Analysis(
    [ENTRY],
    pathex=["."],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=QT_EXCLUDES + OTHER_EXCLUDES,
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

if ONEDIR:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
        console=False,
        disable_windowed_traceback=False,
        icon=icon,
    )
    coll = COLLECT(
        exe, a.binaries, a.datas,
        strip=False, upx=False, upx_exclude=[], name=APP_NAME,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,          # 开 UPX 能再小一半，但容易被杀软误报，默认不开
        runtime_tmpdir=None,
        console=False,      # GUI 模式，不弹控制台
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon,
    )
    coll = None
