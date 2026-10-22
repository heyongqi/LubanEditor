# -*- coding: utf-8 -*-
"""LubanEditor 启动入口。

打包为 GUI（--noconsole）后没有控制台，一旦出错用户只会看到"双击没反应"。
所以这里做两件事：
1. 把 stdout / stderr 落到日志文件，未捕获异常也写进去；
2. 提供 --selftest，打包后可直接验证 Qt / openpyxl / 界面是否正常。
"""
import os
import sys
import traceback

APP_NAME = "LubanEditor"
_LINES = []


def _log_dir() -> str:
    """挑一个可写目录放日志：优先 exe 同目录（便携），否则用户目录 / 临时目录。"""
    cands = []
    if getattr(sys, "frozen", False):
        cands.append(os.path.dirname(sys.executable))
    cands.append(os.path.dirname(os.path.abspath(__file__)))
    cands.append(os.path.expanduser("~"))
    cands.append(os.environ.get("TEMP", "") or "/tmp")
    for d in cands:
        if not d:
            continue
        try:
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, APP_NAME + ".log")
            with open(p, "a", encoding="utf-8"):
                pass
            return d
        except Exception:
            continue
    return os.path.expanduser("~")


LOG_PATH = os.path.join(_log_dir(), APP_NAME + ".log")


def _emit(msg: str):
    _LINES.append(msg)
    try:
        print(msg)
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def _install_excepthook():
    def hook(etype, value, tb):
        _emit("=" * 60)
        _emit("未捕获异常：%s: %s" % (etype.__name__, value))
        _emit("".join(traceback.format_exception(etype, value, tb)))
        try:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(
                None, APP_NAME + " 出错了",
                "%s: %s\n\n完整日志：%s" % (etype.__name__, value, LOG_PATH))
        except Exception:
            pass
        sys.__excepthook__(etype, value, tb)

    sys.excepthook = hook


def _redirect_streams():
    """GUI 模式下 sys.stdout 可能是 None，换成写文件的伪流，避免第三方库 print 崩掉。"""
    class _W:
        def write(self, s):
            if s and s.strip():
                try:
                    with open(LOG_PATH, "a", encoding="utf-8") as f:
                        f.write(s if s.endswith("\n") else s + "\n")
                except Exception:
                    pass
            return len(s)

        def flush(self):
            pass

        def isatty(self):
            return False

    for name in ("stdout", "stderr"):
        s = getattr(sys, name, None)
        if s is None:
            setattr(sys, name, _W())


def _selftest(root: str) -> int:
    _redirect_streams()
    ok = True

    def chk(cond, msg):
        nonlocal ok
        _emit(("  OK    " if cond else "  FAIL  ") + msg)
        if not cond:
            ok = False

    _emit(APP_NAME + " self-test")
    _emit("  exe    : " + sys.executable)
    _emit("  frozen : %s" % getattr(sys, "frozen", False))
    _emit("  log    : " + LOG_PATH)

    try:
        from PySide6 import __version__ as qv
        chk(True, "PySide6 可用 (%s)" % qv)
    except Exception as e:
        chk(False, "PySide6 不可用：%s" % e)
        return 1
    try:
        import openpyxl
        chk(True, "openpyxl 可用 (%s)" % openpyxl.__version__)
    except Exception as e:
        chk(False, "openpyxl 不可用：%s" % e)
        return 1
    try:
        from lubankit import knowledge
        chk(len(knowledge.KNOWLEDGE) >= 20, "知识点库可用（%d 条）" % len(knowledge.KNOWLEDGE))
    except Exception as e:
        chk(False, "知识点库不可用：%s" % e)
        ok = False

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication(sys.argv)
        from lubankit.ui import MainWindow
        w = MainWindow()
        chk(True, "主窗口构造成功")
        if root and os.path.isdir(root):
            w.open_project(root, silent=True)
            chk(w.tree.topLevelItemCount() >= 1,
                "打开工程成功：树顶层 %d 个 / 表 %d / bean %d / 枚举 %d"
                % (w.tree.topLevelItemCount(), len(w.model.tables),
                   len(w.model.beans), len(w.model.enums)))
            chk(w.stack.currentWidget() is w.work, "已切到工作区")
            chk(w.windowTitle() == "LubanEditor",
                "窗口标题显示工具名：%s" % w.windowTitle())
            chk(w.bt_folder.text() == "打开", "工程目录区右侧有「打开」按钮")
            chk([w.tabs.tabText(i) for i in range(w.tabs.count())]
                == ["配置方式", "模板预览"],
                "右侧页签：%s" % [w.tabs.tabText(i) for i in range(w.tabs.count())])
            shown = [a.text() for a in w.toolbar.actions() if a.isVisible()]
            chk(all(t not in shown for t in
                    ("导出数据表模板", "静态自检", "生成并校验", "设置")),
                "精简后的工具条：%s" % shown)
        else:
            _emit("  SKIP  未指定工程目录，跳过导入测试")
    except Exception:
        _emit(traceback.format_exc())
        chk(False, "界面初始化失败")
    _emit("SELFTEST " + ("PASSED" if ok else "FAILED"))
    return 0 if ok else 1


def main():
    _install_excepthook()
    _redirect_streams()
    args = sys.argv[1:]
    if args and args[0] == "--selftest":
        return _selftest(args[1] if len(args) > 1 else "")

    from lubankit.ui import main as ui_main
    return ui_main()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        try:
            _emit("启动阶段异常：\n" + traceback.format_exc())
        except Exception:
            pass
        raise
