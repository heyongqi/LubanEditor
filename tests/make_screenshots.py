# -*- coding: utf-8 -*-
"""重新生成 docs/screenshots/ 里的界面截图。

用法（**必须用本机原生平台**，不要设 QT_QPA_PLATFORM）：
    python tests/make_screenshots.py

为什么不能用 offscreen：那个平台下 Qt 拿不到系统字体（QFontDatabase.families()
返回 0），中文会全部渲染成方框，截出来的图没法看。所以这里走原生平台、
窗口一闪而过、脚本不调 exec()。跑之前先关掉正在运行的 LubanEditor，
否则改动的数据表可能被占用。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication

from lubankit import ui
from lubankit.model import FieldDef

app = QApplication.instance() or QApplication(sys.argv)
app.setStyleSheet(ui.QSS)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.join(BASE, "sample_project")
OUT = os.path.join(BASE, "docs", "screenshots")
os.makedirs(OUT, exist_ok=True)


def shot(widget, name, w=0, h=0):
    if w and h:
        widget.resize(w, h)
    app.processEvents()
    app.processEvents()
    p = os.path.join(OUT, name)
    widget.grab().save(p)
    print("saved", name, widget.size())


w = ui.MainWindow()
w.resize(1500, 900)
w.show()
app.processEvents()
w.open_project(ROOT)
app.processEvents()

name = w.model.tables[0].full_name
w._load_table_editor(name)
fs = w._edit_fields
fs.append(FieldDef("items", "list,Item", "", "道具列表"))
fs.append(FieldDef("dict", "map,string,Item", "", "多语言名字"))
w._reload_fields_table()
# 最后一行：给 map 字段选上 json 单元格格式（分组也选一个带中文名的）
fs[-1].fmt = "json"
fs[-1].group = "c"
w._reload_fields_table()
w.tbl_fields.selectRow(len(fs) - 1)
app.processEvents()
print("表:", name, "字段:", [(f.name, f.type, f.fmt) for f in fs])
shot(w, "10-字段表-容器类型-分组与单元格格式.png", 2010, 900)

w._refresh_right()
app.processEvents()
shot(w.view_opts, "11-配置方式-字段横跨多列(合并格).png", 620, 760)

d = ui.CreateTypeDialog(w, w.model, name)
d.show()
app.processEvents()
shot(d, "12-新建自定义类型-两组单选独立与底部按钮.png", 760, 660)
d.close()

w._open_new_table(ask_leave=False)
app.processEvents()
shot(w.bt_nt_create.parent(), "13-新建配置表-底部按钮居中放大.png", 900, 96)

# 发布工程：用一个「工程目录里放着 Luban.dll、还没有 gen 脚本」的临时工程来截，
# 这样能看出执行方式 / 可执行文件 / 配置文件是怎么自动带出来的
import shutil
import tempfile

tdir = os.path.join(tempfile.gettempdir(), "lubankit_shot_proj")
shutil.rmtree(tdir, ignore_errors=True)
shutil.copytree(ROOT, tdir)
for f in ("gen.bat", "gen.sh"):
    fp = os.path.join(tdir, f)
    if os.path.exists(fp):
        os.remove(fp)
with open(os.path.join(tdir, "Luban.dll"), "wb") as fh:
    fh.write(b"")
w2 = ui.MainWindow()
w2.show()
app.processEvents()
w2.open_project(tdir)
app.processEvents()
p = ui.PublishDialog(w2, w2.model, w2.settings)
p.show()
app.processEvents()
app.processEvents()
shot(p, "14-发布工程-基础参数.png", 1000, 720)
p.bt_more.click()
app.processEvents()
app.processEvents()
shot(p, "15-发布工程-高级参数(点更多展开).png", 1000, 900)
p.close()
w2.close()
shutil.rmtree(tdir, ignore_errors=True)

mt = ui.ManageTypesDialog(w, w.model, w.fields)
mt.show()
app.processEvents()
app.processEvents()
shot(mt, "16-管理自定义类型-引用次数与位置.png", 1020, 660)
mt.close()

print("done")
