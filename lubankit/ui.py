"""LubanEditor 主界面（PySide6）。

布局：
- 没有打开工程时：首页只有「打开 / 创建 / 初始化 Luban 工程」三个按钮；
- 打开工程后分四个区域：
  1. 顶部功能区 —— 工具条（打开工程 / 发布工程 / 新建表 / 类型助手 / 管理自定义类型 /
     知识点；导出数据表模板、静态自检、生成并校验、设置四个入口在界面隐藏、代码保留）；
  2. 左侧工程区 —— 标题「工程目录」+「打开」按钮 + 完整目录结构；
  3. 中间文件区 —— 配置表编辑器（可编辑）/ 只读预览 / 新建配置表（内联，不弹窗）；
  4. 右侧上下文区 —— 配置方式 / 模板预览
     （「知识点」搬到顶部工具条开独立窗口，「问题 / 校验」不再显示）。
工程完整路径占一行太浪费，改成写在窗口标题上；资源管理器入口收在左侧标题右侧。
「知识点」是独立（非模态）窗口：顶部工具条打开，带上当前表字段推导出的关键词。

配置表编辑器采用「缓冲 + 提交」模型：改动先进缓冲区并标红，点底部
「保存（写入定义文件）」才落到 model 与 xlsx，「取消修改」丢弃缓冲；
底部的保存/取消按钮**只在有改动时才出现**。保存时只校验当前这张表，
不会因为别的表有问题而拦你。
新建配置表也不再弹独立窗口，而是显示在中间区域。
「类型助手」是个独立对话框，分四块：顶部标题、左侧类型、右上常用标签、
右下配置示例；常用标签随所选类型变化，类型列表含通用的 bean / 多态bean。
字段表的「配置方式」与它取的是**同一份标签定义**（`typecatalog.TAGS_BY_KIND`），
所以下拉里能选的东西和右侧说明、类型助手永远对得上；「配置方式」页签里
当前选中项标绿（左竖条 + 浅绿底 + ★ 当前选择）。
新建表页的初始字段区与编辑器共用同一套构建代码（`_append_field_row` 等）。
「类型」下拉可以直接手写（容器 `list,Reward`、带标签的 `int#ref=…`），
回车或失焦即落到字段上。

「发布工程」在工具栏加粗入口（紧跟「打开工程」）：把 gen.bat / gen.sh 的参数
摊开成一张可编辑的表 —— 基础参数照官方模板脚本，高级参数照官方 CLI 全参数
（点「更多」展开），每个参数都带备注；`-t / -c / -d` 下拉选择，文件 / 目录
既能输也能选；发布时可选是否覆盖脚本，不覆盖就生成临时脚本跑完即删；
点「发布」运行 gen.bat / gen.sh。参数清单 / 解析 / 生成都在 `publish.py`。

左侧目录树：单击在中间区预览 / 编辑，**双击文件用系统默认程序打开**（Excel 表就是
Excel），右键菜单也补了这项。「管理自定义类型」是工具栏上的深色按钮（只在打开了
Luban 工程时出现）：左侧「所有 / 公共 / 每张表」选范围，右侧列出类型、**引用次数**
（真扫整个工程统计，为 0 标红）、引用位置和删除按钮；删除有二次确认，被引用的类型
会列出全部引用位置，还有子结构继承时直接拒绝。

「点开的文件算不算配置表」由 `_table_for_file` 决定，三级判定：
① 按 `__tables__.xlsx` 里声明的 `input` 路径精确匹配；② `input` 留空时按 Luban
规则取表名（`TbItem` → `TbItem.xlsx`）；③ 再按文件名兜底（撞名就不猜，宁可只读）。
Data 目录里没被声明成配置表的 xlsx 仍然只读，但会**说清原因和出路**，
免得看起来像「配置表不能编辑了」。
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import Qt, QSettings, QTimer, Signal, QUrl
from PySide6.QtGui import (QAction, QBrush, QColor, QFont, QIcon,
                           QDesktopServices)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QStackedWidget, QSplitter, QTreeWidget,
    QTreeWidgetItem, QVBoxLayout, QHBoxLayout, QFormLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QTabWidget, QTextBrowser, QListWidget, QListWidgetItem,
    QMessageBox, QFileDialog, QDialog, QDialogButtonBox, QCheckBox, QGroupBox,
    QToolBar,
    QTextEdit, QMenu, QAbstractItemView, QFrame, QSizePolicy, QButtonGroup,
    QRadioButton, QSpinBox, QToolBar, QToolButton, QStatusBar, QScrollArea,
)

from .model import ProjectModel, TableDef, BeanDef, EnumDef, FieldDef, EnumItem
from . import (projectconf, exceldef, xmldef, typecatalog, template, runner,
               knowledge, projectwizard, publish)

# ------------------------------------------------------------------ 样式
QSS = """
* { font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif; }
QWidget { font-size: 12px; }
QMainWindow, QDialog { background: #F4F6FA; }
QToolBar { background: #FFFFFF; border-bottom: 1px solid #E2E8F0; padding: 6px 8px; }
QToolBar QToolButton {
    background: #FFFFFF; border: 1px solid #D5DDE8; border-radius: 6px;
    padding: 6px 12px; margin-right: 4px; color:#1F2937;
}
QToolBar QToolButton:hover { background: #EEF4FF; border-color: #7EA6E8; }
QToolBar QToolButton:pressed { background: #DCE8FB; }
/* 深色工具条按钮：给「用了会有副作用」的功能（比如删类型）用，普通按钮仍是浅色 */
QToolButton#deep {
    background: #1D4ED8; border: 1px solid #1740AE; color: #FFFFFF; font-weight: bold;
}
QToolButton#deep:hover { background: #2559E8; border-color: #1B48BE; }
QToolButton#deep:pressed { background: #163FB0; }
/* 加粗文字的工具条按钮：给主流程入口用 */
QToolButton#bold { font-weight: bold; color: #0F172A; }
/* 像链接一样的小按钮（展开 / 收起高级参数） */
QPushButton#linkbtn {
    border: 0; background: transparent; color: #2F6FEB; padding: 2px 0;
    font-weight: bold;
}
QPushButton#linkbtn:hover { color: #1D4ED8; }
QPushButton {
    background: #FFFFFF; border: 1px solid #D5DDE8; border-radius: 6px;
    padding: 5px 12px; color:#1F2937;
}
QPushButton:hover { background: #EEF4FF; border-color: #7EA6E8; }
QPushButton:pressed { background: #DCE8FB; }
QPushButton#primary { background:#2F6FEB; border-color:#2F6FEB; color:#fff; }
QPushButton#primary:hover { background:#3B7BF5; }
QPushButton#danger { color:#B42318; border-color:#F0B4AE; }
QPushButton#big {
    font-size: 14px; font-weight: bold; padding: 16px 22px; border-radius: 10px;
    border: 1px solid #C9D6E8; background:#FFFFFF;
}
QPushButton#big:hover { background:#EEF4FF; border-color:#7EA6E8; }
QPushButton#bigprimary {
    font-size: 14px; font-weight: bold; padding: 16px 22px; border-radius: 10px;
    background:#2F6FEB; border-color:#2F6FEB; color:#fff;
}
QPushButton#bigprimary:hover { background:#3B7BF5; }
/* 对话框/表单底部的主按钮：居中、加宽加大 */
QPushButton#foot {
    font-size: 14px; font-weight: bold; padding: 10px 30px; border-radius: 8px;
    min-width: 120px; border: 1px solid #C9D6E8; background:#FFFFFF;
}
QPushButton#foot:hover { background:#EEF4FF; border-color:#7EA6E8; }
QPushButton#footprimary {
    font-size: 14px; font-weight: bold; padding: 10px 30px; border-radius: 8px;
    min-width: 120px; background:#2F6FEB; border-color:#2F6FEB; color:#fff;
}
QPushButton#footprimary:hover { background:#3B7BF5; }
QGroupBox {
    border: 1px solid #E2E8F0; border-radius: 8px; margin-top: 12px;
    background:#FFFFFF; padding: 10px 12px 12px 12px;
}
QGroupBox::title {
    subcontrol-origin: margin; left: 12px; padding: 0 6px;
    color:#334155; font-weight: bold;
}
QLineEdit, QComboBox, QSpinBox {
    border: 1px solid #D5DDE8; border-radius: 6px; padding: 5px 8px;
    background:#FFFFFF; selection-background-color:#2F6FEB;
}
QLineEdit:focus, QComboBox:focus { border-color:#7EA6E8; }
QTableWidget { border:1px solid #E2E8F0; border-radius:8px; background:#FFFFFF;
    gridline-color:#EDF1F6; }
QHeaderView::section {
    background:#F0F4FA; border:0; border-right:1px solid #E2E8F0;
    border-bottom:1px solid #E2E8F0; padding:6px; font-weight:bold; color:#334155;
}
QTreeWidget { border:1px solid #E2E8F0; border-radius:8px; background:#FFFFFF; }
QTabWidget::pane { border:1px solid #E2E8F0; border-top:0; background:#FFFFFF;
    border-bottom-left-radius:8px; border-bottom-right-radius:8px; }
QTabBar::tab { padding:7px 16px; background:#F0F4FA; border:1px solid #E2E8F0;
    border-bottom:0; margin-right:2px; }
QTabBar::tab:selected { background:#FFFFFF; color:#2F6FEB; font-weight:bold; }
QTextBrowser { border:1px solid #E2E8F0; border-radius:8px; background:#FFFFFF; }
QStatusBar { background:#FFFFFF; border-top:1px solid #E2E8F0; }
QLabel#hint { color:#64748B; }
QLabel#banner { background:#FFF7E6; border:1px solid #F5D9A8; border-radius:8px;
    padding:10px 12px; color:#7A4B00; }
QFrame#savebar { background:#FFFFFF; border:1px solid #E2E8F0; border-radius:8px; }
QLabel#dirty { color:#B42318; font-weight:bold; }
QLabel#toast { background:#0F7B3F; color:#FFFFFF; font-size:13px; font-weight:bold;
    border-radius:16px; padding:9px 20px; }
QLabel#cap { color:#334155; font-weight:bold; }
QFrame#dlghead { background:#FFFFFF; border:1px solid #E2E8F0; border-radius:8px; }
QLabel#dlgtitle { font-size:15px; font-weight:bold; color:#1F2937; }
QListWidget { border:1px solid #E2E8F0; border-radius:8px; background:#FFFFFF; }
"""

# ------------------------------------------------------------------ 通用控件
class ValueCombo(QComboBox):
    """显示中文名、实际取值为代码的下拉框。"""

    def __init__(self, items: List[Tuple[str, str]], parent=None, editable=True):
        super().__init__(parent)
        self.setEditable(editable)
        self._disp: Dict[str, str] = {}
        for code, disp in items:
            self.addItem(disp, code)
            self._disp[code] = disp

    def set_code(self, code: str, disp: str = None):
        """选中 code；清单里没有就先插一个（显示名可另给，如带标签的元素类型）。"""
        code = (code or "").strip()
        i = self.findData(code)
        if i >= 0:
            self.setCurrentIndex(i)
            return
        self.insertItem(0, disp or self._disp.get(code, code), code)
        self.setCurrentIndex(0)

    def code(self) -> str:
        d = self.currentData()
        if d:
            return str(d)
        return typecatalog.strip_label(self.currentText().strip())


def _expand_cells(row) -> List:
    """把 (文本, 跨列数) 展开成 [文本, PAD, PAD…]，方便统一渲染。"""
    out: List = []
    for c in row:
        if isinstance(c, tuple):
            txt, span = (c[0], int(c[1])) if len(c) > 1 else (c[0], 1)
            out.append(txt)
            out.extend([typecatalog.PAD] * (max(1, span) - 1))
        else:
            out.append(c)
    return out


# 跨多列字段的填充色，与生成的数据表模板（template.MULTI_FILLS）一致：
# 相邻两个多列字段取不同深浅，一眼就能分清哪几列属于同一个字段。
MULTI_SPAN_COLORS = ("#FFF3B0", "#FFE380", "#FFD451", "#FFC107")

# 工具条上存在、但不显示在界面上的入口。留住 action 是为了让内部流程
# （比如「新建表」之后自动生成数据表模板）照旧能用，产品上又不占顶部空间。
HIDDEN_TOOLBAR_ACTIONS = ("导出数据表模板", "静态自检", "生成并校验", "设置")


def _html_table(grid: List[List[str]]) -> str:
    """渲染表头示例。跨列合并的单元格画成一张跨列的格子 ——

    复合字段在 Luban 数据表里会横跨多列，续列本身不写字段名。如果照原样渲成
    空格子，看起来就像「一个字段只占一列，旁边还有几个空列」，初学者极易误解。

    跨列的格子染成黄色，且从左到右按多列字段的出现顺序由浅到深循环：
    连续两个字段都是多列时颜色必然不同，不会看成一整块。
    """
    if not grid:
        return "<p>（无内容）</p>"
    rows_cells = [_expand_cells(r) for r in grid]
    ncol = max(len(r) for r in rows_cells)
    css = ("<style>table{border-collapse:collapse;font-size:12px}"
           "td{border:1px solid #C9D4E2;padding:3px 8px;white-space:nowrap}"
           "tr:nth-child(1){background:#D9E7F7;font-weight:bold}"
           "tr:nth-child(2){background:#EAF1FB}"
           ".m{background:#FBE3D6;font-weight:bold}</style>")

    def td(text, span, seq):
        cls = " class='m'" if str(text).startswith("##") else ""
        bg = ""
        if span > 1:
            bg = " background:%s;" % MULTI_SPAN_COLORS[seq % len(MULTI_SPAN_COLORS)]
        st = f" style='{bg}'" if bg else ""
        cs = f" colspan='{span}'" if span > 1 else ""
        tip = f" title='这一格跨 {span} 列'" if span > 1 else ""
        body = text if text not in (None, "") else "&nbsp;"
        return f"<td{cls}{st}{cs}{tip}>{body}</td>"

    rows = []
    for r in rows_cells:
        cells = r + [""] * (ncol - len(r))
        out, cur, seq = [], None, 0
        for c in cells:
            if c is typecatalog.PAD:
                if cur is not None:
                    cur[1] += 1
                continue
            if cur is not None:
                out.append(td(cur[0], cur[1], seq))
                if cur[1] > 1:
                    seq += 1
            cur = [c, 1]
        if cur is not None:
            out.append(td(cur[0], cur[1], seq))
        rows.append("<tr>%s</tr>" % "".join(out))
    return css + "<table>" + "".join(rows) + "</table>"


def _esc(s) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _widget_text(w) -> str:
    """QLineEdit / QComboBox / QCheckBox 的当前值，统一成字符串。"""
    if isinstance(w, QCheckBox):
        return "1" if w.isChecked() else ""
    if isinstance(w, QComboBox):
        return (w.currentText() or "").strip()
    return (w.text() or "").strip()


def _set_widget_text(w, val):
    """把值塞回 QLineEdit / QComboBox / QCheckBox。"""
    v = "" if val is None else str(val)
    if isinstance(w, QCheckBox):
        w.blockSignals(True)
        w.setChecked(v.strip().lower() in ("1", "true", "yes", "on"))
        w.blockSignals(False)
        return
    if isinstance(w, QComboBox):
        w.blockSignals(True)
        w.setEditText(v.strip())
        w.blockSignals(False)
        return
    w.blockSignals(True)
    w.setText(v)
    w.blockSignals(False)


GROUP_ITEMS = [
    ("", "（留空 = 各端都导出）"),
    ("c,s", "c,s（客户端 + 服务器）"),
    ("c", "c（仅客户端）"),
    ("s", "s（仅服务器）"),
    ("e", "e（仅编辑器）"),
    ("c,s,e", "c,s,e（全部）"),
]

GROUP_CN = {"c": "客户端", "s": "服务器", "e": "编辑器"}


def group_items(model) -> List[Tuple[str, str]]:
    """字段「分组」下拉的选项：留空 + 每个分组（带中文名）+ 全部组合。

    分组名来自工程 luban.conf 的 groups；没读到就用 Luban 默认的 c / s / e。
    字段要进多个分组时用逗号写（c,s），所以下拉保留可编辑。
    """
    gs = list(getattr(model, "groups", None) or []) or ["c", "s", "e"]
    out = [("", "（留空 = 各端都导出）")]
    for g in gs:
        cn = GROUP_CN.get(g, "")
        out.append((g, f"{g}（{cn}）" if cn else g))
    if len(gs) > 1:
        allv = ",".join(gs)
        label = " + ".join(GROUP_CN.get(g, g) for g in gs)
        out.append((allv, f"{allv}（{label}）"))
    return out

MODE_ITEMS = [
    ("map", "map（按主键查找，最常用）"),
    ("list", "list（列表，无主键 / 多主键）"),
    ("one", "one（全局单例）"),
    ("", "（留空 = 按 index 自动推断）"),
]

READ_ITEMS = [
    ("true", "true（从数据表头推断字段）"),
    ("false", "false（字段在结构 bean 里定义，支持继承多态）"),
    ("", "（留空）"),
]


def apply_style_to_type(f: FieldDef, style: str) -> str:
    """可空 / 禁止默认值 / 校验标签 这类配置方式会改写类型串。"""
    core, tags = typecatalog.split_tags(f.type)
    p = typecatalog.parse_type(f.type)
    if style.startswith("tag:"):
        # 「引用校验 / 范围校验 / 资源路径 / 本地化 / 长度约束」——往类型串上挂标签
        key = style.split(":", 1)[1]
        t = next((x for x in typecatalog.TAGS if x["key"] == key), None)
        if not t:
            return f.type
        tags = [x for x in tags if x.lstrip("#").split("=")[0] != key]
        tags.append(t["form"])
        return typecatalog.join_type(core, tags, nullable=p["nullable"])
    if style == "nullable":
        if p["kind"] == "container":
            return f.type
        return typecatalog.join_type(core, tags, nullable=True)
    if style == "bang":
        tags = [t for t in tags if t != "!"] + ["!"]
        return typecatalog.join_type(core, tags, nullable=False)
    tags = [t for t in tags if t != "!"]
    return typecatalog.join_type(core, tags, nullable=False)


def options_html(opts: List[Dict], cur: str = "") -> str:
    """渲染「配置方式」列表：标题加大，当前选择整块标绿（左竖条 + 浅绿底）。"""
    out: List[str] = []
    if opts:
        out.append("<p style='font-size:12px;color:#64748B;margin:2px 0 6px 0'>"
                   "表格第 1 列就是 Excel 的 A 列（写 <code>##var</code> / "
                   "<code>##type</code> 这些标记）；<b>黄色合并的格子</b>表示这一个字段"
                   "横跨了多列，它们在 Excel 里会合并成一个单元格，不是几个独立空列；"
                   "相邻两个多列字段颜色深浅不同，方便区分。"
                   "最后那行 <code>##</code> 备注写明每一列是第几个元素。</p>")
    for o in opts:
        sel = bool(cur) and o.get("key") == cur
        color = "#15803D" if sel else "#1F2937"
        bar = "#15803D" if sel else "#CBD5E1"
        bg = "#ECFDF3" if sel else "transparent"
        star = (" <span style='font-size:12px;color:#15803D;font-weight:bold'>"
                "★ 当前选择</span>") if sel else ""
        out.append("<div style='margin:12px 0 2px 0;padding:5px 8px;border-left:4px solid "
                   + bar + ";background:" + bg + ";'>"
                   + "<span style='font-size:15px;font-weight:bold;color:" + color + "'>"
                   + _esc(o.get("title", "")) + "</span>" + star + "</div>")
        out.append("<p style='margin:2px 0 4px 14px'><i>适用："
                   + _esc(o.get("suitable", "")) + "</i></p>")
        out.append(_html_table(o.get("grid") or []))
        if o.get("notes"):
            out.append("<ul>" + "".join(f"<li>{_esc(n)}</li>" for n in o["notes"]) + "</ul>")
    return "".join(out)


# ------------------------------------------------------------------ 首页
class HomePage(QWidget):
    """没打开工程时的首页：只放三个入口按钮，不放别的内容。"""

    openRequested = Signal()
    createRequested = Signal()
    initRequested = Signal()
    recentOpenRequested = Signal(str)
    recentDeleteRequested = Signal(str)

    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.addStretch(1)
        box = QFrame()
        box.setObjectName("card")
        box.setStyleSheet("QFrame#card{background:#FFFFFF;border:1px solid #E2E8F0;"
                          "border-radius:14px;}")
        lay = QVBoxLayout(box)
        lay.setContentsMargins(40, 36, 40, 32)
        title = QLabel("LubanEditor")
        title.setStyleSheet("font-size:26px;font-weight:bold;color:#1F2937;")
        title.setAlignment(Qt.AlignCenter)
        sub = QLabel("让使用 Luban 更 easy!")
        sub.setStyleSheet("font-size:13px;color:#64748B;")
        sub.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)
        lay.addSpacing(4)
        lay.addWidget(sub)
        lay.addSpacing(26)

        self.bt_open = QPushButton("打开工程")
        self.bt_open.setObjectName("bigprimary")
        self.bt_open.setToolTip("选择一个已有工程的目录（目录没有要求，任意文件夹都能打开）")
        self.bt_create = QPushButton("创建工程")
        self.bt_create.setObjectName("big")
        self.bt_create.setToolTip("选一个新目录，生成一整套工程骨架"
                                 "（luban.conf / gen.bat / Data / Defines 与结构定义文件）")
        # 「初始化工程」按钮在首页隐藏：只有打开了一个非完整目录时才需要，
        # 那个场景的入口放在预览区中间（见 _show_overview）。
        self.bt_init = QPushButton("初始化工程")
        self.bt_init.setObjectName("big")
        self.bt_init.setToolTip("在一个已经存在的目录里补上 Luban 配置，"
                                "不新建目录、不动已有文件")
        self.bt_init.hide()
        self.bt_open.clicked.connect(self.openRequested.emit)
        self.bt_create.clicked.connect(self.createRequested.emit)
        self.bt_init.clicked.connect(self.initRequested.emit)

        hl = QHBoxLayout()
        hl.setSpacing(12)
        for b in (self.bt_open, self.bt_create):
            hl.addWidget(b)
        lay.addLayout(hl)
        lay.addSpacing(16)

        hint = QLabel("打开 = 加载一个已有工程　·　创建 = 在新目录里从零搭好工程")
        hint.setStyleSheet("font-size:12px;color:#94A3B8;")
        hint.setAlignment(Qt.AlignCenter)
        lay.addWidget(hint)
        lay.addSpacing(18)

        # 最近打开的工程（可点击打开）
        self.lab_recent = QLabel("最近打开")
        self.lab_recent.setStyleSheet("font-size:13px;color:#64748B;font-weight:bold;")
        self.lab_recent.setAlignment(Qt.AlignCenter)
        self.lst_recent = QListWidget()
        self.lst_recent.setMaximumHeight(160)
        # 只能通过每行的「打开」按钮打开工程（点击整行不触发）
        lay.addWidget(self.lab_recent)
        lay.addWidget(self.lst_recent)
        self.lab_recent.hide()
        self.lst_recent.hide()

        outer.addWidget(box, 0, Qt.AlignCenter)
        outer.addStretch(1)

    def set_recent(self, paths: List[str]):
        """刷新「最近打开」列表；空则隐藏。每项显示工程名 + 路径，带打开/删除按钮。"""
        self.lst_recent.clear()
        paths = [p for p in (paths or []) if p]
        if not paths:
            self.lab_recent.hide()
            self.lst_recent.hide()
            return
        self.lab_recent.show()
        self.lst_recent.show()
        for p in paths:
            item = QListWidgetItem()
            item.setData(Qt.UserRole, p)
            row = self._make_recent_row(p)
            item.setSizeHint(row.sizeHint())
            self.lst_recent.addItem(item)
            self.lst_recent.setItemWidget(item, row)

    def _make_recent_row(self, p: str) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(8, 6, 8, 6)
        lay.setSpacing(8)
        # 左侧：工程名 + 路径分行显示（不左右滑动）
        txt = QWidget()
        tl = QVBoxLayout(txt)
        tl.setContentsMargins(0, 0, 0, 0)
        tl.setSpacing(2)
        name = QLabel(os.path.basename(os.path.normpath(p)) or p)
        name.setStyleSheet("font-weight:bold;color:#1F2937;")
        path = QLabel(p)
        path.setStyleSheet("color:#64748B;")
        path.setToolTip(p)
        path.setWordWrap(True)
        tl.addWidget(name)
        tl.addWidget(path)
        # 右侧：打开 / 删除按钮
        bt_open = QPushButton("打开")
        bt_open.setCursor(Qt.PointingHandCursor)
        bt_open.clicked.connect(lambda _=False, pp=p: self.recentOpenRequested.emit(pp))
        bt_del = QPushButton("删除")
        bt_del.setCursor(Qt.PointingHandCursor)
        bt_del.clicked.connect(lambda _=False, pp=p: self.recentDeleteRequested.emit(pp))
        lay.addWidget(txt, 1)
        lay.addWidget(bt_open)
        lay.addWidget(bt_del)
        return w


# ------------------------------------------------------------------ 建工程
class CreateProjectDialog(QDialog):
    """创建 / 初始化 Luban 工程。mode: create（新目录）| init（已有目录）。"""

    def __init__(self, parent, default_root: str = "", mode: str = "create"):
        super().__init__(parent)
        self.mode = mode if mode in ("create", "init") else "create"
        creating = self.mode == "create"
        self.setWindowTitle("创建 Luban 工程" if creating else "初始化 Luban 工程")
        self.setMinimumWidth(560)
        self.ed_root = QLineEdit(default_root)
        bt = QPushButton("选择目录…")
        bt.clicked.connect(self._pick)
        hb = QHBoxLayout()
        hb.addWidget(self.ed_root, 1)
        hb.addWidget(bt)

        self.ed_tools = QLineEdit("")
        self.ed_tools.setPlaceholderText("可选：指定本地 Luban 工具目录，创建时复制到工程的 Tools")
        bt_tools = QPushButton("选择目录…")
        bt_tools.clicked.connect(self._pick_tools)
        hb_tools = QHBoxLayout()
        hb_tools.addWidget(self.ed_tools, 1)
        hb_tools.addWidget(bt_tools)

        self.cb_xlsx = QCheckBox("Excel Schema（Schemas/__tables__.xlsx、__beans__.xlsx、"
                                 "__enums__.xlsx）")
        self.cb_xlsx.setChecked(True)
        self.cb_xml = QCheckBox("XML Schema（Defines/tables.xml、beans.xml、enums.xml，可并存）")
        self.cb_xlsx.stateChanged.connect(self._guard)
        self.cb_xml.stateChanged.connect(self._guard)

        if creating:
            info = QLabel(
                "将在所选目录下创建：luban.conf、gen.bat、gen.sh、\n"
                "Schemas 目录（xlsx 定义文件）、Defines 目录（xml 定义文件）、"
                "Data 目录（数据表）。\n"
                "已有同名文件不会被覆盖。创建完成后会自动打开该目录。")
        else:
            info = QLabel(
                "在所选目录里补上 Luban 配置：luban.conf、gen.bat、gen.sh、\n"
                "Schemas 目录（xlsx 定义文件）、Defines 目录（xml 定义文件）、"
                "Data 目录（数据表）。\n"
                "已有同名文件不会被覆盖，目录里原有的表和数据都会保留。\n"
                "完成后会自动打开该目录。")
        info.setStyleSheet("color:#64748B;")
        info.setWordWrap(True)

        # 只保留一个「创建」按钮（去掉「关闭」），居中放大
        self.bt_ok = QPushButton("创建")
        self.bt_ok.setObjectName("primary")
        self.bt_ok.setMinimumSize(160, 40)
        self.bt_ok.setCursor(Qt.PointingHandCursor)
        self.bt_ok.clicked.connect(self._ok)

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("工程根目录"))
        lay.addLayout(hb)
        lay.addSpacing(8)
        lay.addWidget(QLabel("Luban 工具目录（Tools，可选）"))
        lay.addLayout(hb_tools)
        lay.addSpacing(8)
        gb = QGroupBox("Schema 文件类型")
        gl = QVBoxLayout(gb)
        gl.addWidget(self.cb_xlsx)
        gl.addWidget(self.cb_xml)
        lay.addWidget(gb)
        lay.addSpacing(6)
        lay.addWidget(info)
        lay.addSpacing(6)
        link = QLabel('<a href="https://www.datable.cn/docs/guide/install">安装与获取工具</a>')
        link.setOpenExternalLinks(True)
        link.setAlignment(Qt.AlignCenter)
        lay.addWidget(link)
        lay.addSpacing(10)
        hbox_bt = QHBoxLayout()
        hbox_bt.addStretch(1)
        hbox_bt.addWidget(self.bt_ok)
        hbox_bt.addStretch(1)
        lay.addLayout(hbox_bt)

    def _guard(self):
        pass

    def _pick(self):
        d = QFileDialog.getExistingDirectory(self, "选择工程根目录", self.ed_root.text() or "")
        if d:
            self.ed_root.setText(d)

    def _pick_tools(self):
        d = QFileDialog.getExistingDirectory(self, "选择 Luban 工具目录", self.ed_tools.text() or "")
        if d:
            self.ed_tools.setText(d)

    def _ok(self):
        root = self.ed_root.text().strip()
        if not root:
            QMessageBox.warning(self, "缺少目录", "请先选择工程根目录")
            return
        types = []
        if self.cb_xlsx.isChecked():
            types.append("xlsx")
        if self.cb_xml.isChecked():
            types.append("xml")
        if not types:
            QMessageBox.warning(self, "缺少选项", "请至少勾选一种 Schema 文件类型")
            return
        tools = self.ed_tools.text().strip()
        if not tools:
            if QMessageBox.question(
                    self, "未指定 Tools",
                    "没有指定 Luban 工具目录（Tools）。\n"
                    "创建后工程的 Tools 目录会留空，导表前需手动放入 Luban 工具。\n\n"
                    "仍要创建吗？") != QMessageBox.Yes:
                return
        self.accept()

    def result(self) -> Tuple[str, List[str], str]:
        types = []
        if self.cb_xlsx.isChecked():
            types.append("xlsx")
        if self.cb_xml.isChecked():
            types.append("xml")
        return self.ed_root.text().strip(), types, self.ed_tools.text().strip()


# ------------------------------------------------------------------ 发布工程
class ExeModeSwitch(QWidget):
    """「执行方式」二选一：dll（dotnet 跑）/ exe（直接跑）。

    为什么单开一个控件：它不是命令行参数，只决定「可执行文件」怎么填、
    命令前面要不要补 `dotnet`。两个单选框比下拉更直观，也不会被当成可手输文本。
    """

    changed = Signal(str)

    def __init__(self, parent=None, mode: str = "dll"):
        super().__init__(parent)
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(14)
        self.grp = QButtonGroup(self)
        self.grp.setExclusive(True)
        self.rb_dll = QRadioButton("dll（dotnet 运行 Luban.dll）")
        self.rb_exe = QRadioButton("exe（直接运行 Luban.exe）")
        self.rb_dll.setToolTip('dotnet 版：命令写成 dotnet "Luban.dll" …，跨平台通用')
        self.rb_exe.setToolTip('Windows 独立版：命令直接写 "Luban.exe" …')
        for rb in (self.rb_dll, self.rb_exe):
            self.grp.addButton(rb)
            h.addWidget(rb)
        h.addStretch(1)
        self.grp.buttonClicked.connect(lambda _b: self.changed.emit(self.mode()))
        self.set_mode(mode)

    def mode(self) -> str:
        return "exe" if self.rb_exe.isChecked() else "dll"

    def set_mode(self, mode: str):
        rb = self.rb_exe if str(mode).strip().lower() == "exe" else self.rb_dll
        rb.blockSignals(True)
        rb.setChecked(True)
        rb.blockSignals(False)


class PublishDialog(QDialog):
    """发布工程：把 gen.bat / gen.sh 的参数摊开成一张表，改完直接跑脚本。

    三个区域：
    1. **顶部标题** —— 一句话说清当前脚本与工程目录；
    2. **参数区**（可滚动）—— 基础参数（照官方模板脚本里的那几个）默认展开，
       高级参数默认折叠，点「更多」展开；每个参数都带一句「干什么用的」备注，
       悬停还能看完整说明；文件 / 目录类参数既能手输也能点「选择…」挑；
    3. **命令预览 + 发布** —— 实时看到会跑什么；勾了「覆盖」就把脚本写回工程
       根目录的 gen.bat / gen.sh（写前备份），没勾就生成临时脚本跑完删掉。
    """

    def __init__(self, parent, model: ProjectModel, settings: Dict[str, str]):
        super().__init__(parent)
        self.model = model
        self.settings = settings or {}
        self.setWindowTitle("发布工程")
        self.resize(1000, 720)

        self.kind = publish.preferred_kind()
        self._script_path = os.path.join(
            model.root, "gen." + publish.script_ext(self.kind))
        self._params = publish.defaults(
            model, self.settings, publish.read_script(model.root, self.kind),
            self.kind)
        self._ensure_output_dirs()
        self._fields: Dict[str, QWidget] = {}

        v = QVBoxLayout(self)
        v.setContentsMargins(14, 12, 14, 12)
        v.setSpacing(9)

        # ---------------- 区域 1：顶部标题 ----------------
        head = QFrame()
        head.setObjectName("dlghead")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(12, 8, 12, 8)
        t = QLabel("发布工程　＝　跑一次 Luban 导表脚本")
        t.setObjectName("dlgtitle")
        hl.addWidget(t)
        sub = QLabel(f"脚本：gen.{publish.script_ext(self.kind)}　·　工程：{model.root}")
        sub.setObjectName("hint")
        hl.addWidget(sub)
        hl.addStretch(1)
        b_view = QPushButton("查看现有脚本")
        b_view.setToolTip("显示工程里现在这份 gen 脚本的原文")
        b_view.clicked.connect(self._view_script)
        hl.addWidget(b_view)
        v.addWidget(head)

        # ---------------- 区域 2：参数（可滚动） ----------------
        host = QWidget()
        hv = QVBoxLayout(host)
        hv.setContentsMargins(0, 0, 0, 0)
        hv.setSpacing(9)

        gb_b = QGroupBox("基础参数（官方模板脚本里的那几个）")
        fb = QFormLayout(gb_b)
        fb.setContentsMargins(12, 10, 12, 10)
        fb.setSpacing(6)
        for key in publish.BASIC_KEYS:
            self._add_param(fb, publish.PARAM_BY_KEY[key])
        hv.addWidget(gb_b)

        self.bt_more = QPushButton("▸ 更多（高级参数）")
        self.bt_more.setObjectName("linkbtn")
        self.bt_more.setToolTip("官方 CLI 的其余开关与常用 -x 扩展项，平时不用动")
        self.bt_more.clicked.connect(self._toggle_more)
        hv.addWidget(self.bt_more, 0, Qt.AlignLeft)

        self.gb_a = QGroupBox("高级参数（官方文档 · CLI 全参数 / 常用 -x）")
        fa = QFormLayout(self.gb_a)
        fa.setContentsMargins(12, 10, 12, 10)
        fa.setSpacing(6)
        last = None
        for key in publish.ADV_KEYS:
            p = publish.PARAM_BY_KEY[key]
            if p.group and p.group != last:
                if last is not None:
                    sp = QLabel("")
                    sp.setMinimumHeight(2)
                    fa.addRow(sp)
                cap = QLabel(p.group)
                cap.setObjectName("cap")
                fa.addRow(cap)
            last = p.group
            self._add_param(fa, p)
        self.gb_a.hide()
        hv.addWidget(self.gb_a)
        hv.addStretch(1)

        sc = QScrollArea()
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        sc.setWidget(host)
        v.addWidget(sc, 1)

        # ---------------- 区域 3：命令预览 ----------------
        self.lab_output_warn = QLabel("")
        self.lab_output_warn.setStyleSheet("color:#B42318; font-size:12px;")
        self.lab_output_warn.setWordWrap(True)
        v.addWidget(self.lab_output_warn)
        gb_cmd = QGroupBox("命令预览（点「发布」就执行它）")
        cl = QHBoxLayout(gb_cmd)
        cl.setContentsMargins(12, 10, 12, 10)
        self.ed_cmd = QLineEdit()
        self.ed_cmd.setReadOnly(True)
        self.ed_cmd.setStyleSheet('font-family:"Consolas","Cascadia Mono",monospace;')
        cl.addWidget(self.ed_cmd, 1)
        b_copy = QPushButton("复制命令")
        b_copy.clicked.connect(self._copy_cmd)
        cl.addWidget(b_copy)
        v.addWidget(gb_cmd)

        self.ck_overwrite = QCheckBox(
            "发布时覆盖写入工程根目录的 gen.%s（不勾选＝在工程根目录生成 gen_时间戳.%s 运行，"
            "成功后问你是否删掉）" % (publish.script_ext(self.kind),
                                publish.script_ext(self.kind)))
        self.ck_overwrite.setToolTip(
            "工程里通常已经有 gen 脚本了。勾上 = 用现在的参数重写它（自动备份），再跑它；\n"
            "不勾 = 在工程根目录另存一份 gen_时间戳 的临时脚本再跑，原脚本原样保留，\n"
            "导表成功后会问你要不要删掉这份临时脚本（失败则保留，方便双击看报错）。")
        v.addWidget(self.ck_overwrite)

        hbtn = QHBoxLayout()
        hbtn.addStretch(1)
        b_close = QPushButton("关闭")
        b_close.setObjectName("foot")
        b_close.clicked.connect(self.reject)
        self.bt_run = QPushButton("发布")
        self.bt_run.setObjectName("footprimary")
        self.bt_run.setToolTip("直接调用系统命令执行脚本（Windows 就是 cmd /c gen.%s）"
                               % publish.script_ext(self.kind))
        self.bt_run.clicked.connect(self._run)
        hbtn.addWidget(b_close)
        hbtn.addWidget(self.bt_run)
        hbtn.addStretch(1)
        v.addLayout(hbtn)

        self._sync_all()
        self._refresh_cmd()

    # ---------------------------------------------------------------- 控件
    def _choices_of(self, p: publish.ParamDef) -> List[str]:
        """-t 的选项取自本工程 luban.conf 的 targets（再加一个 all）。"""
        if p.key == "target":
            items = [t for t in self.model.targets if t]
            if "all" not in [x.lower() for x in items]:
                items.append("all")
            return items
        return list(p.choices)

    def _add_param(self, form: QFormLayout, p: publish.ParamDef):
        lab = QLabel(p.label)
        lab.setToolTip(p.meaning)
        cell = QWidget()
        col = QVBoxLayout(cell)
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(2)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)
        if p.kind == "flag":
            w: QWidget = QCheckBox("启用")
        elif p.kind == "mode":
            w = ExeModeSwitch(mode=self._params.get(p.key, "dll"))
        elif p.kind == "combo":
            w = QComboBox()
            w.setEditable(True)
            w.addItems([""] + self._choices_of(p))
        else:
            w = QLineEdit()
            w.setPlaceholderText(p.placeholder)
        row.addWidget(w, 1)
        if p.kind in ("file", "dir"):
            b = QPushButton("选择…")
            b.setMaximumWidth(62)
            b.clicked.connect(lambda _=False, pp=p, ww=w: self._pick_path(pp, ww))
            row.addWidget(b)
        if p.key == "exe":
            b2 = QPushButton("查找")
            b2.setMaximumWidth(62)
            b2.setToolTip("去工程目录里找 Luban.dll / Luban.exe（也会往下一层子目录找）")
            b2.clicked.connect(self._find_exe)
            row.addWidget(b2)
        col.addLayout(row)

        hint = QLabel(p.meaning)
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#64748B; font-size:11px;")
        col.addWidget(hint)
        form.addRow(lab, cell)
        self._fields[p.key] = w
        w.setToolTip(p.meaning)     # 悬停输入框也能看到「这是干什么的」

        if p.kind == "mode":
            w.changed.connect(self._on_exe_mode)
        elif isinstance(w, QCheckBox):
            w.stateChanged.connect(self._refresh_cmd)
        elif isinstance(w, QComboBox):
            w.currentTextChanged.connect(self._refresh_cmd)
        else:
            w.textChanged.connect(self._refresh_cmd)

    # ------------------------------------------------------- 可执行文件
    def _exe_mode(self) -> str:
        w = self._fields.get("exeMode")
        return w.mode() if isinstance(w, ExeModeSwitch) else "dll"

    def _on_exe_mode(self, mode: str = ""):
        """换执行方式：顺手把「可执行文件」换成对得上的那个。"""
        mode = mode or self._exe_mode()
        w = self._fields.get("exe")
        if w is not None and isinstance(w, QLineEdit):
            w.setPlaceholderText("Luban.dll" if mode == "dll" else "Luban.exe")
            cur = _widget_text(w).strip().lower()
            # 路径后缀和当前执行方式对不上（或还空着）→ 重新去工程目录找
            if not cur or not cur.endswith("." + mode):
                self._find_exe(silent=True)
        self._refresh_cmd()

    def _find_exe(self, silent: bool = False):
        """在工程目录里找**当前执行方式对应**的 Luban 并填进「可执行文件」。

        只找当前这一种：用户亲手选了 exe 就别再把 dll 塞回去（那会让人以为
        切换没生效）。找不到就保持原样，由「选择…」手动指定。
        """
        mode = self._exe_mode()
        found = publish.find_exe(self.model.root, mode)
        w = self._fields.get("exe")
        if not found:
            if not silent:
                QMessageBox.information(
                    self, "没找到",
                    "工程目录（含下一层子目录）里没有 Luban.dll / Luban.exe。\n"
                    "可以点「选择…」手动指定，或去工具栏「设置」里填一次。")
            return
        _set_widget_text(w, found)
        # 找到的是另一种形式就同步切换执行方式，避免「选 dll 却填了 exe」
        m2 = publish.exe_mode_of(found)
        if m2 and m2 != mode:
            ms = self._fields.get("exeMode")
            if isinstance(ms, ExeModeSwitch):
                ms.set_mode(m2)
        self._refresh_cmd()
        if not silent:
            self.toast_note("已找到 %s" % found)

    def toast_note(self, msg: str):
        p = self.parent()
        if p is not None and hasattr(p, "toast"):
            p.toast(msg)
        else:
            self.setWindowTitle("发布工程　·　" + msg)

    def _pick_path(self, p: publish.ParamDef, w: QWidget):
        cur = _widget_text(w)
        base = cur if (cur and os.path.isdir(cur)) else (self.model.root or ".")
        if p.kind == "dir":
            d = QFileDialog.getExistingDirectory(self, f"选择{p.label}", base)
        else:
            if p.key == "exe":
                ext = "dll" if self._exe_mode() == "dll" else "exe"
                flt = f"Luban.{ext} (*.{ext});;可执行文件 (*.exe *.dll);;所有文件 (*.*)"
            else:
                flt = "所有文件 (*.*)"
            d, _ = QFileDialog.getOpenFileName(self, f"选择{p.label}", base, flt)
        if d:
            if p.key == "exe":
                d = self._display_exe(d)   # 工程内 → 相对路径；工程外 → 全路径
            _set_widget_text(w, d)
            if p.key == "exe":
                m2 = publish.exe_mode_of(d)
                ms = self._fields.get("exeMode")
                if m2 and isinstance(ms, ExeModeSwitch) and m2 != self._exe_mode():
                    ms.set_mode(m2)
            self._refresh_cmd()

    def _display_exe(self, p: str) -> str:
        """执行文件路径：在当前工程里 → 相对路径/文件名；在别处 → 全路径。"""
        p = (p or "").strip()
        root = os.path.abspath(self.model.root or "")
        if not p or not root or not os.path.isabs(p):
            return p
        try:
            if os.path.normcase(os.path.commonpath(
                    [os.path.abspath(p), root])) == os.path.normcase(root):
                return os.path.relpath(os.path.abspath(p), root).replace("\\", "/")
        except Exception:
            pass
        return p

    def _ensure_output_dirs(self):
        """确保默认输出目录（工程目录下 Output/Code、Output/Data）存在。"""
        root = self.model.root or ""
        if not root or not os.path.isdir(root):
            return
        for key in ("outputCodeDir", "outputDataDir"):
            d = (self._params.get(key) or "").strip()
            if d and not os.path.isabs(d):
                try:
                    os.makedirs(os.path.join(root, d.replace("/", os.sep)),
                                exist_ok=True)
                except Exception:
                    pass

    def _toggle_more(self):
        on = not self.gb_a.isVisible()
        self.gb_a.setVisible(on)
        self.bt_more.setText("▾ 收起高级参数" if on else "▸ 更多（高级参数）")

    # ---------------------------------------------------------------- 数据
    def _sync_all(self):
        for k, w in self._fields.items():
            v = self._params.get(k, "")
            if isinstance(w, ExeModeSwitch):
                w.set_mode(v or "dll")
            elif isinstance(w, QCheckBox):
                w.setChecked(str(v or "").strip().lower() in ("1", "true", "yes"))
            else:
                _set_widget_text(w, v)

    def _collect(self) -> Dict[str, str]:
        out: Dict[str, str] = {}
        for k, w in self._fields.items():
            if isinstance(w, ExeModeSwitch):
                out[k] = w.mode()
            elif isinstance(w, QCheckBox):
                out[k] = "1" if w.isChecked() else ""
            else:
                out[k] = _widget_text(w).strip()
        return out

    def _refresh_cmd(self, *_a):
        self._params = self._collect()
        self.ed_cmd.setText(publish.build_command(self._params))
        self._update_output_warn()

    def _update_output_warn(self):
        cd = (self._params.get("outputCodeDir") or "").strip()
        dd = (self._params.get("outputDataDir") or "").strip()
        if cd and dd and os.path.normcase(os.path.normpath(cd)) == \
                os.path.normcase(os.path.normpath(dd)):
            self.lab_output_warn.setText(
                "⚠ 代码输出目录与数据输出目录相同，导表时后写的一方会覆盖先写的一方，"
                "请改成不同的目录。")
        else:
            self.lab_output_warn.clear()

    # ---------------------------------------------------------------- 动作
    def _copy_cmd(self):
        QApplication.clipboard().setText(self.ed_cmd.text())
        p = self.parent()
        if p is not None and hasattr(p, "toast"):
            p.toast("导表命令已复制到剪贴板")

    def _view_script(self):
        text = publish.read_script(self.model.root, self.kind)
        d = QDialog(self)
        d.setWindowTitle(f"gen.{publish.script_ext(self.kind)} 原文")
        d.resize(760, 520)
        lay = QVBoxLayout(d)
        ed = QTextEdit()
        ed.setReadOnly(True)
        ed.setPlainText(text or "（工程里还没有这个脚本）")
        ed.setStyleSheet('font-family:"Consolas","Cascadia Mono",monospace; font-size:12px;')
        lay.addWidget(ed, 1)
        hb = QHBoxLayout()
        hb.addStretch(1)
        b = QPushButton("关闭")
        b.setObjectName("foot")
        b.clicked.connect(d.reject)
        hb.addWidget(b)
        hb.addStretch(1)
        lay.addLayout(hb)
        d.exec()

    def _run(self):
        """生成脚本 →（可选）覆盖写入 → 运行 → 展示输出。"""
        p = self._collect()
        if not p.get("exe"):
            QMessageBox.warning(
                self, "还没指定 Luban",
                "先在上面「执行文件」里选好 Luban.exe 或 Luban.dll，再点「发布」。\n"
                "也可以点旁边的「查找」，或去工具栏「设置」里填一次，之后每次都会带出来。")
            return
        # 代码 / 数据输出目录相同（且都非空）→ 二次确认，防止数据被覆盖
        cd = (p.get("outputCodeDir") or "").strip()
        dd = (p.get("outputDataDir") or "").strip()
        if cd and dd and os.path.normcase(os.path.normpath(cd)) == \
                os.path.normcase(os.path.normpath(dd)):
            if QMessageBox.question(
                    self, "输出目录相同",
                    "代码输出目录和数据输出目录是同一个：\n%s\n\n"
                    "导表时后写的一方会覆盖先写的一方。仍要继续吗？" % cd
                    ) != QMessageBox.Yes:
                return
        # 回写到设置：工具栏「生成并校验」走的是同一套参数
        publish.sync_settings(p, self.settings)
        pp = self.parent()
        if pp is not None and hasattr(pp, "save_settings"):
            pp.save_settings()
        ext = publish.script_ext(self.kind)
        enc = "gbk" if self.kind == "bat" else "utf-8"
        nl = "\r\n" if self.kind == "bat" else "\n"
        temp = ""
        if self.ck_overwrite.isChecked():
            # ① 覆盖：直接写回工程根目录的 gen.bat / gen.sh，然后就跑它
            target = self._script_path
            bak = ""
            if os.path.exists(target):
                try:
                    bak = exceldef.backup_any(target)
                except Exception as e:
                    QMessageBox.warning(self, "备份失败",
                                        f"{e}\n已中止，没有改动脚本。")
                    return
            text = publish.build_script(p, self.kind, self.model.root,
                                        portable=True)
            try:
                with open(target, "w", encoding=enc, newline=nl) as f:
                    f.write(text)
            except Exception as e:
                QMessageBox.warning(self, "写入脚本失败", str(e))
                return
            note = "已覆盖写入 %s%s" % (os.path.basename(target),
                                   ("（备份：" + bak + "）") if bak else "（新建）")
        else:
            # ② 不覆盖：在工程根目录生成一个带时间戳的临时脚本再跑
            #    （放在工程目录里，`%~dp0` 天然指向工程，脚本不用改写路径）
            target = os.path.join(
                self.model.root,
                "gen_%s.%s" % (time.strftime("%Y%m%d_%H%M%S"), ext))
            text = publish.build_script(p, self.kind, self.model.root,
                                        portable=True)
            try:
                with open(target, "w", encoding=enc, newline=nl) as f:
                    f.write(text)
            except Exception as e:
                QMessageBox.warning(self, "生成临时脚本失败", str(e))
                return
            note = "未覆盖：已在工程根目录生成临时脚本 %s（原脚本没动）" \
                   % os.path.basename(target)
            temp = target

        # 直接调系统命令执行脚本（Windows 就是 cmd /c gen.bat，等于双击它）
        self.bt_run.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            QApplication.processEvents()
            res = runner.run_script(target, self.model.root)
        finally:
            QApplication.restoreOverrideCursor()
            self.bt_run.setEnabled(True)

        issues = runner.parse_issues(res["output"], self.model)
        p = self.parent()
        if p is not None and hasattr(p, "_show_issues"):
            p._show_issues(issues or [{"level": "info", "table": "",
                                       "msg": "无报错", "hint": ""}])
        if res["ok"]:
            if temp:
                self._ask_remove_temp(temp)
            self.accept()
            if p is not None and hasattr(p, "toast"):
                p.toast("发布完成：导表成功")
            return
        box = QMessageBox(self)
        box.setWindowTitle("发布失败")
        box.setIcon(QMessageBox.Warning)
        box.setText(f"导表失败（退出码 {res['code']}）")
        box.setInformativeText(note + ("；临时脚本保留着，可以直接双击它看报错"
                                       if temp else ""))
        box.setDetailedText(f"$ {res['cmd']}\n\n{res['output'][:6000]}")
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.button(QMessageBox.Yes).setText("关闭")
        box.button(QMessageBox.No).setText("返回修改参数")
        box.setDefaultButton(QMessageBox.No)
        if box.exec() == QMessageBox.Yes:
            self.reject()

    def _ask_remove_temp(self, path: str):
        """跑成功了才问：临时脚本要不要删掉。"""
        box = QMessageBox(self)
        box.setWindowTitle("发布完成")
        box.setIcon(QMessageBox.Question)
        box.setText("导表成功")
        box.setInformativeText("是否删除临时脚本 %s？" % os.path.basename(path))
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.button(QMessageBox.Yes).setText("删除")
        box.button(QMessageBox.No).setText("保留")
        box.setDefaultButton(QMessageBox.Yes)
        if box.exec() != QMessageBox.Yes:
            return
        try:
            os.remove(path)
        except Exception as e:
            QMessageBox.warning(self, "删除失败", str(e))


# ------------------------------------------------------------------ 自定义类型
class CreateTypeDialog(QDialog):
    def __init__(self, parent, model: ProjectModel, current_table: Optional[str]):
        super().__init__(parent)
        self.model = model
        self.current_table = current_table
        self.setWindowTitle("新建自定义类型")
        self.resize(680, 560)

        self.rb_bean = QRadioButton("结构 bean（复合类型，有多个字段）")
        self.rb_enum = QRadioButton("枚举 enum（只能填固定的几项）")
        # 两组单选框各自独立：以前它们共用一个隐式互斥组，点「归属」会把
        # 上面的「结构 / 枚举」取消掉，而且初始能同时选中两个。
        self.grp_kind = QButtonGroup(self)
        self.grp_kind.addButton(self.rb_bean)
        self.grp_kind.addButton(self.rb_enum)
        self.rb_bean.setChecked(True)
        self.rb_bean.toggled.connect(self._switch)
        self.rb_enum.toggled.connect(self._switch)
        hg = QHBoxLayout()
        hg.addWidget(self.rb_bean)
        hg.addWidget(self.rb_enum)
        hg.addStretch(1)

        self.ed_name = QLineEdit()
        self.ed_name.setPlaceholderText("英文，如 Reward / Quality（建议大驼峰）")
        self.ed_comment = QLineEdit()
        self.ed_comment.setPlaceholderText("中文说明，会显示在下拉框里")
        self.rb_common = QRadioButton("公共（所有表都能用）")
        self.rb_table = QRadioButton("当前表专用")
        self.grp_scope = QButtonGroup(self)
        self.grp_scope.addButton(self.rb_common)
        self.grp_scope.addButton(self.rb_table)
        self.rb_common.setChecked(True)
        if current_table:
            self.rb_table.setText(f"当前表专用（{current_table}）")
        else:
            # 别静默灰掉 —— 说清楚为什么不能选、怎么才能选
            self.rb_table.setText("当前表专用（先把表名填了）")
            self.rb_table.setEnabled(False)
            self.rb_table.setToolTip(
                "还没有表名：在「新建配置表」页填好表名，或先选中一张已存在的表，"
                "就能建只属于该表的类型。")
        hs = QHBoxLayout()
        hs.addWidget(self.rb_common)
        hs.addWidget(self.rb_table)
        hs.addStretch(1)

        # bean 区
        self.w_bean = QWidget()
        vl = QVBoxLayout(self.w_bean)
        hl1 = QHBoxLayout()
        self.cb_parent = QComboBox()
        self.cb_parent.setEditable(True)
        self.cb_parent.addItem("（无父类）", "")
        self.cb_parent.addItem("", "")  # 分隔占位
        for b in sorted(model.beans, key=lambda x: x.full_name):
            self.cb_parent.addItem(typecatalog.type_label(model, b.full_name), b.full_name)
        self.ed_sep = QLineEdit(",")
        self.ed_sep.setMaximumWidth(70)
        hl1.addWidget(QLabel("父类型"))
        hl1.addWidget(self.cb_parent, 1)
        hl1.addWidget(QLabel("分隔符 sep"))
        hl1.addWidget(self.ed_sep)
        vl.addLayout(hl1)
        self.tbl = QTableWidget(0, 4)
        self.tbl.setHorizontalHeaderLabels(["字段名", "类型", "分组", "注释"])
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        bt_add = QPushButton("+ 添加字段")
        bt_del = QPushButton("- 删除")
        bt_add.clicked.connect(lambda: self._add_row())
        bt_del.clicked.connect(self._del_row)
        hb = QHBoxLayout()
        hb.addWidget(bt_add)
        hb.addWidget(bt_del)
        hb.addStretch(1)
        vl.addWidget(self.tbl)
        vl.addLayout(hb)

        # enum 区
        self.w_enum = QWidget()
        el = QVBoxLayout(self.w_enum)
        self.cb_flags = QCheckBox("标志位枚举（可多选，填 A|B）")
        el.addWidget(self.cb_flags)
        self.tbl_item = QTableWidget(0, 4)
        self.tbl_item.setHorizontalHeaderLabels(["项名", "别名", "值", "说明"])
        self.tbl_item.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        bt_add2 = QPushButton("+ 添加枚举项")
        bt_del2 = QPushButton("- 删除")
        bt_add2.clicked.connect(lambda: self._add_item())
        bt_del2.clicked.connect(self._del_item)
        hb2 = QHBoxLayout()
        hb2.addWidget(bt_add2)
        hb2.addWidget(bt_del2)
        hb2.addStretch(1)
        el.addWidget(self.tbl_item)
        el.addLayout(hb2)

        note = QLabel(
            "提示：选「公共」写进 Defines/beans.xml / enums.xml，所有表共用；"
            "选「当前表专用」则写到 Defines/<u>以该表命名</u>的 xml（如 TbEquip.xml），"
            "并自动登记进 luban.conf，不会污染公共定义。"
            "写入前都会自动备份。")
        note.setStyleSheet("color:#64748B;")
        note.setWordWrap(True)

        # 底部按钮：居中 + 放大（放右边的小按钮新手容易点错）
        self.bt_ok = QPushButton("确定")
        self.bt_ok.setObjectName("footprimary")
        self.bt_ok.setToolTip("按上面的定义创建这个自定义类型")
        self.bt_ok.setDefault(True)
        self.bt_ok.clicked.connect(self._ok)
        self.bt_cancel = QPushButton("取消")
        self.bt_cancel.setObjectName("foot")
        self.bt_cancel.setToolTip("放弃，不创建任何东西")
        self.bt_cancel.clicked.connect(self.reject)
        hbtn = QHBoxLayout()
        hbtn.addStretch(1)
        hbtn.addWidget(self.bt_cancel)
        hbtn.addWidget(self.bt_ok)
        hbtn.addStretch(1)

        lay = QVBoxLayout(self)
        lay.addLayout(hg)
        form = QFormLayout()
        form.addRow("类型名称 *", self.ed_name)
        form.addRow("中文说明", self.ed_comment)
        form.addRow("归属", hs)
        lay.addLayout(form)
        lay.addWidget(self.w_bean, 1)
        lay.addWidget(self.w_enum, 1)
        lay.addWidget(note)
        lay.addSpacing(4)
        lay.addLayout(hbtn)
        self._switch()
        self._add_row()
        self._add_item()

    def _switch(self):
        is_bean = self.rb_bean.isChecked()
        self.w_bean.setVisible(is_bean)
        self.w_enum.setVisible(not is_bean)

    def _type_combo(self, value="int") -> QComboBox:
        cb = QComboBox()
        cb.setEditable(True)
        cb.addItems(_type_names(self.model, self.current_table))
        cb.setCurrentText(value)
        return cb

    def _group_combo(self) -> QWidget:
        """bean 字段的「分组」格：与主字段表同一套带中文名的下拉。"""
        cell = QWidget()
        lay = QHBoxLayout(cell)
        lay.setContentsMargins(2, 0, 2, 0)
        cb = ValueCombo(group_items(self.model), editable=True)
        cb.setToolTip("该字段导出到哪些分组；留空 = 各端都导出，多个用逗号：c,s")
        lay.addWidget(cb, 1)
        return cell

    def _add_row(self):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        self.tbl.setItem(r, 0, QTableWidgetItem("field_%d" % (r + 1)))
        self.tbl.setCellWidget(r, 1, self._type_combo())
        self.tbl.setCellWidget(r, 2, self._group_combo())
        self.tbl.setItem(r, 3, QTableWidgetItem(""))

    def _del_row(self):
        for r in sorted({i.row() for i in self.tbl.selectedIndexes()}, reverse=True):
            self.tbl.removeRow(r)

    def _add_item(self):
        r = self.tbl_item.rowCount()
        self.tbl_item.insertRow(r)
        self.tbl_item.setItem(r, 0, QTableWidgetItem("ITEM_%d" % (r + 1)))
        self.tbl_item.setItem(r, 1, QTableWidgetItem(""))
        self.tbl_item.setItem(r, 2, QTableWidgetItem(str(r + 1)))
        self.tbl_item.setItem(r, 3, QTableWidgetItem(""))

    def _del_item(self):
        for r in sorted({i.row() for i in self.tbl_item.selectedIndexes()}, reverse=True):
            self.tbl_item.removeRow(r)

    def _ok(self):
        name = self.ed_name.text().strip()
        if not name:
            QMessageBox.warning(self, "缺少名称", "请填写类型名称")
            return
        if not re.match(r"^[A-Za-z][A-Za-z0-9_]*$", name):
            QMessageBox.warning(self, "名称不合规",
                                "类型名只能用英文字母、数字、下划线，且以字母开头")
            return
        if self.model.find_bean(name) or self.model.find_enum(name):
            QMessageBox.warning(self, "重名", f"已存在同名类型：{name}")
            return
        self.accept()

    def table_scope(self) -> bool:
        """用户选的是「当前表专用」吗（=类型名要带模块前缀）。"""
        return self.rb_table.isChecked() and bool(self.current_table)

    def full_name(self) -> str:
        name = self.ed_name.text().strip()
        if self.rb_table.isChecked() and self.current_table:
            t = self.model.find_table(self.current_table)
            # 表可能还没创建（新建表页里建的），这时按表名推模块，
            # 与 exceldef.derive_scopes 的规则一致，否则重开会认成公共类型
            mod = t.module if t else TableDef(full_name=self.current_table).module
            return f"{mod}.{name}"
        return name

    def build(self):
        if self.rb_bean.isChecked():
            b = BeanDef(full_name=self.full_name(),
                        parent=self.cb_parent.currentData() or "",
                        sep=self.ed_sep.text().strip(),
                        comment=self.ed_comment.text().strip())
            for r in range(self.tbl.rowCount()):
                n = self.tbl.item(r, 0)
                cb = self.tbl.cellWidget(r, 1)
                gcell = self.tbl.cellWidget(r, 2)
                gcombo = gcell.findChild(QComboBox) if gcell is not None else None
                c = self.tbl.item(r, 3)
                nm = (n.text().strip() if n else "")
                if not nm:
                    continue
                b.fields.append(FieldDef(
                    name=nm,
                    type=typecatalog.strip_label(cb.currentText()).strip() if cb else "int",
                    group=((gcombo.code() if gcombo is not None else "") or "").strip(),
                    comment=(c.text().strip() if c else "")))
            return b
        e = EnumDef(full_name=self.full_name(),
                    flags="true" if self.cb_flags.isChecked() else "false",
                    unique="true",
                    comment=self.ed_comment.text().strip())
        for r in range(self.tbl_item.rowCount()):
            n = self.tbl_item.item(r, 0)
            if not n or not n.text().strip():
                continue
            e.items.append(EnumItem(
                name=n.text().strip(),
                alias=(self.tbl_item.item(r, 1).text().strip()
                       if self.tbl_item.item(r, 1) else ""),
                value=(self.tbl_item.item(r, 2).text().strip()
                       if self.tbl_item.item(r, 2) else str(r + 1)),
                comment=(self.tbl_item.item(r, 3).text().strip()
                         if self.tbl_item.item(r, 3) else "")))
        return e


def container_stub(name: str) -> str:
    """从「类型」下拉选了容器后的初始类型串（子类型先给一个能跑通的默认值）。"""
    if name == "map":
        return "map,int,string"
    return f"{name},int"


def _type_choices(model: ProjectModel, current_table: Optional[str],
                  containers: bool = True) -> List[Tuple[str, str]]:
    """(代码, 显示名) 列表：基础类型 + 容器 + 公共/当前表的枚举与结构。

    「类型」下拉与「子类型」下拉共用这一份，保证两处看到的东西一致。
    containers=False 时不含容器（容器不能再套容器，Luban 不支持 list,list,int）。
    """
    # 表名可能是手填的简写（新建表页还没落盘），先归一到真正的 full_name，
    # 否则「本表专用」的类型在这一页的下拉里看不到
    t = model.find_table(current_table) if current_table else None
    scope_key = t.full_name if t else current_table
    out: List[Tuple[str, str]] = [
        (t2["name"], typecatalog.type_label(model, t2["name"]))
        for t2 in typecatalog.BASE_TYPES]
    if containers:
        for c in typecatalog.CONTAINERS:
            code = container_stub(c["name"])
            out.append((code, typecatalog.type_label(model, code)))
    for e in model.enums:
        if e.scope == "common" or e.scope == scope_key:
            out.append((e.full_name, typecatalog.type_label(model, e.full_name)))
    for b in model.beans:
        if b.scope == "common" or b.scope == scope_key:
            out.append((b.full_name, typecatalog.type_label(model, b.full_name)))
    return out


def _type_names(model: ProjectModel, current_table: Optional[str]) -> List[str]:
    """下拉框类型列表：基础类型 + 容器 + 公共/当前表的枚举与结构。"""
    return [label for _code, label in _type_choices(model, current_table)]


# 字段表的列定义（编辑器与新建表页共用，改这里两边一起变）
FIELD_COLS = ["字段名", "类型", "子类型", "配置方式", "单元格格式", "分组", "注释", "助手"]
(COL_NAME, COL_TYPE, COL_SUB, COL_STYLE, COL_FMT,
 COL_GROUP, COL_COMMENT, COL_HELP) = range(8)
FIELD_WIDTHS = (130, 165, 200, 190, 118, 72, 130, 52)

FMT_ITEMS = [
    ("", "stream（默认·流式）"),
    ("lite", "lite（花括号）"),
    ("json", "json（单元格内 JSON）"),
    ("lua", "lua（单元格内 Lua 表）"),
]


# ------------------------------------------------------------------ 类型助手
ADD_NEW_TYPE = "＋ 新建自定义类型…"

# 各类类型可选用的标签 —— 与「配置方式」下拉共用 typecatalog 里的同一份定义
TAGS_BY_KIND: Dict[str, List[str]] = typecatalog.TAGS_BY_KIND

KIND_CN: Dict[str, str] = {
    "atom": "基础类型", "enum": "枚举", "bean": "结构",
    "poly": "多态结构", "container": "容器", "map": "映射容器", "unknown": "未知类型",
}


class TypePickerDialog(QDialog):
    """类型助手：四个区域 —— 顶部标题、左侧类型、右上常用标签、右下配置示例。

    - 常用标签**随所选类型变化**（基础类型给校验标签、容器给 sep/size…）；
    - 类型列表含通用的 **bean / 多态bean**（不依赖具体工程，纯讲解），
      并列出本工程里可用的结构与枚举；
    - 不再在列表里放「新建自定义类型」。
    """

    def __init__(self, parent, model: ProjectModel, field_name: str,
                 current: str, current_table: Optional[str]):
        super().__init__(parent)
        self.model = model
        self.field_name = field_name or "x"
        self.current_table = current_table
        self._updating = False
        self._vis_keys: List[str] = []
        self._null_vis = True
        self.setWindowTitle("类型助手")
        self.resize(980, 672)

        self.ed = QLineEdit(current)
        self.ed.setToolTip("就是字段「类型」那一格的内容，可以手写")
        self.lst = QListWidget()
        self.view = QTextBrowser()

        # ---------------- 区域 1：顶部标题 ----------------
        head = QFrame()
        head.setObjectName("dlghead")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(12, 8, 12, 8)
        ttl = QLabel("类型助手")
        ttl.setObjectName("dlgtitle")
        hl.addWidget(ttl)
        sub = QLabel("① 选左侧类型　② 右上勾标签　③ 右下看示例")
        sub.setObjectName("hint")
        hl.addWidget(sub)
        hl.addStretch(1)
        hl.addWidget(QLabel("类型字符串"))
        self.ed.setMinimumWidth(280)
        hl.addWidget(self.ed, 1)

        # ---------------- 区域 3：右上 · 常用标签 ----------------
        self.gb_tag = QGroupBox("常用标签（随所选类型变化）")
        tl = QVBoxLayout(self.gb_tag)
        self.lab_tag_hint = QLabel("")
        self.lab_tag_hint.setObjectName("hint")
        self.lab_tag_hint.setWordWrap(True)
        tl.addWidget(self.lab_tag_hint)
        self.ck_null = QCheckBox("可空 ?（留空 / null 表示没有值；容器不支持）")
        self.ck_null.setToolTip("写在类型最后：int -> int?")
        tl.addWidget(self.ck_null)
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        self.tag_cks: Dict[str, QCheckBox] = {}
        self.tag_eds: Dict[str, QLineEdit] = {}
        row = 0
        for t in typecatalog.TAGS:
            ck = QCheckBox(t["cn"])
            ck.setToolTip(f"{t['desc']}\n写在：{t['where']}\n注意：{t['note']}")
            self.tag_cks[t["key"]] = ck
            grid.addWidget(ck, row, 0)
            if t.get("editable") == "1":
                ed = QLineEdit(t.get("default", ""))
                ed.setMaximumWidth(170)
                ed.setPlaceholderText(t["form"].split("=")[-1])
                ed.setVisible(False)
                self.tag_eds[t["key"]] = ed
                grid.addWidget(ed, row, 1)
            row += 1
        grid.setColumnStretch(2, 1)
        tl.addLayout(grid)
        tl.addStretch(1)

        # ---------------- 区域 4：右下 · 配置示例 ----------------
        gb_ex = QGroupBox("配置示例")
        el = QVBoxLayout(gb_ex)
        el.setContentsMargins(10, 8, 10, 10)
        el.addWidget(self.view, 1)

        # ---------------- 区域 2：左侧 · 类型 ----------------
        gb_type = QGroupBox("类型")
        tpl = QVBoxLayout(gb_type)
        tpl.setContentsMargins(10, 8, 10, 10)
        tpl.addWidget(self.lst, 1)
        pick = QLabel("选一个类型；也可直接在顶部手写类型字符串")
        pick.setObjectName("hint")
        pick.setWordWrap(True)
        tpl.addWidget(pick)

        # ---------------- 组装 ----------------
        right = QSplitter(Qt.Vertical)
        right.addWidget(self.gb_tag)
        right.addWidget(gb_ex)
        right.setStretchFactor(0, 0)
        right.setStretchFactor(1, 1)
        right.setSizes([250, 330])
        right.setChildrenCollapsible(False)

        mid = QSplitter(Qt.Horizontal)
        mid.addWidget(gb_type)
        mid.addWidget(right)
        mid.setStretchFactor(0, 4)
        mid.setStretchFactor(1, 6)
        mid.setSizes([350, 600])
        mid.setChildrenCollapsible(False)

        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.button(QDialogButtonBox.Ok).setText("确定")
        box.button(QDialogButtonBox.Cancel).setText("关闭")
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)
        lay.addWidget(head)
        lay.addWidget(mid, 1)
        lay.addWidget(box)

        self._fill_list(current)

        self.lst.itemClicked.connect(self._pick)
        self.lst.itemDoubleClicked.connect(self._dbl)
        self.ed.textChanged.connect(self._on_text_changed)
        self.ck_null.stateChanged.connect(self._rebuild)
        for ck in self.tag_cks.values():
            ck.stateChanged.connect(self._on_tag_toggle)
        for ed in self.tag_eds.values():
            ed.textChanged.connect(self._rebuild)

        self._apply_tag_visibility()
        self._sync_controls(current)
        self._sync_tag_editors()
        self._refresh()

    # ---------------- 区域 2：类型列表 ----------------
    def _add_group(self, title: str):
        it = QListWidgetItem(title)
        it.setFlags(Qt.NoItemFlags)
        f = it.font()
        f.setBold(True)
        it.setFont(f)
        it.setForeground(QColor("#64748B"))
        it.setData(Qt.UserRole, {"group": True})
        self.lst.addItem(it)

    def _add_item(self, text: str, code: str, kind: str, base: str, tip: str = ""):
        it = QListWidgetItem("   " + text)
        it.setData(Qt.UserRole, {"code": code, "kind": kind, "base": base})
        if tip:
            it.setToolTip(tip)
        self.lst.addItem(it)

    def _fill_list(self, current: str):
        self.lst.clear()
        self._add_group("基础类型")
        for t in typecatalog.BASE_TYPES:
            self._add_item(f"{t['name']}（{t['cn']}）", t["name"], "atom", t["name"],
                           f"{t['name']}｜{t['cn']}｜填法：{t['fill']}")
        self._add_group("容器")
        for c in typecatalog.CONTAINERS:
            self._add_item(f"{c['name']}<T>（{c['cn']}容器）", f"{c['name']},T",
                           "container", c["name"], c["note"])
        self._add_group("结构")
        self._add_item("bean（通用结构）", "bean", "bean", "bean",
                       "通用讲解：结构（bean）是什么、单元格怎么填")
        self._add_item("多态bean（多态结构）", "bean", "poly", "bean",
                       "通用讲解：父结构 + 子结构，单元格里写具体子类名")
        beans = [b for b in self.model.beans
                 if b.scope == "common" or b.scope == self.current_table]
        if beans:
            self._add_group("结构（本工程）")
            for b in beans:
                k = "poly" if self.model.is_polymorphic(b.full_name) else "bean"
                self._add_item(typecatalog.type_label(self.model, b.full_name),
                               b.full_name, k, b.full_name,
                               f"{b.full_name}｜{b.comment or ''}")
        enums = [e for e in self.model.enums
                 if e.scope == "common" or e.scope == self.current_table]
        if enums:
            self._add_group("枚举（本工程）")
            for e in enums:
                self._add_item(typecatalog.type_label(self.model, e.full_name),
                               e.full_name, "enum", e.full_name,
                               f"{e.full_name}｜{e.comment or ''}")
        # 选中与当前类型串匹配的那一项
        want = typecatalog.strip_label(current or "")
        for i in range(self.lst.count()):
            d = self.lst.item(i).data(Qt.UserRole) or {}
            if d.get("code") and d["code"] == want:
                self.lst.setCurrentItem(self.lst.item(i))
                break

    def _pick(self, item: QListWidgetItem):
        d = item.data(Qt.UserRole) or {}
        code = d.get("code") or ""
        if not code:
            return
        if code.endswith(",T"):
            cur = self.ed.text().strip()
            elem = "int"
            m = re.match(r"^(list|array|set|map)\s*,\s*(.+)$", cur)
            if m:
                elem = m.group(2).split(",")[0]
            code = f"{code[:-2]},{elem}"
            if code.startswith("map,"):
                code = "map,int,string"
        self.ed.setText(code)
        self._sync_controls(code)          # 让标签勾选跟着类型走
        self._apply_tag_visibility()
        self._sync_tag_editors()
        self._refresh()

    def _dbl(self, item: QListWidgetItem):
        d = item.data(Qt.UserRole) or {}
        if d.get("code"):
            self.accept()

    # ---------------- 类型判定 ----------------
    def _kind_of(self, t: str) -> Tuple[str, str]:
        s = (t or "").strip()
        if not s:
            return "atom", ""
        p = typecatalog.parse_type(s)
        if p["kind"] == "container":
            return ("map" if p["container"] == "map" else "container"), p["container"]
        base = p["base"]
        if base == "bean":
            return "bean", base
        if base in typecatalog.BASE_TYPE_NAMES:
            return "atom", base
        if self.model.find_enum(base):
            return "enum", base
        b = self.model.find_bean(base)
        if b:
            return ("poly" if self.model.is_polymorphic(b.full_name) else "bean"), base
        return "unknown", base

    def _eff_kind(self) -> Tuple[str, str]:
        """优先采用左侧选中项的语义（通用 bean / 多态bean 靠它区分）。"""
        t = self.ed.text().strip()
        it = self.lst.currentItem()
        d = (it.data(Qt.UserRole) if it is not None else None) or {}
        if d.get("code") == t and d.get("kind"):
            return d["kind"], d.get("base") or t
        return self._kind_of(t)

    def _tags_for(self, kind: str, base: str) -> List[str]:
        return typecatalog.tags_for_kind(kind, base)

    def _apply_tag_visibility(self):
        """常用标签随类型变化：只显示这一类型真的用得上的标签。"""
        kind, base = self._eff_kind()
        self._vis_keys = self._tags_for(kind, base)
        self._updating = True
        for key, ck in self.tag_cks.items():
            ck.setVisible(key in self._vis_keys)
        self._null_vis = kind in ("atom", "enum", "bean", "poly")
        self.ck_null.setVisible(self._null_vis)
        self._updating = False
        self.lab_tag_hint.setText(
            f"当前：{KIND_CN.get(kind, kind)}"
            + (f"（{base}）" if base else "")
            + "　—— 下面只列出该类型可用的标签。")

    def _sync_tag_editors(self):
        for key, ed in self.tag_eds.items():
            ck = self.tag_cks.get(key)
            ed.setVisible(key in self._vis_keys and bool(ck and ck.isChecked()))

    # ---------------- 区域 3：标签 ----------------
    def _sync_controls(self, t: str):
        _core, tags = typecatalog.split_tags(t)
        p = typecatalog.parse_type(t)
        self._updating = True
        self.ck_null.setChecked(p["nullable"])
        for key, ck in self.tag_cks.items():
            on = False
            if key == "!":
                on = "!" in tags
            else:
                for tg in tags:
                    if tg.startswith("#" + key + "=") or tg == "#" + key:
                        on = True
                        if key in self.tag_eds:
                            self.tag_eds[key].setText(tg.split("=", 1)[1])
                        break
            ck.setChecked(on)
        self._updating = False

    def _on_tag_toggle(self, *_):
        self._sync_tag_editors()
        self._rebuild()

    def _on_text_changed(self, _=None):
        self._apply_tag_visibility()
        self._sync_controls(self.ed.text())
        self._sync_tag_editors()
        self._refresh()

    def _rebuild(self, *_):
        if getattr(self, "_updating", False):
            return
        core, _tags = typecatalog.split_tags(self.ed.text())
        tags: List[str] = []
        for t in typecatalog.TAGS:
            key = t["key"]
            if key not in self._vis_keys:
                continue
            if not self.tag_cks[key].isChecked():
                continue
            if key == "!":
                tags.append("!")
            elif key in self.tag_eds:
                v = (self.tag_eds[key].text() or "").strip() \
                    or t.get("default", "") or t["form"].split("=", 1)[-1]
                tags.append(f"#{key}={v}")
            else:
                tags.append(t["form"])
        cur = self.ed.text()
        new = typecatalog.join_type(core, tags,
                                    nullable=self.ck_null.isChecked() and self._null_vis)
        if new != cur:
            self._updating = True
            pos = self.ed.cursorPosition()
            self.ed.setText(new)
            self.ed.setCursorPosition(min(pos, len(new)))
            self._updating = False
        self._refresh()

    # ---------------- 区域 4：配置示例 ----------------
    def _refresh(self):
        kind, base = self._eff_kind()
        t = self.ed.text().strip()
        if kind == "poly" or (kind == "bean" and base == "bean"):
            self.view.setHtml(self._generic_guide("poly" if kind == "poly" else "bean"))
            return
        h = f"<h4>{_esc(t)}</h4>"
        if not t:
            h += "<p>先选左侧的类型，或在顶部手写类型字符串。</p>"
            self.view.setHtml(h)
            return
        issues = typecatalog.validate_type(self.model, t)
        if issues:
            h += "<p style='color:#c00'>问题：" + "；".join(
                f"[{i['level']}] {_esc(i['msg'])}" for i in issues) + "</p>"
        opts = typecatalog.fill_options(self.model, self.field_name, t)
        h += options_html(opts)
        if not opts and not issues:
            h += "<p>这个类型没有额外的示例。</p>"
        self.view.setHtml(h)

    def _generic_guide(self, kind: str) -> str:
        """通用的 bean / 多态bean 讲解：不依赖具体工程，纯说明 + 示例表。"""
        if kind == "poly":
            body = (
                "<h4>多态bean（多态结构）</h4>"
                "<p>多态用来表达「这个字段可能是几种形状里的任意一种」。"
                "分三步：</p>"
                "<ol>"
                "<li>在 <code>__beans__.xlsx</code> 里定义<b>父结构</b>（可以只有名字，没有字段）；</li>"
                "<li>再定义若干<b>子结构</b>，把它们的 <code>parent</code> 列填成父结构名；</li>"
                "<li>字段的<b>类型写父结构名</b>，数据单元格里写<b>具体子结构名 + 参数</b>。</li>"
                "</ol>"
                "<p><b>单元格写法</b>：<code>子结构名,参数1,参数2</code>（分隔符取父结构的 sep）</p>"
                + _html_table([
                    ["##var", "id", "effect"],
                    ["##type", "int", "Effect"],
                    ["##", "ID", "效果（多态）"],
                    ["", "1", "Damage,100"],
                    ["", "2", "Heal,50"],
                ])
                + "<ul>"
                "<li>类型写父结构名；写具体子结构名会丢掉多态能力</li>"
                "<li>父结构若也定义了字段，那些字段会成为所有子结构共有的前缀</li>"
                "<li>子结构与父结构同名会报「定义重复」</li>"
                "<li>可选空时写 <code>Effect?</code>，子类内部仍不可空</li>"
                "</ul>")
            return body
        return (
            "<h4>通用结构 bean</h4>"
            "<p>结构（bean）＝「一行里的一组字段」，在 <code>__beans__.xlsx</code> 里定义。"
            "字段类型填结构名后，这一列在数据表里会展开成若干子列，"
            "或压成一个格子用 <code>sep</code> 拆开。</p>"
            "<p><b>真正使用时</b>：把下面的 <code>bean</code> 换成你工程里的结构名，"
            "例如 <code>vec3</code>、<code>item.TbItem</code>。"
            "公共结构写全名；表专用结构写 <code>模块.结构名</code>。</p>"
            "<p><b>常见填法</b></p>"
            + _html_table([
                ["##var", "id", ("pos", 2)],
                ["##var", "", "x", "y"],
                ["##type", "int", ("vec3", 2)],
                ["##", "ID", "坐标 x", "坐标 y"],
                ["", "1", "10", "20"],
            ])
            + "<ul>"
            "<li>结构名不存在时 Luban 报「未知类型」，先去 __beans__.xlsx 定义</li>"
            "<li>结构里字段的顺序 = 流式读取顺序，别随意调整</li>"
            "<li>字段多、表太宽时，改「配置方式」为单格 sep 更省列</li>"
            "</ul>")

    def value(self) -> str:
        return self.ed.text().strip()

    def accept(self):
        """通用 bean / 多态bean 只是讲解，不是真类型名 —— 用它之前先提醒一次。"""
        kind, _base = self._eff_kind()
        if self.ed.text().strip() == "bean" and kind in ("bean", "poly"):
            if QMessageBox.question(
                    self, "这是通用说明",
                    "「bean / 多态bean」只是通用讲解，不是你工程里的真实类型名。\n"
                    "请把它换成 __beans__.xlsx 里的结构名，例如 vec3、item.TbItem。\n\n"
                    "仍要返回类型字符串 “bean”？") != QMessageBox.Yes:
                return
        super().accept()


def referenced_types(type_str: str) -> List[str]:
    """类型串里直接引用到的类型名（剥掉容器关键字与 `#tag`、外层括号）。

    `map,string,(Reward#sep=,)` -> `["string", "Reward"]`。
    基础类型（`string` 之类）也照样返回，由调用方去自定义类型表里查，
    查不到就不是自定义类型 —— 这样统计不会漏掉容器里的元素类型。
    """
    out: List[str] = []

    def walk(s: str, depth: int = 0):
        if depth > 8 or not (s or "").strip():
            return
        p = typecatalog.parse_type(s)
        if p["kind"] == "container":
            for sub in (p["elem"], p["value"]):
                if sub:
                    walk(sub, depth + 1)
            return
        base = typecatalog.strip_label(p["base"] or "")
        if base:
            out.append(base)

    walk(type_str or "")
    return out


def _type_scope(model: ProjectModel, full_name: str) -> str:
    b = model.find_bean(full_name)
    if b:
        return b.scope or "common"
    e = model.find_enum(full_name)
    if e:
        return e.scope or "common"
    return "common"


# ------------------------------------------------------------------ 管理自定义类型
class ManageTypesDialog(QDialog):
    """管理自定义类型：按「所有 / 公共 / 某张表」查看，并删除没人用的类型。

    三个区域：
    1. 顶部标题 —— 一句话说清怎么看、删了会怎样；
    2. 左侧范围 —— 所有 / 公共 / 每张配置表（后面带该范围的类型数）；
    3. 右侧列表 —— 类型、引用次数、引用位置、删除按钮。

    「引用次数」不是拍脑袋：真的去扫整个工程的每一张配置表 —— 表头字段
    （含容器元素类型）、每个结构定义自己的字段、以及表声明里的「行结构类型」。
    为 0 标红，因为那意味着删掉它不会让任何一张表变坏。

    引用位置的显示规则：公共类型显示「表名.字段名」；表专用类型默认只显示
    字段名（归属表已经由左侧确定了，再写一遍是噪音），但它被**别的表**引用时
    会把表名补回来并标「跨表」—— 这种情况往往是用错了，必须让人看见。
    """

    changed = Signal()      # 删过东西了，主窗口要重新读定义文件

    def __init__(self, parent, model: ProjectModel,
                 fields: Dict[str, List[FieldDef]]):
        super().__init__(parent)
        self.model = model
        self.fields = fields or {}
        self._entries: List[Dict[str, Any]] = []
        self._refs: Dict[str, List[Tuple[str, str, bool]]] = {}
        self.setWindowTitle("管理自定义类型")
        self.resize(1020, 680)
        self._build()
        self._rebuild()

    # ---------------------------------------------------------------- 界面
    def _build(self):
        head = QFrame()
        head.setObjectName("dlghead")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(12, 8, 12, 8)
        t = QLabel("管理自定义类型")
        t.setObjectName("dlgtitle")
        hl.addWidget(t)
        sub = QLabel("① 左侧选范围　② 右侧看「引用次数」　"
                     "③ 引用次数为 0（红色）＝ 没人用，可以安全删")
        sub.setObjectName("hint")
        hl.addWidget(sub)
        hl.addStretch(1)

        self.lst = QListWidget()
        self.lst.setMinimumWidth(200)
        self.lst.setMaximumWidth(340)
        gb_l = QGroupBox("范围")
        ll = QVBoxLayout(gb_l)
        ll.setContentsMargins(10, 8, 10, 10)
        ll.addWidget(self.lst, 1)
        tip = QLabel("「公共」＝ 写在 Defines/beans.xml / enums.xml 里的类型；"
                     "表专用类型单独放在 Defines/&lt;表名&gt;.xml。")
        tip.setObjectName("hint")
        tip.setWordWrap(True)
        ll.addWidget(tip)

        self.ck_unused = QCheckBox("只看没被引用的（引用次数 = 0）")
        self.ck_unused.setToolTip("勾上就只剩可以直接删掉的类型，清理垃圾定义很好用")
        self.ck_unused.stateChanged.connect(lambda _s: self._fill())

        self.tbl = QTableWidget(0, 4)
        self.tbl.setHorizontalHeaderLabels(["类型", "引用次数", "引用位置", "操作"])
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setWordWrap(True)
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        hh.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        gb_r = QGroupBox("自定义类型")
        rl = QVBoxLayout(gb_r)
        rl.setContentsMargins(10, 8, 10, 10)
        rl.addWidget(self.tbl, 1)
        rl.addWidget(self.ck_unused)

        self.lab_sum = QLabel("")
        self.lab_sum.setObjectName("hint")
        self.lab_sum.setWordWrap(True)

        sp = QSplitter(Qt.Horizontal)
        sp.addWidget(gb_l)
        sp.addWidget(gb_r)
        sp.setStretchFactor(0, 3)
        sp.setStretchFactor(1, 7)
        sp.setSizes([290, 700])
        sp.setChildrenCollapsible(False)

        hb = QHBoxLayout()
        hb.addStretch(1)
        b_ref = QPushButton("重新扫描")
        b_ref.setObjectName("foot")
        b_ref.setToolTip("改过字段后点一下，重新统计引用次数")
        b_close = QPushButton("关闭")
        b_close.setObjectName("foot")
        b_ref.clicked.connect(self._rebuild)
        b_close.clicked.connect(self.reject)
        hb.addWidget(b_ref)
        hb.addWidget(b_close)
        hb.addStretch(1)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)
        lay.addWidget(head)
        lay.addWidget(sp, 1)
        lay.addWidget(self.lab_sum)
        lay.addLayout(hb)

        self.lst.currentItemChanged.connect(lambda *_a: self._fill())

    # ---------------------------------------------------------------- 数据
    def _collect(self) -> List[Dict[str, Any]]:
        ents: List[Dict[str, Any]] = []
        for b in self.model.beans:
            if not b.full_name:
                continue
            ents.append({"kind": "bean", "name": b.full_name, "scope": b.scope,
                         "file": b._file, "comment": b.comment,
                         "extra": f"{len(b.fields)} 个字段"})
        for e in self.model.enums:
            if not e.full_name:
                continue
            ents.append({"kind": "enum", "name": e.full_name, "scope": e.scope,
                         "file": e._file, "comment": e.comment,
                         "extra": f"{len(e.items)} 个枚举项"})
        ents.sort(key=lambda x: (0 if x["scope"] == "common" else 1,
                                 x["scope"], x["kind"], x["name"]))
        return ents

    def _scan(self) -> Dict[str, List[Tuple[str, str, bool]]]:
        """扫整个工程，统计每个自定义类型被哪些「表名 + 字段名」用到。

        只记**直接引用**（字段类型串 / 表结构类型里写到的名字）：结构内部
        字段引用了谁，由结构自己的定义负责记一次，这样不会同一个引用被
        重复计两遍。
        """
        m = self.model
        names = {b.full_name for b in m.beans} | {e.full_name for e in m.enums}
        short: Dict[str, List[str]] = {}
        for n in names:
            short.setdefault(n.split(".")[-1], []).append(n)

        def resolve(ref: str) -> Optional[str]:
            if ref in names:
                return ref
            c = short.get(ref, [])
            return c[0] if len(c) == 1 else None

        refs: Dict[str, List[Tuple[str, str, bool]]] = {}

        def note(ref: str, where: str, path: str):
            full = resolve(ref)
            if not full:
                return
            scope = _type_scope(m, full)
            # 表专用类型被**别的表**（或公共结构）引用 = 跨表，得把表名写出来
            cross = scope not in ("", "common") and where != scope
            refs.setdefault(full, []).append((where, path, cross))

        def note_field(where: str, prefix: str, f: FieldDef):
            for ref in referenced_types(f.type):
                note(ref, where, prefix + (f.name or "?"))

        def same_as_bean(t: "TableDef") -> bool:
            """表的字段是不是「抄」自它的行结构类型。

            read_schema_from_file=false 的表，字段来自 __beans__.xlsx；打开工程时
            工具会把这批字段也读进 self.fields。两边都扫就会把同一个引用数两遍，
            所以字段与结构完全一致时只算结构那一份。
            """
            b = m.find_bean(t.value_type)
            fs = self.fields.get(t.full_name) or []
            if not (b and b.fields and fs):
                return False
            return ([(f.name, f.type) for f in fs]
                    == [(f.name, f.type) for f in b.fields])

        for t in m.tables:
            if not t.full_name:
                continue
            if t.value_type:
                note(t.value_type, t.full_name, "<表结构>")
            if not same_as_bean(t):
                for f in (self.fields.get(t.full_name) or []):
                    note_field(t.full_name, "", f)
        for b in m.beans:
            where = b.scope if (b.scope and b.scope != "common") else "公共"
            for f in b.fields:
                note_field(where, b.short_name + ".", f)
        return refs

    def _count(self, key: str) -> int:
        return len(self._pick(key))

    def _pick(self, key: str) -> List[Dict[str, Any]]:
        if key == "all":
            return list(self._entries)
        if key == "common":
            return [e for e in self._entries if e["scope"] == "common"]
        return [e for e in self._entries if e["scope"] == key]

    # ---------------------------------------------------------------- 左侧
    def _add_group(self, title: str):
        it = QListWidgetItem(title)
        it.setFlags(Qt.NoItemFlags)
        f = it.font()
        f.setBold(True)
        it.setFont(f)
        it.setForeground(QColor("#64748B"))
        self.lst.addItem(it)

    def _add_scope(self, key: str, label: str):
        it = QListWidgetItem(f"   {label}  （{self._count(key)}）")
        it.setData(Qt.UserRole, key)
        self.lst.addItem(it)

    def _fill_left(self):
        cur = self._cur_key()
        self.lst.blockSignals(True)
        self.lst.clear()
        self._add_group("范围")
        self._add_scope("all", "所有")
        self._add_scope("common", "公共")
        self._add_group("配置表")
        for t in sorted((x for x in self.model.tables if x.full_name),
                        key=lambda x: x.full_name):
            self._add_scope(t.full_name, t.full_name)
        self.lst.blockSignals(False)
        idx = self._index_of(cur)
        self.lst.setCurrentRow(idx if idx >= 0 else 0)

    def _index_of(self, key: str) -> int:
        for i in range(self.lst.count()):
            it = self.lst.item(i)
            if it is not None and it.data(Qt.UserRole) == key:
                return i
        return -1

    def _cur_key(self) -> str:
        it = self.lst.currentItem()
        if it is None:
            return "all"
        return it.data(Qt.UserRole) or "all"

    # ---------------------------------------------------------------- 右侧
    def _ref_texts(self, e: Dict[str, Any]) -> List[str]:
        """引用位置文案：公共类型带表名，表专用类型只在跨表时带表名。"""
        out: List[str] = []
        for where, path, cross in self._refs.get(e["name"], []):
            if path == "<表结构>":
                # 表的「行结构类型」不是某个字段，单独一种写法，别显示成 TbItem.<表结构>
                txt = f"{where}（表结构）"
            elif e["scope"] == "common" or cross:
                txt = f"{where}.{path}" if where else path
            else:
                txt = path
            if cross:
                txt += "（跨表）"
            out.append(txt)
        return out

    def _fill(self):
        key = self._cur_key()
        ents = self._pick(key)
        if self.ck_unused.isChecked():
            ents = [e for e in ents if not self._refs.get(e["name"])]
        self.tbl.setRowCount(0)
        self.tbl.setRowCount(len(ents))
        unused = 0
        for r, e in enumerate(ents):
            refs = self._refs.get(e["name"], [])
            n = len(refs)
            if n == 0:
                unused += 1
            tag = "结构" if e["kind"] == "bean" else "枚举"
            it = QTableWidgetItem(f"[{tag}] {e['name']}")
            tip = ""
            if e["comment"]:
                tip += e["comment"]
            if e["file"]:
                tip += ("\n" if tip else "") + "定义文件：" + os.path.basename(e["file"])
            if e["extra"]:
                tip += ("\n" if tip else "") + e["extra"]
            it.setToolTip(tip)
            self.tbl.setItem(r, 0, it)

            itn = QTableWidgetItem(str(n))
            itn.setTextAlignment(Qt.AlignCenter)
            if n == 0:
                itn.setForeground(QColor("#B42318"))
                f = itn.font()
                f.setBold(True)
                itn.setFont(f)
            self.tbl.setItem(r, 1, itn)

            texts = self._ref_texts(e)
            itp = QTableWidgetItem("；".join(texts) if texts else "（没有字段引用它）")
            itp.setForeground(QColor("#1F2937") if texts else QColor("#64748B"))
            itp.setToolTip("\n".join(texts) if texts else "可以直接删掉")
            self.tbl.setItem(r, 2, itp)

            bt = QPushButton("删除")
            bt.setObjectName("danger")
            bt.setToolTip("删除这个自定义类型（会先弹确认框）")
            bt.clicked.connect(lambda _c=False, ee=e: self._del(ee))
            self.tbl.setCellWidget(r, 3, bt)
        self.tbl.resizeRowsToContents()

        scope = {"all": "所有", "common": "公共"}.get(key, key)
        self.lab_sum.setText(
            f"范围「{_esc(scope)}」共 {len(ents)} 个自定义类型，"
            f"其中 <b>{unused}</b> 个没有被任何配置表引用（引用次数 0，可安全删除）。"
            f"引用位置里，公共类型显示「表名.字段名」；表专用类型只在跨表引用时补表名。")

    # ---------------------------------------------------------------- 删除
    def _del(self, e: Dict[str, Any]):
        kind_cn = "结构" if e["kind"] == "bean" else "枚举"
        refs = self._refs.get(e["name"], [])
        # 先挡掉「删了会连带坏一片」的情况：还有子结构继承它
        if e["kind"] == "bean":
            subs = [b.full_name for b in self.model.beans
                    if (b.parent or "").strip() == e["name"]]
            if subs:
                QMessageBox.warning(
                    self, "删不了",
                    f"{e['name']} 还被 {len(subs)} 个子结构继承着：\n"
                    + "\n".join("  · " + s for s in subs)
                    + "\n\n先把这些子结构的「父结构」改掉或删掉，再来删它。")
                return
        lines = self._ref_texts(e)
        box = QMessageBox(self)
        box.setWindowTitle("删除确认")
        box.setIcon(QMessageBox.Warning if refs else QMessageBox.Question)
        box.setText(f"确定删除{kind_cn} {e['name']} 吗？")
        if refs:
            box.setInformativeText(
                f"它有 <b>{len(refs)}</b> 处引用。删除后这些字段会指向一个不存在的类型，"
                "下次导表会直接报错，需要你手动把字段类型改掉。")
            box.setDetailedText("\n".join(lines))
        else:
            box.setInformativeText(
                "目前没有任何字段引用它，删除是安全的。<br>删除不可撤销"
                "（工具会在 _lubankit_backup 里留一份备份）。")
        box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        box.button(QMessageBox.Yes).setText("删除")
        box.button(QMessageBox.No).setText("取消")
        box.setDefaultButton(QMessageBox.No)
        if box.exec() != QMessageBox.Yes:
            return
        if refs:
            # 引用不为 0 时再拦一道，避免手滑把正在用的类型删了
            box2 = QMessageBox(self)
            box2.setWindowTitle("再确认一次")
            box2.setIcon(QMessageBox.Critical)
            box2.setText(f"{e['name']} 正在被 {len(refs)} 处引用。")
            box2.setInformativeText("真的要删吗？相关配置表会立刻变成「导不出来」。")
            box2.setDetailedText("\n".join(lines))
            box2.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            box2.button(QMessageBox.Yes).setText("仍然删除")
            box2.button(QMessageBox.No).setText("取消")
            box2.setDefaultButton(QMessageBox.No)
            if box2.exec() != QMessageBox.Yes:
                return
        ok, msg = (exceldef.delete_bean(self.model, e["name"]) if e["kind"] == "bean"
                   else exceldef.delete_enum(self.model, e["name"]))
        if not ok:
            QMessageBox.warning(self, "删除失败", msg)
            return
        if e["kind"] == "bean":
            self.model.beans = [b for b in self.model.beans
                                if b.full_name != e["name"]]
        else:
            self.model.enums = [x for x in self.model.enums
                                if x.full_name != e["name"]]
        self.changed.emit()
        self._rebuild()
        QMessageBox.information(self, "已删除", msg)

    def _rebuild(self):
        exceldef.derive_scopes(self.model)
        self._entries = self._collect()
        self._refs = self._scan()
        self._fill_left()
        self._fill()


# ------------------------------------------------------------------ 设置
class SettingsDialog(QDialog):
    def __init__(self, parent, settings: Dict[str, str], model: ProjectModel):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setMinimumWidth(600)
        self.settings = settings
        self.model = model
        self.ed_exe = QLineEdit(settings.get("luban_exe", ""))
        bt = QPushButton("浏览…")
        bt.clicked.connect(self._browse)
        hb = QHBoxLayout()
        hb.addWidget(self.ed_exe, 1)
        hb.addWidget(bt)
        w = QWidget()
        w.setLayout(hb)
        self.ed_conf = QLineEdit(settings.get("conf_path", "") or model.conf_path)
        self.cb_target = QComboBox()
        self.cb_target.setEditable(True)
        self.cb_target.addItems([""] + model.targets)
        self.cb_target.setCurrentText(settings.get("target", ""))
        self.ed_extra = QLineEdit(settings.get("extra_args", ""))
        self.ed_cmd = QTextEdit(settings.get("cmd_template") or runner.DEFAULT_CMD)
        self.ed_cmd.setFixedHeight(70)
        self.ed_work = QLineEdit(settings.get("work_dir", ""))
        form = QFormLayout()
        form.addRow("Luban 可执行文件", w)
        form.addRow("luban.conf", self.ed_conf)
        form.addRow("导出目标 -t", self.cb_target)
        form.addRow("附加参数", self.ed_extra)
        form.addRow("命令模板", self.ed_cmd)
        form.addRow("工作目录", self.ed_work)
        lab = QLabel("占位符：{exe} {conf} {target} {extra} {data_dir} {root}\n"
                     "不确定参数时，把 gen.bat 里的命令抄进来，只替换占位符即可。")
        lab.setWordWrap(True)
        lab.setStyleSheet("color:#64748B;")
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.button(QDialogButtonBox.Ok).setText("确定")
        box.button(QDialogButtonBox.Cancel).setText("关闭")
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        lay = QVBoxLayout(self)
        lay.addLayout(form)
        lay.addWidget(lab)
        lay.addWidget(box)

    def _browse(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 Luban.exe", "",
                                           "可执行文件 (*.exe *.dll);;所有文件 (*.*)")
        if p:
            self.ed_exe.setText(p)

    def apply(self):
        s = self.settings
        s["luban_exe"] = self.ed_exe.text().strip()
        s["conf_path"] = self.ed_conf.text().strip()
        s["target"] = self.cb_target.currentText().strip()
        s["extra_args"] = self.ed_extra.text().strip()
        s["cmd_template"] = self.ed_cmd.toPlainText().strip()
        s["work_dir"] = self.ed_work.text().strip()
        return s


# ------------------------------------------------------------------ 主窗口
class KnowledgeDialog(QDialog):
    """知识点浏览器（顶部工具栏「知识点」打开）。

    三块：搜索框 / 条目列表 / 正文。以前是右侧的一个页签 —— 那样既占着
    右侧本就不宽的地方，又只有在右侧那一栏才看得见。搬到工具条之后是个独立
    窗口，可以随手开、随手关，正文里还能点开官方文档链接。

    打开时可以用 `seed()` 塞一串关键词进来（主窗口会根据当前表的字段类型
    预设，比如 sep / 多态 / 可空），省得自己想该搜什么。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("知识点")
        self.resize(860, 640)
        # 非模态：看知识点的时候还想在左边继续改表，别把主窗口锁住
        self.setWindowModality(Qt.NonModal)

        head = QFrame()
        head.setObjectName("dlghead")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(12, 8, 12, 8)
        t = QLabel("知识点")
        t.setObjectName("dlgtitle")
        hl.addWidget(t)
        sub = QLabel("搜规则 / 术语 / 踩坑经验；点右下「清空」看全部")
        sub.setObjectName("hint")
        hl.addWidget(sub)
        hl.addStretch(1)

        self.ed_search = QLineEdit()
        self.ed_search.setPlaceholderText("搜索知识点：如 sep / 多态 / 可空 / group / 表头")
        # 列表区分两栏：核心知识 / 常见问题
        self.lst_core = QListWidget()
        self.lst_faq = QListWidget()
        self.tabs_k = QTabWidget()
        self.tabs_k.setMaximumWidth(340)
        self.tabs_k.addTab(self.lst_core, "核心知识")
        self.tabs_k.addTab(self.lst_faq, "常见问题")
        self.view = QTextBrowser()
        self.view.setOpenExternalLinks(True)
        bt_all = QPushButton("清空")
        bt_all.setToolTip("清空搜索条件，列出全部知识点")
        bt_all.clicked.connect(self.ed_search.clear)
        bt_close = QPushButton("关闭")
        bt_close.clicked.connect(self.close)

        self.ed_search.textChanged.connect(self._refresh)
        self.lst_core.currentItemChanged.connect(self._show)
        self.lst_faq.currentItemChanged.connect(self._show)

        hl2 = QHBoxLayout()
        hl2.setContentsMargins(0, 0, 0, 0)
        hl2.addWidget(self.ed_search, 1)
        hl2.addWidget(bt_all)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(6)
        rl.addLayout(hl2)
        rl.addWidget(self.view, 1)

        sp = QSplitter(Qt.Horizontal)
        sp.addWidget(self.tabs_k)
        sp.addWidget(right)
        sp.setStretchFactor(0, 3)
        sp.setStretchFactor(1, 7)
        sp.setChildrenCollapsible(False)

        foot = QDialogButtonBox()
        foot.addButton(bt_close, QDialogButtonBox.RejectRole)

        lay = QVBoxLayout(self)
        lay.addWidget(head)
        lay.addWidget(sp, 1)
        lay.addWidget(foot)

    # ---------------------------------------------------------------- 数据
    def seed(self, query: str):
        """打开前塞一串关键词（由当前选中表的字段类型推导出来）。"""
        self.ed_search.setText(query or "")
        # setText 的内容和原来一样时 textChanged 不会发，列表就空着 —— 直接刷一次
        self._refresh()

    def focus_entry(self, kid: str):
        """直接跳到某一条知识点（问题报表里的「看知识点」用它）。"""
        k = knowledge.by_id(kid)
        if k:
            self.ed_search.setText(k["title"])

    def _refresh(self):
        hits = knowledge.search(self.ed_search.text())
        core = [k for k in hits if k["cat"] in ("概念", "规则")]
        faq = [k for k in hits if k["cat"] == "常见错误"]
        self.lst_core.clear()
        self.lst_faq.clear()
        for k in core:
            it = QListWidgetItem(f"[{k['cat']}] {k['title']}")
            it.setData(Qt.UserRole, k["id"])
            self.lst_core.addItem(it)
        for k in faq:
            it = QListWidgetItem(k["title"])
            it.setData(Qt.UserRole, k["id"])
            self.lst_faq.addItem(it)
        # 默认选中核心知识第一条；核心知识空则选常见问题第一条
        if self.lst_core.count():
            self.tabs_k.setCurrentIndex(0)
            self.lst_core.setCurrentRow(0)
        elif self.lst_faq.count():
            self.tabs_k.setCurrentIndex(1)
            self.lst_faq.setCurrentRow(0)

    def _show(self, cur, prev=None):
        if not cur:
            return
        k = knowledge.by_id(cur.data(Qt.UserRole))
        if not k:
            return
        cat_cn = "核心知识" if k["cat"] in ("概念", "规则") else "常见问题"
        self.view.setHtml(
            f"<h4>{_esc(k['title'])}</h4>"
            f"<p style='color:#64748B'>分类：{_esc(cat_cn)}</p>"
            f"<p>{_esc(k['body']).replace(chr(10), '<br>')}</p>"
            f"<p style='font-size:11px'>来源：<a href='{k['source']}'>{k['source']}</a></p>")

    def closeEvent(self, e):
        # 主窗口持有的是单例引用，关掉要把它清掉，否则下次点开是具「僵尸」
        p = self.parent()
        if p is not None and getattr(p, "_knowledge_dlg", None) is self:
            p._knowledge_dlg = None
        super().closeEvent(e)


class MainWindow(QMainWindow):
    # 表定义表单里的字段（与 TableDef 属性同名，顺序即表单顺序）
    FORM_KEYS = ("full_name", "read_schema_from_file", "input",
                 "index", "mode", "group", "comment")
    # 文件树里跳过的目录（噪音，不是工程内容）
    SKIP_DIRS = {"_lubankit_backup", "__pycache__", ".git", ".svn", ".hg",
                 ".idea", "node_modules"}
    MAX_TREE_DEPTH = 8
    TEXT_EXT = {".txt", ".md", ".markdown", ".json", ".jsonc", ".conf", ".log",
                ".xml", ".csv", ".tsv", ".bat", ".cmd", ".sh", ".py", ".cs",
                ".java", ".cpp", ".h", ".hpp", ".lua", ".ts", ".js", ".yml",
                ".yaml", ".ini", ".sql", ".gitignore"}
    DEF_FILES = ("__tables__.xlsx", "__beans__.xlsx", "__enums__.xlsx")

    def __init__(self):
        super().__init__()
        self.setWindowTitle("LubanEditor")
        self.resize(1440, 900)
        self.qt_settings = QSettings("LubanEditor", "LubanEditor")
        self.settings: Dict[str, str] = {
            "luban_exe": self.qt_settings.value("luban_exe", ""),
            "conf_path": self.qt_settings.value("conf_path", ""),
            "target": self.qt_settings.value("target", ""),
            "extra_args": self.qt_settings.value("extra_args", ""),
            "cmd_template": self.qt_settings.value("cmd_template", runner.DEFAULT_CMD),
            "work_dir": self.qt_settings.value("work_dir", ""),
        }
        self.model: ProjectModel = ProjectModel()
        self.fields: Dict[str, List[FieldDef]] = {}
        self.current: Optional[str] = None
        self._loading = False
        self._painting = False          # 正在刷红/去红，避免 itemChanged 递归
        self._editing = False           # 右栏当前是不是配置表编辑器
        self._form: Dict[str, str] = {}         # 表定义缓冲区（未提交）
        self._base_form: Dict[str, str] = {}    # 表定义基线（用于判断是否改动）
        self._edit_fields: List[FieldDef] = []  # 字段缓冲区
        self._base_fields: List[FieldDef] = []  # 字段基线
        self._input_map: Dict[str, str] = {}    # 规范化数据文件路径 -> 表名
        self._stem_map: Dict[str, str] = {}     # 小写文件名 -> 表名（兜底匹配）
        self._nt_buf: List[FieldDef] = []       # 新建表页的初始字段缓冲区
        self._nt_vertical_hint: Optional[bool] = None  # 新建表时勾的「纵表模板」透传给导出
        self._nt_schema: str = "xlsx"           # 新建表的表声明文件类型：xlsx / xml
        self._k_seed: str = ""                  # 给「知识点」窗口的预设搜索词
        self._knowledge_dlg = None              # 知识点窗口（非模态，单例）
        self._last_issues: List[Dict[str, str]] = []
        self._build_ui()
        self._bind_form()
        last = self.qt_settings.value("last_root", "")
        if last and os.path.isdir(last):
            self.open_project(last)
        else:
            self._show_home()

    # ---------------------------------------------------------------- UI
    def _build_ui(self):
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.home = HomePage()
        self.home.openRequested.connect(self.act_open)
        self.home.createRequested.connect(self.act_create_project)
        self.home.initRequested.connect(self.act_init_project)
        self.home.recentOpenRequested.connect(self._open_recent)
        self.home.recentDeleteRequested.connect(self._remove_recent)
        self.stack.addWidget(self.home)

        self.work = QWidget()
        self._build_work()
        self.stack.addWidget(self.work)

    def _build_work(self):
        tb = QToolBar("主工具条")
        tb.setMovable(False)
        self.addToolBar(tb)
        for text, fn, tip, style in (
            ("打开工程", self.act_open, "选择工程根目录（目录没有限制，任意文件夹都能打开）", ""),
            ("发布工程", self.act_publish,
             "设置 gen.bat / gen.sh 的参数并运行它（＝ 导表发布）", "bold"),
            ("新建表", self._new_table, "新建一张配置表并生成数据表模板", ""),
            ("管理自定义类型", self.act_manage_types,
             "查看工程里所有结构 / 枚举被哪些表引用，并删除没人用的类型", ""),
            # ---- 下面几个入口按产品要求从界面隐藏，代码留着给内部流程用 ----
            # 「导出数据表模板」新建表后会自动调一次；其余是排查问题时的备用手段。
            ("导出数据表模板", self.act_export, "生成带正确表头的 .xlsx 数据表", ""),
            ("静态自检", self.act_check, "不跑 Luban，先查低级错误", ""),
            ("生成并校验", self.act_gen, "调用 Luban.exe 并解析报错", ""),
            ("设置", self.act_settings, "Luban.exe 路径与命令模板", ""),
        ):
            a = QAction(text, self)
            a.triggered.connect(fn)
            a.setToolTip(tip)
            tb.addAction(a)
            if style:
                btn = tb.widgetForAction(a)
                if btn is not None:
                    btn.setObjectName(style)
            if text in HIDDEN_TOOLBAR_ACTIONS:
                a.setVisible(False)
            if text == "发布工程":
                self.action_publish = a
            if text == "新建表":
                self.action_new_table = a
            if text == "管理自定义类型":
                # 注意：动作对象存 action_manage_types，别把同名的槽方法覆盖掉
                self.action_manage_types = a
        # 「帮助」按钮：下拉菜单，收纳类型助手 / 核心知识 / Luban 文档 / 工具源码
        help_btn = QToolButton()
        help_btn.setText("帮助")
        help_btn.setPopupMode(QToolButton.InstantPopup)
        help_btn.setStyleSheet("QToolButton::menu-indicator { image: none; }")
        help_btn.setToolTip("类型助手 / 核心知识 / Luban 文档 / 工具源码")
        help_menu = QMenu(help_btn)
        help_menu.addAction("类型助手", self.act_type_helper)
        help_menu.addAction("核心知识", self.act_knowledge)
        help_menu.addSeparator()
        a_doc = help_menu.addAction("Luban 文档")
        a_doc.triggered.connect(
            lambda: QDesktopServices.openUrl(QUrl("https://www.datable.cn/docs/intro")))
        a_src = help_menu.addAction("工具源码")
        a_src.triggered.connect(
            lambda: QDesktopServices.openUrl(QUrl("https://github.com/heyongqi/hello-world")))
        help_btn.setMenu(help_menu)
        tb.addWidget(help_btn)
        self.help_btn = help_btn
        self.toolbar = tb
        tb.setVisible(False)

        root = QVBoxLayout(self.work)
        root.setContentsMargins(10, 8, 10, 8)
        root.setSpacing(8)

        # 顶部不再放「未初始化」条幅：非完整工程的「初始化工程」按钮放在预览区中间
        # （见 page_init 与 _show_overview）。

        # ---- 左侧 · 工程区：标题「工程目录」+ 打开按钮 + 目录树 ----
        self.tree = QTreeWidget()
        # 表头改为上面那行「工程目录」标题（右侧要放「打开」按钮），
        # 树自己的表头再留一条就是重复信息了。
        self.tree.setHeaderHidden(True)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemClicked.connect(self._on_tree)
        self.tree.itemDoubleClicked.connect(self._on_tree_dbl)
        self.tree.setToolTip("单击：在中间区预览 / 编辑；双击文件：用系统默认程序（如 Excel）打开")
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_menu)

        # 工程区的标题行：「工程目录」+ 右侧「打开」按钮（打开资源管理器）。
        # 以前这一排在窗口最上面连着完整路径一起显示，现在只留按钮放在这里。
        self.left_panel = QWidget()
        lp = QVBoxLayout(self.left_panel)
        lp.setContentsMargins(0, 0, 0, 0)
        lp.setSpacing(6)
        head_row = QWidget()
        hr = QHBoxLayout(head_row)
        hr.setContentsMargins(0, 0, 0, 0)
        lab_proj = QLabel("工程目录")
        lab_proj.setObjectName("cap")
        self.bt_folder = QPushButton("打开")
        self.bt_folder.setToolTip("用资源管理器打开工程目录（选中文件时会直接定位到该文件）")
        self.bt_folder.clicked.connect(lambda: self.act_show_folder())
        self.bt_refresh = QPushButton("刷新")
        self.bt_refresh.setToolTip("重新读取工程（luban.conf / 定义文件 / 数据表）")
        self.bt_refresh.clicked.connect(self.act_refresh)
        self.bt_remove = QPushButton("移除")
        self.bt_remove.setToolTip("从工具中移除当前工程（不删除磁盘文件夹），回到首页")
        self.bt_remove.clicked.connect(self.act_remove)
        hr.addWidget(lab_proj)
        hr.addStretch(1)
        hr.addWidget(self.bt_folder)
        hr.addWidget(self.bt_refresh)
        hr.addWidget(self.bt_remove)
        lp.addWidget(head_row)
        lp.addWidget(self.tree, 1)

        # ---- 中间 · 文件预览 / 创建区 ----
        self.editor_stack = QStackedWidget()
        self.page_table = self._build_table_editor()      # 配置表编辑器（可编辑，改动标红）
        self.page_file = QTextBrowser()                   # 只读预览
        self.page_file.setOpenExternalLinks(True)
        self.page_new = self._build_new_table_page()      # 新建配置表（内联，不弹窗）
        # 配置表预览页：先只读看配置内容，底部「编辑」按钮进入可编辑状态
        self.page_preview = QWidget()
        pv = QVBoxLayout(self.page_preview)
        pv.setContentsMargins(0, 0, 0, 0)
        pv.setSpacing(6)
        self.view_table_preview = QTextBrowser()
        self.view_table_preview.setOpenExternalLinks(True)
        self.view_table_preview.anchorClicked.connect(self._on_preview_anchor)
        pv.addWidget(self.view_table_preview, 1)
        self.bt_edit_table = QPushButton("编辑")
        self.bt_edit_table.setObjectName("primary")
        self.bt_edit_table.setMinimumHeight(36)
        self.bt_edit_table.setMinimumWidth(160)
        self.bt_edit_table.setToolTip("进入可编辑状态，编辑这张配置表的表定义与字段")
        self.bt_edit_table.clicked.connect(self._on_edit_table)
        pvb = QHBoxLayout()
        pvb.addStretch(1)
        pvb.addWidget(self.bt_edit_table)
        pvb.addStretch(1)
        pv.addLayout(pvb)
        self._pending_table = None
        self._pending_path = None
        self._show_all_data = False
        # 非完整工程：预览区中间的「初始化工程」按钮
        self.page_init = QWidget()
        pil = QVBoxLayout(self.page_init)
        pil.setContentsMargins(0, 0, 0, 0)
        pil.addStretch(1)
        lab_init = QLabel("这个目录还没有 Luban 配置")
        lab_init.setStyleSheet("font-size:16px;color:#64748B;")
        lab_init.setAlignment(Qt.AlignCenter)
        pil.addWidget(lab_init)
        pil.addSpacing(8)
        self.bt_init_center = QPushButton("初始化工程")
        self.bt_init_center.setObjectName("primary")
        self.bt_init_center.setMinimumSize(160, 40)
        self.bt_init_center.setToolTip("在这个目录里补上 luban.conf / gen.bat / Schemas / Defines / Data")
        self.bt_init_center.clicked.connect(self.act_init_project)
        pil.addWidget(self.bt_init_center, 0, Qt.AlignCenter)
        pil.addStretch(1)
        # 自定义类型预览页：只读看定义 + 底部「新增自定义类型」按钮
        self.page_def = QWidget()
        pd = QVBoxLayout(self.page_def)
        pd.setContentsMargins(0, 0, 0, 0)
        pd.setSpacing(6)
        self.view_def = QTextBrowser()
        self.view_def.setOpenExternalLinks(True)
        pd.addWidget(self.view_def, 1)
        self.bt_add_type = QPushButton("新增自定义类型")
        self.bt_add_type.setObjectName("primary")
        self.bt_add_type.setMinimumHeight(36)
        self.bt_add_type.setMinimumWidth(160)
        self.bt_add_type.setToolTip("新建一个结构 bean 或枚举（公共或当前表专用）")
        self.bt_add_type.clicked.connect(self._on_add_type)
        pdb = QHBoxLayout()
        pdb.addStretch(1)
        pdb.addWidget(self.bt_add_type)
        pdb.addStretch(1)
        pd.addLayout(pdb)
        self._def_path = None
        for pg in (self.page_table, self.page_file, self.page_new, self.page_preview,
                   self.page_init, self.page_def):
            self.editor_stack.addWidget(pg)

        # ---- 右侧 · 上下文面板 ----
        self.tabs = QTabWidget()
        self.view_opts = QTextBrowser()
        self.view_tpl = QTextBrowser()
        self.view_issue = QListWidget()
        self.view_issue.itemDoubleClicked.connect(self._goto_issue_table)
        w_tpl = QWidget()
        tl = QVBoxLayout(w_tpl)
        tl.setContentsMargins(0, 0, 0, 0)
        hl2 = QHBoxLayout()
        hl2.addWidget(QLabel("模板形式"))
        self.cb_tpl = QComboBox()
        self.cb_tpl.addItems(["按每个字段的配置方式生成", "全部用单格 sep（最保守）"])
        self.cb_tpl.currentIndexChanged.connect(self._refresh_right)
        hl2.addWidget(self.cb_tpl, 1)
        tl.addLayout(hl2)
        tl.addWidget(self.view_tpl, 1)

        # 只用「配置方式 / 模板预览」两个页签：
        # 「知识点」搬到顶部工具栏（点开独立窗口），「问题 / 校验」不再显示。
        self.tabs.addTab(self.view_opts, "配置方式")
        self.tabs.addTab(w_tpl, "模板预览")

        # 右侧面板：标题行（含折叠按钮）+ tabs，支持折叠
        self.right_panel = QWidget()
        rp = QHBoxLayout(self.right_panel)
        rp.setContentsMargins(0, 0, 0, 0)
        rp.setSpacing(0)
        # 折叠按钮：左侧窄条，垂直居中，只有箭头图标、无背景
        self.bt_fold = QToolButton()
        self.bt_fold.setText("»")
        self.bt_fold.setFixedWidth(22)
        self.bt_fold.setStyleSheet(
            "QToolButton { border: none; background: transparent; "
            "font-size: 18px; font-weight: bold; color: #64748B; }"
            "QToolButton:hover { color: #1D4ED8; }")
        self.bt_fold.setToolTip("折叠 / 展开右侧面板")
        self.bt_fold.clicked.connect(self._toggle_right)
        fold_bar = QWidget()
        fb = QVBoxLayout(fold_bar)
        fb.setContentsMargins(0, 0, 0, 0)
        fb.addStretch(1)
        fb.addWidget(self.bt_fold)
        fb.addStretch(1)
        rp.addWidget(fold_bar)
        rp.addWidget(self.tabs, 1)

        sp = QSplitter(Qt.Horizontal)
        sp.addWidget(self.left_panel)     # 左：工程区
        sp.addWidget(self.editor_stack)   # 中：文件预览 / 创建区
        sp.addWidget(self.right_panel)    # 右：配置方式 / 模板预览（可折叠）
        sp.setStretchFactor(0, 2)
        sp.setStretchFactor(1, 6)
        sp.setStretchFactor(2, 4)
        sp.setSizes([250, 640, 430])
        sp.setChildrenCollapsible(False)
        self.workspace = sp
        root.addWidget(sp, 1)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.hide()

    # ---------------------------------------------------------------- 配置表编辑器
    def _build_table_editor(self) -> QWidget:
        """中间区的配置表编辑器：表定义 + 字段 + 底部保存/取消。

        改动先进缓冲区（_form / _edit_fields）并标红，点底部「保存」才写回定义文件。
        """
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        self.gb_form = QGroupBox("表定义")
        f = QFormLayout(self.gb_form)
        self.ed_input = QLineEdit()
        self.ed_input.setPlaceholderText("英文/英文+数字，如 reward 或 reward.xlsx（默认补 .xlsx）")
        self.ed_module = QComboBox()
        self.ed_module.setEditable(True)
        self.ed_module.addItem("")
        self.ed_full = QLineEdit()
        self.ed_full.setPlaceholderText("默认 Tb+配置表名，如 TbReward；可手改")
        self.lab_def_type = QLabel("")
        self.lab_def_type.setObjectName("hint")
        self.cb_mode = ValueCombo(MODE_ITEMS, editable=False)
        self.ed_index = QLineEdit()
        self.ed_index.setPlaceholderText("map 表填主键字段名；多主键：联合 a+b / 独立 a,b")
        self.cb_group = ValueCombo(group_items(self.model))
        self.cb_read = ValueCombo(READ_ITEMS, editable=False)
        self.ed_comment = QLineEdit()
        self.lab_hint = QLabel("")
        self.lab_hint.setObjectName("hint")
        self.lab_hint.setWordWrap(True)
        f.addRow("配置表名 *", self.ed_input)
        f.addRow("模块名", self.ed_module)
        f.addRow("导出表名", self.ed_full)
        f.addRow("表声明文件类型", self.lab_def_type)
        f.addRow("表模式 mode", self.cb_mode)
        f.addRow("主键 index", self.ed_index)
        f.addRow("导出分组 group", self.cb_group)
        f.addRow("从表头推断表结构(bean)", self.cb_read)
        f.addRow("注释 comment", self.ed_comment)
        f.addRow(self.lab_hint)
        v.addWidget(self.gb_form)

        cap = QLabel("字段　（选中某行后，下方「配置方式」会列出该类型的所有填法与示例）")
        cap.setObjectName("cap")
        v.addWidget(cap)

        self.tbl_fields = QTableWidget(0, len(FIELD_COLS))
        self.tbl_fields.setHorizontalHeaderLabels(FIELD_COLS)
        hh = self.tbl_fields.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.Interactive)
        for i, w in enumerate(FIELD_WIDTHS):
            self.tbl_fields.setColumnWidth(i, w)
        hh.setSectionResizeMode(COL_COMMENT, QHeaderView.Stretch)
        self.tbl_fields.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl_fields.setAlternatingRowColors(True)
        self.tbl_fields.verticalHeader().setDefaultSectionSize(30)
        self.tbl_fields.verticalHeader().setVisible(False)
        self.tbl_fields.itemChanged.connect(self._on_field_item_changed)
        self.tbl_fields.itemSelectionChanged.connect(self._on_field_selected)
        v.addWidget(self.tbl_fields, 1)

        hb = QHBoxLayout()
        for text, fn, tip in (("+ 添加字段", self._add_field, "在末尾新增一个字段"),
                              ("- 删除字段", self._del_field, "删除选中的字段行"),
                              ("↑ 上移", lambda: self._move(-1), "与上一行交换顺序"),
                              ("↓ 下移", lambda: self._move(1), "与下一行交换顺序")):
            b = QPushButton(text)
            b.setToolTip(tip)
            b.clicked.connect(fn)
            hb.addWidget(b)
        hb.addStretch(1)
        v.addLayout(hb)

        bar = QFrame()
        bar.setObjectName("savebar")
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(10, 6, 10, 6)
        self.lab_dirty = QLabel("")
        self.lab_dirty.setObjectName("dirty")
        self.lab_dirty.setWordWrap(True)
        self.bt_cancel_ed = QPushButton("取消修改")
        self.bt_cancel_ed.setObjectName("danger")
        self.bt_cancel_ed.setToolTip("丢弃所有未保存的改动，恢复到上次保存的状态")
        self.bt_cancel_ed.clicked.connect(self._cancel_editor)
        self.bt_save_ed = QPushButton("保存（写入定义文件）")
        self.bt_save_ed.setObjectName("primary")
        self.bt_save_ed.setToolTip("把标红的改动写回定义文件（写前自动备份）")
        self.bt_save_ed.clicked.connect(self._save_editor)
        bl.addWidget(self.lab_dirty, 1)
        bl.addWidget(self.bt_cancel_ed)
        bl.addWidget(self.bt_save_ed)
        self.bar_editor = bar
        bar.setVisible(False)        # 没有改动时不占位
        v.addWidget(bar)
        return page

    # ---------------------------------------------------------------- 新建表（内联）
    def _build_new_table_page(self) -> QWidget:
        """中间区的「新建配置表」表单 —— 不再弹独立窗口。"""
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        title = QLabel("新建配置表　（填好后点底部「创建」）")
        title.setObjectName("cap")
        v.addWidget(title)

        self.gb_new = QGroupBox("表定义")
        nf = QFormLayout(self.gb_new)
        self.nt_input = QLineEdit()
        self.nt_input.setPlaceholderText("英文/英文+数字，如 reward 或 reward.xlsx（默认补 .xlsx）")
        self.nt_module = QComboBox()
        self.nt_module.setEditable(True)
        self.nt_module.addItem("")
        self.nt_module.addItems(sorted({t.full_name.split(".")[0] for t in self.model.tables
                                        if t.full_name and "." in t.full_name}))
        self.nt_module.setCurrentText("")
        self.nt_name = QLineEdit()
        self.nt_name.setPlaceholderText("默认 Tb+配置表名，如 TbReward；可手改")
        self.nt_mode = ValueCombo(MODE_ITEMS, editable=False)
        self.nt_mode.set_code("map")
        self.nt_index = QLineEdit()
        self.nt_index.setPlaceholderText("map 表填主键字段名；多主键：联合 a+b / 独立 a,b")
        self.nt_group = ValueCombo(group_items(self.model))
        self.nt_read = ValueCombo(READ_ITEMS, editable=False)
        self.nt_read.set_code("false")
        self.nt_comment = QLineEdit()
        self.nt_vertical = QCheckBox("生成纵表模板（适合单例配置）")
        self.lab_nt_err = QLabel("")
        self.lab_nt_err.setStyleSheet("color:#B42318;")
        self.lab_nt_err.setWordWrap(True)
        self._nt_name_manual = False

        # 表声明文件类型：决定这张表的声明最后落在哪个文件里
        self.rb_nt_xlsx = QRadioButton("Excel Schema")
        self.rb_nt_xml = QRadioButton("XML Schema")
        self.nt_defgrp = QButtonGroup(page)
        self.nt_defgrp.setExclusive(True)
        for rb in (self.rb_nt_xlsx, self.rb_nt_xml):
            self.nt_defgrp.addButton(rb)
        self.rb_nt_xlsx.setToolTip("表声明写进 Excel 定义文件（__tables__.xlsx），用表格维护")
        self.rb_nt_xml.setToolTip("表声明写进 XML 定义文件（如 Defines/tables.xml），用文本维护")
        self.lab_nt_target = QLabel("")
        self.lab_nt_target.setObjectName("hint")
        w_deftype = QWidget()
        dh = QHBoxLayout(w_deftype)
        dh.setContentsMargins(0, 0, 0, 0)
        dh.addWidget(self.rb_nt_xlsx)
        dh.addWidget(self.rb_nt_xml)
        dh.addSpacing(12)
        dh.addWidget(self.lab_nt_target, 1)

        nf.addRow("配置表名 *", self.nt_input)
        nf.addRow("模块名", self.nt_module)
        nf.addRow("", self.lab_nt_err)
        nf.addRow("导出表名", self.nt_name)
        # 表名一改，「当前表专用」的类型清单就变了，类型下拉要跟着重建
        self.nt_name.textChanged.connect(self._on_nt_name_changed)
        nf.addRow("表声明文件类型", w_deftype)
        nf.addRow("表模式 mode", self.nt_mode)
        nf.addRow("主键 index", self.nt_index)
        nf.addRow("导出分组 group", self.nt_group)
        nf.addRow("从表头推断表结构(bean)", self.nt_read)
        nf.addRow("注释 comment", self.nt_comment)
        nf.addRow(self.nt_vertical)
        v.addWidget(self.gb_new)
        # 配置表名 / 模块名改变：校验 + 刷新导出表名默认值 + 重建类型下拉
        self.nt_input.textChanged.connect(self._on_nt_name_fields_changed)
        self.nt_module.currentTextChanged.connect(self._on_nt_name_fields_changed)
        self.rb_nt_xlsx.toggled.connect(lambda _c: self._nt_def_type_changed())
        self.rb_nt_xml.toggled.connect(lambda _c: self._nt_def_type_changed())

        cap = QLabel("初始字段　（与预览表时同一套功能：类型下拉可选容器、"
                     "容器再选子类型、配置方式随类型变化、右侧「助手」查示例）")
        cap.setObjectName("cap")
        cap.setWordWrap(True)
        v.addWidget(cap)

        self.nt_fields = QTableWidget(0, len(FIELD_COLS))
        self.nt_fields.setHorizontalHeaderLabels(FIELD_COLS)
        nh = self.nt_fields.horizontalHeader()
        nh.setSectionResizeMode(QHeaderView.Interactive)
        for i, w in enumerate(FIELD_WIDTHS):
            self.nt_fields.setColumnWidth(i, w)
        nh.setSectionResizeMode(COL_COMMENT, QHeaderView.Stretch)
        self.nt_fields.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.nt_fields.setAlternatingRowColors(True)
        self.nt_fields.verticalHeader().setDefaultSectionSize(30)
        self.nt_fields.verticalHeader().setVisible(False)
        self.nt_fields.itemChanged.connect(self._on_field_item_changed)
        self.nt_fields.itemSelectionChanged.connect(self._refresh_new_table_right)
        v.addWidget(self.nt_fields, 1)

        hb = QHBoxLayout()
        for text, fn, tip in (("+ 添加字段", self._nt_add_field, "在末尾新增一个字段"),
                              ("- 删除字段", self._nt_del_field, "删除选中的字段行"),
                              ("↑ 上移", lambda: self._nt_move(-1), "与上一行交换顺序"),
                              ("↓ 下移", lambda: self._nt_move(1), "与下一行交换顺序")):
            b = QPushButton(text)
            b.setToolTip(tip)
            b.clicked.connect(fn)
            hb.addWidget(b)
        hb.addStretch(1)
        v.addLayout(hb)

        bar = QFrame()
        bar.setObjectName("savebar")
        self.bt_nt_cancel = QPushButton("取消")
        self.bt_nt_cancel.setObjectName("foot")
        self.bt_nt_cancel.setToolTip("放弃这次新建，回到工程概览")
        self.bt_nt_cancel.clicked.connect(self._nt_cancel)
        self.bt_nt_create = QPushButton("创建")
        self.bt_nt_create.setObjectName("footprimary")
        self.bt_nt_create.setToolTip("把这张表加进工程（数据表模板可在创建后导出）")
        self.bt_nt_create.clicked.connect(self._nt_create)
        hbt = QHBoxLayout()
        hbt.addStretch(1)
        hbt.addWidget(self.bt_nt_cancel)
        hbt.addWidget(self.bt_nt_create)
        hbt.addStretch(1)
        bl = QVBoxLayout(bar)
        bl.setContentsMargins(10, 8, 10, 8)
        bl.setSpacing(6)
        hint = QLabel("创建后表定义与字段都能在中间区继续改；点它底部的「保存」才写回表声明文件。")
        hint.setObjectName("hint")
        hint.setWordWrap(True)
        hint.setAlignment(Qt.AlignCenter)
        bl.addWidget(hint)
        bl.addLayout(hbt)
        v.addWidget(bar)
        return page

    def _open_new_table(self, ask_leave: bool = True):
        """把中间区切到「新建配置表」表单。"""
        if not self.model.initialized:
            QMessageBox.information(self, "提示",
                                    "当前目录还没有 Luban 配置，请先「初始化 Luban 工程」。")
            self.act_init_project()
            return
        if ask_leave and not self._confirm_leave_editor():
            return
        self._reset_new_table_page()
        self.editor_stack.setCurrentWidget(self.page_new)
        self.nt_fields.selectRow(0)          # 直接显示第一个字段的配置方式
        self._refresh_new_table_right()
        self.nt_name.setFocus()

    @property
    def _nt_def_type(self) -> str:
        """新建表页勾的表声明文件类型：'xlsx'（Excel Schema）/ 'xml'（XML Schema）。"""
        return "xml" if self.rb_nt_xml.isChecked() else "xlsx"

    def _nt_def_type_changed(self):
        """切了「表声明文件类型」就把目标文件写在右边的提示里。

        不写出来的话，勾 XML Schema 的人根本不知道这张表待会儿会落到哪个文件。
        """
        self._nt_schema = self._nt_def_type
        kind = "Excel Schema" if self._nt_schema == "xlsx" else "XML Schema"
        try:
            p = exceldef.plan_table_file(self.model, self._nt_schema)
        except Exception:
            p = ""
        if not p:
            self.lab_nt_target.setText(kind)
            return
        try:
            rel = os.path.relpath(p, self.model.root or os.path.dirname(p))
        except Exception:
            rel = p
        self.lab_nt_target.setText(
            f"{kind}：表声明写入 {rel}（{'已存在' if os.path.exists(p) else '将新建'}）")

    def _reset_new_table_page(self):
        self.nt_input.clear()
        # 刷新模块下拉：收集当前工程已有的模块名（保留空项）
        self.nt_module.blockSignals(True)
        self.nt_module.clear()
        self.nt_module.addItem("")
        self.nt_module.addItems(sorted({t.full_name.split(".")[0] for t in self.model.tables
                                        if t.full_name and "." in t.full_name}))
        self.nt_module.setCurrentText("")
        self.nt_module.blockSignals(False)
        self.nt_name.clear()
        self.nt_mode.set_code("map")
        self.nt_index.clear()
        self.nt_group.set_code("")
        self.nt_read.set_code("false")
        self.nt_comment.clear()
        self.nt_vertical.setChecked(False)
        self.lab_nt_err.clear()
        self.rb_nt_xlsx.setChecked(True)      # 默认 Excel Schema
        self._nt_def_type_changed()
        self.nt_fields.setRowCount(0)
        self._nt_buf = []
        self._nt_add_field()
        self._nt_name_manual = False

    def _nt_config_name(self) -> str:
        """配置表名（去掉扩展名）。"""
        s = (self.nt_input.text() or "").strip()
        return os.path.splitext(s)[0] if s else ""

    def _nt_input_with_ext(self) -> str:
        """配置表名补默认扩展名 .xlsx（除非已带扩展名）。"""
        s = (self.nt_input.text() or "").strip()
        if not s:
            return ""
        return s if os.path.splitext(s)[1] else s + ".xlsx"

    def _nt_default_full_name(self) -> str:
        """导出表名默认值：Tb + 配置表名（首字母大写），带模块前缀。"""
        cfg = self._nt_config_name()
        if not cfg:
            return ""
        base = "Tb" + cfg[:1].upper() + cfg[1:]
        mod = (self.nt_module.currentText() or "").strip()
        return f"{mod}.{base}" if mod else base

    def _validate_nt_names(self) -> bool:
        """校验配置表名 / 模块名，非法时红色提示「名字不合法：***」。"""
        errs = []
        cfg = (self.nt_input.text() or "").strip()
        if cfg and not re.match(r"^[A-Za-z][A-Za-z0-9_]*$", os.path.splitext(cfg)[0]):
            errs.append(f"名字不合法：{cfg}")
        mod = (self.nt_module.currentText() or "").strip()
        if mod and not re.match(r"^[A-Za-z][A-Za-z0-9_]*$", mod):
            errs.append(f"名字不合法：{mod}")
        self.lab_nt_err.setText("；".join(errs))
        return not errs

    def _sync_nt_full_name(self):
        """按配置表名 / 模块名刷新导出表名默认值（用户手动改过则不动）。"""
        if self._nt_name_manual:
            return
        cfg = self._nt_config_name()
        if not re.match(r"^[A-Za-z][A-Za-z0-9_]*$", cfg):
            return
        new = self._nt_default_full_name()
        if new and self.nt_name.text().strip() != new:
            prev = self._loading
            self._loading = True
            try:
                self.nt_name.setText(new)
            finally:
                self._loading = prev

    def _on_nt_name_fields_changed(self):
        """配置表名 / 模块名变了：校验 + 刷新导出表名默认值。"""
        if not hasattr(self, "_nt_buf"):
            return
        self._validate_nt_names()
        self._sync_nt_full_name()

    def _on_nt_name_changed(self):
        """表名改了 -> 重建各行的「类型 / 子类型」下拉。

        清单里含「本表专用」的枚举与结构，表名一换这份清单就变了；
        不重建的话刚为本表建的类型在下拉里看不到。
        """
        if not hasattr(self, "_nt_buf"):
            return
        if not self._loading:
            self._nt_name_manual = True      # 用户手动编辑了导出表名
        prev = self._loading
        self._loading = True
        try:
            for r in range(min(self.nt_fields.rowCount(), len(self._nt_buf))):
                self._set_type_combo(r, self._nt_buf[r].type, self.nt_fields)
        finally:
            self._loading = prev

    def _nt_add_field(self):
        self._nt_buf.append(FieldDef(name="field_%d" % (len(self._nt_buf) + 1), type="int"))
        self._append_field_row(self._nt_buf[-1], self.nt_fields)

    def _nt_del_field(self):
        rows = sorted({i.row() for i in self.nt_fields.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(self._nt_buf):
                self._nt_buf.pop(r)
            self.nt_fields.removeRow(r)
        self._refresh_new_table_right()

    def _nt_move(self, delta: int):
        rows = sorted({i.row() for i in self.nt_fields.selectedIndexes()})
        if len(rows) != 1:
            return
        r = rows[0]
        n = r + delta
        if not (0 <= n < len(self._nt_buf)):
            return
        self._nt_buf[r], self._nt_buf[n] = self._nt_buf[n], self._nt_buf[r]
        self._reload_fields_table(self.nt_fields, self._nt_buf)
        self.nt_fields.selectRow(n)
        self._refresh_new_table_right()

    def _nt_result_table(self) -> TableDef:
        name = self.nt_name.text().strip()
        cfg = self._nt_input_with_ext()
        mod = (self.nt_module.currentText() or "").strip()
        input_path = f"{mod}/{cfg}" if (mod and cfg) else cfg
        mode = self.nt_mode.code()
        t = TableDef(
            full_name=name,
            value_type=template.derive_value_type(name),
            read_schema_from_file=self.nt_read.code(),
            input=input_path,
            index=self.nt_index.text().strip(),
            mode=mode,
            group=self.nt_group.code(),
            comment=self.nt_comment.text().strip(),
        )
        if mode == "one":
            t.index = ""
        return t

    def _nt_result_fields(self) -> List[FieldDef]:
        self._sync_fields_from_cells(self.nt_fields, self._nt_buf)
        return [self._copy_field(f) for f in self._nt_buf]

    def _refresh_new_table_right(self):
        """新建表页：右侧「配置方式」跟着选中的初始字段走，与预览表时一致。"""
        if self.editor_stack.currentWidget() is not self.page_new:
            return
        r = self.nt_fields.currentRow()
        if not (0 <= r < len(self._nt_buf)):
            self._show_right_placeholder(
                "正在新建配置表。点一个初始字段行，右侧会列出该类型的所有配置方式与示例。")
            return
        f = self._nt_buf[r]
        cb2 = self.nt_fields.cellWidget(r, COL_STYLE)
        cur = ((cb2.currentData() if cb2 else "") or "")
        h = (f"<h3>新建表 · 字段 <code>{_esc(f.name)}</code>　类型 "
             f"<code>{_esc(f.type)}</code></h3>")
        issues = typecatalog.validate_type(self.model, f.type)
        if issues:
            h += "<p style='color:#B42318'>" + "；".join(
                f"[{i['level']}] {_esc(i['msg'])}" for i in issues) + "</p>"
        h += options_html(typecatalog.fill_options(self.model, f.name, f.type), cur)
        self.view_opts.setHtml(h)
        try:
            cols = template.build_cols(self.model, self._nt_buf, default="auto")
            self.view_tpl.setHtml("<p>按各字段「配置方式」生成的表头预览：</p>"
                                  + _html_table(template.header_grid(cols)))
        except Exception as e:
            self.view_tpl.setHtml(f"<p>预览失败：{_esc(str(e))}</p>")

    def _nt_create(self):
        cfg = (self.nt_input.text() or "").strip()
        if not cfg:
            QMessageBox.warning(self, "缺少配置表名", "请填写配置表名")
            self.nt_input.setFocus()
            return
        if not self._validate_nt_names():
            QMessageBox.warning(self, "名字不合法", self.lab_nt_err.text())
            return
        name = self.nt_name.text().strip() or self._nt_default_full_name()
        if not name:
            QMessageBox.warning(self, "缺少导出表名", "请填写导出表名")
            self.nt_name.setFocus()
            return
        if self.model.find_table(name):
            QMessageBox.warning(self, "重名", f"已存在同名表：{name}")
            self.nt_name.setFocus()
            return
        # ① 先定好「这张表的声明写进哪个文件」：勾的就是用户选的那一种，
        #    文件不在就先建（xml 建 tables.xlsx 的空模板），并登记进 luban.conf
        schema_kind = self._nt_def_type
        ok, target, wmsg = exceldef.ensure_table_file(self.model, schema_kind)
        if not ok:
            QMessageBox.warning(self, "没法写表声明", wmsg)
            return
        t = self._nt_result_table()
        # 填了模块名：在 Data 下建模块文件夹，数据文件放里面
        mod = t.full_name.split(".")[0] if "." in t.full_name else ""
        if mod:
            os.makedirs(os.path.join(self.model.data_dir, mod), exist_ok=True)
        t._file = target
        fs = self._nt_result_fields()
        want_vertical = self.nt_vertical.isChecked()
        self.model.tables.append(t)
        # 表刚出现，之前给它建的表专用类型这会儿才认得出归属，重算一次
        exceldef.derive_scopes(self.model)
        self.fields[t.full_name] = fs
        self._nt_vertical_hint = want_vertical
        self._reindex_inputs()
        self._refresh_tree()
        self._load_table_editor(t.full_name)
        tipfile = os.path.basename(target or "") or "定义文件"
        if wmsg:
            # 文件建好了、表也进来了，但没能自动登记进 luban.conf —— 这种事
            # 必须当面说清，不然 Luban 导表时会当这张表不存在。
            QMessageBox.warning(self, "定义文件没能自动登记", wmsg)
        if fs:
            self.act_export()
        # 创建后立即写盘：表声明 + 表结构 bean（read_schema_from_file=false 时），
        # 让「新建表」这个动作本身就把 bean 落进 xml，不用再点一次保存。
        self._save_definitions(t.full_name)
        self.status.showMessage(
            f"已新建 {t.full_name}（声明写在 {tipfile}）", 10000)

    def _nt_cancel(self):
        self._show_overview()

    def _bind_form(self):
        """表定义表单 -> 缓冲区（不直接改 model）。"""
        self._form_widgets: Dict[str, QWidget] = {
            "full_name": self.ed_full,
            "read_schema_from_file": self.cb_read,
            "input": self.ed_input,
            "index": self.ed_index,
            "mode": self.cb_mode,
            "group": self.cb_group,
            "comment": self.ed_comment,
        }
        # 模块名不映射 TableDef 属性，但变化时要刷新 input 的目录部分
        self.ed_module.currentTextChanged.connect(self._on_module_changed)
        for key, w in self._form_widgets.items():
            if isinstance(w, QComboBox):
                w.currentIndexChanged.connect(
                    lambda _=None, k=key: self._on_form_changed(k))
                if w.isEditable():
                    w.editTextChanged.connect(
                        lambda _=None, k=key: self._on_form_changed(k))
            else:
                w.textChanged.connect(lambda _=None, k=key: self._on_form_changed(k))

    # ---------------------------------------------------------------- 打开/创建
    def _show_home(self):
        self.home.set_recent(self._recent_projects())
        self.stack.setCurrentWidget(self.home)
        # 首页也显示工具栏，但屏蔽「发布工程」「管理自定义类型」「新建表」
        self.toolbar.setVisible(True)
        self.action_publish.setVisible(False)
        self.action_manage_types.setVisible(False)
        self.action_new_table.setVisible(False)
        self.status.hide()

    def _recent_projects(self) -> List[str]:
        # 独立 key：只存「手动打开」的工程，不会被程序内部打开/测试数据污染
        v = self.qt_settings.value("recent_manual", []) or []
        return [p for p in v if isinstance(p, str) and p] if isinstance(v, list) else []

    def _remember(self, root: str):
        """记住上次打开的工程（启动时自动打开用），不污染「最近打开」列表。"""
        self.qt_settings.setValue("last_root", root)

    def _remember_manual(self, root: str):
        """手动打开的工程加入「最近打开」列表（去重、最新在前、最多 10 条）。"""
        recent = self._recent_projects()
        recent = [p for p in recent if os.path.normcase(p) != os.path.normcase(root)]
        recent.insert(0, root)
        self.qt_settings.setValue("recent_manual", recent[:10])

    def _remove_recent(self, root: str):
        """从「最近打开」列表删除某条记录。"""
        recent = [p for p in self._recent_projects()
                  if os.path.normcase(p) != os.path.normcase(root)]
        self.qt_settings.setValue("recent_manual", recent)
        if self.stack.currentWidget() is self.home:
            self._show_home()

    def _open_recent(self, root: str):
        """点击「最近打开」的工程：移到最前并打开。"""
        self._remember_manual(root)
        self.open_project(root)

    def act_open(self):
        d = QFileDialog.getExistingDirectory(
            self, "选择工程根目录（任意目录都可以）",
            self.qt_settings.value("last_root", "") or "")
        if d:
            self._remember_manual(d)   # 手动打开，记录进「最近打开」
            self.open_project(d)

    def open_project(self, root: str, silent: bool = False):
        if not silent and not self._confirm_leave_editor():
            return
        try:
            self.model = projectconf.discover(root)
            exceldef.load_project(self.model)
        except Exception as e:
            if not silent:
                QMessageBox.critical(self, "打开失败", str(e))
            return
        self._remember(root)
        self._reindex_inputs()
        self.fields = {}
        for t in self.model.tables:
            fs = self._read_fields_from_datafile(t)
            if fs is None and t.value_type:
                b = self.model.find_bean(t.value_type)
                if b and b.fields:
                    fs = [FieldDef(f.name, f.type, f.group, f.comment) for f in b.fields]
            self.fields[t.full_name] = fs or []
        self._reset_editor()
        self._refresh_group_choices()
        self._refresh_tree()
        self.stack.setCurrentWidget(self.work)
        self.toolbar.setVisible(True)
        self.action_publish.setVisible(True)
        self.action_new_table.setVisible(True)
        self.status.show()
        # 工程完整路径不显示在顶部（窗口标题保持固定，路径在概览页/状态栏里能看到）
        self.setWindowTitle("LubanEditor")
        self.status.showMessage(
            f"{root}　表 {len(self.model.tables)} / 结构 {len(self.model.beans)} / "
            f"枚举 {len(self.model.enums)}", 8000)
        self._update_banner()
        self._show_overview()
        self._repaint_dirty()
        self._show_issues([{"level": "info", "table": "", "msg": w, "hint": ""}
                           for w in self.model.warnings])

    def _reset_editor(self):
        """清空编辑器状态（未选中任何表）。"""
        self.current = None
        self._editing = False
        self._form = {k: "" for k in self.FORM_KEYS}
        self._base_form = dict(self._form)
        self._edit_fields = []
        self._base_fields = []
        self._loading = True
        self.tbl_fields.setRowCount(0)
        self._loading = False

    def _refresh_group_choices(self):
        """换工程后，分组下拉的选项要跟着新工程的 luban.conf 走。"""
        items = group_items(self.model)
        prev = self._loading
        self._loading = True
        for cb in (getattr(self, "cb_group", None), getattr(self, "nt_group", None)):
            if cb is None:
                continue
            cur = cb.code()
            cb.blockSignals(True)
            cb.clear()
            cb._disp = {}
            for code, disp in items:
                cb.addItem(disp, code)
                cb._disp[code] = disp
            cb.blockSignals(False)
            cb.set_code(cur)
        self._loading = prev
        # 已填好的字段行也要重建（旧选项可能已不存在于新工程）
        for tbl in (self.tbl_fields, self.nt_fields):
            fields, _name = self._field_ctx(tbl)
            for r, f in enumerate(fields):
                if r < tbl.rowCount():
                    self._fill_group_cell(r, f.group, tbl)

    def toast(self, text: str, msec: int = 2800):
        """窗口底部浮一条轻提示：不阻塞、不打断正在填的内容。

        「保存成功后要有个提示」——但弹模态对话框会让连续编辑很难受，
        所以做成自动消失的浮层，同时状态栏也留一条。
        """
        if not hasattr(self, "_toast"):
            self._toast = QLabel(self)
            self._toast.setObjectName("toast")
            self._toast.setAlignment(Qt.AlignCenter)
            self._toast.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            self._toast.hide()
        lab = self._toast
        lab.setText(text)
        lab.adjustSize()
        lab.move(max(8, (self.width() - lab.width()) // 2),
                 max(8, self.height() - lab.height() - 64))
        lab.raise_()
        lab.show()
        QTimer.singleShot(msec, lab.hide)

    def _update_banner(self):
        # 「管理自定义类型」依赖 luban.conf 里的定义文件，不是 Luban 工程就藏起来。
        # 顶部已不放「未初始化」条幅，非完整工程的初始化入口在预览区中间（page_init）。
        act = getattr(self, "action_manage_types", None)
        if act is not None:
            act.setVisible(self.model.initialized)

    def act_create_project(self):
        self._create_project_at("", "create")

    def act_init_project(self):
        self._create_project_at(self.model.root or "", "init")

    def _create_project_at(self, root: str, mode: str = "create"):
        d = CreateProjectDialog(self, root, mode)
        if d.exec() != QDialog.Accepted:
            return
        target, types, tools = d.result()
        ok, msg = projectwizard.create_project(target, types, tools)
        if not ok:
            QMessageBox.critical(self, "创建失败", msg)
            return
        QMessageBox.information(self, "完成", msg)
        self.open_project(target)

    def act_show_folder(self, path: str = ""):
        """在资源管理器里打开工程目录；给了文件就定位到该文件。"""
        root = self.model.root
        if not root or not os.path.isdir(root):
            QMessageBox.information(self, "提示", "还没有打开工程")
            return
        target = str(path or "")
        try:
            if os.name == "nt":
                if target and os.path.exists(target):
                    subprocess.Popen(["explorer", "/select,", os.path.normpath(target)])
                else:
                    os.startfile(root)          # noqa: S606 - 本地 GUI 程序，打开目录
            elif sys.platform == "darwin":
                subprocess.Popen(["open", target or root])
            else:
                subprocess.Popen(["xdg-open", target or root])
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开资源管理器：{e}")

    def act_refresh(self):
        """重新读取工程（luban.conf / 定义文件 / 数据表）。"""
        if not self.model.root or not os.path.isdir(self.model.root):
            QMessageBox.information(self, "提示", "还没有打开工程")
            return
        self.open_project(self.model.root)

    def act_remove(self):
        """从工具中移除当前工程（不删除磁盘文件夹），回到首页。"""
        if not self.model.root:
            return
        if QMessageBox.question(
                self, "移除工程",
                "确定要从工具中移除当前工程吗？\n\n%s\n\n"
                "（只是从工具移除，不会删除磁盘上的文件夹）" % self.model.root
                ) != QMessageBox.Yes:
            return
        self.model = ProjectModel()
        self.fields = {}
        self._reset_editor()
        self.tree.clear()
        self._show_home()

    def _toggle_right(self):
        """折叠 / 展开右侧上下文面板。折叠后右侧只留箭头按钮，中间区自适应变大。"""
        if self.tabs.isVisible():
            # 折叠：隐藏 tabs，右侧缩成窄条（只留箭头）
            self.tabs.setVisible(False)
            self.right_panel.setMaximumWidth(22)
            self.bt_fold.setText("«")
        else:
            # 展开：恢复 tabs，解除宽度限制
            self.tabs.setVisible(True)
            self.right_panel.setMaximumWidth(16777215)
            self.bt_fold.setText("»")

    # ---------------------------------------------------------------- 读取字段
    def _read_fields_from_datafile(self, t: TableDef) -> Optional[List[FieldDef]]:
        p = template.default_input_path(self.model, t)
        if not os.path.exists(p):
            return None
        try:
            from openpyxl import load_workbook
            wb = load_workbook(p, data_only=True)
            ws = wb.worksheets[0]
            out: List[FieldDef] = []
            vrow = trow = None
            for r in range(1, min(ws.max_row or 1, 12) + 1):
                a = str(ws.cell(row=r, column=1).value or "").strip()
                if a == "##var" and vrow is None:
                    vrow = r
                elif a == "##type" and trow is None:
                    trow = r
            if vrow and trow:
                crow = None
                # 中文说明行（## / ##comment）在 ##type、##group 之后，数据行之前。
                # 以前只在 ##type 之前找，永远找不到 —— 重新打开表时注释就丢了。
                for r in range(trow + 1, (ws.max_row or 1) + 1):
                    a = str(ws.cell(row=r, column=1).value or "").strip()
                    if a in ("##", "##comment"):
                        crow = r
                        break
                    if not a.startswith("##"):
                        break
                for c in range(2, (ws.max_column or 1) + 1):
                    n = ws.cell(row=vrow, column=c).value
                    if n in (None, ""):
                        continue
                    # 字段名上可能挂着 #format=lite（官方「紧凑格式」写法），要拆出来
                    nm, fmt = typecatalog.split_format(str(n))
                    out.append(FieldDef(
                        name=nm,
                        type=str(ws.cell(row=trow, column=c).value or "int").strip(),
                        comment=str(ws.cell(row=crow, column=c).value or "").strip() if crow else "",
                        fmt=fmt))
            elif str(ws.cell(row=1, column=1).value or "").strip().lower().startswith("##column"):
                for r in range(3, (ws.max_row or 2) + 1):
                    n = ws.cell(row=r, column=1).value
                    if n in (None, ""):
                        continue
                    out.append(FieldDef(
                        name=str(n).strip(),
                        type=str(ws.cell(row=r, column=2).value or "int").strip(),
                        comment=str(ws.cell(row=r, column=3).value or "").strip()))
            wb.close()
            return out or None
        except Exception:
            return None

    # ---------------------------------------------------------------- 工程目录树
    def _reindex_inputs(self):
        """建立「数据文件 -> 表名」索引，判断某个文件是不是某张配置表的数据文件。

        两张索引：
        - ``_input_map``：规范化绝对路径 -> 表名（按声明里的路径精确匹配）；
        - ``_stem_map``：小写文件名（去目录、去 @sheet）-> 表名。

        为什么要第二张：声明里的 input 可能留空（Luban 取表名）、带目录前缀、
        或大小写/扩展名不一致，只按精确路径匹配会把能编辑的表判成只读。
        """
        self._input_map = {}
        stems: Dict[str, Optional[str]] = {}
        for t in self.model.tables:
            for p in self._input_paths(t):
                self._input_map.setdefault(os.path.normcase(os.path.abspath(p)), t.full_name)
                stem = self._stem_of(p)
                if not stem:
                    continue
                if stem in stems and stems[stem] != t.full_name:
                    stems[stem] = None          # 同名文件属于两张表 -> 放弃兜底，不猜
                else:
                    stems.setdefault(stem, t.full_name)
        self._stem_map = {k: v for k, v in stems.items() if v}

    @staticmethod
    def _stem_of(path: str) -> str:
        """取文件名主干（小写）：a/b/C.xlsx@Sheet1 -> c。"""
        base = os.path.basename((path or "").split("@")[0].strip())
        return os.path.splitext(base)[0].lower()

    def _input_paths(self, t: TableDef) -> List[str]:
        """把 TableDef.input 拆成候选文件路径（支持 a.xlsx、a.xlsx@sheet、a.xlsx,b.xlsx）。

        input 为空时按 Luban 的默认规则取表名（与 template.default_input_path 一致）：
        否则「声明了表但没写 input」的表永远匹配不到自己的数据文件，点开只能是只读。
        """
        raw = (t.input or "").strip()
        base = self.model.data_dir or self.model.root
        if not raw:
            return [os.path.normpath(template.default_input_path(self.model, t))]
        out: List[str] = []
        for part in re.split(r"[,;]", raw):
            part = part.strip()
            if not part:
                continue
            for cand in (part, part.split("@")[0].strip()):
                if not cand:
                    continue
                p = cand if os.path.isabs(cand) else os.path.join(base, cand)
                p = os.path.normpath(p)
                if p not in out:
                    out.append(p)
        # 声明里有 input，但那儿没有文件时，顺手把默认文件名也当候选（写作习惯不同也能认出来）
        if not any(os.path.exists(p) for p in out):
            d = os.path.normpath(template.default_input_path(self.model, t))
            if d not in out:
                out.append(d)
        return out

    def _table_for_file(self, path: str) -> str:
        """点击的文件 -> 表名。先精确匹配声明路径，再按文件名兜底；都没有返回空串。"""
        if not path:
            return ""
        hit = self._input_map.get(os.path.normcase(os.path.abspath(path)), "")
        if hit:
            return hit
        return self._stem_map.get(self._stem_of(path), "")

    @staticmethod
    def _under(child: str, parent: str) -> bool:
        try:
            child = os.path.abspath(child)
            parent = os.path.abspath(parent)
            return os.path.commonpath([child, parent]) == parent
        except Exception:
            return False

    def _is_def_file(self, path: str) -> bool:
        """是不是本工程的「定义文件」：__*.xlsx 或 luban.conf 登记过的 xml 定义。"""
        if os.path.basename(path).lower() in self.DEF_FILES:
            return True
        pool = (list(self.model.table_files) + [self.model.bean_file,
                self.model.enum_file] + list(self.model.extra_bean_files)
                + list(self.model.extra_enum_files))
        want = os.path.normcase(os.path.abspath(path))
        return any(p and os.path.normcase(os.path.abspath(p)) == want for p in pool)

    def _file_tag(self, path: str) -> str:
        if self._is_def_file(path):
            return "[定义]"
        tname = self._table_for_file(path)
        if tname:
            t = self.model.find_table(tname)
            # 配置文件不再标「[配置表]」，有注释就显示注释
            return (t.comment or "").strip() if t else ""
        return ""

    def _fill_dir(self, parent: QTreeWidgetItem, path: str, depth: int):
        if depth >= self.MAX_TREE_DEPTH:
            return
        try:
            entries = list(os.scandir(path))
        except OSError:
            return
        entries.sort(key=lambda e: (not e.is_dir(), e.name.lower()))
        for e in entries:
            if e.name.startswith("~$"):
                continue
            try:
                is_dir = e.is_dir()
            except OSError:
                continue
            if is_dir:
                if e.name in self.SKIP_DIRS:
                    continue
                it = QTreeWidgetItem([e.name])
                it.setData(0, Qt.UserRole, ("dir", e.path))
                it.setToolTip(0, e.path)
                parent.addChild(it)
                self._fill_dir(it, e.path, depth + 1)
            else:
                tag = self._file_tag(e.path)
                it = QTreeWidgetItem([e.name + (f"   {tag}" if tag else "")])
                it.setData(0, Qt.UserRole, ("file", e.path))
                it.setToolTip(0, e.path)
                if tag and tag != "[定义]":
                    it.setForeground(0, QColor("#64748B"))   # 配置文件注释用灰色
                parent.addChild(it)

    def _find_dir_item(self, path: str) -> Optional[QTreeWidgetItem]:
        want = os.path.normcase(os.path.abspath(path or "")) if path else ""
        if not want:
            return None
        stack = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
        while stack:
            n = stack.pop()
            d = n.data(0, Qt.UserRole)
            if d and d[0] == "dir" and os.path.normcase(os.path.abspath(d[1])) == want:
                return n
            stack += [n.child(i) for i in range(n.childCount())]
        return None

    def _add_missing_tables(self, root_item: QTreeWidgetItem):
        """已在 __tables__.xlsx 声明、但数据文件还没生成的表，挂到 Data 目录下。"""
        missing = [t for t in self.model.tables
                   if not any(os.path.exists(p) for p in self._input_paths(t))]
        if not missing:
            return
        holder = self._find_dir_item(self.model.data_dir) or root_item
        for t in missing:
            label = (t.input or f"{t.full_name.split('.')[-1]}.xlsx") + "   [数据表未生成]"
            it = QTreeWidgetItem([label])
            it.setData(0, Qt.UserRole, ("table", t.full_name))
            it.setToolTip(0, "这个表已经在 __tables__.xlsx 里声明，但数据文件还不存在。\n"
                             "点它可以编辑表定义与字段；点工具栏「导出数据表模板」生成文件。")
            it.setForeground(0, QColor("#B54708"))
            holder.addChild(it)

    def _refresh_tree(self):
        """左侧树 = 工程目录名 + 工程目录结构。"""
        self.tree.clear()
        root_path = self.model.root
        if not root_path:
            return
        name = os.path.basename(os.path.normpath(root_path)) or root_path
        troot = QTreeWidgetItem([name])
        troot.setData(0, Qt.UserRole, ("dir", root_path))
        troot.setToolTip(0, root_path)
        self.tree.addTopLevelItem(troot)
        if os.path.isdir(root_path):
            self._fill_dir(troot, root_path, 0)

        # dataDir 在工程目录之外时，补一个顶层节点，否则数据表点不到
        dd = self.model.data_dir
        if dd and os.path.isdir(dd) and not self._under(dd, root_path):
            dnode = QTreeWidgetItem([os.path.basename(os.path.normpath(dd)) + "   [dataDir]"])
            dnode.setData(0, Qt.UserRole, ("dir", dd))
            dnode.setToolTip(0, dd)
            self.tree.addTopLevelItem(dnode)
            self._fill_dir(dnode, dd, 0)
            dnode.setExpanded(True)

        if self.model.initialized:
            self._add_missing_tables(troot)
        else:
            it = QTreeWidgetItem(["（这个目录还没有 Luban 配置）"])
            it.setForeground(0, QColor("#B5470F"))
            troot.addChild(it)
        troot.setExpanded(True)
        self.tree.expandToDepth(1)

    # ---------------------------------------------------------------- 点击树
    def _on_tree(self, item, col=0):
        d = item.data(0, Qt.UserRole)
        if not d:
            return
        kind, val = d
        if kind == "table":
            self._load_table_editor(val)
        elif kind == "file":
            self._open_file(val)

    def _on_tree_dbl(self, item, col=0):
        """双击文件 = 用系统默认程序打开（Excel 表就用 Excel 打开）。

        单击已经负责「在中间区预览/编辑」，双击不再抢它的活，只额外唤起外部程序，
        方便直接在 Excel 里改数据。目录节点保持 Qt 默认的展开/收起。
        """
        d = item.data(0, Qt.UserRole)
        if not d:
            return
        kind, val = d
        if kind == "file":
            self._open_external(val)

    def _open_external(self, path: str):
        """用系统默认程序打开文件。"""
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "文件不存在",
                                f"{path}\n目录可能已经变化，已刷新左侧列表。")
            self._refresh_tree()
            return
        err = ""
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa: S606 —— 就是让用户用 Excel 打开
                self.status.showMessage(f"已用系统默认程序打开：{os.path.basename(path)}", 6000)
                return
            if QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
                self.status.showMessage(f"已用系统默认程序打开：{os.path.basename(path)}", 6000)
                return
            err = "系统没有为这个后缀关联程序"
        except Exception as e:
            err = str(e)
        QMessageBox.warning(self, "打不开",
                            f"{path}\n{err}\n\n可以右键「在资源管理器中显示」，再手动打开。")

    def _table_name_of_file(self, path: str) -> str:
        return self._table_for_file(path)

    def _is_data_file(self, path: str) -> bool:
        """是不是 Data 目录下的数据文件（xlsx/xls/csv）。"""
        dd = self.model.data_dir
        return bool(dd) and self._under(path, dd) and \
            os.path.splitext(path)[1].lower() in (".xlsx", ".xls", ".xlsm", ".csv")

    def _open_file(self, path: str):
        """按文件类型分流：配置表 -> 可编辑；定义文件 -> 只读摘要；其它 -> 只读预览。"""
        if not os.path.exists(path):
            QMessageBox.warning(self, "文件不存在",
                                f"{path}\n目录可能已经变化，已刷新左侧列表。")
            self._refresh_tree()
            return
        base = os.path.basename(path).lower()
        if base in self.DEF_FILES:
            self._show_def_summary(base)
            return
        if self._is_def_file(path):
            self._show_def_file_summary(path)
            return
        name = self._table_for_file(path)
        if name:
            self._show_table_preview(name, path)
            return
        if self._is_data_file(path):
            # 别让用户以为「配置表不能编辑了」——说清楚为什么，以及怎么变成可编辑
            decl = "、".join(t.full_name for t in self.model.tables) or "（还没有声明任何表）"
            defs = "、".join(os.path.basename(p) for p in
                             (self.model.table_files or
                              [self.model.table_file]) if p) or "（没找到表定义文件）"
            self._show_readonly(path, (
                "这个文件在 Data 目录里，但<b>没有在表定义文件里声明成配置表</b>"
                f"（当前工程的声明文件：<code>{_esc(defs)}</code>），所以只能只读预览。<br>"
                "想编辑它：点工具栏「<b>新建表</b>」，把「表名」和「数据文件」填成这个名字"
                f"（<code>{_esc(os.path.basename(path))}</code>），创建并保存后它就能编辑了。<br>"
                f"当前已声明的表：{_esc(decl)}"))
            return
        self._show_readonly(path)

    def _tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item:
            return
        d = item.data(0, Qt.UserRole)
        if not d:
            return
        kind, val = d
        path = val if kind == "file" else ""
        if kind == "table":
            tname = val
        elif kind == "file":
            tname = self._table_name_of_file(val)
        else:
            tname = ""
        m = QMenu(self)
        a_open = m.addAction("用系统默认程序打开") if path else None
        a_del = m.addAction("删除该表定义…") if tname else None
        a_show = m.addAction("在资源管理器中显示")
        a_ref = m.addAction("刷新目录树")
        r = m.exec(self.tree.viewport().mapToGlobal(pos))
        if a_open is not None and r is a_open:
            self._open_external(path)
        elif a_del is not None and r is a_del:
            self._delete_table(tname)
        elif r is a_show:
            self.act_show_folder(path)
        elif r is a_ref:
            self._reindex_inputs()
            self._refresh_tree()

    # ---------------------------------------------------------------- 只读页面
    def _show_page_html(self, html: str):
        if not self._confirm_leave_editor():
            return
        self.page_file.setHtml(html)
        self.editor_stack.setCurrentWidget(self.page_file)
        self._show_right_placeholder(
            "中间区现在是只读预览。点左侧 Data 目录里的配置文件，"
            "这里会列出该类型的所有配置方式、表头模板与相关知识点。")

    def _show_def_page(self, html: str, path: str):
        """渲染到「自定义类型预览页」（底部带「新增自定义类型」按钮）。"""
        if not self._confirm_leave_editor():
            return
        self._def_path = path
        self.view_def.setHtml(html)
        self.editor_stack.setCurrentWidget(self.page_def)
        self._show_right_placeholder(
            "中间区是定义文件预览。底部「新增自定义类型」可直接新建结构 / 枚举。")

    def _table_for_def_file(self, path: str) -> Optional[str]:
        """定义文件是否表专用：模块名.xml / 配置表名.xml → 对应表名；公共 → None。"""
        if not path or not path.lower().endswith(".xml"):
            return None
        stem = os.path.splitext(os.path.basename(path))[0]
        if stem == "common":
            return None
        for t in self.model.tables:
            if "." in t.full_name and t.full_name.split(".")[0] == stem:
                return t.full_name
        for t in self.model.tables:
            if os.path.splitext(os.path.basename(t.input or ""))[0] == stem:
                return t.full_name
            if t.full_name.split(".")[-1] == stem:
                return t.full_name
        return None

    def _on_add_type(self):
        """预览定义文件时的「新增自定义类型」：根据文件类型默认公共 / 专用。"""
        path = getattr(self, "_def_path", "") or ""
        table_name = self._table_for_def_file(path)
        d = CreateTypeDialog(self, self.model, table_name)
        if d.exec() != QDialog.Accepted:
            return
        obj = d.build()
        is_bean = isinstance(obj, BeanDef)
        save_path = None
        if d.table_scope() and table_name:
            t0 = self.model.find_table(table_name)
            inp = t0.input if (t0 and t0.input) else ""
            save_path = exceldef.table_def_file(self.model, table_name, inp)
        ok, msg = (exceldef.save_bean(self.model, obj, save_path) if is_bean
                   else exceldef.save_enum(self.model, obj, save_path))
        if not ok:
            QMessageBox.warning(self, "新建失败", msg)
            return
        exceldef.load_project(self.model)
        self._refresh_tree()
        # 重新渲染当前预览，让新类型立即出现
        if os.path.exists(path):
            self._open_file(path)
        written = os.path.basename(save_path or (exceldef.public_bean_file(self.model)
                                                if is_bean else exceldef.public_enum_file(self.model)) or "")
        self.status.showMessage(f"已创建 {obj.full_name} → {written}", 8000)
        self.toast(f"已创建 {obj.full_name} → {written}")

    def _show_overview(self):
        m = self.model
        if not m.initialized:
            # 非完整工程：预览区中间显示「初始化工程」按钮
            self.editor_stack.setCurrentWidget(self.page_init)
            self._show_right_placeholder("这个目录还没有 Luban 配置，点中间按钮一键补全。")
            return
        name = os.path.basename(os.path.normpath(m.root)) or m.root
        h = [f"<h3>{_esc(name)}　<span style='color:#64748B;font-weight:normal'>工程概览</span></h3>",
             f"<p>工程目录：<code>{_esc(m.root)}</code></p>"]
        h.append("<ul>"
                 f"<li>配置表 <b>{len(m.tables)}</b> 张</li>"
                 f"<li>结构 bean <b>{len(m.beans)}</b> 个</li>"
                 f"<li>枚举 enum <b>{len(m.enums)}</b> 个</li>"
                 f"<li>数据目录：<code>{_esc(m.data_dir)}</code></li>"
                 f"<li>导出分组：<code>{_esc('、'.join(m.groups) or '（未读到）')}</code></li>"
                 f"<li>导出目标：<code>{_esc('、'.join(m.targets) or '（未读到）')}</code></li>"
                 "</ul>")
        if m.tables:
            rows = [["表名", "mode", "主键", "数据文件", "分组", "注释"]]
            for t in m.tables:
                rows.append([t.full_name, t.real_mode, t.index, t.input, t.group, t.comment])
            h.append("<h4>配置表一览</h4>" + _html_table(rows))
        h.append("<p style='color:#64748B'>在左侧点开 <b>Data 目录里的配置表文件</b>即可编辑"
                 "（改动标红，底部保存）；点其它文件只做只读预览。</p>")
        self._show_page_html("".join(h))

    def _show_def_summary(self, fname: str):
        m = self.model
        pathmap = {"__tables__.xlsx": m.table_file, "__beans__.xlsx": m.bean_file,
                   "__enums__.xlsx": m.enum_file}
        target = pathmap.get(fname, "")
        h = (f"<h3>{_esc(fname)}　"
             f"<span style='color:#64748B;font-weight:normal'>定义文件 · 只读</span></h3>"
             f"<p style='color:#64748B'><code>{_esc(target)}</code></p>")

        def same(p):
            return (os.path.normcase(os.path.abspath(p or ""))
                    == os.path.normcase(os.path.abspath(target or "")))

        if fname == "__tables__.xlsx":
            h += "<p>表的声明：表名 / 行结构类型 / 主键 / 表模式 / 数据文件 / 导出分组。</p>"
            tables = [t for t in m.tables if same(getattr(t, "_file", ""))]
            if tables:
                rows = [["full_name", "value_type", "index", "mode", "input", "group",
                         "从表头读字段", "注释"]]
                for t in tables:
                    rows.append([t.full_name, t.value_type, t.index, t.real_mode, t.input,
                                 t.group, t.read_schema_from_file, t.comment])
                h += _html_table(rows)
            else:
                h += "<p>（这个文件里还没有表定义）</p>"
            h += ("<p style='color:#64748B'>要改某张表，请在左侧点开它在 Data 目录下的数据文件；"
                  "标红的改动会用底部「保存」写回这个文件。</p>")
        elif fname == "__beans__.xlsx":
            h += "<p>复合结构（bean）及其字段。</p>"
            beans = [b for b in m.beans if same(getattr(b, "_file", ""))]
            if not beans:
                h += "<p>（这个文件里没有 bean 定义）</p>"
            for b in beans:
                tag = "公共" if b.scope == "common" else "表专用"
                h += (f"<h4>{_esc(b.full_name)}　<span style='color:#64748B;font-weight:normal'>"
                      f"[{tag}]　父类型 {_esc(b.parent or '无')}　sep {_esc(b.sep or '默认')}"
                      f"　{_esc(b.comment or '')}</span></h4>")
                if b.fields:
                    h += _html_table([["字段名", "类型", "分组", "注释"]] +
                                     [[f.name, f.type, f.group, f.comment] for f in b.fields])
                else:
                    h += "<p>（无字段）</p>"
        else:
            h += "<p>枚举项：name / alias / value。</p>"
            enums = [e for e in m.enums if same(getattr(e, "_file", ""))]
            if not enums:
                h += "<p>（这个文件里没有枚举定义）</p>"
            for e in enums:
                tag = "公共" if e.scope == "common" else "表专用"
                h += (f"<h4>{_esc(e.full_name)}　<span style='color:#64748B;font-weight:normal'>"
                      f"[{tag}]　标志位 {_esc(e.flags or 'false')}　{_esc(e.comment or '')}"
                      f"</span></h4>")
                if e.items:
                    h += _html_table([["项名", "别名", "值", "说明"]] +
                                     [[i.name, i.alias, i.value, i.comment] for i in e.items])
                else:
                    h += "<p>（无枚举项）</p>"
        if fname == "__tables__.xlsx":
            self._show_page_html(h)
        else:
            self._show_def_page(h, target)

    def _show_def_file_summary(self, path: str):
        """解析定义文件（xml）里的结构按顺序显示，而不是显示源文件内容。"""
        h = [f"<h3>{_esc(os.path.basename(path))}　"
             f"<span style='color:#64748B;font-weight:normal'>定义文件 · 只读</span></h3>",
             f"<p style='color:#64748B'><code>{_esc(path)}</code></p>"]
        tables, beans, enums = [], [], []
        if path.lower().endswith(".xml"):
            try:
                tables, beans, enums = xmldef.read_defs(path)
            except Exception as e:
                h.append(f"<p style='color:#B42318'>解析失败：{_esc(str(e))}</p>")
        if tables:
            h.append("<h4>表声明</h4>")
            h.append(_html_table(
                [["表名", "value", "input", "index", "mode", "group", "注释"]] +
                [[t.full_name, t.value_type, t.input, t.index, t.mode, t.group, t.comment]
                 for t in tables]))
        for b in beans:
            h.append(f"<h4>{_esc(b.full_name)}　<span style='color:#64748B;font-weight:normal'>"
                     f"结构　父类型 {_esc(b.parent or '无')}　sep {_esc(b.sep or '默认')}　"
                     f"{_esc(b.comment or '')}</span></h4>")
            if b.fields:
                h.append(_html_table([["字段名", "类型", "分组", "注释"]] +
                                     [[f.name, f.type, f.group, f.comment] for f in b.fields]))
            else:
                h.append("<p>（无字段）</p>")
        for e in enums:
            h.append(f"<h4>{_esc(e.full_name)}　<span style='color:#64748B;font-weight:normal'>"
                     f"枚举　标志位 {_esc(e.flags or 'false')}　{_esc(e.comment or '')}</span></h4>")
            if e.items:
                h.append(_html_table([["项名", "别名", "值", "说明"]] +
                                     [[i.name, i.alias, i.value, i.comment] for i in e.items]))
            else:
                h.append("<p>（无枚举项）</p>")
        if not tables and not beans and not enums:
            h.append("<p>（这个定义文件里没有解析到表 / 结构 / 枚举）</p>")
        if os.path.basename(path).lower() == "tables.xml":
            self._show_page_html("".join(h))
        else:
            self._show_def_page("".join(h), path)

    def _show_table_preview(self, name: str, path: str):
        """配置表只读预览：显示表定义、字段与数据内容，底部「编辑」进入可编辑状态。"""
        t = self.model.find_table(name)
        if not t:
            return
        self._pending_table = name
        self._pending_path = path
        h = [f"<h3>{_esc(name)}　<span style='color:#64748B;font-weight:normal'>配置表预览</span></h3>",
             f"<p style='color:#64748B'><code>{_esc(path)}</code></p>"]
        h.append("<h4>表定义</h4>")
        h.append(_html_table(
            [["导出表名", "行结构类型", "从表头推断", "数据文件", "主键", "模式", "分组", "注释"]] +
            [[t.full_name, t.value_type, t.read_schema_from_file, t.input, t.index,
              t.real_mode, t.group, t.comment]]))
        fs = self.fields.get(name, [])
        if fs:
            h.append("<h4>字段</h4>")
            h.append(_html_table([["字段名", "类型", "分组", "注释"]] +
                                 [[f.name, f.type, f.group, f.comment] for f in fs]))
        if path and os.path.exists(path):
            is_xlsx = path.lower().endswith((".xlsx", ".xlsm"))
            total = self._xlsx_row_count(path) if is_xlsx else 0
            show_all = self._show_all_data
            max_rows = 999999 if show_all else 100
            title = "数据内容"
            if is_xlsx and not show_all and total > 100:
                title += ("　<a href='#showall' style='font-weight:normal;font-size:12px'>"
                          f"显示全部（共 {total} 行）</a>")
            h.append(f"<h4>{title}</h4>")
            if is_xlsx:
                h.append(self._preview_xlsx(path, max_rows=max_rows))
            else:
                h.append(self._preview_text(path))
        self.view_table_preview.setHtml("".join(h))
        self.editor_stack.setCurrentWidget(self.page_preview)
        self._show_right_placeholder("这是配置表只读预览。点底部「编辑」进入可编辑状态。")

    def _on_edit_table(self):
        """从预览进入可编辑状态。"""
        name = self._pending_table
        if name:
            self._load_table_editor(name)

    def _on_preview_anchor(self, url):
        """预览页里的链接点击：「显示全部」重新渲染完整数据内容。"""
        if url.toString() == "#showall":
            self._show_all_data = True
            name, path = self._pending_table, self._pending_path
            if name and path:
                self._show_table_preview(name, path)

    def _show_readonly(self, path: str, extra: str = ""):
        ext = os.path.splitext(path)[1].lower()
        try:
            size = os.path.getsize(path)
        except OSError:
            size = 0
        h = [f"<h3>{_esc(os.path.basename(path))}　"
             f"<span style='color:#64748B;font-weight:normal'>只读预览</span></h3>",
             f"<p style='color:#64748B'><code>{_esc(path)}</code><br>{size:,} 字节</p>"]
        if extra:
            h.append(f"<p style='background:#FFF7E6;border:1px solid #F5D9A8;"
                     f"border-radius:6px;padding:8px 10px'>{extra}</p>")
        if ext in (".xlsx", ".xlsm"):
            h.append(self._preview_xlsx(path))
        elif ext in self.TEXT_EXT or size <= 262144:
            h.append(self._preview_text(path))
        else:
            h.append("<p>这个文件类型没有内置预览。点右上「显示文件夹」，"
                     "用系统程序打开它。</p>")
        self._show_page_html("".join(h))

    @staticmethod
    def _preview_text(path: str, max_bytes: int = 200000, max_lines: int = 4000) -> str:
        try:
            with open(path, "rb") as f:
                data = f.read(max_bytes + 1)
        except OSError as e:
            return f"<p>读不了这个文件：{_esc(str(e))}</p>"
        text = None
        for enc in ("utf-8-sig", "gbk"):
            try:
                text = data[:max_bytes].decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            return "<p>这个文件不是文本（或编码无法识别），没有内置预览。</p>"
        all_lines = text.splitlines()
        lines = all_lines[:max_lines]
        more = "（内容过长，只显示前 4000 行）" if len(all_lines) > max_lines else ""
        return (f"<pre style='font-size:12px'>{_esc(chr(10).join(lines))}</pre>"
                f"<p style='color:#64748B'>{_esc(more)}</p>")

    @staticmethod
    def _preview_xlsx(path: str, max_rows: int = 100, max_cols: int = 24) -> str:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, data_only=True)
        except Exception as e:
            return f"<p>打不开这个工作簿：{_esc(str(e))}</p>"
        try:
            sheets = list(wb.sheetnames)
            if not sheets:
                return "<p>（空工作簿）</p>"
            ws = wb[sheets[0]]
            # 合并单元格（除左上角外）标记为 PAD，让多列字段名跨列染色
            merged_after = set()
            for rng in ws.merged_cells.ranges:
                for rr in range(rng.min_row, rng.max_row + 1):
                    for cc in range(rng.min_col, rng.max_col + 1):
                        if rr != rng.min_row or cc != rng.min_col:
                            merged_after.add((rr, cc))
            grid: List[List[str]] = []
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows,
                                                 max_col=max_cols, values_only=True)):
                # 跳过整行为空的行（表头行至少有 ##var；数据行整行空 = 空行）
                if all(v is None or str(v).strip() == "" for v in row):
                    continue
                grid_row = []
                for j, v in enumerate(row):
                    if (i + 1, j + 1) in merged_after:
                        grid_row.append(typecatalog.PAD)
                    else:
                        grid_row.append("" if v is None else str(v))
                grid.append(grid_row)
            body = _html_table(grid) if grid else "<p>（这张表是空的）</p>"
            total = ws.max_row or 0
            if total <= max_rows:
                note = f"共 {total} 行 × {max_cols} 列"
            else:
                note = f"前 {max_rows} 行 × {max_cols} 列（共 {total} 行）"
            return (f"<p style='color:#64748B'>sheet：{_esc('、'.join(sheets))}"
                    f"（只预览第一张，{note}）</p>" + body)
        finally:
            try:
                wb.close()
            except Exception:
                pass

    @staticmethod
    def _xlsx_row_count(path: str) -> int:
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, data_only=True)
        except Exception:
            return 0
        try:
            if not wb.sheetnames:
                return 0
            return wb[wb.sheetnames[0]].max_row or 0
        finally:
            try:
                wb.close()
            except Exception:
                pass

    # ---------------------------------------------------------------- 表编辑（缓冲 + 提交）
    @staticmethod
    def _copy_field(f: FieldDef) -> FieldDef:
        return FieldDef.from_dict(f.to_dict())

    def _select_table(self, name: str):
        self._load_table_editor(name)

    def _has_unsaved(self) -> bool:
        if not self._editing:
            return False
        for k in self.FORM_KEYS:
            if (self._form.get(k, "") or "") != (self._base_form.get(k, "") or ""):
                return True
        if len(self._edit_fields) != len(self._base_fields):
            return True
        for i, f in enumerate(self._edit_fields):
            b = self._base_fields[i]
            for attr in ("name", "type", "fill_style", "group", "comment"):
                if (getattr(f, attr, "") or "") != (getattr(b, attr, "") or ""):
                    return True
        return False

    def _confirm_leave_editor(self) -> bool:
        """离开当前配置表前，给未保存的改动一个交代。返回 False 表示别切。"""
        if not self._has_unsaved():
            return True
        box = QMessageBox(self)
        box.setWindowTitle("有未保存的修改")
        box.setIcon(QMessageBox.Warning)
        box.setText(f"配置表 {self.current} 有未保存的修改。")
        box.setInformativeText("保存后继续，或者放弃这些修改。")
        b_save = box.addButton("保存后继续", QMessageBox.AcceptRole)
        b_drop = box.addButton("放弃修改", QMessageBox.DestructiveRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(b_save)
        box.exec()
        clicked = box.clickedButton()
        if clicked is b_save:
            self._commit_editor()
            return self._save_definitions(self.current)
        return clicked is b_drop

    def _load_table_editor(self, name: str):
        """把表载入编辑器：读进缓冲区，并记录基线用于标红与「取消修改」。"""
        t = self.model.find_table(name)
        if not t:
            return
        if self._editing and self.current != name and not self._confirm_leave_editor():
            return
        self.current = name
        self._editing = True
        self._form = {k: (getattr(t, k, "") or "") for k in self.FORM_KEYS}
        self._base_form = dict(self._form)
        cur = self.fields.get(name, [])
        self._edit_fields = [self._copy_field(f) for f in cur]
        self._base_fields = [self._copy_field(f) for f in cur]
        # 分解 full_name / input：模块名、配置表名、导出表名（与新建表一致）
        full = self._form["full_name"] or ""
        module = full.split(".")[0] if "." in full else ""
        config = os.path.basename(self._form["input"] or "")
        # 表声明文件类型：由该表当前所在文件决定（只读）
        tfile = getattr(t, "_file", "") or ""
        if tfile.lower().endswith(".xlsx"):
            self.lab_def_type.setText("Excel Schema（__tables__.xlsx）")
        else:
            self.lab_def_type.setText("XML Schema")
        self._loading = True
        self.ed_input.setText(config)
        self.ed_module.setCurrentText(module)
        self.ed_full.setText(full)
        self.cb_read.set_code(self._form["read_schema_from_file"])
        self.ed_index.setText(self._form["index"])
        self.cb_mode.set_code(self._form["mode"])
        self.cb_group.set_code(self._form["group"])
        self.ed_comment.setText(self._form["comment"])
        self._loading = False
        self._reload_fields_table()
        self.editor_stack.setCurrentWidget(self.page_table)
        self._repaint_dirty()
        self._refresh_right()

    def _reload_fields_table(self, tbl: QTableWidget = None,
                            fields: List[FieldDef] = None):
        tbl = tbl or self.tbl_fields
        fields = self._edit_fields if fields is None else fields
        prev = self._loading
        self._loading = True
        tbl.setRowCount(0)
        for f in fields:
            self._append_field_row(f, tbl)
        self._loading = prev

    # ---------------- 字段表（编辑器与新建表页共用同一套） ----------------
    def _nt_current_name(self) -> str:
        """新建表页里**正在填**的表名。

        表还没创建出来，但「归属：当前表专用」和类型下拉都得用它 ——
        以前这里固定返回 None，于是新建表页里建自定义类型时
        「当前表专用」是灰的，也没法引用本表刚建的类型。
        """
        return (self.nt_name.text() or "").strip() if hasattr(self, "nt_name") else ""

    def _field_ctx(self, tbl: QTableWidget) -> Tuple[List[FieldDef], Optional[str]]:
        """字段表控件 -> (字段列表, 该表所属表名)。新建表页用正在填的表名。"""
        if tbl is self.nt_fields:
            return self._nt_buf, (self._nt_current_name() or None)
        return self._edit_fields, self.current

    @staticmethod
    def _sync_fields_from_cells(tbl: QTableWidget, fields: List[FieldDef]):
        for r in range(min(tbl.rowCount(), len(fields))):
            for c, attr in ((COL_NAME, "name"), (COL_COMMENT, "comment")):
                it = tbl.item(r, c)
                if it is not None:
                    setattr(fields[r], attr, it.text().strip())
            # 分组现在是下拉控件，取不到 item，从控件里读
            cell = tbl.cellWidget(r, COL_GROUP)
            cb = cell.findChild(QComboBox) if cell is not None else None
            if cb is not None:
                fields[r].group = (cb.code() or "").strip()
            # 单元格格式同理（基础类型的格子是「—」，没有下拉）
            fcell = tbl.cellWidget(r, COL_FMT)
            fcb = fcell.findChild(QComboBox) if fcell is not None else None
            if fcb is not None:
                fields[r].fmt = ((fcb.code() or "").strip().lower()
                                 if hasattr(fcb, "code") else "")

    def _style_key(self, f: FieldDef) -> str:
        """当前该选中哪一项配置方式：显式选的优先，其次看类型串里挂了哪个标签。"""
        s = (f.fill_style or "auto").strip()
        if s and s != "auto":
            return s
        _core, tags = typecatalog.split_tags(f.type)
        for tg in tags:
            key = tg.lstrip("#").split("=")[0]
            if key in typecatalog.STYLE_TAG_KEYS:
                return "tag:" + key
        if typecatalog.parse_type(f.type)["kind"] == "container":
            return "sep"
        return template.auto_style(self.model, f.type)

    def _append_field_row(self, f: FieldDef, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        _fields, table_name = self._field_ctx(tbl)
        prev = self._loading
        self._loading = True
        r = tbl.rowCount()
        tbl.insertRow(r)
        tbl.setItem(r, COL_NAME, QTableWidgetItem(f.name))
        tbl.item(r, COL_NAME).setToolTip("字段名：英文，建议 xx_yy；以 # 开头是注释列，不导出")

        cb = QComboBox()
        cb.setEditable(True)
        cb.setToolTip("基础类型 / 容器（list、array、set、map）/ 公共类型 / "
                      "当前表专用类型；最后一项可新建自定义类型")
        cb.addItems(_type_names(self.model, table_name))
        cb.addItem(ADD_NEW_TYPE)
        tbl.setCellWidget(r, COL_TYPE, cb)
        cb.setCurrentText(typecatalog.type_label(self.model, f.type))
        cb.currentIndexChanged.connect(
            lambda _=None, w=cb, t=tbl: self._on_type_combo(self._row_of(w), t))
        le = cb.lineEdit()
        if le is not None:
            # 下拉可以手写（容器、带标签的类型），回车 / 失焦时同步进缓冲区
            le.editingFinished.connect(
                lambda t=tbl, w=cb: self._on_type_text(self._row_of(w), t))

        # 子类型：容器才有内容（list/array/set 给元素；map 给 key + value）
        tbl.setCellWidget(r, COL_SUB, QWidget())
        self._fill_subtype_cell(r, f.type, tbl)

        cb2 = QComboBox()
        cb2.setToolTip("这一列在数据表里怎么展开（流式 / sep / 列限定 / 多行），"
                       "以及校验标签（引用 / 范围 / 路径 / 本地化），随「类型」变化")
        tbl.setCellWidget(r, COL_STYLE, cb2)
        self._fill_style_combo(r, f, tbl)
        cb2.currentIndexChanged.connect(
            lambda _=None, w=cb2, t=tbl: self._on_style_combo(self._row_of(w), t))

        # 单元格格式：stream / lite / json / lua（只有结构 / 容器才有意义）
        tbl.setCellWidget(r, COL_FMT, QWidget())
        self._fill_format_cell(r, f, tbl)

        tbl.setCellWidget(r, COL_GROUP, QWidget())
        self._fill_group_cell(r, f.group, tbl)
        tbl.setItem(r, COL_COMMENT, QTableWidgetItem(f.comment))
        b = QPushButton("…")
        b.setToolTip("打开类型助手：挑类型、勾标签、看该类型的所有配置方式与示例")
        b.clicked.connect(lambda _=False, w=b, t=tbl: self._pick_type(self._row_of(w), t))
        tbl.setCellWidget(r, COL_HELP, b)
        tbl.setRowHeight(r, 30)
        self._loading = prev

    # ---------------- 子类型（容器专用） ----------------
    def _fill_subtype_cell(self, r: int, code: str, tbl: QTableWidget = None):
        """重建第 r 行的「子类型」格：随类型是容器与否变化。

        list / array / set → 一个「元素」下拉；map → 「键」+「值」两个下拉。
        下拉内容与「类型」下拉同源（不含容器 —— Luban 不允许容器套容器）。
        """
        tbl = tbl or self.tbl_fields
        if not (0 <= r < tbl.rowCount()):
            return
        _fields, table_name = self._field_ctx(tbl)
        p = typecatalog.parse_type(code)
        prev = self._loading
        self._loading = True
        cell = QWidget()
        lay = QHBoxLayout(cell)
        lay.setContentsMargins(5, 0, 5, 0)
        lay.setSpacing(4)
        if p["kind"] != "container":
            lab = QLabel("—")
            lab.setStyleSheet("color:#94A3B8;")
            lab.setToolTip("只有容器类型（list / array / set / map）才需要填子类型")
            lay.addWidget(lab)
            lay.addStretch(1)
        else:
            choices = _type_choices(self.model, table_name, containers=False)
            names = ([("键", p["elem"] or "int"), ("值", p["value"] or "string")]
                     if p["container"] == "map" else [("元素", p["elem"] or "int")])
            for label, cur in names:
                lab = QLabel(label)
                lab.setStyleSheet("color:#475467;")
                lab.setToolTip("map 的 key 一般是基础类型或 enum；"
                               "value 可以是基础类型 / 枚举 / 结构"
                               if p["container"] == "map" else "容器里每个元素的类型")
                cb = ValueCombo(choices, editable=True)
                cb.setToolTip("与「类型」下拉同一份清单；容器不能再套容器。"
                              "当前完整写法：%s" % cur)
                cb.set_code(cur, typecatalog.elem_label(self.model, cur))
                cb.currentIndexChanged.connect(
                    lambda _=None, w=cb, t=tbl: self._on_subtype(self._row_of(w), t))
                le = cb.lineEdit()
                if le is not None:
                    le.editingFinished.connect(
                        lambda t=tbl, w=cb: self._on_subtype(self._row_of(w), t))
                lay.addWidget(lab)
                lay.addWidget(cb, 1)
        tbl.setCellWidget(r, COL_SUB, cell)
        self._loading = prev

    def _on_subtype(self, r: int, tbl: QTableWidget = None):
        """子类型改了 → 重算类型串（保留容器上已有的 #sep= / #size=），并刷新周边。"""
        tbl = tbl or self.tbl_fields
        if self._loading or r < 0:
            return
        fields, _table_name = self._field_ctx(tbl)
        if not (0 <= r < len(fields)):
            return
        f = fields[r]
        p = typecatalog.parse_type(f.type)
        if p["kind"] != "container":
            return
        cell = tbl.cellWidget(r, COL_SUB)
        combos = cell.findChildren(QComboBox) if cell is not None else []
        vals = [(c.code() or "").strip() for c in combos]
        if not vals:
            return
        cont = p["container"]
        if cont == "map":
            elem = vals[0] or "int"
            value = (vals[1] if len(vals) > 1 else "") or "string"
            rest = f"{elem},{value}"
        else:
            rest = vals[0] or "int"
        core = f"({cont}{p['sep']}),{rest}" if p["sep"] else f"{cont},{rest}"
        if core == f.type:
            return
        f.type = core
        self._set_type_text(r, core, tbl)
        self._fill_style_combo(r, f, tbl)
        self._after_field_change(tbl)

    def _row_of(self, w) -> int:
        """由单元格控件反查当前行号，避免行号漂移后改错字段。"""
        for tbl in (self.tbl_fields, self.nt_fields):
            for r in range(tbl.rowCount()):
                for c in range(tbl.columnCount()):
                    cw = tbl.cellWidget(r, c)
                    if cw is None:
                        continue
                    if cw is w or (w is not None and w.parent() is cw):
                        return r
        return -1

    def _after_field_change(self, tbl: QTableWidget):
        """字段改了以后统一刷新：编辑器要重算标红，新建表页只刷新右侧。"""
        if tbl is self.nt_fields:
            self._refresh_new_table_right()
        else:
            self._repaint_dirty()
            self._refresh_right()

    def _on_field_selected(self):
        """在编辑器里点到另一行字段时，右侧「配置方式」要跟着换 ——

        否则选中的是 A 行、右边讲的是 B 行，看示例时会一直对不上。
        """
        if self._loading or not self._editing:
            return
        self._refresh_right()

    # ---------------- 单元格格式（stream / lite / json / lua） ----------------
    @staticmethod
    def _is_complex(model: ProjectModel, type_str: str) -> bool:
        """结构 / 容器才有「单元格格式」；基础类型固定是 stream。"""
        p = typecatalog.parse_type(type_str)
        if p["kind"] == "container":
            return True
        return typecatalog.resolve_kind(model, p["base"]) == "bean"

    def _fill_format_cell(self, r: int, f: FieldDef, tbl: QTableWidget = None):
        """第 r 行的「单元格格式」格。

        Luban 的 #format= 只对**非原子**字段有意义（官方文档：写在字段名上，
        写在 ##type 上不生效）。所以基础类型这一格显示「—」，不给选。
        """
        tbl = tbl or self.tbl_fields
        if not (0 <= r < tbl.rowCount()):
            return
        prev = self._loading
        self._loading = True
        cell = QWidget()
        lay = QHBoxLayout(cell)
        lay.setContentsMargins(4, 0, 4, 0)
        lay.setSpacing(0)
        if not self._is_complex(self.model, f.type):
            lab = QLabel("—")
            lab.setStyleSheet("color:#94A3B8;")
            lab.setToolTip("单元格格式只对结构 / 容器有意义；基础类型固定按 stream 读")
            lay.addWidget(lab)
            lay.addStretch(1)
        else:
            cb = ValueCombo(FMT_ITEMS, editable=False)
            cb.setToolTip(
                "这一格怎么写：\n"
                "· stream（默认）：按左边「配置方式」流式 / sep / 分列填\n"
                "· lite：一格一个花括号 {1,2,3}\n"
                "· json：一格一个 JSON {\"x\":1,\"y\":2}\n"
                "· lua：一格一个 Lua 表 {x=1,y=2}\n"
                "选 lite/json/lua 后，整格写一个值，配置方式会自动禁用。")
            cb.set_code(f.fmt or "")
            cb.currentIndexChanged.connect(
                lambda _=None, w=cb, t=tbl: self._on_format(w, t))
            lay.addWidget(cb, 1)
        tbl.setCellWidget(r, COL_FMT, cell)
        self._loading = prev

    def _on_format(self, w, tbl: QTableWidget = None):
        """选了单元格格式 → 写回字段，并让「配置方式」跟着可用 / 禁用。"""
        if self._loading:
            return
        tbl = tbl or self.tbl_fields
        r = self._row_of(w)
        if r < 0:
            return
        fields, _table_name = self._field_ctx(tbl)
        if not (0 <= r < len(fields)):
            return
        fmt = (w.code() or "").strip().lower()
        if fmt == "stream":
            fmt = ""
        if fields[r].fmt == fmt:
            return
        fields[r].fmt = fmt
        self._fill_style_combo(r, fields[r], tbl)
        self._after_field_change(tbl)

    # ---------------- 分组（带中文名的下拉） ----------------
    def _fill_group_cell(self, r: int, group: str, tbl: QTableWidget = None):
        """第 r 行的「分组」格：下拉选，选项来自工程 luban.conf 的 groups。

        直接写 c / s 对新手不友好，所以显示成「c（客户端）」。要进多个分组时
        下拉本身可编辑，手写 `c,s` 即可（回车 / 失焦生效）。
        """
        tbl = tbl or self.tbl_fields
        if not (0 <= r < tbl.rowCount()):
            return
        prev = self._loading
        self._loading = True
        cell = QWidget()
        lay = QHBoxLayout(cell)
        lay.setContentsMargins(4, 0, 4, 0)
        lay.setSpacing(0)
        cb = ValueCombo(group_items(self.model), editable=True)
        cb.setToolTip("该字段导出到哪些分组（分组名来自 luban.conf 的 groups）。\n"
                      "留空 = 各端都导出；要进多个分组就手写 c,s")
        cb.set_code(group)
        cb.currentIndexChanged.connect(
            lambda _=None, w=cb, t=tbl: self._on_group(w, t))
        le = cb.lineEdit()
        if le is not None:
            le.editingFinished.connect(
                lambda w=cb, t=tbl: self._on_group(w, t))
        lay.addWidget(cb, 1)
        tbl.setCellWidget(r, COL_GROUP, cell)
        self._loading = prev

    def _on_group(self, w, tbl: QTableWidget = None):
        """分组下拉改了 → 写回缓冲并刷新（配置方式预览里会用到 group 示例）。"""
        if self._loading:
            return
        tbl = tbl or self.tbl_fields
        r = self._row_of(w)
        if r < 0:
            return
        fields, _table_name = self._field_ctx(tbl)
        if not (0 <= r < len(fields)):
            return
        val = (w.code() or "").strip()
        if fields[r].group == val:
            return
        fields[r].group = val
        self._after_field_change(tbl)

    def _fill_style_combo(self, r: int, f: FieldDef, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        cb2 = tbl.cellWidget(r, COL_STYLE)
        if not cb2:
            return
        prev = self._loading
        self._loading = True
        cb2.blockSignals(True)
        cb2.clear()
        for key, label in typecatalog.fill_style_options(self.model, f.type):
            cb2.addItem(label, key)
        cur = self._style_key(f)
        i = cb2.findData(cur)
        cb2.setCurrentIndex(i if i >= 0 else 0)
        # 选了 lite / json / lua：整格就是一个 {...}，跟「流式 / sep / 分列」互斥，
        # 这时配置方式再选也没意义，直接禁用免得两边打架。
        use_fmt = (f.fmt or "").strip().lower() in ("lite", "json", "lua")
        cb2.setEnabled(not use_fmt)
        cb2.setToolTip(
            f"已选「{f.fmt}」单元格格式：整格写一个 {f.fmt} 值，不需要再挑配置方式"
            if use_fmt else
            "这一列在数据表里怎么展开（流式 / sep / 列限定 / 多行），"
            "以及校验标签（引用 / 范围 / 路径 / 本地化），随「类型」变化")
        cb2.blockSignals(False)
        self._loading = prev

    def _set_type_text(self, r: int, code: str, tbl: QTableWidget = None):
        """只改「类型」框里显示的文字，不重建列表（不会打断正在操作的下拉）。"""
        tbl = tbl or self.tbl_fields
        cb = tbl.cellWidget(r, COL_TYPE)
        if not cb:
            return
        prev = self._loading
        self._loading = True
        cb.blockSignals(True)
        cb.setCurrentText(typecatalog.type_label(self.model, code))
        cb.blockSignals(False)
        self._loading = prev

    def _set_type_combo(self, r: int, code: str, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        cb = tbl.cellWidget(r, COL_TYPE)
        if not cb:
            return
        _fields, table_name = self._field_ctx(tbl)
        prev = self._loading
        self._loading = True
        cb.blockSignals(True)
        cb.clear()
        cb.addItems(_type_names(self.model, table_name))
        cb.addItem(ADD_NEW_TYPE)
        cb.setCurrentText(typecatalog.type_label(self.model, code))
        cb.blockSignals(False)
        self._loading = prev
        self._fill_subtype_cell(r, code, tbl)

    def _on_form_changed(self, key: str):
        if self._loading or not self._editing:
            return
        w = self._form_widgets.get(key)
        if w is None:
            return
        if isinstance(w, ValueCombo):
            val = w.code()
        elif isinstance(w, QComboBox):
            val = w.currentText().strip()
        else:
            val = w.text()
        if key == "input":
            # 配置表名（不含模块目录）+ 模块名 → 完整 input
            self._sync_input_from_edits()
        else:
            self._form[key] = (val or "").strip()
        self._repaint_dirty()
        self._refresh_right()

    def _sync_input_from_edits(self):
        """配置表名 + 模块名 → 完整 input（模块目录/配置表名.xlsx）。"""
        config = (self.ed_input.text() or "").strip()
        if config and not os.path.splitext(config)[1]:
            config = config + ".xlsx"
        module = (self.ed_module.currentText() or "").strip()
        self._form["input"] = f"{module}/{config}" if (module and config) else config

    def _on_module_changed(self):
        """模块名变化：刷新 input 的目录部分（不改导出表名）。"""
        if self._loading or not self._editing:
            return
        self._sync_input_from_edits()
        self._repaint_dirty()
        self._refresh_right()

    def _on_field_item_changed(self, item):
        if self._loading or self._painting:
            return
        tbl = item.tableWidget()
        if tbl is None:
            return
        if tbl is self.tbl_fields and not self._editing:
            return
        fields, _table_name = self._field_ctx(tbl)
        r, c = item.row(), item.column()
        if not (0 <= r < len(fields)):
            return
        f = fields[r]
        if c == COL_NAME:
            f.name = item.text().strip()
        elif c == COL_COMMENT:
            f.comment = item.text().strip()
        else:
            return
        self._after_field_change(tbl)

    def _on_type_combo(self, r: int, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        if self._loading or r < 0:
            return
        cb = tbl.cellWidget(r, COL_TYPE)
        fields, _table_name = self._field_ctx(tbl)
        if not cb:
            return
        txt = cb.currentText()
        if txt == ADD_NEW_TYPE:
            self._create_type_inline(r, tbl)
            return
        if not (0 <= r < len(fields)):
            return
        f = fields[r]
        f.type = typecatalog.strip_label(txt) or "int"
        f.fill_style = "auto"
        if not self._is_complex(self.model, f.type):
            f.fmt = ""                              # 换成基础类型，单元格格式就没意义了
        self._fill_subtype_cell(r, f.type, tbl)     # 选到容器就要露出子类型
        self._fill_format_cell(r, f, tbl)
        self._fill_style_combo(r, f, tbl)
        self._after_field_change(tbl)

    def _on_type_text(self, r: int, tbl: QTableWidget = None):
        """类型下拉允许手写（list,Item、int#ref=… 这类），回车或失焦时落到字段上。"""
        tbl = tbl or self.tbl_fields
        if self._loading or r < 0:
            return
        cb = tbl.cellWidget(r, COL_TYPE)
        fields, _table_name = self._field_ctx(tbl)
        if not cb or not (0 <= r < len(fields)):
            return
        txt = cb.currentText().strip()
        if not txt or txt == ADD_NEW_TYPE:
            return
        code = typecatalog.strip_label(txt) or "int"
        f = fields[r]
        if code == f.type:
            return
        f.type = code
        f.fill_style = "auto"
        if not self._is_complex(self.model, f.type):
            f.fmt = ""
        self._set_type_combo(r, code, tbl)      # 回显成带中文名的形式
        self._fill_format_cell(r, f, tbl)
        self._fill_style_combo(r, f, tbl)
        self._after_field_change(tbl)

    def _create_type_inline(self, r: int, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        fields, table_name = self._field_ctx(tbl)
        d = CreateTypeDialog(self, self.model, table_name)
        if d.exec() != QDialog.Accepted:
            if 0 <= r < len(fields):
                self._set_type_combo(r, fields[r].type, tbl)
            return
        obj = d.build()
        # 「当前表专用」的类型写进该表自己的 xml（模块名.xml 或 配置表名.xml），
        # 「公共」的写进 common.xml —— 都不再碰 __beans__.xlsx 等 xlsx。
        is_bean = isinstance(obj, BeanDef)
        kind = "bean" if is_bean else "enum"
        path = None
        if d.table_scope() and table_name:
            t0 = self.model.find_table(table_name)
            inp = t0.input if (t0 and t0.input) else (
                self._nt_input_with_ext() if hasattr(self, "nt_input") else "")
            path = exceldef.table_def_file(self.model, table_name, inp)
        ok, msg = (exceldef.save_bean(self.model, obj, path) if is_bean
                   else exceldef.save_enum(self.model, obj, path))
        if not ok:
            QMessageBox.warning(self, "新建失败", msg)
            if 0 <= r < len(fields):
                self._set_type_combo(r, fields[r].type, tbl)
            return
        exceldef.load_project(self.model)
        self._refresh_tree()
        if 0 <= r < len(fields):
            f = fields[r]
            f.type = obj.full_name
            f.fill_style = "auto"
            self._set_type_combo(r, f.type, tbl)
            self._fill_style_combo(r, f, tbl)
        written = os.path.basename(path or (exceldef.public_bean_file(self.model) if is_bean
                                            else exceldef.public_enum_file(self.model)) or "")
        self.status.showMessage(f"已创建 {obj.full_name} → {written}", 8000)
        self.toast(f"已创建 {obj.full_name} → {written}")
        self._after_field_change(tbl)

    def _on_style_combo(self, r: int, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        if self._loading or r < 0:
            return
        cb2 = tbl.cellWidget(r, COL_STYLE)
        fields, _table_name = self._field_ctx(tbl)
        if not cb2 or not (0 <= r < len(fields)):
            return
        f = fields[r]
        key = cb2.currentData() or "auto"
        new_type = apply_style_to_type(f, key)
        f.fill_style = key
        if new_type != f.type:
            f.type = new_type
            self._set_type_combo(r, f.type, tbl)
        self._after_field_change(tbl)

    def _pick_type(self, r: int, tbl: QTableWidget = None):
        tbl = tbl or self.tbl_fields
        fields, table_name = self._field_ctx(tbl)
        if not (0 <= r < len(fields)):
            return
        cb = tbl.cellWidget(r, COL_TYPE)
        cur = cb.currentText() if cb else fields[r].type
        it = tbl.item(r, COL_NAME)
        name = it.text() if it else "x"
        dlg = TypePickerDialog(self, self.model, typecatalog.strip_label(name),
                              typecatalog.strip_label(cur), table_name)
        if dlg.exec() == QDialog.Accepted:
            fields[r].type = dlg.value()
            self._set_type_combo(r, fields[r].type, tbl)
            self._fill_style_combo(r, fields[r], tbl)
            self._after_field_change(tbl)

    def _add_field(self):
        if not self._editing:
            QMessageBox.information(self, "提示",
                                    "请先在左侧选中一张配置表，或用工具栏「新建表」新建一张")
            return
        f = FieldDef(name="field_%d" % (len(self._edit_fields) + 1), type="int")
        self._edit_fields.append(f)
        prev = self._loading
        self._loading = True
        self._append_field_row(f)
        self._loading = prev
        self.tbl_fields.selectRow(self.tbl_fields.rowCount() - 1)
        self._repaint_dirty()
        self._refresh_right()

    def _del_field(self):
        if not self._editing:
            return
        rows = sorted({i.row() for i in self.tbl_fields.selectedIndexes()}, reverse=True)
        for r in rows:
            if 0 <= r < len(self._edit_fields):
                self._edit_fields.pop(r)
        self._reload_fields_table()
        self._repaint_dirty()
        self._refresh_right()

    def _move(self, delta: int):
        if not self._editing:
            return
        rows = sorted({i.row() for i in self.tbl_fields.selectedIndexes()})
        if len(rows) != 1:
            return
        r = rows[0]
        n = r + delta
        if not (0 <= n < len(self._edit_fields)):
            return
        self._edit_fields[r], self._edit_fields[n] = \
            self._edit_fields[n], self._edit_fields[r]
        self._reload_fields_table()
        self.tbl_fields.selectRow(n)
        self._repaint_dirty()
        self._refresh_right()

    # ---------------------------------------------------------------- 标红 / 脏状态
    def _mark_widget(self, w, on: bool):
        if w is None:
            return
        w.setProperty("dirty", bool(on))
        w.setStyleSheet(
            "color:#B42318;border:1px solid #F0B4AE;border-radius:6px;"
            "padding:5px 8px;background:#FFF6F5;" if on else "")

    def _mark_cell(self, r: int, col: int, on: bool):
        it = self.tbl_fields.item(r, col)
        if it is None:
            return
        self._painting = True
        if on:
            it.setForeground(QBrush(QColor("#B42318")))
            it.setBackground(QBrush(QColor("#FDECEA")))
        else:
            it.setForeground(QBrush(QColor("#1F2937")))
            it.setBackground(QBrush(Qt.transparent))
        self._painting = False

    def _update_hint(self):
        if not self._editing:
            self.lab_hint.setText("")
            return
        tmp = TableDef(full_name=self._form.get("full_name", ""),
                       mode=self._form.get("mode", ""),
                       index=self._form.get("index", ""))
        self.lab_hint.setText({
            "map": "map：按主键查找，代码里是 Get(id)。index 为空时取第一个字段当主键。",
            "list": "list：列表表，只能遍历。多主键：联合写 a+b、独立索引写 a,b。",
            "one": "one：全局单例，通常只有一行，也可用纵表。",
        }.get(tmp.real_mode, ""))

    def _repaint_dirty(self):
        """按「缓冲区 vs 基线」重新算一遍，把改动过的表单项和单元格标红。

        底部保存条**只在有改动时出现**（没有改动时整条隐藏，不占位）。
        """
        if not self._editing:
            self.lab_dirty.setText("")
            self.lab_dirty.setVisible(False)
            self.bt_save_ed.setEnabled(False)
            self.bt_cancel_ed.setEnabled(False)
            self.bar_editor.setVisible(False)
            self.lab_hint.setText("")
            return
        dirty = False
        for key, w in self._form_widgets.items():
            d = (self._form.get(key, "") or "") != (self._base_form.get(key, "") or "")
            self._mark_widget(w, d)
            dirty = dirty or d

        same = len(self._edit_fields) == len(self._base_fields)
        for r, f in enumerate(self._edit_fields):
            b = self._base_fields[r] if same else None
            for col, attr in ((COL_NAME, "name"), (COL_COMMENT, "comment")):
                d = (b is None) or \
                    ((getattr(f, attr, "") or "") != (getattr(b, attr, "") or ""))
                self._mark_cell(r, col, d)
                dirty = dirty or d
            d = (b is None) or ((f.group or "") != (b.group or ""))
            self._mark_widget(self.tbl_fields.cellWidget(r, COL_GROUP), d)
            dirty = dirty or d
            d = (b is None) or ((f.fmt or "") != (b.fmt or ""))
            self._mark_widget(self.tbl_fields.cellWidget(r, COL_FMT), d)
            dirty = dirty or d
            for col, attr in ((COL_TYPE, "type"), (COL_STYLE, "fill_style")):
                d = (b is None) or \
                    ((getattr(f, attr, "") or "") != (getattr(b, attr, "") or ""))
                self._mark_widget(self.tbl_fields.cellWidget(r, col), d)
                dirty = dirty or d
        if not same:
            dirty = True

        self.lab_dirty.setText(
            "有未保存的修改（红色部分）；点右侧「保存」写进定义文件，或「取消修改」放弃"
            if dirty else "")
        self.lab_dirty.setVisible(dirty)
        self.bt_save_ed.setEnabled(dirty)
        self.bt_cancel_ed.setEnabled(dirty)
        self.bt_save_ed.setVisible(dirty)
        self.bt_cancel_ed.setVisible(dirty)
        self.bar_editor.setVisible(dirty)      # 无改动时整条隐藏
        self._update_hint()

    # ---------------------------------------------------------------- 保存 / 取消
    def _commit_editor(self) -> bool:
        """把缓冲区落到 model 与 self.fields（不写盘）。"""
        if not self._editing or not self.current:
            return False
        t = self.model.find_table(self.current)
        if not t:
            return False
        old = t.full_name
        # 先同步一次 input（防止配置表名/模块名改了但 input 键还没触发）
        self._sync_input_from_edits()
        for k in self.FORM_KEYS:
            setattr(t, k, self._form.get(k, "") or "")
        # value_type 自动跟随导出表名推导（与新建表一致，不再单独编辑）
        t.value_type = template.derive_value_type(t.full_name)
        if t.full_name != old:
            if old in self.fields:
                self.fields[t.full_name] = self.fields.pop(old)
            self.current = t.full_name
        self.fields[self.current] = [self._copy_field(f) for f in self._edit_fields]
        self._reindex_inputs()
        self._base_form = dict(self._form)
        self._base_fields = [self._copy_field(f) for f in self._edit_fields]
        return True

    def _write_header_formats(self, table_name: str,
                              fields: List[FieldDef]) -> Optional[str]:
        """把 `#format=` 写回数据表 ##var 的**字段名**上（Luban 只认这里）。

        定义文件里没有「单元格格式」这一列 —— 它本来就属于数据表表头。所以保存时
        顺手同步一次，否则重新打开工程这个设置就丢了。只改表头文字，不动数据；
        写前备份；找不到表头 / 文件不是 xlsx 就安静跳过。
        """
        t = self.model.find_table(table_name)
        if not t or not fields:
            return None
        p = template.default_input_path(self.model, t)
        if not p or not os.path.exists(p) or not p.lower().endswith(".xlsx"):
            return None
        try:
            from openpyxl import load_workbook
            wb = load_workbook(p)
            ws = wb.worksheets[0]
            vrow = None
            for r in range(1, min(ws.max_row or 1, 12) + 1):
                if str(ws.cell(row=r, column=1).value or "").strip() == "##var":
                    vrow = r
                    break
            if vrow is None:
                wb.close()
                return None
            by_name = {f.name: f for f in fields if f.name}
            changed = False
            for c in range(2, (ws.max_column or 1) + 1):
                cell = ws.cell(row=vrow, column=c)
                old = str(cell.value or "").strip()
                if not old:
                    continue
                nm, _old_fmt = typecatalog.split_format(old)
                f = by_name.get(nm)
                if f is None:
                    continue
                new = typecatalog.join_format(nm, f.fmt)
                if new != old:
                    cell.value = new
                    changed = True
            if changed:
                exceldef.backup_file(p)     # 先备份再写，和写定义文件一个规矩
                wb.save(p)
            wb.close()
            return None
        except Exception as e:
            return f"单元格格式没能写回数据表头：{e}"

    def _save_editor(self):
        if not self._editing:
            return
        self._commit_editor()
        if self._save_definitions(self.current):
            msg = self._write_header_formats(self.current,
                                             self.fields.get(self.current, []))
            if msg:
                self.status.showMessage(msg, 8000)
            self._load_table_editor(self.current)
            self.status.showMessage("配置表已保存", 6000)
            self.toast("✓ 已保存：Schema 定义与数据表表头都写回了")

    def _cancel_editor(self):
        if not self._editing or not self.current:
            return
        if not self.model.find_table(self.current):
            return
        self._load_table_editor(self.current)
        self.status.showMessage("已放弃未保存的修改", 5000)

    def _current_fields(self) -> List[FieldDef]:
        if self._editing and self.current:
            return self._edit_fields
        return self.fields.get(self.current or "", [])

    def _fields_snapshot(self) -> Dict[str, List[FieldDef]]:
        """自检用：所有表的字段，其中当前编辑中的表用缓冲区（含未保存改动）。"""
        d = {k: list(v) for k, v in self.fields.items()}
        if self._editing and self.current:
            d[self.current] = list(self._edit_fields)
        return d

    # ---------------------------------------------------------------- 右栏
    def _show_right_placeholder(self, msg: str):
        """中间区不是配置表编辑器时，右侧上下文面板给个说明，别留上一张表的残影。"""
        html = f"<p style='color:#64748B'>{_esc(msg)}</p>"
        self.view_opts.setHtml(html)
        self.view_tpl.setHtml(html)
        self._k_seed = ""

    def _refresh_right(self):
        if not self._editing or not self.current:
            return
        tmp = TableDef(full_name=self._form.get("full_name", ""),
                       mode=self._form.get("mode", ""),
                       index=self._form.get("index", ""))
        fs = self._edit_fields
        row = self.tbl_fields.currentRow()
        r = row if 0 <= row < len(fs) else (0 if fs else -1)
        cur_style = ""
        if r >= 0:
            cb2 = self.tbl_fields.cellWidget(r, COL_STYLE)
            cur_style = (cb2.currentData() if cb2 else "") or ""

        h = (f"<h3>{_esc(tmp.full_name)}　"
             f"<span style='color:#64748B;font-weight:normal'>mode={_esc(tmp.real_mode)}"
             f"　字段 {len(fs)} 个</span></h3>")
        h += (f"<p>value_type=<code>"
              f"{_esc(template.derive_value_type(self._form.get('full_name', '')))}</code>　"
              f"从数据表头读字段=<code>"
              f"{_esc(self._form.get('read_schema_from_file') or '空')}</code></p>")

        if r >= 0:
            f = fs[r]
            issues = typecatalog.validate_type(self.model, f.type)
            h += (f"<h4>字段 <code>{_esc(f.name)}</code>　类型 "
                  f"<code>{_esc(f.type)}</code></h4>")
            if issues:
                h += "<p style='color:#B42318'>" + "；".join(
                    f"[{i['level']}] {_esc(i['msg'])}" for i in issues) + "</p>"
            h += options_html(typecatalog.fill_options(self.model, f.name, f.type),
                              cur_style)
        else:
            h += "<p>（还没有字段，点「+ 添加字段」）</p>"
        self.view_opts.setHtml(h)

        try:
            style = "auto" if self.cb_tpl.currentIndex() == 0 else "sep"
            cols = template.build_cols(self.model, fs, default=style)
            note = ("按每个字段在「配置方式」里选的生成。复合类型默认单格 sep，"
                    "一格装下，几乎不会填错。" if style == "auto" else
                    "全部强制用单格 sep：最保守，适合第一次用。")
            self.view_tpl.setHtml(f"<p>{note}</p><p>这是「导出数据表模板」会生成的表头：</p>"
                                  + _html_table(template.header_grid(cols)))
        except Exception as e:
            self.view_tpl.setHtml(f"<p>预览失败：{_esc(str(e))}</p>")

        # 顺手算出这批字段相关的知识点关键词：以后点顶部「知识点」时
        # 直接用它当搜索条件，省得自己想该搜什么（见 _k_seed）。
        kws = set()
        for f in fs:
            p = typecatalog.parse_type(f.type)
            if p["kind"] == "container":
                kws.update(["sep", "容器", "list"])
            if p["nullable"]:
                kws.add("可空")
            kd = typecatalog.resolve_kind(self.model, p["base"])
            if kd == "bean":
                kws.add("bean")
            if kd == "enum":
                kws.add("枚举")
        if tmp.real_mode == "one":
            kws.add("纵表")
        idx = self._form.get("index", "")
        if idx and ("+" in idx or "," in idx):
            kws.add("mode")
        self._k_seed = " ".join(sorted(kws))

    def act_knowledge(self, kid: str = ""):
        """工具栏「知识点」：打开知识点浏览器（非模态，可以开着继续改表）。

        打开时带上当前表的关键词（_k_seed）；传 kid 则是「直接跳到某一条」，
        给別处「看知识点」用。
        """
        d = getattr(self, "_knowledge_dlg", None)
        if d is None:
            d = KnowledgeDialog(self)
            self._knowledge_dlg = d
        if kid:
            d.focus_entry(kid)
        else:
            d.seed(getattr(self, "_k_seed", ""))
        d.show()
        d.raise_()
        d.activateWindow()

    # ---------------------------------------------------------------- 动作
    def act_type_helper(self):
        """工具栏「类型助手」：打开四区对话框。

        编辑器里选中了字段行 → 直接把结果应用到该字段；否则复制到剪贴板备用。
        """
        r = self.tbl_fields.currentRow() if self._editing else -1
        has_row = self._editing and 0 <= r < len(self._edit_fields)
        cur = self._edit_fields[r].type if has_row else "int"
        name = "x"
        if has_row:
            it = self.tbl_fields.item(r, COL_NAME)
            name = typecatalog.strip_label(it.text()) if it else "x"
        dlg = TypePickerDialog(self, self.model, name, cur, self.current)
        if dlg.exec() != QDialog.Accepted:
            return
        v = dlg.value()
        if has_row:
            self._edit_fields[r].type = v
            self._set_type_combo(r, v)
            self._fill_style_combo(r, self._edit_fields[r])
            self._after_field_change(self.tbl_fields)
            self.status.showMessage(f"字段 {name} 的类型已设为 {v}", 6000)
        else:
            QApplication.clipboard().setText(v)
            QMessageBox.information(self, "类型助手",
                                    f"类型字符串：{v}\n已复制到剪贴板。")

    def _save_definitions(self, only: Optional[str] = None) -> bool:
        """写回定义文件。

        only 不为空时，只校验/提示这张表的问题 —— 保存一张表不该被别的表的问题拦下。
        """
        if not self.model.tables:
            QMessageBox.information(self, "提示", "还没有可保存的表")
            return False
        issues = runner.validate_project(self.model, self._fields_snapshot())
        if only:
            issues = [i for i in issues if (i.get("table") or "") == only]
        errs = [i for i in issues if i["level"] == "error"]
        if errs:
            txt = "\n".join(f"[{i['table']}] {i['msg']}" for i in errs[:15])
            title = f"配置表 {only} 存在问题，仍要保存？" if only else "存在错误，仍要保存？"
            if QMessageBox.question(self, title, txt + "\n\n继续保存？") != QMessageBox.Yes:
                self._show_issues(issues)
                return False
        self._show_issues(issues)
        ok, msg = exceldef.save_tables(self.model)
        if not ok:
            QMessageBox.critical(self, "保存失败", msg)
            return False
        self.status.showMessage(msg, 8000)
        t = self.model.find_table(self.current or "")
        if t and (t.read_schema_from_file or "").strip().lower() in ("false", "0", "no"):
            fs = self.fields.get(t.full_name, [])
            if fs and t.value_type:
                b = self.model.find_bean(t.value_type) or BeanDef(full_name=t.value_type)
                b.fields = fs
                # 表结构类型（value_type 的 bean）写进该表自己的 xml，不写公共 common.xml
                ok2, msg2 = exceldef.save_bean(
                    self.model, b,
                    exceldef.table_def_file(self.model, t.full_name, t.input))
                if not ok2:
                    QMessageBox.warning(self, "结构定义未写入", msg2)
                else:
                    self.status.showMessage(msg2, 8000)
        self._reindex_inputs()
        self._refresh_tree()
        return True

    def act_export(self):
        if not self._editing or not self.current:
            QMessageBox.information(self, "提示", "请先在左侧选中一张配置表")
            return
        fs = self._edit_fields
        if not fs:
            QMessageBox.information(self, "提示", "该表还没有字段，先添加字段")
            return
        # 用表单里的（可能还没保存的）值生成，所见即所得
        tmp = TableDef(full_name=self._form.get("full_name", ""),
                       mode=self._form.get("mode", ""),
                       index=self._form.get("index", ""),
                       input=self._form.get("input", ""))
        if not tmp.full_name:
            QMessageBox.information(self, "提示", "表名 full_name 不能为空")
            return
        path = template.default_input_path(self.model, tmp)
        if os.path.exists(path):
            if QMessageBox.question(self, "文件已存在",
                                    f"{path}\n已存在，覆盖会丢失已填数据。继续？") != QMessageBox.Yes:
                return
        vertical_hint = self._nt_vertical_hint
        self._nt_vertical_hint = None
        if tmp.real_mode != "one":
            vertical = False
        elif vertical_hint is not None:
            vertical = vertical_hint
        else:
            vertical = QMessageBox.question(
                self, "模板形式", "这是单例表，是否生成纵表模板？",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes
        style = "auto" if self.cb_tpl.currentIndex() == 0 else "sep"
        ok, msg = template.generate_data_xlsx(self.model, tmp, fs, path,
                                              vertical=vertical, style=style)
        if not ok:
            QMessageBox.critical(self, "失败", str(msg))
            return
        self.status.showMessage(f"已生成模板：{path}", 8000)
        QMessageBox.information(self, "完成",
                                f"模板已生成：\n{path}\n\n"
                                f"建议先跑一次「生成并校验」确认表头能被 Luban 正确识别。")
        try:
            rel = os.path.relpath(path, self.model.data_dir).replace("\\", "/")
        except ValueError:
            rel = path.replace("\\", "/")
        self._form["input"] = rel
        prev = self._loading
        self._loading = True
        # 编辑器里「配置表名」只显示文件名（不含模块目录）
        self.ed_input.setText(os.path.basename(rel))
        self._loading = prev
        self._reindex_inputs()
        self._repaint_dirty()
        self._refresh_tree()

    def act_check(self):
        issues = runner.validate_project(self.model, self._fields_snapshot())
        self._show_issues(issues)

    def act_publish(self):
        """发布工程：设置 gen.bat / gen.sh 的参数并运行它（＝ 导表发布）。"""
        if not self.model.root:
            QMessageBox.information(self, "还没打开工程", "先打开或创建一个 Luban 工程")
            return
        if not self._confirm_leave_editor():
            return
        d = PublishDialog(self, self.model, self.settings)
        d.exec()

    def act_gen(self):
        if not self.settings.get("luban_exe"):
            QMessageBox.information(self, "还没配置", "请先在「设置」里指定 Luban.exe 路径")
            self.act_settings()
            if not self.settings.get("luban_exe"):
                return
        res = runner.run_luban(self.settings, self.model)
        self._show_issues(res["issues"] or
                          [{"level": "info", "table": "", "msg": "无报错", "hint": ""}])
        self.view_opts.setHtml(
            f"<h4>生成结果：{'成功' if res['ok'] else '失败'}（code={res['code']}）</h4>"
            f"<p><code>{_esc(res['cmd'])}</code></p><pre>{_esc(res['output'][:6000])}</pre>")

    def _show_issues(self, issues: List[Dict[str, str]]):
        """记下这批问题。

        「问题 / 校验」页签已经不显示了，所以不再切页签 —— 改成在状态栏给一句
        结论（几条 error / warn），让人知道有没有事。列表控件还留着（没挂到界面上），
        真要细看可以跑顶部「静态自检」。
        """
        self._last_issues = issues or []
        self.view_issue.clear()
        for i in issues:
            color = {"error": "#B42318", "warn": "#B54708", "info": "#475467"}.get(
                i["level"], "#475467")
            txt = f"[{i['level']}] {i.get('table', '')}：{i['msg']}"
            if i.get("hint"):
                txt += f"\n    → {i['hint']}"
            it = QListWidgetItem(txt)
            it.setForeground(QColor(color))
            it.setData(Qt.UserRole, i.get("table", ""))
            it.setData(Qt.UserRole + 1, i.get("knowledge", ""))
            self.view_issue.addItem(it)
        n_err = sum(1 for i in self._last_issues if i["level"] == "error")
        n_warn = sum(1 for i in self._last_issues if i["level"] == "warn")
        if n_err or n_warn:
            self.status.showMessage(f"错误 {n_err} 条 / 警告 {n_warn} 条", 8000)
        elif self._last_issues:
            self.status.showMessage("没有发现问题", 6000)

    def _goto_issue_table(self, item):
        name = item.data(Qt.UserRole)
        if name and self.model.find_table(name):
            self._select_table(name)
        kid = item.data(Qt.UserRole + 1)
        if kid and knowledge.by_id(kid):
            self.act_knowledge(kid)

    def _ask_delete_table(self, name: str) -> Optional[bool]:
        """问「只删声明」还是「连数据文件一起删」。返回 None = 取消。"""
        box = QMessageBox(self)
        box.setWindowTitle("删除配置表")
        box.setText(f"删除配置表 {name}？")
        box.setInformativeText(
            "「只删声明」：从 __tables__.xlsx 里删掉这一行，Schema 不再导出这张表，"
            "Data 下的数据文件原样保留。\n"
            "「连数据文件」：数据文件也一起删掉，删了找不回来。")
        b_decl = box.addButton("只删声明（保留数据文件）", QMessageBox.AcceptRole)
        b_all = box.addButton("连数据文件一起删", QMessageBox.DestructiveRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.setDefaultButton(b_decl)
        box.exec()
        clicked = box.clickedButton()
        if clicked is b_decl:
            return False
        if clicked is b_all:
            return True
        return None

    def _delete_table(self, name: str, with_file: Optional[bool] = None):
        """删除配置表：声明（Schema 行）+ 可选的数据文件。

        只删声明的话，数据文件还留在 Data 里，下次可以重新登记回来；连文件一起删
        则是真删掉 —— 两个选项写清楚，别让人以为「删表 = 文件没了」。
        with_file=None 时弹框问；测试可以直接传 False / True 跳过对话框。
        """
        t = self.model.find_table(name)
        if not t:
            return
        if with_file is not None:
            do_file = with_file
        else:
            do_file = self._ask_delete_table(name)
            if do_file is None:
                return
        removed: List[str] = []
        if do_file:
            for p in self._input_paths(t):
                if p and os.path.isfile(p):
                    try:
                        os.remove(p)
                        removed.append(os.path.basename(p))
                    except Exception as e:
                        QMessageBox.warning(self, "数据文件删不掉", f"{p}\n{e}")

        self.model.tables = [x for x in self.model.tables if x.full_name != name]
        self.fields.pop(name, None)
        if self.current == name:
            self._reset_editor()
            self._show_overview()
            self._repaint_dirty()
        self._reindex_inputs()
        self._refresh_tree()
        # 工具栏不再有「保存定义」，删除直接落盘，避免用户以为删了其实没生效
        if self._save_definitions():
            self._refresh_tree()
            tail = f"（同时删除数据文件：{', '.join(removed)}）" if removed else ""
            self.status.showMessage(f"已删除配置表 {name}{tail}", 8000)
            self.toast(f"已删除配置表 {name}{tail}")
        else:
            self.status.showMessage(f"已从内存移除 {name}，但未写入定义文件", 8000)

    def act_manage_types(self):
        """工具栏「管理自定义类型」：查看引用情况、删除没人用的类型。"""
        if not self.model.initialized:
            QMessageBox.information(
                self, "提示",
                "当前目录还不是 Luban 工程（没读到 luban.conf）。\n"
                "先「打开工程」选一个 Luban 工程根目录，或用「初始化 Luban 工程」把它变成工程。")
            return
        if not self._confirm_leave_editor():
            return
        d = ManageTypesDialog(self, self.model, self.fields)
        d.changed.connect(self._on_types_changed)
        d.exec()

    def _on_types_changed(self):
        """自定义类型被删了：重新读定义文件，刷新树与右侧面板。"""
        try:
            exceldef.load_project(self.model)
        except Exception as e:
            self.status.showMessage(f"重新读取定义文件失败：{e}", 8000)
        self._reindex_inputs()
        self._refresh_tree()
        if hasattr(self, "_refresh_right"):
            try:
                self._refresh_right()
            except Exception:
                pass

    def save_settings(self):
        for k, v in self.settings.items():
            self.qt_settings.setValue(k, v)

    def act_settings(self):
        d = SettingsDialog(self, self.settings, self.model)
        if d.exec() == QDialog.Accepted:
            d.apply()
            self.save_settings()

    def _new_table(self):
        """工具栏「新建表」：在中间区打开新建表单，不弹独立窗口。"""
        self._open_new_table()


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
