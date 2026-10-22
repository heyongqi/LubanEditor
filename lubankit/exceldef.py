"""Luban 定义文件（__tables__.xlsx / __beans__.xlsx / __enums__.xlsx）的通用读写。

三条硬规则：
1. 只读/写目标单元格，绝不整体重建工作簿 —— 保住样式、批注、其它行列；
2. 每次写盘前先备份到 _lubankit_backup/；
3. 列名不做硬编码，用"语义别名表"模糊匹配，匹配不到就上报给用户。
"""
from __future__ import annotations

import json
import os
import shutil
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border

from . import xmldef
from . import projectwizard
from .model import ProjectModel, TableDef, BeanDef, EnumDef, FieldDef, EnumItem

BACKUP_DIRNAME = "_lubankit_backup"
MAX_BACKUPS = 20


def _norm(s) -> str:
    """列名归一化：小写 + 只保留字母数字与中文。"""
    t = str(s if s is not None else "").strip().lower()
    return "".join(ch for ch in t if ch.isalnum() or "\u4e00" <= ch <= "\u9fff")


# ---------------------------------------------------------------- 别名表
TABLE_ALIASES: Dict[str, List[str]] = {
    "full_name": ["full_name", "fullname", "name", "tablename", "表名", "表全名"],
    "value_type": ["value_type", "valuetype", "valuetypename", "value_type_name",
                   "记录类型", "行类型", "bean", "bean名"],
    "read_schema_from_file": ["read_schema_from_file", "readschemafromfile",
                              "define_from_file", "definefromfile",
                              "define_from_excel", "read_schema_from_excel",
                              "从文件读取定义", "从数据表读取定义", "自动读取表头"],
    "input": ["input", "inputs", "inputfiles", "input_files", "datafile",
              "数据文件", "数据源", "输入文件", "输入"],
    "index": ["index", "indexes", "key", "keys", "索引", "主键"],
    "mode": ["mode", "tablemode", "表模式"],
    "group": ["group", "groups", "分组", "导出分组"],
    "comment": ["comment", "desc", "description", "注释", "说明", "备注"],
    "tags": ["tags", "tag", "标签"],
    "output": ["output", "outputfile", "outputfilename", "output_file_name", "输出文件", "输出"],
}

BEAN_ALIASES: Dict[str, List[str]] = {
    "full_name": ["full_name", "fullname", "name", "bean名", "结构名", "类型名"],
    "parent": ["parent", "parentname", "父类", "父结构"],
    "value_type": ["valuetype", "value_type", "isvaluetype", "值类型"],
    "sep": ["sep", "separator", "分隔符"],
    "alias": ["alias", "别名"],
    "group": ["group", "groups", "分组", "导出分组"],
    "comment": ["comment", "desc", "description", "注释", "说明", "备注"],
    "fields": ["fields", "field", "字段", "字段列表"],
}

FIELD_ALIASES: Dict[str, List[str]] = {
    "name": ["name", "fieldname", "字段名", "字段", "变量"],
    "type": ["type", "fieldtype", "类型"],
    "group": ["group", "groups", "分组", "导出分组"],
    "comment": ["comment", "desc", "description", "注释", "说明", "备注"],
    "tags": ["tags", "tag", "标签"],
}

ENUM_ALIASES: Dict[str, List[str]] = {
    "full_name": ["full_name", "fullname", "name", "枚举名", "类型名"],
    "flags": ["flags", "isflags", "标志位"],
    "unique": ["unique", "值唯一"],
    "comment": ["comment", "desc", "description", "注释", "说明", "备注"],
    "items": ["items", "item", "枚举项", "项"],
}

ITEM_ALIASES: Dict[str, List[str]] = {
    "name": ["name", "itemname", "名字", "项名"],
    "alias": ["alias", "别名"],
    "value": ["value", "值"],
    "comment": ["comment", "desc", "description", "注释", "说明", "备注"],
}


def map_columns(actual_names: List[str], aliases: Dict[str, List[str]]
                ) -> Tuple[Dict[str, str], List[str]]:
    """把实际列名映射到语义名。返回 (语义名->实际列名, 未识别的实际列名)。"""
    norm_alias = {}
    for sem, names in aliases.items():
        for n in names:
            norm_alias[_norm(n)] = sem
    mapping: Dict[str, str] = {}
    used = set()
    for a in actual_names:
        if not a:
            continue
        sem = norm_alias.get(_norm(a))
        if sem and sem not in used:
            mapping[sem] = a
            used.add(sem)
    unmatched = [a for a in actual_names if a and a not in mapping.values()]
    return mapping, unmatched


# ---------------------------------------------------------------- 定义表解析
class DefSheet:
    """一张定义 sheet 的结构化视图（含多级表头）。"""

    def __init__(self, ws):
        self.ws = ws
        self.name = ws.title
        self.header_row = 0
        self.marker_rows: List[int] = []
        self.columns: List[Dict] = []     # {idx, l0, sub}
        self.data_rows: List[int] = []
        self._parse()

    # -- 合并单元格取值 --
    def _anchor_value(self, row: int, col: int):
        for rng in self.ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                v = self.ws.cell(row=rng.min_row, column=rng.min_col).value
                return v
        return self.ws.cell(row=row, column=col).value

    def _parent_of(self, row: int, col: int) -> str:
        """嵌套表头单元格属于哪个一级列（靠合并区或向左查找最近的非空一级列）。"""
        for rng in self.ws.merged_cells.ranges:
            if rng.min_row <= self.header_row <= rng.max_row and \
               rng.min_col <= col <= rng.max_col and rng.min_col <= rng.max_col:
                v = self.ws.cell(row=self.header_row, column=rng.min_col).value
                if v:
                    return str(v).strip()
        c = col
        while c >= 1:
            v = self._anchor_value(self.header_row, c)
            if v:
                return str(v).strip()
            c -= 1
        return ""

    def _has_grouped_header(self) -> bool:
        """表头行里有没有跨列合并的分组单元格 —— 有才可能出现二级表头。"""
        for rng in self.ws.merged_cells.ranges:
            if rng.min_row <= self.header_row <= rng.max_row and rng.max_col > rng.min_col:
                return True
        return False

    def _parse(self):
        ws = self.ws
        max_row, max_col = ws.max_row or 1, ws.max_column or 1
        markers = []
        for r in range(1, max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v is not None and str(v).strip().startswith("##"):
                markers.append(r)
        if not markers:
            # 退化：把第一行当表头
            markers = [1]
        self.marker_rows = markers

        # 一级表头：跳过 ##comment / ##type / ##group 之类，取第一个含 >=2 个有效列名的标记行
        header_row = markers[0]
        for r in markers:
            a = str(ws.cell(row=r, column=1).value or "").strip().lower()
            if a in ("##comment", "##type", "##group"):
                continue
            cnt = sum(1 for c in range(2, max_col + 1)
                      if self._anchor_value(r, c) not in (None, ""))
            if cnt >= 1:
                header_row = r
                break
        self.header_row = header_row

        # 一级列
        cols: List[Dict] = []
        for c in range(2, max_col + 1):
            v = self._anchor_value(header_row, c)
            if v in (None, ""):
                continue
            cols.append({"idx": c, "l0": str(v).strip(), "sub": ""})
        # 二级表头（多级字段，如 bean 的 fields / enum 的 items）
        #
        # 关键区分（踩过坑）：工具自己的建工程模板会在表头下面写一行「行说明」，
        # A1 是「##」、后面是中文列名（##、表名、行结构类型…）。那**不是**二级表头，
        # 是整行的注释行。以前把它当二级表头，于是每一列都被安上 c["sub"]，
        # 而 value()/set_value() 都要求 not c["sub"] —— 结果整张表既读不出来也写不进去，
        # 还一律报告成功。真正的二级表头是**再写一行 ##var**（bean / enum 就是这么干的）。
        sub_rows = []
        for r in markers:
            if r <= header_row:
                continue
            a = str(ws.cell(row=r, column=1).value or "").strip().lower()
            if a == "##var":
                sub_rows.append(r)
            elif a == "##" and self._has_grouped_header():
                sub_rows.append(r)
        for r in sub_rows:
            for col in cols:
                sv = self._anchor_value(r, col["idx"])
                if sv in (None, ""):
                    sv = ws.cell(row=r, column=col["idx"]).value
                if sv in (None, ""):
                    continue
                col["sub"] = str(sv).strip()
        self.columns = cols

        last_marker = max(markers)
        for r in range(last_marker + 1, max_row + 1):
            a = ws.cell(row=r, column=1).value
            if a is not None and str(a).strip().startswith("##"):
                continue
            if any(ws.cell(row=r, column=c["idx"]).value not in (None, "")
                   for c in cols):
                self.data_rows.append(r)

    # -- 取值 --
    def value(self, row: int, semantic: str, mapping: Dict[str, str]) -> str:
        colname = mapping.get(semantic)
        if not colname:
            return ""
        idx = None
        for c in self.columns:
            if c["l0"] == colname and not c["sub"]:
                idx = c["idx"]
                break
        if idx is None:
            return ""
        v = self.ws.cell(row=row, column=idx).value
        return "" if v is None else str(v).strip()

    def set_value(self, row: int, semantic: str, mapping: Dict[str, str], val) -> bool:
        colname = mapping.get(semantic)
        if not colname:
            return False
        idx = None
        for c in self.columns:
            if c["l0"] == colname and not c["sub"]:
                idx = c["idx"]
                break
        if idx is None:
            return False
        self.ws.cell(row=row, column=idx).value = val
        return True

    def column_index(self, colname: str) -> Optional[int]:
        for c in self.columns:
            if c["l0"] == colname and not c["sub"]:
                return c["idx"]
        return None

    def sub_groups(self, parent_semantic: str, mapping: Dict[str, str],
                   sub_aliases: Dict[str, List[str]]) -> List[Dict[str, int]]:
        """找出嵌套列分组：每遇到一个"name"子类列就开一组。"""
        parent_name = mapping.get(parent_semantic, "")
        groups: List[Dict[str, int]] = []
        cur: Dict[str, int] = {}
        for c in self.columns:
            if not c["sub"]:
                continue
            if parent_name and c["l0"] != parent_name:
                continue
            sem, _ = map_columns([c["sub"]], sub_aliases)
            if not sem:
                continue
            key = list(sem.keys())[0]
            if key == "name":
                if cur:
                    groups.append(cur)
                cur = {}
            cur.setdefault(key, c["idx"])
        if cur:
            groups.append(cur)
        return groups

    def read_sub(self, row: int, group: Dict[str, int], key: str) -> str:
        idx = group.get(key)
        if not idx:
            return ""
        v = self.ws.cell(row=row, column=idx).value
        return "" if v is None else str(v).strip()


def pick_sheet(wb, aliases: Dict[str, List[str]]) -> Optional[DefSheet]:
    """在工作簿里挑出最像定义表的 sheet。"""
    best, best_score = None, -1
    for ws in wb.worksheets:
        a1 = ws.cell(row=1, column=1).value
        if a1 is None or not str(a1).strip().startswith("##"):
            continue
        try:
            ds = DefSheet(ws)
        except Exception:
            continue
        names = [c["l0"] for c in ds.columns]
        mapping, _ = map_columns(names, aliases)
        score = len(mapping)
        if score > best_score:
            best, best_score = ds, score
    if best is None and wb.worksheets:
        best = DefSheet(wb.worksheets[0])
    return best


def open_def_sheet(path: str, aliases: Dict[str, List[str]], for_write: bool = False
                   ) -> Tuple[Optional[DefSheet], Optional[object], List[str]]:
    warns: List[str] = []
    if not path or not os.path.exists(path):
        return None, None, ["文件不存在: %s" % path]
    try:
        wb = load_workbook(path, data_only=False)
    except Exception as e:
        return None, None, ["无法打开 %s: %s" % (path, e)]
    ds = pick_sheet(wb, aliases)
    if ds is None:
        return None, None, ["%s 中没有可用 sheet" % path]
    if not ds.columns:
        warns.append(f"{os.path.basename(path)} 的 sheet「{ds.name}」没解析到列名，"
                     f"请确认 A1 以 ## 开头。")
    return ds, wb, warns


# ---------------------------------------------------------------- 读取
def _has(items, full_name: str) -> bool:
    """同名定义只留第一份（同一个类型被两个定义文件声明时）。"""
    return any(x.full_name == full_name for x in items)


def _def_files(files: List[str], primary: str) -> List[str]:
    """要读的定义文件列表：主文件排第一，其余去重跟上。"""
    out: List[str] = []
    for p in [primary] + list(files or []):
        if p and p not in out:
            out.append(p)
    return out


def _merge_xml(model: ProjectModel, path: str) -> None:
    """xml 定义文件里的表 / 结构 / 枚举并进 model（同名不重复加）。"""
    for t, seq in zip(xmldef.read_defs(path), (model.tables, model.beans, model.enums)):
        for item in t:
            if item.full_name and not _has(seq, item.full_name):
                seq.append(item)


def _load_table_file(model: ProjectModel, path: str, primary: bool = False) -> int:
    """读一个 xlsx 表定义文件。返回读到的表数。"""
    ds, wb, warns = open_def_sheet(path, TABLE_ALIASES)
    model.warnings += warns
    if not ds:
        return 0
    names = [c["l0"] for c in ds.columns]
    mapping, unmatched = map_columns(names, TABLE_ALIASES)
    if primary or not model.table_columns:
        model.table_columns = mapping
    if "full_name" not in mapping:
        # 主定义文件认不出表头要报（否则用户以为表都没了）；
        # 顺带读到的其它 xlsx（被目录型 schemaFile 一起圈进来的 __beans__.xlsx 等）
        # 认不出就是「它本来不是表定义文件」，不用刷告警。
        if primary:
            model.warnings.append("表定义文件没识别出 full_name 列，请在「列映射」中手动指定。")
        wb.close()
        return 0
    got = 0
    for r in ds.data_rows:
        t = TableDef(
            full_name=ds.value(r, "full_name", mapping),
            value_type=ds.value(r, "value_type", mapping),
            read_schema_from_file=ds.value(r, "read_schema_from_file", mapping),
            input=ds.value(r, "input", mapping),
            index=ds.value(r, "index", mapping),
            mode=ds.value(r, "mode", mapping),
            group=ds.value(r, "group", mapping),
            comment=ds.value(r, "comment", mapping),
            tags=ds.value(r, "tags", mapping),
            output=ds.value(r, "output", mapping),
        )
        t._row, t._sheet = r, ds.name
        t._file = path
        if t.full_name and not _has(model.tables, t.full_name):
            model.tables.append(t)
            got += 1
    if ds.data_rows and not got:
        # 有数据行却一张表都没读出来 —— 以前是完全静默的，
        # 表现就是「Data 里的配置文件点开全是只读」，以为是功能没了
        model.warnings.append(
            "表定义文件 %s 有 %d 行数据，但一行表都没解析出来。"
            "请检查它的表头（第一行应以 ##var 开头、"
            "并且有 full_name / value_type / input 这些列名）。"
            % (os.path.basename(path), len(ds.data_rows)))
    wb.close()
    return got


def _load_bean_file(model: ProjectModel, path: str, primary: bool) -> int:
    """读一个 bean 定义文件，把结果并进 model.beans。返回识别到的字段列组数。"""
    if path.lower().endswith(".xml"):
        _merge_xml(model, path)
        return 0
    ds, wb, warns = open_def_sheet(path, BEAN_ALIASES)
    model.warnings += warns
    if not ds:
        return 0
    names = [c["l0"] for c in ds.columns]
    mapping, _unmatched = map_columns(names, BEAN_ALIASES)
    if primary:
        model.bean_columns = mapping
    groups = ds.sub_groups("fields", mapping, FIELD_ALIASES)
    for r in ds.data_rows:
        b = BeanDef(
            full_name=ds.value(r, "full_name", mapping),
            parent=ds.value(r, "parent", mapping),
            value_type=ds.value(r, "value_type", mapping),
            sep=ds.value(r, "sep", mapping),
            alias=ds.value(r, "alias", mapping),
            group=ds.value(r, "group", mapping),
            comment=ds.value(r, "comment", mapping),
        )
        b._row, b._sheet, b._layout = r, ds.name, ("nested" if groups else "unknown")
        b._file = path
        for g in groups:
            fn = ds.read_sub(r, g, "name")
            if not fn:
                continue
            b.fields.append(FieldDef(
                name=fn,
                type=ds.read_sub(r, g, "type") or "int",
                group=ds.read_sub(r, g, "group"),
                comment=ds.read_sub(r, g, "comment"),
                tags=ds.read_sub(r, g, "tags"),
            ))
        if b.full_name and not _has(model.beans, b.full_name):
            model.beans.append(b)
    wb.close()
    return len(groups)


def _load_enum_file(model: ProjectModel, path: str, primary: bool) -> None:
    """读一个 enum 定义文件，把结果并进 model.enums。"""
    if path.lower().endswith(".xml"):
        _merge_xml(model, path)
        return
    ds, wb, warns = open_def_sheet(path, ENUM_ALIASES)
    model.warnings += warns
    if not ds:
        return
    names = [c["l0"] for c in ds.columns]
    mapping, _unmatched = map_columns(names, ENUM_ALIASES)
    if primary:
        model.enum_columns = mapping
    groups = ds.sub_groups("items", mapping, ITEM_ALIASES)
    for r in ds.data_rows:
        e = EnumDef(
            full_name=ds.value(r, "full_name", mapping),
            flags=ds.value(r, "flags", mapping),
            unique=ds.value(r, "unique", mapping),
            comment=ds.value(r, "comment", mapping),
        )
        # 和 bean 一样记住来源文件：表专用枚举要靠文件名判断归属
        # （TbEquip.enums.xlsx 里的类型就归 TbEquip），删的时候也要知道去哪删
        e._file = path
        for g in groups:
            nm = ds.read_sub(r, g, "name")
            if not nm:
                continue
            e.items.append(EnumItem(
                name=nm,
                alias=ds.read_sub(r, g, "alias"),
                value=ds.read_sub(r, g, "value"),
                comment=ds.read_sub(r, g, "comment"),
            ))
        if e.full_name and not _has(model.enums, e.full_name):
            model.enums.append(e)
    wb.close()


def load_project(model: ProjectModel) -> ProjectModel:
    model.warnings = [w for w in model.warnings if w.startswith("未找到") or "不存在" in w]
    # 重新读一遍 = 完全以文件为准，先清空（以前不清，重复调用会读出一堆重复项）
    model.tables = []
    model.beans = []
    model.enums = []
    model.loaded_tables = []
    # tables：可能分散在多个文件里（__tables__.xlsx + Defines/*.xml），全都要读，
    # 只挑一个读的话「声明在 xml 里的表」会被当成没声明，数据文件全变只读预览。
    for p in _def_files(model.table_files, model.table_file):
        if not p or not os.path.exists(p):
            continue
        if p.lower().endswith(".xml"):
            _merge_xml(model, p)
        else:
            _load_table_file(model, p, primary=(p == model.table_file))
    # beans：主定义文件 + 各配置表专用的独立文件
    bean_groups = 0
    if model.bean_file and os.path.exists(model.bean_file):
        bean_groups = _load_bean_file(model, model.bean_file, True)
    for p in list(model.extra_bean_files):
        if os.path.exists(p):
            _load_bean_file(model, p, False)
        else:
            model.extra_bean_files.remove(p)     # 文件被删了，别留着一个死路径
    if bean_groups and any(b.fields for b in model.beans):
        model.warnings.append(
            f"Bean 定义采用嵌套列布局，当前识别到 {bean_groups} 组字段列。"
            f"新建结构时字段数不超过 {bean_groups} 个可直接写入，"
            f"超过请先用 Excel 补列，或改用「从数据表头读字段」。")
    # enums：同上
    if model.enum_file and os.path.exists(model.enum_file):
        _load_enum_file(model, model.enum_file, True)
    for p in list(model.extra_enum_files):
        if os.path.exists(p):
            _load_enum_file(model, p, False)
        else:
            model.extra_enum_files.remove(p)
    model.loaded_tables = [t.full_name for t in model.tables if t.full_name]
    derive_scopes(model)
    return model


# ---------------------------------------------------------------- 写回
def backup_file(path: str) -> str:
    d = os.path.join(os.path.dirname(path), BACKUP_DIRNAME)
    os.makedirs(d, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(d, f"{os.path.splitext(os.path.basename(path))[0]}_{stamp}.xlsx")
    shutil.copy2(path, dst)
    # 只保留最近 N 份
    try:
        files = sorted([os.path.join(d, f) for f in os.listdir(d)
                        if f.startswith(os.path.splitext(os.path.basename(path))[0])],
                       key=os.path.getmtime)
        for f in files[:-MAX_BACKUPS]:
            os.remove(f)
    except Exception:
        pass
    return dst


def backup_any(path: str) -> str:
    """任意文件（luban.conf 这类非 xlsx 的）备份，统一放进同目录的 _lubankit_backup/。"""
    d = os.path.join(os.path.dirname(path), BACKUP_DIRNAME)
    os.makedirs(d, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(d, f"{stem}_{stamp}{ext}")
    shutil.copy2(path, dst)
    return dst


def define_dir(model: ProjectModel) -> str:
    """定义文件所在目录（用户工程里通常叫 Defines，示例工程里是 Data）。"""
    for p in (model.bean_file, model.enum_file, model.table_file):
        if p:
            return os.path.dirname(p)
    return os.path.join(model.root, "Defines")


def defines_dir(model: ProjectModel) -> str:
    """xml 定义文件所在目录（bean / enum 写 xml 时都放这）。

    优先跟已有 xml 定义文件走（老工程定义在 Defines 或 Data 都行），
    工程里还没有 xml 定义文件时统一用 <root>/Defines。
    """
    pool = (_def_files(model.table_files, model.table_file)
            + list(model.extra_bean_files) + list(model.extra_enum_files))
    for p in pool:
        if p and p.lower().endswith(".xml"):
            return os.path.dirname(p)
    return os.path.join(model.root, "Defines")


def common_types_file(model: ProjectModel) -> str:
    """公共类型（bean / enum）统一的 xml 文件：Defines/common.xml。"""
    return os.path.join(defines_dir(model), "common.xml")


def public_bean_file(model: ProjectModel) -> str:
    """公共结构定义的 xml 文件（与公共枚举共用一个 common.xml）。"""
    return common_types_file(model)


def public_enum_file(model: ProjectModel) -> str:
    """公共枚举定义的 xml 文件（与公共结构共用一个 common.xml）。"""
    return common_types_file(model)


def ensure_defines_registered(model: ProjectModel) -> Tuple[bool, str]:
    """把 xml 定义目录登记进 luban.conf 的 schemaFiles（type=""）。"""
    d = defines_dir(model)
    try:
        os.makedirs(d, exist_ok=True)
    except Exception as e:
        return False, f"创建定义目录失败：{e}"
    return ensure_schema_registered(model, d, "")


def table_def_file(model: ProjectModel, table_name: str, input_name: str = "") -> str:
    """表结构 bean 与表专属类型所在的 xml 文件。

    有模块名 → `<定义目录>/<模块名>.xml`；
    无模块名 → `<定义目录>/<配置表名>.xml`（配置表名取 input 去掉扩展名）。
    表和它的 bean / enum 都放同一个文件里（Luban 一个 xml 可以同时有
    table / bean / enum，不用像 xlsx 那样按 type 拆文件）。
    """
    full = (table_name or "").strip()
    mod = full.split(".")[0] if "." in full else ""
    if mod:
        return os.path.join(defines_dir(model), "%s.xml" % mod)
    base = os.path.splitext(os.path.basename((input_name or "").strip()))[0]
    if not base:
        base = full.split(".")[-1]
    return os.path.join(defines_dir(model), "%s.xml" % (base or "Table"))


def ensure_schema_registered(model: ProjectModel, path: str, kind: str) -> Tuple[bool, str]:
    """把新定义文件登记进 luban.conf 的 schemaFiles —— 不登记 Luban 会当它不存在。"""
    conf = model.conf_path
    if not conf or not os.path.exists(conf):
        return False, (f"工程里没有 luban.conf，没法自动登记 {os.path.basename(path)}；"
                       f"请在导表配置里手动把它加进 schemaFiles（type={kind}）。")
    try:
        with open(conf, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as e:
        return False, f"luban.conf 读不出来：{e}"
    if not isinstance(data, dict):
        return False, "luban.conf 不是 JSON 对象，无法自动登记定义文件"
    files = data.setdefault("schemaFiles", [])
    rel = os.path.relpath(path, os.path.dirname(conf)).replace(os.sep, "/")
    want = os.path.normcase(os.path.normpath(rel))
    conf_dir = os.path.dirname(conf)
    for item in files:
        if not isinstance(item, dict):
            continue
        fn = str(item.get("fileName", "")).replace("\\", "/")
        # 已经登记过同一个文件就不用再加了；Luban 报「定义重复」大多是被登了两次
        if os.path.normcase(os.path.normpath(fn)) == want:
            return True, ""
        # swagger：schemaFiles 里登记了整个目录时，目录下的文件不用再单登一条
        _p = os.path.normpath(os.path.join(conf_dir, fn))
        if os.path.isdir(_p) and os.path.normcase(os.path.abspath(path)).startswith(
                os.path.normcase(os.path.abspath(_p)) + os.sep):
            return True, ""
    files.append({"fileName": rel, "type": kind})
    try:
        backup_any(conf)
        with open(conf, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")
    except Exception as e:
        return False, f"luban.conf 写入失败：{e}"
    return True, ""


def schema_dir(model: ProjectModel) -> str:
    """Schema 目录：定义文件（__tables__.xlsx / __beans__.xlsx / tables.xml…）放哪。

    新建的工程统一用 `<root>/Schemas`；老工程的定义文件在 Data 或 Defines 里，
    就跟着它们走 —— 新建文件时不要凭空再开一个目录，否则 luban.conf 里
    一半在这儿一半在那儿。
    """
    for p in (model.table_file, model.bean_file, model.enum_file):
        if p and os.path.isdir(os.path.dirname(p)):
            return os.path.dirname(p)
    d = os.path.join(model.root, "Schemas")
    return d if os.path.isdir(d) else define_dir(model)


def plan_table_file(model: ProjectModel, kind: str) -> str:
    """新建表时，表声明会写到哪个文件（只算路径，不建文件）。

    kind='xlsx' → `__tables__.xlsx`（Excel Schema）
    kind='xml'  → 工程里已有的 xml 表定义文件，否则 `<Schema 目录>/tables.xml`
    """
    kind = "xml" if str(kind).lower() == "xml" else "xlsx"
    d = schema_dir(model)
    if kind == "xml":
        for p in _def_files(model.table_files, model.table_file):
            if p.lower().endswith(".xml"):
                return p
        return os.path.join(d, "tables.xml")
    main = model.table_file
    if main and not main.lower().endswith(".xml"):
        return main
    return os.path.join(d, "__tables__.xlsx")


def ensure_table_file(model: ProjectModel, kind: str) -> Tuple[bool, str, str]:
    """保证「表声明要写进去的那个文件」存在并已登记进 luban.conf。

    xlsx → `__tables__.xlsx`（工程里本来就有；没有就照标准表头建一个）；
    xml  → 工程里已有的 xml 表定义文件；没有就新建 `tables.xml`。
    返回 (ok, path, msg)。
    """
    path = plan_table_file(model, kind)
    existed = os.path.exists(path)
    if kind == "xml":
        if not existed:
            ok, msg = xmldef.ensure_file(path)
            if not ok:
                return False, "", msg
    else:
        if not existed:
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                projectwizard.write_tables_xlsx(path)
            except Exception as e:
                return False, "", f"创建 __tables__.xlsx 失败：{e}"
    ok, msg = ensure_schema_registered(
        model, path, "enum" if os.path.basename(path).startswith("__enums__")
        else ("bean" if os.path.basename(path).startswith("__beans__") else "table"))
    if not ok:
        # 登记失败（比如没有 luban.conf）不致命：文件已经建好了，写回去也能成，
        # 只是 Luban 导表前要手动把这条路加进 schemaFiles，所以必须说清楚。
        return True, path, msg
    if path not in model.table_files:
        model.table_files.append(path)
    if not model.table_file or (kind == "xlsx" and not existed):
        model.table_file = path
    return True, path, msg


def _copy_style(src, dst):
    try:
        if src.has_style:
            dst._style = src._style
    except Exception:
        pass


def _xml_defs(model: ProjectModel) -> List[str]:
    """工程里所有 xml 定义文件（表 / 结构 / 枚举）。"""
    pool = (_def_files(model.table_files, model.table_file)
            + list(model.extra_bean_files) + list(model.extra_enum_files))
    out: List[str] = []
    for p in pool:
        if p and p.lower().endswith(".xml") and p not in out:
            out.append(p)
    return out


def _src_file(t: TableDef, main: str) -> str:
    """这张表该写回哪个文件：读的时候记的 _file（还在就用），否则主定义文件。"""
    return t._file if (t._file and os.path.exists(t._file)) else main


def save_tables(model: ProjectModel) -> Tuple[bool, str]:
    """把 model.tables 写回表定义文件（新增/修改/删除）。

    表可能分散在好几个定义文件里（__tables__.xlsx + Defines/*.xml），
    所以按「这张表是从哪个文件读出来的」分组写回，别把 xml 里的表塞进 xlsx。
    """
    files = [p for p in _def_files(model.table_files, model.table_file)
             if p and os.path.exists(p)
             and BACKUP_DIRNAME not in os.path.normpath(p).split(os.sep)]
    main = model.table_file or (files[0] if files else "")
    groups: Dict[str, List[TableDef]] = {p: [] for p in files}
    rest: List[TableDef] = []
    for t in model.tables:
        if not t.full_name:
            continue
        f = _src_file(t, main)
        if f in groups:
            groups[f].append(t)
        else:
            rest.append(t)
    if rest:
        if not main:
            return False, "没有表定义文件可写（工程里既没有 __tables__.xlsx 也没有表定义 xml）"
        groups.setdefault(main, []).extend(rest)
    if not groups:
        return False, "表定义文件不存在，无法保存"

    msgs: List[str] = []
    ok_all = True
    for path, subset in groups.items():
        if not os.path.exists(path):
            ok_all = False
            msgs.append("表定义文件不存在，无法保存：%s" % path)
            continue
        if path.lower().endswith(".xml"):
            try:
                bak = backup_any(path)
            except Exception as e:
                ok_all = False
                msgs.append(f"备份失败，已中止保存：{e}")
                continue
            ok, msg = xmldef.write_tables(path, subset, set(model.loaded_tables))
            if ok:
                msg += f"，备份：{bak}"
        else:
            ok, msg = _save_tables_xlsx(model, path, subset)
        ok_all = ok_all and ok
        if msg:
            msgs.append(msg)
    return ok_all, "；".join(msgs)


def _save_tables_xlsx(model: ProjectModel, path: str,
                      subset: List[TableDef]) -> Tuple[bool, str]:
    """把 subset 这些表写回一个 xlsx 表定义文件（新增/修改/删除）。"""
    try:
        bak = backup_file(path)
    except Exception as e:
        return False, f"备份失败，已中止保存：{e}"
    try:
        wb = load_workbook(path, data_only=False)
        ws = wb[subset[0]._sheet] if (subset and subset[0]._sheet in wb.sheetnames) \
            else None
        if ws is None:
            ds = pick_sheet(wb, TABLE_ALIASES)
            ws = ds.ws
        sheet = DefSheet(ws)
        names = [c["l0"] for c in sheet.columns]
        mapping, _ = map_columns(names, TABLE_ALIASES)
        if "full_name" not in mapping:
            return False, "表定义文件缺少 full_name 列，未做修改"

        existing: Dict[str, int] = {}
        for r in sheet.data_rows:
            v = sheet.value(r, "full_name", mapping)
            if v:
                existing[v] = r

        wanted = {t.full_name: t for t in subset if t.full_name}
        style_src = sheet.data_rows[-1] if sheet.data_rows else None

        # 1) 更新 / 新增
        next_row = (max(sheet.data_rows) + 1) if sheet.data_rows else (max(sheet.marker_rows) + 1)
        wrote = 0
        for name, t in wanted.items():
            if name in existing:
                r = existing[name]
            else:
                r = next_row
                next_row += 1
                if style_src:
                    for c in sheet.columns:
                        _copy_style(ws.cell(row=style_src, column=c["idx"]),
                                    ws.cell(row=r, column=c["idx"]))
            for sem in ("full_name", "value_type", "read_schema_from_file", "input",
                        "index", "mode", "group", "comment", "tags", "output"):
                if sheet.set_value(r, sem, mapping, getattr(t, sem)) and sem == "full_name":
                    wrote += 1
        if wanted and not wrote:
            # 表头认不出来时 set_value 会全False，以前这里照样报「已保存」= 静默丢数据
            wb.close()
            return False, ("表定义文件的表头无法识别，%d 张表一行都没写进去，已中止保存。"
                           "请检查 %s 第一行是否以 ##var 开头、"
                           "列名是否含 full_name。" % (len(wanted), os.path.basename(path)))

        # 2) 删除（自下而上，避免行号漂移）
        removed = [r for name, r in existing.items() if name not in wanted]
        if removed:
            if not wanted:
                # 之前只拦「删得太多」，表少的时候（existing<=4）拦不住，会把整张表定义抹空。
                # 但只要这些表确实是本次从文件里读进来的，就说明是用户主动删光的，放行。
                loaded = set(model.loaded_tables or [])
                if not set(existing) <= loaded:
                    wb.close()
                    return False, ("这次保存会把 %s 里 %d 行表定义全部删掉，疑似导入异常，"
                                   "已中止保存。确实要清空请直接在 Excel 里删。"
                                   % (os.path.basename(path), len(removed)))
            if len(removed) > len(existing) // 2 and len(existing) > 4:
                wb.close()
                return False, ("本次删除行数过多（%d/%d），疑似导入异常，已中止保存。"
                               % (len(removed), len(existing)))
            for r in sorted(removed, reverse=True):
                ws.delete_rows(r, 1)

        wb.save(path)
        wb.close()
    except Exception as e:
        return False, f"保存失败：{e}（备份在 {bak}）"
    return True, f"已保存，备份：{bak}"


def _save_type_xml(model: ProjectModel, kind: str, path: str,
                   obj) -> Tuple[bool, str]:
    """把 bean / enum 写进一个 xml 文件：确保目录 / 文件 / 登记，备份后写盘。"""
    try:
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    except Exception as e:
        return False, f"创建定义目录失败：{e}"
    ok_reg, reg_msg = ensure_defines_registered(model)
    ok, msg = xmldef.ensure_types_file(path)
    if not ok:
        return False, msg
    try:
        bak = backup_any(path)
    except Exception as e:
        return False, f"备份失败，已中止：{e}"
    if kind == "bean":
        ok, msg = xmldef.write_types(path, beans=[obj])
    else:
        ok, msg = xmldef.write_types(path, enums=[obj])
    if not ok:
        return False, msg
    bucket = model.extra_bean_files if kind == "bean" else model.extra_enum_files
    if path not in bucket:
        bucket.append(path)
    tail = "" if ok_reg else ("；但登记进 luban.conf 失败：%s" % reg_msg)
    return True, f"已写入 {os.path.basename(path)}（{msg}），备份：{bak}{tail}"


def save_bean(model: ProjectModel, bean: BeanDef,
              path: Optional[str] = None) -> Tuple[bool, str]:
    """把结构 bean 写回 xml 定义文件。

    path 为空 = 公共结构 → <Defines>/beans.xml；
    传了 path = 写进指定文件（表专用 → <Defines>/<表名>.xml）。
    """
    path = path or public_bean_file(model)
    return _save_type_xml(model, "bean", path, bean)


def _delete_def_row(path: str, full_name: str, aliases: Dict[str, List[str]],
                    what: str) -> Tuple[bool, str]:
    """在定义文件里删掉 full_name 那一行（备份后删行保存）。

    和 save_bean / save_tables 一样走 DefSheet 语义列，避免写死列号；
    找不到就明确报错，不静默「删了个寂寞」。
    """
    if not path or not os.path.exists(path):
        return False, f"定义文件不存在，删不掉 {what}：{path}"
    try:
        bak = backup_file(path)
    except Exception as e:
        return False, f"备份失败，已中止删除：{e}"
    try:
        wb = load_workbook(path, data_only=False)
        ds = pick_sheet(wb, aliases)
        if ds is None:
            wb.close()
            return False, f"无法识别 {os.path.basename(path)} 的定义 sheet"
        mapping, _ = map_columns([c["l0"] for c in ds.columns], aliases)
        if "full_name" not in mapping:
            wb.close()
            return False, f"{os.path.basename(path)} 缺少 full_name 列，未做修改"
        row = None
        for r in ds.data_rows:
            if ds.value(r, "full_name", mapping) == full_name:
                row = r
                break
        if row is None:
            wb.close()
            return False, f"在 {os.path.basename(path)} 里没找到 {what} {full_name}，未做修改"
        ds.ws.delete_rows(row, 1)
        wb.save(path)
        wb.close()
    except Exception as e:
        return False, f"删除 {what} 失败：{e}（备份在 {bak}）"
    return True, f"已从 {os.path.basename(path)} 删除 {what} {full_name}，备份：{bak}"


def _delete_type_xml(path: str, tag: str, full_name: str) -> Tuple[bool, str]:
    """从 xml 定义文件里删掉一个 bean / enum（备份后删）。"""
    if not path or not os.path.exists(path):
        return False, f"定义文件不存在，删不掉 {tag}：{path}"
    try:
        bak = backup_any(path)
    except Exception as e:
        return False, f"备份失败，已中止删除：{e}"
    ok, msg = xmldef.delete_def(path, tag, full_name)
    if not ok:
        return False, msg
    return True, f"{msg}，备份：{bak}"


def delete_bean(model: ProjectModel, full_name: str,
                path: Optional[str] = None) -> Tuple[bool, str]:
    """删除一个结构（bean）。path 为空时按 bean 自己记的文件找，再退回公共 beans.xml。"""
    p = path
    if not p:
        b = model.find_bean(full_name)
        p = (b._file if b and b._file else "") or public_bean_file(model)
    if p.lower().endswith((".xlsx", ".xlsm")):
        return _delete_def_row(p, full_name, BEAN_ALIASES, "结构")
    return _delete_type_xml(p, "bean", full_name)


def delete_enum(model: ProjectModel, full_name: str,
                path: Optional[str] = None) -> Tuple[bool, str]:
    """删除一个枚举（enum）。path 为空时按 enum 自己记的文件找，再退回公共 enums.xml。"""
    p = path
    if not p:
        e = model.find_enum(full_name)
        p = (e._file if e and e._file else "") or public_enum_file(model)
    if p.lower().endswith((".xlsx", ".xlsm")):
        return _delete_def_row(p, full_name, ENUM_ALIASES, "枚举")
    return _delete_type_xml(p, "enum", full_name)


def _after_save(model: ProjectModel, path: str, kind: str,
                full_name: str, bak: str) -> Tuple[bool, str]:
    """表专用定义文件写完后的收尾：登记进 luban.conf + 记到 model 上。"""
    name = os.path.basename(path)
    if path == (model.bean_file if kind == "bean" else model.enum_file):
        return True, f"已写入 {name}，备份：{bak}"
    ok, msg = ensure_schema_registered(model, path, kind)
    if not ok:
        return False, f"{full_name} 已写入 {name}，但登记失败：{msg}"
    bucket = model.extra_bean_files if kind == "bean" else model.extra_enum_files
    if path not in bucket:
        bucket.append(path)
    return True, f"已写入表专用定义文件 {name}（并登记进 luban.conf），备份：{bak}"


# ---------------------------------------------------------------- 归属推导
def derive_scopes(model: ProjectModel) -> None:
    """推断自定义类型的归属：common（公共）或某张表专有。

    规则：类型名带模块前缀且该前缀与某张表的模块同名 -> 归该表；否则为公共。
    """
    mod2table: Dict[str, str] = {}
    # 两趟：显式模块前缀（item.TbEquip）优先，再补「TbItem -> item」这种推导名。
    # 否则 TbItem 会把 item 抢走，item.TbEquip 的专用类型就被认成 TbItem 的了。
    for t in model.tables:
        if t.full_name and "." in t.full_name:
            mod2table[t.full_name.split(".")[0]] = t.full_name
    for t in model.tables:
        if not t.full_name:
            continue
        mod2table.setdefault(t.module, t.full_name)

    def scope_of(full: str) -> str:
        if "." in full:
            return mod2table.get(full.split(".")[0], "common")
        return "common"

    def from_file(path: str) -> str:
        """从类型来源文件名推导归属表。

        - common.xml → 公共（返回 ""）
        - <模块名>.xml → 归该模块（取第一张带该前缀的表）
        - <配置表名>.xml / <表名> → 归 input 去扩展名 / 表名匹配的表
        - 旧版 TbEquip.beans.xlsx / .enums.xlsx → 归 TbEquip
        """
        if not path:
            return ""
        stem = os.path.splitext(os.path.basename(path))[0]
        for suf in (".beans", ".enums"):
            if stem.endswith(suf):
                stem = stem[:-len(suf)]
                break
        if stem == "common":
            return ""
        for t in model.tables:
            if "." in t.full_name and t.full_name.split(".")[0] == stem:
                return t.full_name
        for t in model.tables:
            if t.full_name.split(".")[-1] == stem:
                return t.full_name
            if os.path.splitext(os.path.basename(t.input or ""))[0] == stem:
                return t.full_name
        return ""

    for b in model.beans:
        b.scope = from_file(b._file) or scope_of(b.full_name)
    for e in model.enums:
        e.scope = from_file(e._file) or scope_of(e.full_name)


def save_enum(model: ProjectModel, enum: EnumDef,
              path: Optional[str] = None) -> Tuple[bool, str]:
    """把枚举写回 xml 定义文件。

    path 为空 = 公共枚举 → <Defines>/enums.xml；
    传了 path = 写进指定文件（表专用 → <Defines>/<表名>.xml）。
    """
    path = path or public_enum_file(model)
    return _save_type_xml(model, "enum", path, enum)
