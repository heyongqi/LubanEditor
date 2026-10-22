"""工程探测：找到 luban.conf，解析出 schemaFiles / dataDir / groups / targets。

通用性策略（不绑定任何具体工程）：
1. 优先在用户选择的根目录及其 3 层子目录内找 *.conf（JSON 且含 schemaFiles）；
2. 找不到就退回扫描 __tables__.xlsx / __beans__.xlsx / __enums__.xlsx；
3. 定义文件的真实列名以"语义别名表"模糊匹配，匹配不到就弹给用户确认。
"""
from __future__ import annotations

import json
import os
from typing import Dict, List, Optional, Tuple

from . import xmldef
from .model import ProjectModel

CONF_CANDIDATE_NAMES = ("luban.conf", "luban.json", "gen.conf", "config.conf")
MAX_SCAN_DEPTH = 3
# 这些目录里的文件一律不当定义文件：备份目录里躺着的是旧版本定义，
# 当成定义文件读会把已经删掉的表又读回来（还能一路套娃备份下去）。
SKIP_DIRS = {"_lubankit_backup", ".git", ".svn", "__pycache__", "node_modules"}


def _is_conf_json(path: str) -> bool:
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception:
        return False
    return isinstance(data, dict) and ("schemaFiles" in data or "dataDir" in data)


def find_conf(root: str) -> Optional[str]:
    if os.path.isfile(root) and _is_conf_json(root):
        return root
    for dirpath, dirnames, filenames in os.walk(root):
        depth = dirpath[len(root):].strip(os.sep).count(os.sep)
        if depth >= MAX_SCAN_DEPTH:
            dirnames[:] = []
        for fn in filenames:
            if not fn.lower().endswith((".conf", ".json")):
                continue
            p = os.path.join(dirpath, fn)
            if _is_conf_json(p):
                return p
    return None


def _norm(s: str) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def _pick_by_type(schema_files: List, want: str) -> Optional[str]:
    """接受 (path, type) 元组或 {fileName,type} 字典。"""
    for item in schema_files:
        if isinstance(item, tuple):
            path, typ = item
        else:
            path, typ = item.get("fileName", ""), item.get("type", "")
        if _norm(typ) == want:
            return path
    return None


def _pick_by_name(files: List[str], keys: Tuple[str, ...]) -> Optional[str]:
    for f in files:
        base = _norm(os.path.basename(f))
        for k in keys:
            if k in base:
                return f
    return None


def _pick_primary(paths: List[str], key: str) -> str:
    """从同类定义文件里挑出「主」文件：__beans__.xlsx / __enums__.xlsx。

    工程里除了主文件，还可能有配置表专用的 TbEquip.beans.xlsx 这种，
    它们不是主定义文件 —— 这里优先认以 __ 开头的那个。
    """
    for p in paths:
        if os.path.basename(p).lower().startswith("__" + key):
            return p
    for p in paths:
        if key in os.path.basename(p).lower():
            return p
    return paths[0] if paths else ""


def _skip_dirs(dirnames: List[str]) -> None:
    dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith(".")]


def scan_xlsx(root: str) -> List[str]:
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        depth = dirpath[len(root):].strip(os.sep).count(os.sep)
        if depth >= MAX_SCAN_DEPTH:
            dirnames[:] = []
        _skip_dirs(dirnames)
        for fn in filenames:
            if fn.lower().endswith((".xlsx", ".xlsm")) and not fn.startswith("~$"):
                out.append(os.path.join(dirpath, fn))
    return out


def scan_xml(root: str) -> List[str]:
    """Luban 也可以把定义写在 Defines/*.xml 里，没有 conf 时也要能找到。"""
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        depth = dirpath[len(root):].strip(os.sep).count(os.sep)
        if depth >= MAX_SCAN_DEPTH:
            dirnames[:] = []
        _skip_dirs(dirnames)
        for fn in filenames:
            if fn.lower().endswith(".xml") and not fn.startswith("~$"):
                out.append(os.path.join(dirpath, fn))
    return out


def _pick_writable(paths: List[str], key: str) -> str:
    """挑「主」定义文件时优先挑 xlsx：xml 也能读，但工具写回 xlsx 最稳。"""
    xl = [p for p in paths if p.lower().endswith((".xlsx", ".xlsm"))]
    return _pick_primary(xl, key) if xl else _pick_primary(paths, key)


def _sniff_unclaimed(paths: List[str], claimed: set) -> Tuple[List[str], List[str], List[str]]:
    """名字看不出类型的 xml，按内容认是表 / 结构 / 枚举。

    有些工程 luban.conf 只登记了 Defines 目录（type 留空），里面的文件也不叫
    tables/beans/enums —— 不嗅探的话这些定义等于没被工具看见。
    """
    t, b, e = [], [], []
    for p in paths:
        if p in claimed or not p.lower().endswith(".xml"):
            continue
        kinds = xmldef.sniff(p)
        if "table" in kinds:
            t.append(p)
        if "bean" in kinds:
            b.append(p)
        if "enum" in kinds:
            e.append(p)
    return t, b, e


def _dedup(paths: List[str]) -> List[str]:
    out = []
    for p in paths:
        if p and p not in out:
            out.append(p)
    return out


def discover(root: str) -> ProjectModel:
    """探测工程，返回未加载定义的 ProjectModel 骨架。

    对目录没有要求：随便选一个目录都能打开。找不到 Luban 痕迹时返回空模型，
    由界面提示「初始化 Luban 工程」，而不是弹一堆警告。
    """
    m = ProjectModel(root=os.path.abspath(root))
    if not os.path.isdir(m.root):
        m.warnings.append(f"目录不存在：{m.root}")
        m.data_dir = m.root
        return m
    conf = find_conf(root)
    xlsx_files = scan_xlsx(root)

    if conf:
        m.conf_path = conf
        conf_dir = os.path.dirname(conf)
        with open(conf, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        schema_files = [s for s in data.get("schemaFiles", []) if isinstance(s, dict)]
        rel = []
        for s in schema_files:
            fn = s.get("fileName", "")
            if not fn:
                continue
            rel.append((os.path.join(conf_dir, fn.replace("/", os.sep)), s.get("type", "")))
        # 目录型 schemaFile：递归收集
        expanded: List[Tuple[str, str]] = []
        for path, typ in rel:
            if os.path.isdir(path):
                for dp, dn, fns in os.walk(path):
                    _skip_dirs(dn)
                    for fn in fns:
                        if fn.lower().endswith((".xlsx", ".xml")) and not fn.startswith("~$"):
                            expanded.append((os.path.join(dp, fn), typ))
            else:
                expanded.append((path, typ))
        paths = [p for p, _ in expanded]

        table_paths = [p for p, typ in expanded if _norm(typ) == "table"]
        bean_paths = [p for p, typ in expanded if _norm(typ) == "bean"]
        enum_paths = [p for p, typ in expanded if _norm(typ) == "enum"]
        if not table_paths:
            table_paths = [p for p in paths if "tables" in _norm(os.path.basename(p))]
        if not bean_paths:
            bean_paths = [p for p in paths if "beans" in _norm(os.path.basename(p))]
        if not enum_paths:
            enum_paths = [p for p in paths if "enums" in _norm(os.path.basename(p))]
        # 剩下的 xml 按内容认领（Defines 下随手放的 xxx.xml 也能被发现）
        st, sb, se = _sniff_unclaimed(paths, set(table_paths) | set(bean_paths) |
                                      set(enum_paths))
        table_paths += st
        bean_paths += sb
        enum_paths += se
        table_paths, bean_paths, enum_paths = (_dedup(table_paths), _dedup(bean_paths),
                                               _dedup(enum_paths))
        # 表定义可以分散在多个文件里，全部记下来（只读一个会把 xml 里的表漏掉）
        m.table_files = [p for p in table_paths if os.path.exists(p)]
        if not m.table_files:
            # 一条表定义都没登记时的兜底：工程里确实写着表定义（Defines/*.xml），
            # 只是 conf 里漏登了 —— 至少让数据文件能被识别成可编辑
            t3, _b3, _e3 = _sniff_unclaimed(scan_xml(root), set())
            m.table_files = _dedup(t3)
            table_paths = list(m.table_files)
        m.table_file = _pick_writable(table_paths, "tables")
        m.bean_file = _pick_writable(bean_paths, "beans")
        m.enum_file = _pick_writable(enum_paths, "enums")
        # 剩下的都是「配置表专用」类型所在的文件（工具建的 Defines/TbXxx.xml，或旧版残留的 TbXxx.beans.xlsx）
        m.extra_bean_files = [p for p in bean_paths if p != m.bean_file]
        m.extra_enum_files = [p for p in enum_paths if p != m.enum_file]

        data_dir = data.get("dataDir", "")
        m.data_dir = os.path.normpath(os.path.join(conf_dir, data_dir)) if data_dir else \
            os.path.dirname(m.table_file or root)
        m.groups = [g for grp in data.get("groups", []) for g in grp.get("names", [])]
        m.targets = [t.get("name", "") for t in data.get("targets", []) if t.get("name")]
    else:
        m.table_file = _pick_by_name(xlsx_files, ("tables",)) or ""
        m.bean_file = _pick_by_name(xlsx_files, ("beans",)) or ""
        m.enum_file = _pick_by_name(xlsx_files, ("enums",)) or ""
        # 没有 conf 也要认 xml 定义：表可能只写在 Defines/*.xml 里
        t2, b2, e2 = _sniff_unclaimed(
            scan_xml(root), {m.table_file, m.bean_file, m.enum_file} - {""})
        m.table_files = _dedup([m.table_file] + t2)
        m.table_file = m.table_file or (_pick_primary(t2, "tables") if t2 else "")
        m.bean_file = m.bean_file or (_pick_primary(b2, "beans") if b2 else "")
        m.enum_file = m.enum_file or (_pick_primary(e2, "enums") if e2 else "")
        base = os.path.dirname(m.table_file or m.bean_file or m.enum_file or
                               os.path.join(root, "Data"))
        m.data_dir = base
        m.groups = ["c", "s", "e"]
        m.targets = []

    if not m.initialized:
        m.warnings.append("这个目录还没有 Luban 配置。可以点「初始化 Luban 工程」一键创建。")
    elif not m.table_file and not m.table_files:
        m.warnings.append("没有定位到表定义文件（__tables__.xlsx / Defines 下的表定义），"
                          "新建表功能不可用。")
    for label, p in (("表定义", m.table_file), ("结构定义", m.bean_file), ("枚举定义", m.enum_file)):
        if p and not os.path.exists(p):
            m.warnings.append(f"{label}文件不存在：{p}")
    if not os.path.isdir(m.data_dir):
        os.makedirs(m.data_dir, exist_ok=True)
    return m
