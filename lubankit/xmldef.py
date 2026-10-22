"""Luban 的 **xml 定义文件**（Defines/*.xml）读写。

为什么要有这个模块：Luban 允许把表 / 结构 / 枚举的定义写在 xml 里
（luban.conf 的 schemaFiles 指向一个 .xml，或直接指向整个 Defines 目录）。
这类工程以前被工具当成「没有表定义」——点开 Data 里的数据文件只会提示
「没在 __tables__.xlsx 里声明」，其实人家声明在 xml 里，于是好端端能编辑的
配置表全变成只读预览。

读：ElementTree 解析，`<module name="x">` 的 name 当命名空间前缀
    （和 excel 里的 module 一个意思，xml 里写短名 + module 前缀 = 全名）。
写：**文本级**增删改。xml 里常有手写的注释和排版，用 ET 重写会全部丢掉，
    所以只定位 `<table ...>` 这一行做替换 / 插入 / 删除，其它字节一个不动。
"""
from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple

from .model import (ProjectModel, TableDef, BeanDef, EnumDef, FieldDef,
                    EnumItem)

# ---------------------------------------------------------------- 属性取值
def _nk(s: str) -> str:
    """属性名归一：忽略大小写、下划线、连字符。readSchemaFromFile == read_schema_from_file"""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def _get(attrib: Dict[str, str], *names: str, default: str = "") -> str:
    """按候选名取属性，取不到返回 default。"""
    want = {_nk(n) for n in names}
    for k, v in attrib.items():
        if _nk(k) in want:
            return (v or "").strip()
    return default


def _tag(el) -> str:
    """去掉可能存在的 xmlns 前缀。"""
    return str(el.tag).split("}")[-1].lower()


def _full_name(module: str, name: str) -> str:
    """module + 短名 -> 全名。名字里已经带 '.' 就当全名用（有人习惯写全名）。"""
    name = (name or "").strip()
    if not name:
        return ""
    if "." in name:
        return name
    return "%s.%s" % (module, name) if module else name


def _short_name(module: str, full: str) -> str:
    """写回时把全名还原成「module 内的短名」，module 对不上就原样写全名。"""
    full = (full or "").strip()
    if module and full.startswith(module + "."):
        return full[len(module) + 1:]
    return full


# ---------------------------------------------------------------- 读取
def _table_from_el(el, module: str, path: str) -> TableDef:
    a = el.attrib
    t = TableDef(
        full_name=_full_name(module, _get(a, "name")),
        value_type=_get(a, "value", "valueType", "valuetype"),
        # Luban 默认就是从数据文件表头读字段，xml 里不写等于 true
        read_schema_from_file=_get(a, "readSchemaFromFile", "read_schema_from_file",
                                   default="true"),
        input=_get(a, "input", "inputFile", "file"),
        index=_get(a, "index", "indexes", "indexName"),
        mode=_get(a, "mode"),
        group=_get(a, "group", "groups"),
        comment=_get(a, "comment", "desc", "description"),
        tags=_get(a, "tags"),
        output=_get(a, "output", "outputFile"),
    )
    t._file = path
    return t


def _bean_from_el(el, module: str, path: str) -> BeanDef:
    a = el.attrib
    b = BeanDef(
        full_name=_full_name(module, _get(a, "name")),
        parent=_get(a, "parent", "base"),
        value_type=_get(a, "valueType", "value"),
        sep=_get(a, "sep"),
        alias=_get(a, "alias"),
        group=_get(a, "group", "groups"),
        comment=_get(a, "comment", "desc", "description"),
    )
    b._file = path
    for v in el:
        if _tag(v) != "var":
            continue
        va = v.attrib
        nm = _get(va, "name")
        if not nm:
            continue
        b.fields.append(FieldDef(
            name=nm,
            type=_get(va, "type") or "int",
            group=_get(va, "group", "groups"),
            comment=_get(va, "comment", "desc", "description"),
            tags=_get(va, "tags"),
        ))
    return b


def _enum_from_el(el, module: str, path: str) -> EnumDef:
    a = el.attrib
    e = EnumDef(
        full_name=_full_name(module, _get(a, "name")),
        flags=_get(a, "flags"),
        unique=_get(a, "unique"),
        comment=_get(a, "comment", "desc", "description"),
    )
    e._file = path
    for v in el:
        if _tag(v) != "var":
            continue
        va = v.attrib
        nm = _get(va, "name")
        if not nm:
            continue
        e.items.append(EnumItem(
            name=nm,
            alias=_get(va, "alias"),
            value=_get(va, "value"),
            comment=_get(va, "comment", "desc", "description"),
        ))
    return e


def read_defs(path: str) -> Tuple[List[TableDef], List[BeanDef], List[EnumDef]]:
    """解析一个 xml 定义文件，返回 (tables, beans, enums)。解析失败返回三个空表。"""
    try:
        root = ET.parse(path).getroot()
    except Exception:
        return [], [], []
    tables: List[TableDef] = []
    beans: List[BeanDef] = []
    enums: List[EnumDef] = []

    def walk(el, module: str):
        tag = _tag(el)
        if tag == "module":
            module = (el.get("name") or "").strip() or module
        elif tag == "table":
            t = _table_from_el(el, module, path)
            if t.full_name:
                tables.append(t)
        elif tag == "bean":
            b = _bean_from_el(el, module, path)
            if b.full_name:
                beans.append(b)
        elif tag == "enum":
            e = _enum_from_el(el, module, path)
            if e.full_name:
                enums.append(e)
        for c in el:
            walk(c, module)

    walk(root, (root.get("name") or "").strip() if _tag(root) == "module" else "")
    return tables, beans, enums


def sniff(path: str) -> List[str]:
    """看文件内容像哪类定义文件：'table' / 'bean' / 'enum' 的子集（可能同时是几类）。

    给「luban.conf 里只登记了 Defines 目录、文件又没按 tables/beans/enums 命名」
    的工程兜底 —— 名字猜不出就按内容认。一个 xml 里常常 table / bean / enum 混写，
    所以返回的是列表而不是单一类型。
    """
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            text = f.read(200000)
    except Exception:
        return []
    low = text.lower()
    has = lambda tag: re.search(r"<%s\b" % tag, low) is not None
    return [k for k, tag in (("table", "table"), ("bean", "bean"), ("enum", "enum"))
            if has(tag)]


def module_name(path: str) -> str:
    """文件里 `<module name="...">` 的值，用于写回时把全名还原成短名。"""
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            m = re.search(r"<module\b[^>]*?\bname\s*=\s*\"([^\"]*)\"", f.read(200000))
    except Exception:
        return ""
    return (m.group(1) if m else "").strip()


# ---------------------------------------------------------------- 写回
# 一行一个 <table .../>（也容忍属性换行写），连前导缩进和行尾换行一起匹配
_TABLE_RE = re.compile(r"(?P<ind>[ \t]*)<table\b[^<>]*?/?>[ \t]*\r?\n?", re.I)

_TABLE_ATTRS = (
    ("name", "full_name"), ("value", "value_type"), ("input", "input"),
    ("index", "index"), ("mode", "mode"), ("group", "group"),
    ("comment", "comment"), ("tags", "tags"), ("output", "output"),
)


def _xml_escape(s: str) -> str:
    return (str(s or "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _table_line(t: TableDef, module: str, indent: str) -> str:
    parts = []
    for attr, field in _TABLE_ATTRS:
        v = getattr(t, field, "") or ""
        if attr == "name":
            v = _short_name(module, v)
        if not v:
            continue
        parts.append('%s="%s"' % (attr, _xml_escape(v)))
    # read_schema_from_file 只有显式关掉时才写（true 是 Luban 默认，写了反而啰嗦）
    rsf = (t.read_schema_from_file or "").strip().lower()
    if rsf in ("false", "0", "no"):
        parts.append('readSchemaFromFile="false"')
    if not parts:
        return ""
    return "%s<table %s/>\n" % (indent, " ".join(parts))


def ensure_file(path: str) -> Tuple[bool, str]:
    """定义文件不存在就照模板建一个空的；已存在则一个字节都不动。

    新建表时如果工程里还没有 xml 定义文件，就得先凭空造一个能被 Luban 认的
    （`<module>` 包着、里面暂时没有 `<table>`），否则写回函数找不到落脚点。
    """
    if os.path.exists(path):
        return True, ""
    parent = os.path.dirname(path)
    try:
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(XML_TABLE_TEMPLATE)
    except Exception as e:
        return False, "创建 xml 定义文件失败：%s" % e
    return True, ""


XML_TABLE_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8"?>
<!-- 表定义：每张表一行（本文件由 LubanEditor 创建） -->
<module name="">
</module>
'''


def write_tables(path: str, wanted: List[TableDef],
                 known: Optional[set] = None) -> Tuple[bool, str]:
    """把 wanted 写回 xml 定义文件：同名更新、缺的追加、多的删除。

    只动 `<table .../>` 那一行，注释、空行、手写的缩进全部保留。
    ``known`` 是本次从文件里读出来的表名集合 —— 用来区分「用户真把表删光了」
    和「解析失败导致一张表都没读进来」，后者照写会把整份定义抹空。
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        return False, "读不到 xml 定义文件 %s：%s" % (os.path.basename(path), e)

    module = module_name(path)
    # 注释里的 <table .../> 不算数 —— 工程模板里就有一段被注释掉的示例，
    # 当成真表会以为「有 1 张表要删」，保存直接被保护逻辑拦下。
    comments = [(c.start(), c.end()) for c in re.finditer(r"<!--.*?-->", text, re.S)]
    spans = [(m.start("ind"), m.end(), m.group(0))
             for m in _TABLE_RE.finditer(text)
             if not any(a <= m.start() < b for a, b in comments)]
    existing: Dict[str, Tuple[int, int, str]] = {}
    for start, end, raw in spans:
        nm = re.search(r"\bname\s*=\s*\"([^\"]*)\"", raw)
        if not nm:
            continue
        full = _full_name(module, nm.group(1))
        if full and full not in existing:
            existing[full] = (start, end, raw)
    want = {t.full_name: t for t in wanted if t.full_name}

    if not want and existing and (known is None or not set(existing) <= set(known)):
        return False, ("这次保存会把 %s 里 %d 张表全部删掉，疑似异常，已中止。"
                       "确实要清空请直接编辑这个文件。"
                       % (os.path.basename(path), len(existing)))

    # 缩进照抄已有行，没有就用两空格
    indent = "  "
    for _s, _e, raw in spans:
        lead = re.match(r"[ \t]*", raw)
        if lead and lead.group(0):
            indent = lead.group(0)
            break

    edits: List[Tuple[int, int, str]] = []      # (start, end, 替换文本)
    added: List[TableDef] = []
    for name, t in want.items():
        line = _table_line(t, module, indent)
        if not line:
            continue
        if name in existing:
            s, e, _raw = existing[name]
            edits.append((s, e, line))
        else:
            added.append(t)
    gone = 0
    for name, (s, e, _raw) in existing.items():
        if name not in want:
            edits.append((s, e, ""))
            gone += 1

    # 从后往前改，前面的改动不影响后面的下标
    for s, e, rep in sorted(edits, reverse=True):
        text = text[:s] + rep + text[e:]

    if added:
        lines = "".join(_table_line(t, module, indent) for t in added)
        pos = _insert_pos(text)
        text = text[:pos] + lines + text[pos:]

    try:
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(text)
    except Exception as e:
        return False, "写入 xml 定义文件失败：%s" % e
    extra = []
    if added:
        extra.append("新增 %d 张" % len(added))
    if gone:
        extra.append("删除 %d 张" % gone)
    tail = "（%s）" % "，".join(extra) if extra else ""
    return True, "已保存 %d 张表到 %s%s" % (len(want), os.path.basename(path), tail)


def _insert_pos(text: str) -> int:
    """新 <table> 插在哪：最后一个 </module> 前；没有就插到根元素的结束标签前。"""
    last = None
    for m in re.finditer(r"[ \t]*</module\s*>", text, re.I):
        last = m
    if last is not None:
        return last.start()
    closes = list(re.finditer(r"[ \t]*</[A-Za-z_][\w.\-]*\s*>", text))
    if closes:
        return closes[-1].start()
    return len(text)


# ---------------------------------------------------------------- 填进模型
def load_into_model(model: ProjectModel, path: str) -> int:
    """把一个 xml 定义文件的表 / 结构 / 枚举并进 model。返回新增表数。"""
    tables, beans, enums = read_defs(path)
    model.tables.extend(tables)
    model.beans.extend(beans)
    model.enums.extend(enums)
    return len(tables)


# ---------------------------------------------------------------- bean/enum 写回
# 与 write_tables 一样走「文本级」修改：只替换 / 插入 / 删除 <bean> / <enum> 块，
# 注释、缩进、手写的其它内容一个字节不动。bean/enum 写成平铺（parent 属性表达
# 继承），字段用 <var .../> 子元素，与 _bean_from_el / _enum_from_el 的读法一一对应。

XML_TYPES_TEMPLATE = '''<?xml version="1.0" encoding="UTF-8"?>
<!-- 类型定义（bean / enum），由 LubanEditor 创建 -->
<module name="">
</module>
'''


def ensure_types_file(path: str) -> Tuple[bool, str]:
    """类型定义 xml 不存在就建一个空模板；已存在则不动。"""
    if os.path.exists(path):
        return True, ""
    parent = os.path.dirname(path)
    try:
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(XML_TYPES_TEMPLATE)
    except Exception as e:
        return False, "创建 xml 定义文件失败：%s" % e
    return True, ""


def _attr_block(pairs) -> str:
    return " ".join('%s="%s"' % (k, _xml_escape(v)) for k, v in pairs if v)


def _var_block(indent: str, name: str, type_: str = "", group: str = "",
               comment: str = "", tags: str = "") -> str:
    pairs = [("name", name), ("type", type_ or "int")]
    if group:
        pairs.append(("group", group))
    if comment:
        pairs.append(("comment", comment))
    if tags:
        pairs.append(("tags", tags))
    return '%s<var %s/>\n' % (indent, _attr_block(pairs))


def _bean_block(b: BeanDef, module: str, indent: str) -> str:
    name = _short_name(module, b.full_name)
    pairs = [("name", name)]
    if b.parent:
        pairs.append(("parent", b.parent))
    if b.value_type:
        pairs.append(("valueType", b.value_type))
    if b.sep:
        pairs.append(("sep", b.sep))
    if b.alias:
        pairs.append(("alias", b.alias))
    if b.group:
        pairs.append(("group", b.group))
    if b.comment:
        pairs.append(("comment", b.comment))
    if not b.fields:
        return '%s<bean %s/>\n' % (indent, _attr_block(pairs))
    lines = ['%s<bean %s>\n' % (indent, _attr_block(pairs))]
    for f in b.fields:
        lines.append(_var_block(indent + "  ", f.name, f.type, f.group,
                                f.comment, f.tags))
    lines.append('%s</bean>\n' % indent)
    return "".join(lines)


def _enum_block(e: EnumDef, module: str, indent: str) -> str:
    name = _short_name(module, e.full_name)
    pairs = [("name", name)]
    if e.flags:
        pairs.append(("flags", e.flags))
    if e.unique:
        pairs.append(("unique", e.unique))
    if e.comment:
        pairs.append(("comment", e.comment))
    if not e.items:
        return '%s<enum %s/>\n' % (indent, _attr_block(pairs))
    lines = ['%s<enum %s>\n' % (indent, _attr_block(pairs))]
    for it in e.items:
        ip = [("name", it.name)]
        if it.alias:
            ip.append(("alias", it.alias))
        if it.value:
            ip.append(("value", it.value))
        if it.comment:
            ip.append(("comment", it.comment))
        lines.append('%s<var %s/>\n' % (indent + "  ", _attr_block(ip)))
    lines.append('%s</enum>\n' % indent)
    return "".join(lines)


def _find_named_blocks(text: str, tag: str):
    """返回 [(start, end, name)]：text 中顶层 <tag>...</tag> 块（跳过注释，处理自闭合与嵌套）。"""
    comments = [(c.start(), c.end()) for c in re.finditer(r"<!--.*?-->", text, re.S)]

    def in_comment(pos):
        return any(a <= pos < b for a, b in comments)

    start_re = re.compile(r"<%s\b([^>]*)>" % tag, re.I)
    token_re = re.compile(r"<%s\b[^>]*>|</%s\s*>" % (tag, tag), re.I)
    blocks = []
    for m in start_re.finditer(text):
        if in_comment(m.start()):
            continue
        attrs = m.group(1)
        nm = re.search(r"\bname\s*=\s*\"([^\"]*)\"", attrs)
        name = nm.group(1) if nm else ""
        if attrs.rstrip().endswith("/"):
            blocks.append((m.start(), m.end(), name))
            continue
        depth = 1
        end = -1
        for tok in token_re.finditer(text, m.end()):
            if tok.group(0).lower().startswith("</"):
                depth -= 1
                if depth == 0:
                    end = tok.end()
                    break
            else:
                depth += 1
        if end != -1:
            blocks.append((m.start(), end, name))
    return blocks


def write_types(path: str, beans: Optional[List[BeanDef]] = None,
                enums: Optional[List[EnumDef]] = None) -> Tuple[bool, str]:
    """把 bean / enum 写回 xml 定义文件：同名更新、缺的追加，绝不删除其它定义。

    只动 <bean> / <enum> 块，注释与其它内容原样保留。名字按 module 归一成全名
    来匹配（同一个定义可能在 <module name="x"> 里写短名）。
    """
    beans = list(beans or [])
    enums = list(enums or [])
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        return False, "读不到 xml 定义文件 %s：%s" % (os.path.basename(path), e)

    module = module_name(path)
    # 表专属类型文件（非 common.xml）：待写的 bean / enum 带模块前缀时，
    # 用它作为文件的 module（并写回 <module name>），这样写成
    # <module name="item"><bean name="Drop">，而不是 <module name=""><bean name="item.Drop">。
    # common.xml 是公共文件，module 恒为空，不推导。
    if os.path.basename(path) != "common.xml":
        for obj in beans + enums:
            if obj.full_name and "." in obj.full_name and not module:
                mod = obj.full_name.split(".")[0]
                if mod:
                    module = mod
                    text = re.sub(r'(<module\b[^>]*?\bname\s*=\s*)"[^"]*"',
                                  r'\1"%s"' % _xml_escape(mod), text, count=1)
                break
    indent = "  "
    edits: List[Tuple[int, int, str]] = []
    added: List[str] = []

    for tag, wanted, block_fn in (("bean", beans, _bean_block),
                                  ("enum", enums, _enum_block)):
        existing: Dict[str, Tuple[int, int]] = {}
        for s, e, nm in _find_named_blocks(text, tag):
            full = _full_name(module, nm)
            if full and full not in existing:
                existing[full] = (s, e)
        for obj in wanted:
            if not obj.full_name:
                continue
            if obj.full_name in existing:
                s, e = existing[obj.full_name]
                edits.append((s, e, block_fn(obj, module, indent)))
            else:
                added.append(block_fn(obj, module, indent))

    for s, e, rep in sorted(edits, reverse=True):
        text = text[:s] + rep + text[e:]

    if added:
        pos = _insert_pos(text)
        text = text[:pos] + "".join(added) + text[pos:]

    try:
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(text)
    except Exception as e:
        return False, "写入 xml 定义文件失败：%s" % e
    return True, "已保存到 %s（bean %d / enum %d）" % (
        os.path.basename(path), len(beans), len(enums))


def delete_def(path: str, tag: str, full_name: str) -> Tuple[bool, str]:
    """从 xml 定义文件里删掉一个 bean / enum（按全名匹配）。"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            text = f.read()
    except Exception as e:
        return False, "读不到 xml 定义文件 %s：%s" % (os.path.basename(path), e)
    module = module_name(path)
    for s, e, nm in _find_named_blocks(text, tag):
        if _full_name(module, nm) == full_name:
            new_text = text[:s] + text[e:]
            try:
                with open(path, "w", encoding="utf-8", newline="") as f:
                    f.write(new_text)
            except Exception as ex:
                return False, "写入 xml 定义文件失败：%s" % ex
            return True, "已从 %s 删除 %s %s" % (os.path.basename(path), tag, full_name)
    return False, "在 %s 里没找到 %s %s，未做修改" % (os.path.basename(path), tag, full_name)
