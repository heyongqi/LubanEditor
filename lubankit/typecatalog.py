"""类型目录：选中类型后给出「所有可供选择的配置方式 + 可直接抄的 Excel 片段」。

这是本工具最值钱的一屏 —— 新手放弃 Luban 的地方，几乎都在类型字符串和单元格填法上。
所有示例都按你工程里真实的 bean 字段动态生成，不是写死的死例子。

示例严格按官方文档的四种读法组织：
  ① 流式·多单元格   ② 单单元格 + sep   ③ 列限定（多级 ##var）   ④ 多行（仅容器）
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from .model import ProjectModel, BeanDef, FieldDef, EnumDef

# ------------------------------------------------------------------ 基础类型
BASE_TYPES: List[Dict[str, str]] = [
    {"name": "bool", "cn": "布尔", "fill": "true / false、0 / 1、是 / 否（大小写不敏感）",
     "example": "true", "note": "其它值会报错；流式(sep)单元格里必须显式写 false，不能留空"},
    {"name": "int", "cn": "整数", "fill": "数字", "example": "1001",
     "note": "列限定模式下留空当 0；流式模式下必须写 0"},
    {"name": "long", "cn": "长整数", "fill": "数字", "example": "10000000000", "note": "数值可能超过 int 时用"},
    {"name": "short", "cn": "短整数", "fill": "数字", "example": "100", "note": "注意取值范围 -32768~32767"},
    {"name": "byte", "cn": "字节", "fill": "数字", "example": "10", "note": "注意取值范围 0~255"},
    {"name": "float", "cn": "单精度小数", "fill": "数字", "example": "1.5", "note": ""},
    {"name": "double", "cn": "双精度小数", "fill": "数字", "example": "112233.5", "note": "精度要求高时用"},
    {"name": "string", "cn": "字符串", "fill": "任意文字", "example": "hello",
     "note": "列模式下空格子 = 空串；流式(sep)模式下空串必须写 \"\""},
    {"name": "text", "cn": "本地化文本", "fill": "本地化 key", "example": "item_name_1001",
     "note": "等价于 string#text=1，会校验 key 合法性"},
    {"name": "datetime", "cn": "日期时间", "fill": "yyyy-mm-dd hh:mm:ss",
     "example": "2024-01-01 12:00:00", "note": "导出为 UTC 秒(long)；一般不要留空"},
]
BASE_TYPE_NAMES = [t["name"] for t in BASE_TYPES]
_BASE_MAP = {t["name"]: t for t in BASE_TYPES}

CONTAINERS: List[Dict[str, str]] = [
    {"name": "list", "cn": "列表", "note": "元素不可为空，不能写 list,int?"},
    {"name": "array", "cn": "数组", "note": "元素不可为空"},
    {"name": "set", "cn": "集合", "note": "元素不可为空，自动去重"},
    {"name": "map", "cn": "映射", "note": "key / value 都不可为空；key 一般为基类型或 enum"},
]

# ------------------------------------------------------------------ 标签
TAGS: List[Dict[str, str]] = [
    {"key": "sep", "form": "#sep=,", "cn": "分隔符", "editable": "1", "default": ",",
     "desc": "在一个单元格内用分隔符拆分出多个子数据。可含多个字符（\",;\" 表示任一字符都算分隔符）。",
     "where": "可写在 Excel 字段名（x#sep=,）、bean 定义、或类型标签上",
     "note": "分隔符来自类型定义，不要凭单元格里的中文逗号猜测"},
    {"key": "size", "form": "#size=4", "cn": "长度约束", "editable": "1", "default": "4",
     "desc": "容器长度必须等于指定值。", "where": "容器标签：(list#size=4),int",
     "note": "元素个数不符会报错"},
    {"key": "ref", "form": "#ref=TbItem", "cn": "引用校验", "editable": "1", "default": "",
     "desc": "校验该值在目标表中存在。", "where": "字段类型：int#ref=item.TbItem",
     "note": "加 ? 表示可空引用：int#ref=item.TbItem?"},
    {"key": "range", "form": "#range=[0,100]", "cn": "范围校验", "editable": "1", "default": "[0,100]",
     "desc": "数值取值区间校验。", "where": "字段类型", "note": "写法以你工程的校验器为准"},
    {"key": "path", "form": "#path=1", "cn": "资源路径", "editable": "0", "default": "",
     "desc": "校验填写的是存在的资源路径。", "where": "字段类型", "note": ""},
    {"key": "text", "form": "#text=1", "cn": "本地化", "editable": "0", "default": "",
     "desc": "按本地化文本处理并校验 key。", "where": "字段类型", "note": "text 类型等价于 string#text=1"},
    {"key": "!", "form": "!", "cn": "禁止默认值", "editable": "0", "default": "",
     "desc": "不允许取该类型的默认值（如 int 的 0）。", "where": "类型后缀", "note": "用来逼策划显式填值"},
]

_TAG_MAP: Dict[str, Dict[str, str]] = {t["key"]: t for t in TAGS}

# 各类类型可用的标签 —— 类型助手的「常用标签」与字段表的「配置方式」共用这一份，
# 保证两处看到的东西一致（不再各写一套）。
TAGS_BY_KIND: Dict[str, List[str]] = {
    "atom": ["!", "ref", "range", "path"],
    "enum": ["!", "ref"],
    "bean": [],
    "poly": [],
    "container": ["sep", "size"],
    "map": ["sep", "size"],
}

# 这些标签本身就是一种「配置方式」，会出现在字段的「配置方式」下拉里
STYLE_TAG_KEYS = ("ref", "range", "path", "text", "size")


def tags_for_kind(kind: str, base: str = "") -> List[str]:
    """某一类类型用得上的标签键。string 额外给「本地化」。"""
    keys = list(TAGS_BY_KIND.get(kind, []))
    if kind == "atom" and base == "string" and "text" not in keys:
        keys.append("text")
    return keys


def kind_of(model: ProjectModel, type_str: str) -> Tuple[str, str]:
    """把类型串归到 TAGS_BY_KIND 的某一类，返回 (kind, base)。"""
    p = parse_type(type_str)
    if p["kind"] == "container":
        return ("map" if p["container"] == "map" else "container"), p["container"]
    k = resolve_kind(model, p["base"])
    if k == "bean" and model.is_polymorphic(p["base"]):
        return "poly", p["base"]
    return k, p["base"]

# ------------------------------------------------------------------ 解析
_CONTAINER_RE = re.compile(r"^\(?\s*(list|array|set|map)\s*(#[^,)]+)?\s*\)?\s*,?\s*(.*)$", re.I)


def _split_args(s: str) -> List[str]:
    """按逗号拆容器的类型参数，但括号里的逗号不拆。

    `map,string,(Item#sep=,)` 里的 `(Item#sep=,)` 是**一个**参数：它自己带了一个
    逗号。按裸字符串 split(",") 会把它劈成 `Item#sep=` 和 `)`，类型串就被拆坏了。
    """
    out: List[str] = []
    buf = ""
    depth = 0
    for ch in (s or ""):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            out.append(buf)
            buf = ""
        else:
            buf += ch
    out.append(buf)
    return [x.strip() for x in out if x.strip() != "" or len(out) == 0]


def parse_type(t: str) -> Dict[str, Any]:
    s = (t or "").strip()
    out: Dict[str, Any] = {"raw": s, "kind": "atom", "nullable": False, "bang": False,
                           "container": "", "sep": "", "elem": "", "value": "",
                           "tags": [], "base": ""}
    if not s:
        return out
    core = s
    if core.endswith("?"):
        out["nullable"] = True
        core = core[:-1]
    if core.endswith("!"):
        out["bang"] = True
        out["tags"].append("!")
        core = core[:-1]

    m = _CONTAINER_RE.match(core)
    if m and m.group(1).lower() in ("list", "array", "set", "map"):
        out["kind"] = "container"
        out["container"] = m.group(1).lower()
        if m.group(2):
            out["sep"] = m.group(2)
            out["tags"].append(m.group(2))
        parts = _split_args((m.group(3) or "").strip())
        out["elem"] = parts[0] if parts else ""
        out["value"] = parts[1] if len(parts) > 1 else ""
        out["base"] = out["container"]
        return out

    core2 = core.strip("()")
    if "#" in core2:
        head, _, tail = core2.partition("#")
        out["base"] = head.strip()
        out["tags"].append("#" + tail.strip())
        if tail.strip().startswith("sep="):
            out["sep"] = "#" + tail.strip()
    else:
        out["base"] = core2.strip()
    return out


def split_tags(t: str) -> Tuple[str, List[str]]:
    """拆出 (核心类型, 标签列表)。用于标签多选时避免重复追加。

    核心类型原样保留 —— 含括号的元素类型（如 `map,int,(Item#sep=,)`）不能被
    再去掉括号，否则拼回去时类型串就坏了。
    """
    s = (t or "").strip()
    tags: List[str] = []
    if s.endswith("?"):
        s = s[:-1]
    if s.endswith("!"):
        tags.append("!")
        s = s[:-1]
    m = re.match(r"^\(?\s*(list|array|set|map)\s*(#[^,)]+)?\s*\)?\s*,?\s*(.*)$", s, re.I)
    if m and m.group(1).lower() in ("list", "array", "set", "map"):
        cont = m.group(1).lower()
        if m.group(2):
            tags.append(m.group(2))
        rest = (m.group(3) or "").strip()
        return (f"{cont},{rest}" if rest else cont), tags
    if "#" in s:
        head, _, tail = s.partition("#")
        tags.append("#" + tail.strip())
        return head.strip(), tags
    return s, tags


def join_type(core: str, tags: List[str], nullable: bool = False) -> str:
    """按固定顺序拼回类型串：核心 + 容器标签(紧贴容器) + 后缀标签 + ?

    core 里若已带标签会先剥离，只按传入的 tags 重建，因此反复应用同一标签不会叠加。
    """
    core = split_tags(core or "")[0]
    suffix: List[str] = []
    inline: List[str] = []
    seen = set()
    for tg in tags:
        tg = tg.strip()
        if not tg:
            continue
        key = tg.split("=")[0]          # 同一标签只保留最后一次，避免重复追加
        if key in seen:
            suffix = [x for x in suffix if x.split("=")[0] != key]
            inline = [x for x in inline if x.split("=")[0] != key]
        seen.add(key)
        if tg == "!":
            suffix.append("!")
        elif tg.startswith("#sep=") or tg.startswith("#size="):
            inline.append(tg)
        else:
            suffix.append(tg)
    out = core
    if inline:
        m = re.match(r"^\(?\s*(list|array|set|map)\s*\)?\s*,?\s*(.*)$", core, re.I)
        if m:
            cont = m.group(1)
            rest = (m.group(2) or "").strip()      # 括号里的元素类型要原样保留
            out = f"({cont}{''.join(inline)}),{rest}" if rest else f"({cont}{''.join(inline)})"
        else:
            suffix = inline + suffix
    out = out + "".join(suffix)
    if nullable and not out.endswith("?"):
        out += "?"
    return out


def resolve_kind(model: ProjectModel, base: str) -> str:
    if base in BASE_TYPE_NAMES:
        return "atom"
    if model.find_enum(base):
        return "enum"
    if model.find_bean(base):
        return "bean"
    return "unknown"


def type_label(model: ProjectModel, code: str) -> str:
    """给下拉框用的中文显示名：int -> int（整数）"""
    code = (code or "").strip()
    if not code:
        return code
    p = parse_type(code)
    if p["kind"] == "container":
        c = next((x for x in CONTAINERS if x["name"] == p["container"]), None)
        return f"{code}（{c['cn']}）" if c else code
    if code in _BASE_MAP:
        return f"{code}（{_BASE_MAP[code]['cn']}）"
    if model.find_enum(p["base"]):
        return f"{code}（枚举）"
    if model.find_bean(p["base"]):
        b = model.find_bean(p["base"])
        tail = b.comment or ""
        return f"{code}（结构{'·' + tail if tail else ''}）"
    return code


def elem_label(model: ProjectModel, code: str) -> str:
    """「子类型」下拉框里显示的名字。

    元素类型经常带标签 —— 例如 `(Reward#sep=,)` 表示「Reward 这个结构在一个
    格子里用 `,` 分隔它的子字段」。把这种原始写法直接丢进下拉框，新手看不懂，
    也和「类型」下拉显示的 `Reward（结构）` 不是一套风格。
    所以这里剥掉外层括号与 `#tag`，只显示 `Reward（结构）· sep=,`，
    完整写法留给 tooltip。
    """
    code = (code or "").strip()
    if not code:
        return ""
    p = parse_type(code)
    lab = type_label(model, p["base"] or code)
    bits: List[str] = []
    for t in p["tags"]:
        t = t.lstrip("#").strip()
        if t == "!":
            bits.append("不可为空")
        elif t.startswith("sep="):
            bits.append("子字段用 %s 分隔" % t[4:])
        elif t:
            bits.append(t)
    if p["nullable"]:
        bits.append("可空")
    return lab + (" · " + " ".join(bits) if bits else "")


def strip_label(text: str) -> str:
    """从 'int（整数）' 取回 'int'"""
    return re.sub(r"[（(][^）)]*[）)]\s*$", "", (text or "").strip()).strip()


# 单元格格式：官方「紧凑格式」里的四个取值。stream 是默认值，写不写都一样。
CELL_FORMATS = ["stream", "lite", "json", "lua"]


def split_format(name: str) -> Tuple[str, str]:
    """把字段名里的 `#format=xxx` 拆出来。

    Luban 的约定是**写在字段名上**（`pos#format=lite`），不是写在 ##type 上 ——
    写在 ##type 上不生效，这是官方文档点名的常见坑。
    """
    s = (name or "").strip()
    m = re.search(r"#format\s*=\s*([A-Za-z]+)", s)
    if not m:
        return s, ""
    fmt = m.group(1).lower()
    return (s[:m.start()] + s[m.end():]).strip(), fmt


def join_format(name: str, fmt: str) -> str:
    """把 `#format=` 拼回字段名。stream 是默认值，不写出来（表头保持干净）。"""
    base, _old = split_format(name)
    fmt = (fmt or "").strip().lower()
    if fmt not in CELL_FORMATS or fmt == "stream":
        return base
    return f"{base}#format={fmt}"


def validate_type(model: ProjectModel, t: str) -> List[Dict[str, str]]:
    issues: List[Dict[str, str]] = []
    p = parse_type(t)
    if not p["raw"]:
        return [{"level": "error", "msg": "类型不能为空"}]
    if p["kind"] == "container":
        if p["container"] == "map" and not p["value"]:
            issues.append({"level": "error", "msg": "map 需要两个类型参数：map,K,V"})
        if p["container"] != "map" and not p["elem"]:
            issues.append({"level": "error", "msg": f"{p['container']} 缺少元素类型"})
        if "?" in (p["elem"] + p["value"]):
            issues.append({"level": "error",
                           "msg": "容器元素 / key / value 不可为空，不能写 list,int? 这类形式"})
        if p["container"] == "map" and resolve_kind(model, p["elem"]) not in ("atom", "enum"):
            issues.append({"level": "warn", "msg": "map 的 key 一般用基础类型或 enum"})
        if p["elem"] and resolve_kind(model, p["elem"]) == "unknown":
            issues.append({"level": "error", "msg": f"未知类型：{p['elem']}"})
    else:
        k = resolve_kind(model, p["base"])
        if k == "unknown":
            issues.append({"level": "error",
                           "msg": f"未知类型：{p['base']}（不在内置类型、枚举、结构中）"})
        if k == "bean" and model.is_polymorphic(p["base"]) and not p["nullable"]:
            issues.append({"level": "info",
                           "msg": f"{p['base']} 是多态基类，填数据时必须写具体子类名"})
    if p["nullable"] and p["kind"] == "container":
        issues.append({"level": "error", "msg": "容器不支持 list,T? 这种元素可空写法"})
    return issues


# ------------------------------------------------------------------ 样例值
def _sample_value(f: FieldDef, model: ProjectModel) -> str:
    t = f.type
    p = parse_type(t)
    if p["kind"] == "container":
        if p["container"] == "map":
            k = _sample_value(FieldDef(name="k", type=p["elem"] or "int"), model)
            v = _sample_value(FieldDef(name="v", type=p["value"] or "string"), model)
            return "|".join([f"{k}|{v}"] * 2)      # key/value 成对
        return "|".join(["1", "2"])
    # 注意：parse_type 对非容器一律记成 atom，所以这里必须再用 resolve_kind 判一次，
    # 否则枚举 / 结构也会被当成基础类型，示例值全变成 0（对新手是误导）。
    k = resolve_kind(model, p["base"])
    if k == "atom":
        b = _BASE_MAP.get(p["base"])
        return b["example"] if b else "0"
    if k == "enum":
        e = model.find_enum(p["base"])
        if e and e.items:
            return e.items[0].name
        return "ITEM_1"
    if k == "bean":
        b = model.find_bean(p["base"])
        if b and b.fields:
            sep = b.sep or ","
            return sep.join(_sample_value(x, model) for x in b.fields)
        return "{}"
    return "0"


def _bean_fields(model: ProjectModel, base: str) -> List[FieldDef]:
    b = model.find_bean(base)
    return list(b.fields) if b else []


def _bean_sep(model: ProjectModel, base: str) -> str:
    b = model.find_bean(base)
    return (b.sep if b and b.sep else ",")


def elem_of(model: ProjectModel, type_str: str) -> str:
    p = parse_type(type_str)
    return p["elem"] or "int"


# ------------------------------------------------------------------ 配置方式下拉
_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩"


def _circ(n: int) -> str:
    return _CIRCLED[n - 1] if 1 <= n <= len(_CIRCLED) else f"{n}."


def _with_tag_options(opts: List[Tuple[str, str]],
                      kind: str, base: str) -> List[Tuple[str, str]]:
    """把类型助手里同类型可用的「常用标签」也补成可选的配置方式。"""
    out = list(opts)
    keys = {k for k, _ in out}
    for key in tags_for_kind(kind, base):
        if key == "sep":
            continue                      # 已由「单单元格 + sep」覆盖
        if key == "!":
            if "bang" in keys:
                continue
            out.append(("bang", "禁止默认值（!）"))
            keys.add("bang")
            continue
        t = _TAG_MAP.get(key)
        if not t:
            continue
        out.append((f"tag:{key}", f"{t['cn']}（{t['form']}）"))
    return [(k, f"{_circ(i + 1)} {label}") for i, (k, label) in enumerate(out)]


def fill_style_options(model: ProjectModel, type_str: str) -> List[Tuple[str, str]]:
    """「配置方式」下拉框内容，由字段类型决定。返回 [(key, 中文说明)]。

    与类型助手保持一致：除了结构性的填法（流式 / sep / 列限定 / 多行…），
    还会把**同类型用得上的「常用标签」**（#ref / #range / #path / #text / #size）
    一并列出来，避免「类型助手里有、字段表里找不到」。
    """
    p = parse_type(type_str)
    if p["kind"] == "container":
        elem = elem_of(model, type_str)
        ek = resolve_kind(model, elem)
        opts = [("stream", "流式·多单元格（每元素一格）"),
                ("sep", "单单元格 + sep（一格装下）")]
        if p["container"] == "map":
            opts.append(("column", "多级表头（map 的 key 做列名）"))
        elif ek == "bean" and _bean_fields(model, elem):
            opts.append(("column", "列限定（多级 ##var）"))
        opts.append(("multiline", "多行读取（*字段名）"))
        kind = "map" if p["container"] == "map" else "container"
        return _with_tag_options(opts, kind, "")
    k = resolve_kind(model, p["base"])
    if k == "bean":
        poly = model.is_polymorphic(p["base"])
        opts = [("stream", "流式·多单元格"),
                ("sep", "单单元格 + sep"),
                ("column", "列限定（多级 ##var）")]
        if poly:
            opts.append(("poly", "多态：先写具体子类型名"))
        opts.append(("nullable", "可空：留空 / null 表示没有"))
        return _with_tag_options(opts, "poly" if poly else "bean", p["base"])
    if k == "enum":
        e = model.find_enum(p["base"])
        opts = [("plain", "填枚举名 / 别名 / 数字")]
        if e and str(e.flags).lower() in ("true", "1", "yes"):
            opts.append(("flags", "标志位组合（A|B）"))
        opts.append(("nullable", "可空（T?）"))
        return _with_tag_options(opts, "enum", p["base"])
    return _with_tag_options([("plain", "单列直接填"),
                              ("nullable", "可空（T?）"),
                              ("bang", "禁止默认值（!）")], "atom", p["base"])


# ------------------------------------------------------------------ 示例生成
def _opt(key: str, title: str, suitable: str, grid: List[List[str]], notes: List[str]) -> Dict[str, Any]:
    return {"key": key, "title": title, "suitable": suitable, "grid": grid, "notes": notes}


class _Pad:
    """续格：这一格被左边那一个单元格跨列合并（只在预览里用）。"""

    def __repr__(self) -> str:          # 调试时看得懂
        return "<续格>"


PAD = _Pad()


def expand_cells(items) -> List[Any]:
    """把 (文本, 跨列数) 展开成 [文本, PAD, PAD…]；普通字符串占 1 列。

    复合字段在数据表里常常横跨好几列：只有第一格写字段名和类型，后面几格是它
    的延续。以前这些续格被当成空单元格渲染，看起来像一堆独立的空列，初学者很
    容易以为「一个字段只占一列」。用 PAD 明确标出来，预览里就能画成合并单元格。
    """
    out: List[Any] = []
    for it in items or []:
        if isinstance(it, tuple):
            txt, span = (it[0], int(it[1])) if len(it) > 1 else (it[0], 1)
        else:
            txt, span = it, 1
        span = max(1, span)
        out.append(txt)
        out.extend([PAD] * (span - 1))
    return out


def _rows(var, typ, sub=None, group=None, comment=None, data=()) -> List[List[str]]:
    """组装一张示例表。列数取最宽的一行，其余行补空 —— 保证各行对齐、不被截断。"""
    v = expand_cells(var)
    n = max([len(v)] + [len(expand_cells(r)) for r in (typ, sub, group, comment)]
            + [len(expand_cells(d)) for d in data])

    def fit(row, expanded: bool = False):
        row = list(row) if expanded else expand_cells(row)
        return (row + [""] * n)[:n]

    g: List[List[str]] = [["##var"] + fit(v, True)]
    if sub:
        g.append(["##var"] + fit(sub))
    g.append(["##type"] + fit(typ))
    if group:
        g.append(["##group"] + fit(group))
    g.append(["##"] + fit(comment or []))
    for d in data:
        g.append([""] + fit(d))
    return g


def _renumber(opts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """统一编号：把示例标题里原有的 ①②③ 去掉后按顺序重排（多态/标志位等可选块不影响连贯性）。"""
    for i, o in enumerate(opts):
        title = re.sub(r"^[①②③④⑤⑥⑦⑧⑨⑩]\s*", "", str(o.get("title", "")))
        o["title"] = f"{_circ(i + 1)} {title}"
    return opts


def _tag_option(model: ProjectModel, key: str, type_str: str, fname: str,
                group: List[str], sample: str) -> Dict[str, Any]:
    """把一条「常用标签」渲染成配置方式示例（与类型助手同一份标签定义）。"""
    t = _TAG_MAP[key]
    p = parse_type(type_str)
    extra = ""
    if p["kind"] == "container":
        inner = p["container"] + (p["sep"] or "") + t["form"]
        if "#sep=" not in inner:
            inner += "#sep=|"             # 与模板生成保持一致：先一格装下
            extra = "（示例里同时写了 #sep=|，与「导出数据表模板」生成的一致）"
        rest = f"{p['elem']},{p['value']}" if p["value"] else p["elem"]
        if resolve_kind(model, p["elem"]) == "bean" and "#sep=" not in rest:
            rest = f"({rest}#sep={_bean_sep(model, p['elem'])})"   # 元素自己也要 sep 才能读出来
        typ = f"({inner}),{rest}" if rest else f"({inner})"
        if key == "size":
            try:
                n = max(1, int(t.get("default") or 4))
            except ValueError:
                n = 4
            sample = "|".join([sample] * n)
    else:
        typ = (p["base"] or type_str) + t["form"]
    notes = [f"写法：{typ}" + extra, t["desc"]]
    if t["note"]:
        notes.append(t["note"])
    return _opt(f"tag:{key}", f"{t['cn']}（{t['form']}）", t["desc"],
                _rows(["id", fname], ["int", typ], None, group,
                      ["ID", fname], [["1", sample]]),
                notes)


def fill_options(model: ProjectModel, field_name: str, type_str: str) -> List[Dict[str, Any]]:
    p = parse_type(type_str)
    fname = field_name or "x"
    opts: List[Dict[str, Any]] = []

    def add_tag_options(kind: str, base: str, group: List[str], sample: str):
        for key in tags_for_kind(kind, base):
            if key == "sep":
                continue                  # 已由「单单元格 + sep」覆盖
            if key == "!":
                continue                  # 已由「禁止默认值」覆盖
            opts.append(_tag_option(model, key, type_str, fname, group, sample))
        return opts

    # ==================== 容器 ====================
    if p["kind"] == "container":
        cont = p["container"]
        elem = elem_of(model, type_str)
        vtype = p["value"] or "string"
        ek = resolve_kind(model, elem)
        fields = _bean_fields(model, elem) if ek == "bean" else []
        n = len(fields)
        esep = _bean_sep(model, elem) if ek == "bean" else ","
        cn = next((c["cn"] for c in CONTAINERS if c["name"] == cont), "容器")
        esample = _sample_value(FieldDef(name="e", type=elem), model)
        vsample = _sample_value(FieldDef(name="v", type=vtype), model)
        # map 的元素是「key + value」一对，示例都按一对来给
        pair = f"{esample}|{vsample}" if cont == "map" else esample
        vsep = _bean_sep(model, vtype) if resolve_kind(model, vtype) == "bean" else ","
        grp = ["", "c,s"]

        # ① 流式·多单元格
        if ek == "bean" and n and cont != "map":
            one = [_sample_value(f, model) for f in fields]
            sp = n * 2
            opts.append(_opt("stream", "① 流式·多单元格",
                             "元素字段少、希望一眼看清；按从左到右顺序读，遇到空格子即停止",
                             _rows(["id", (fname, sp)],
                                   ["int", (f"{cont},{elem}", sp)],
                                   None, ["", ("c,s", sp)],
                                   ["ID"] + [(f.comment or f.name)
                                             for f in fields] * 2,
                                   [["1"] + one + one,
                                    ["2"] + one]),
                             [f"每元素 {n} 个字段，示例中放 2 个元素共 {sp} 列，"
                              f"这 {sp} 列都属于字段 {fname}",
                              "空格子常被跳过，因此元素个数不定，读到空即结束"]))
        elif cont == "map":
            opts.append(_opt("stream", "① 流式·多单元格",
                             "在合并范围内 key、value 成对填写",
                             _rows(["id", (fname, 3)],
                                   ["int", (f"map,{elem},{vtype}", 3)],
                                   None, ["", ("c,s", 3)],
                                   ["ID", (fname, 3)],
                                   [["1", esample, vsample, esample],
                                    ["2", esample, vsample]]),
                             ["map 按 key、value 成对循环读，读到 } 或结束为止",
                              "类型必须写全 map,key类型,value类型"]))
        else:
            opts.append(_opt("stream", "① 流式·多单元格",
                             "元素个数不固定、希望一个元素占一格",
                             _rows(["id", (fname, 3)],
                                   ["int", (f"{cont},{elem}", 3)],
                                   None, ["", ("c,s", 3)],
                                   ["ID", (fname, 3)],
                                   [["1", esample, esample, esample],
                                    ["2", esample, esample]]),
                             ["从左到右读，空白单元格会被跳过",
                              "元素不可为空，不能写 list,int?"]))

        # ② 单单元格 + sep
        if ek == "bean" and n and cont != "map":
            one = esep.join(_sample_value(f, model) for f in fields)
            opts.append(_opt("sep", "② 单单元格 + sep（最省列）",
                             "字段很多、表格太宽时；一个格子装下整个容器",
                             _rows(["id", fname],
                                   ["int", f"({cont}#sep=|),({elem}#sep={esep})"],
                                   None, grp,
                                   ["ID", fname],
                                   [["1", "|".join([one, one])],
                                    ["2", one]]),
                             [f"容器按 | 拆分元素，元素内部按 {esep} 拆分字段",
                              "流式单元格里不能靠留空表示默认值：bool 写 false、数字写 0、空串写 \"\"",
                              "分隔符来自 bean 定义的 sep，不要凭单元格里的中文逗号猜测"]))
        else:
            if cont == "map":
                # map 的类型必须写全 key 和 value（少了 value Luban 会报参数不足）
                vwrap = (f"({vtype}#sep={vsep})" if resolve_kind(model, vtype) == "bean"
                         else vtype)
                opts.append(_opt("sep", "② 单单元格 + sep（最省列）",
                                 "一格里按 key/value 交替填，最省列",
                                 _rows(["id", fname],
                                       ["int", f"({cont}#sep=|),{elem},{vwrap}"],
                                       None, grp,
                                       ["ID", fname],
                                       [["1", "|".join([pair, pair])],
                                        ["2", pair]]),
                                 [f"容器按 | 拆分后，map 交替读 key、value："
                                  f"{esample} 是 key、{vsample} 是 value",
                                  "类型写全：(map#sep=|),key类型,value类型",
                                  "value 是结构时再给它一个 sep："
                                  "(map#sep=|),string,(Item#sep=,)"]))
            else:
                opts.append(_opt("sep", "② 单单元格 + sep（最省列）",
                                 "元素很多、不想占很多列时",
                                 _rows(["id", fname],
                                       ["int", f"({cont}#sep=|),{elem}"],
                                       None, grp,
                                       ["ID", fname],
                                       [["1", "|".join([esample, esample, esample])],
                                        ["2", "|".join([esample, esample])]]),
                                 ["(list#sep=|) 表示：下一个字符串当作整个容器，按 | 拆分",
                                  "元素类型也可单独带 sep：(list#sep=|),(Vec#sep=,)"]))

        # map 专属：多级表头
        if cont == "map":
            opts.append(_opt("column", "③ 多级表头（map 的 key 做列名）",
                             "key 集合固定时，可读性最好",
                             _rows(["id", (fname, 2)],
                                   ["int", (f"map,{elem},{p['value'] or 'string'}", 2)],
                                   ["", "ch-zn", "en-us"],
                                   ["", ("c,s", 2)],
                                   ["ID", (fname, 2)],
                                   [["1", "苹果", "apple"],
                                    ["2", "香蕉", "banana"]]),
                             ["子字段行用 ##var 写，每个 key 占一列；这两列同属一个字段",
                              "map 的 key 一般用基础类型或 enum"]))

        # ③ 列限定
        if ek == "bean" and n and cont != "map":
            one = [_sample_value(f, model) for f in fields]
            sp = n * 2
            opts.append(_opt("column", "③ 列限定（多级 ##var 子字段行）",
                             "要求每个子字段都钉在固定列上，允许原子字段留空取默认值",
                             _rows(["id", (fname, sp)],
                                   ["int", (f"{cont},{elem}", sp)],
                                   [""] + [f.name for f in fields] * 2,
                                   ["", ("c,s", sp)],
                                   ["ID"] + [(f.comment or f.name)
                                             for f in fields] * 2,
                                   [["1"] + one + one,
                                    ["2"] + one]),
                             ["子字段名写在第二个 ##var 行，Luban 靠它把列和字段对上",
                              "原子类型的子字段留空可取默认值；没有被钉住的部分仍按流式读"]))

        # ④ 多行
        if ek == "bean" and n and cont != "map":
            one = [_sample_value(f, model) for f in fields]
            opts.append(_opt("multiline", "④ 多行读取（*字段名）",
                             "一对多嵌套、元素个数不定且希望纵向书写时",
                             _rows(["id", (f"*{fname}", n)],
                                   ["int", (f"{cont},{elem}", n)],
                                   [""] + [f.name for f in fields],
                                   ["", ("c,s", n)],
                                   ["ID"] + [(f.comment or f.name) for f in fields],
                                   [["1"] + one,
                                    [""] + one,
                                    ["2"] + one]),
                             ["字段名加 * 前缀表示多行；同一条记录的每个元素占一行",
                              "每个元素自身也可以用分列 / sep / 列限定任一方式"]))
        else:
            mtyp = f"map,{elem},{vtype}" if cont == "map" else f"{cont},{elem}"
            opts.append(_opt("multiline", "④ 多行读取（*字段名）",
                             "每行一对或多对 key-value" if cont == "map"
                             else "元素纵向排列，行数不限（仅容器可用）",
                             _rows(["id", (f"*{fname}", 1)],
                                   ["int", (mtyp, 1)],
                                   None, ["", "c,s"],
                                   ["ID", fname],
                                   [["1", pair if cont == "map" else esample],
                                    ["", pair if cont == "map" else esample],
                                    ["2", pair if cont == "map" else esample]]),
                             ["* 只对容器类型有效",
                              "map 多行时每行写 key、value（可写多对）" if cont == "map"
                              else "同一条记录的每个元素占一行，行数不限"]))
        return _renumber(add_tag_options("map" if cont == "map" else "container", "",
                                         grp, pair))

    # ==================== 结构 bean ====================
    k = resolve_kind(model, p["base"])
    if k == "bean":
        fields = _bean_fields(model, p["base"])
        n = max(len(fields), 1)
        sep = _bean_sep(model, p["base"])
        grp = ["", "c,s"]
        one = [_sample_value(f, model) for f in fields]
        poly = model.is_polymorphic(p["base"])
        subs = model.subclasses_of(p["base"])

        opts.append(_opt("stream", "① 流式·多单元格",
                         "字段少、一列一字段；按 bean 字段顺序从左到右读",
                         _rows(["id", (fname, n)],
                               ["int", (p["base"], n)],
                               None, ["", ("c,s", n)],
                               ["ID"] + [(f.comment or f.name) for f in fields],
                               [["1"] + one]),
                         [f"字段名与类型只写一次，横跨这 {n} 列；后面几列都是同一个结构",
                          "空白单元格常被跳过，因此建议放在行的末尾",
                          "想要「留空取默认值」请用下面的列限定方式"]))

        opts.append(_opt("sep", "② 单单元格 + sep",
                         "表格太宽、结构紧凑时；一格装下整个结构",
                         _rows(["id", f"{fname}#sep={sep}"],
                               ["int", p["base"]],
                               None, grp,
                               ["ID", fname],
                               [["1", sep.join(one)]]),
                         [f"分隔符来自 bean 定义的 sep（这里取 {sep}），不要凭感觉改",
                          "流式单元格里 bool / 数字 / 字符串都要显式写值，不能留空"]))

        opts.append(_opt("column", "③ 列限定（多级 ##var 子字段行）",
                         "可读性好、允许部分字段留空取默认值",
                         _rows(["id", (fname, n)],
                               ["int", (p["base"], n)],
                               [""] + [f.name for f in fields],
                               ["", ("c,s", n)],
                               ["ID"] + [(f.comment or f.name) for f in fields],
                               [["1"] + one]),
                         ["子字段名写在第二个 ##var 行",
                          "原子类型的子字段留空可取默认值"]))

        if poly:
            sub_names: List[str] = []
            sub_types: List[str] = []
            for s in subs[:2]:
                for f in s.fields:
                    sub_names.append(f.name)
                    sub_types.append(f.type)
            if not sub_names:
                sub_names, sub_types = ["radius"], ["float"]
            w = len(sub_names)
            opts.append(_opt("poly", "④ 多态：先写具体子类型名",
                             f"{p['base']} 是多态基类，必须标明这一格到底是哪个子类",
                             _rows(["id", (fname, w + 1)],
                                   ["int", (p["base"], w + 1)],
                                   ["", "$type"] + sub_names,
                                   ["", ("c,s", w + 1)],
                                   ["ID"] + (["子类型名"] + sub_names)[:w + 1],
                                   [["1", (subs[0].short_name if subs else "Circle")] +
                                    [_sample_value(FieldDef(name=x, type=sub_types[i]), model)
                                     for i, x in enumerate(sub_names)],
                                    ["2", (subs[1].short_name if len(subs) > 1 else "Rect")] +
                                    [_sample_value(FieldDef(name=x, type=sub_types[i]), model)
                                     for i, x in enumerate(sub_names)]]),
                             ["$type 列写子类名或别名；不用 $type 这个列名时，"
                              "约定仍是「范围内第一个格子 = 类型名」",
                              "常见坑：直接写抽象基类名、子类名或别名拼错"]))
            opts.append(_opt("poly", "④ 多态 · 单单元格写法",
                             "用 sep 或 lite 格式把类型名和字段写在一格",
                             _rows(["id", f"{fname}#sep={sep}"],
                                   ["int", p["base"]],
                                   None, grp,
                                   ["ID", fname],
                                   [["1", sep.join([subs[0].short_name] + one) if subs
                                     else sep.join(["Circle"] + one)]]),
                             ["第一个值必须是具体子类名，其后是该子类的字段",
                              "也可用 #format=lite"]))

        opts.append(_opt("nullable", "⑤ 可空结构（T?）",
                         "这个字段允许「没有值」时",
                         _rows(["id", fname],
                               ["int", p["base"] + "?"],
                               None, grp,
                               ["ID", fname],
                               [["1", "null"],
                                ["2", "{}" + (sep.join(one) if one else "")]]),
                         ["留空或填 null 表示没有，这与填一个全 0 的结构不是一回事",
                          "非多态 bean? 有值时先写 {} 再按字段顺序填"]))
        return _renumber(add_tag_options("poly" if poly else "bean", p["base"], grp,
                                         one[0] if one else "0"))

    # ==================== 枚举 ====================
    if k == "enum":
        e = model.find_enum(p["base"])
        is_flags = bool(e and str(e.flags).lower() in ("true", "1", "yes"))
        names = [i.name for i in (e.items if e else [])][:3] or ["A", "B", "C"]
        aliases = [i.alias for i in (e.items if e else []) if i.alias][:2]
        grp = ["", "c,s"]
        opts.append(_opt("plain", "① 填枚举名 / 别名 / 数字",
                         "绝大多数情况",
                         _rows(["id", fname],
                               ["int", p["base"]],
                               None, grp,
                               ["ID", fname],
                               [["1", names[0]],
                                ["2", aliases[0] if aliases else names[-1]]]),
                         ["名字、别名、数字都能识别；别名以枚举定义里写的为准"]))
        opts.append(_opt("flags", "② 标志位组合"
                         + ("" if is_flags else "（仅当枚举定义里 flags=true）"),
                         "枚举定义里 flags=true 时，可以多选",
                         _rows(["id", fname],
                               ["int", p["base"]],
                               None, grp,
                               ["ID", fname],
                               [["1", "|".join(names[:2])]]),
                         ["多个标志用 | 连接，结果为按位或",
                          "只有 flags 枚举能这么填；普通枚举请改用「引用校验」或逐个字段拆开"]))
        opts.append(_opt("nullable", "③ 可空（T?）",
                         "允许「不填」时",
                         _rows(["id", fname],
                               ["int", p["base"] + "?"],
                               None, grp,
                               ["ID", fname],
                               [["1", "null"]]),
                         ["留空或填 null 表示没有值"]))
        opts.append(_opt("bang", "禁止默认值（!）",
                         "不允许取该枚举的默认值（第一项），逼策划显式填",
                         _rows(["id", fname],
                               ["int", p["base"] + "!"],
                               None, grp,
                               ["ID", fname],
                               [["1", names[0]]]),
                         ["写在类型后缀：EnumName!"]))
        return _renumber(add_tag_options("enum", p["base"], grp, names[0]))

    # ==================== 基础类型 ====================
    b = _BASE_MAP.get(p["base"])
    fill = b["fill"] if b else "按类型填写"
    ex = b["example"] if b else "0"
    cn = b["cn"] if b else p["base"]
    grp = ["", "c,s"]
    notes = [f"合法写法：{fill}"]
    if b and b["note"]:
        notes.append(b["note"])
    opts.append(_opt("plain", f"① 单列直接填（{cn}）",
                     "默认用法",
                     _rows(["id", fname], ["int", type_str], None, grp,
                           ["ID", fname], [["1", ex]]),
                     notes))
    opts.append(_opt("nullable", "② 可空（T?）",
                     "这个字段允许「没有值」时",
                     _rows(["id", fname],
                           ["int", p["base"] + "?"],
                           None, grp,
                           ["ID", fname],
                           [["1", "null"],
                            ["2", ""]]),
                     ["留空或填 null 都表示空；这与填 0 不是一回事"]))
    opts.append(_opt("bang", "③ 禁止默认值（!）",
                     "不允许取该类型的默认值，逼策划显式填",
                     _rows(["id", fname],
                           ["int", p["base"] + "!"],
                           None, grp,
                           ["ID", fname],
                           [["1", ex]]),
                     ["int! 表示不允许为 0"]))
    return _renumber(add_tag_options("atom", p["base"], grp, ex))
