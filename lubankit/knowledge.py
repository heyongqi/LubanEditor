"""Luban 核心知识点库（概念 / 规则 / 填法 / 常见错误）。

内容口径来自 Luban 官方文档（datable.cn），本地缓存，离线可用。
界面做「上下文跟随」：出现 list / 多态 / ? / sep / group 等关键词时就地弹出对应条目。
"""
from __future__ import annotations

from typing import List, Dict, Optional

SRC_CONF = "https://www.datable.cn/docs/concepts/luban-conf"
SRC_HEADER = "https://www.datable.cn/docs/excel/basics"
SRC_TYPES = "https://www.datable.cn/docs/schema/types"
SRC_EXCEL_SCHEMA = "https://www.datable.cn/docs/schema/excel-schema"
SRC_SEP = "https://www.datable.cn/docs/excel/vertical-and-sep"
SRC_POLY = "https://www.datable.cn/docs/excel/polymorphism"
SRC_ADD = "https://www.datable.cn/docs/guide/add-table"

KNOWLEDGE: List[Dict[str, str]] = [
    # ---------------- 概念 ----------------
    {"id": "concept_schema", "cat": "概念", "title": "Schema（定义）与 Data（数据）是两件事",
     "tags": "schema 定义 契约 tables beans enums",
     "body": "Luban 分两层：结构定义（有哪些表、每张表有哪些字段、字段什么类型）与数据（Excel 里的一行行）。\n"
             "定义用 __tables__.xlsx（表声明）与 Defines/*.xml（bean / enum）维护；数据放在 dataDir 下。\n"
             "顺序是：先有定义，再填数据。别指望工具从 Excel 猜出结构再覆盖定义——Schema 是契约。",
     "source": SRC_EXCEL_SCHEMA},
    {"id": "concept_three", "cat": "概念", "title": "三张定义表各管什么",
     "tags": "tables beans enums 定义文件",
     "body": "__tables__.xlsx：有哪些表、数据文件（input）、主键（index）、表模式（mode）、导出分组。\n"
             "结构（bean）与枚举（enum）写在 Defines/*.xml 里：公共的进 common.xml，表专用的进 <模块名>.xml 或 <配置表名>.xml。\n"
             "定义文件都必须在 luban.conf 的 schemaFiles 里登记，漏登记等于没定义。",
     "source": SRC_CONF},
    {"id": "concept_mode", "cat": "概念", "title": "表模式 mode：map / list / one",
     "tags": "mode map list one singleton 表模式",
     "body": "map：普通 key-value 表，按主键查，代码里是 Get(id)。index 为空时自动取第一个字段当主键。\n"
             "list：列表表，没有主键或多个主键时用，只能遍历。\n"
             "one（或 singleton）：全局只有一份的配置，如背包初始容量。\n"
             "留空则由 index 决定：index 为空或 1 个主键 → map；多个主键 → list。\n"
             "多主键规则：联合主键写 key1+key2（多列共同唯一），独立索引写 key1,key2（多个独立查找字典）。",
     "source": SRC_EXCEL_SCHEMA},
    {"id": "concept_group", "cat": "概念", "title": "分组 group：c / s / e 与导出目标",
     "tags": "group c s e 分组 导出 客户端 服务器",
     "body": "group 决定这张表/这个字段进哪个包：c 客户端、s 服务器、e 编辑器。\n"
             "表或类型不写 group：当导出目标的任一 group 在 luban.conf 里 default=true 时就会导出。\n"
             "字段的 group 为空 = 属于所有分组（多数字段如此），这与表级 group 的语义不同。\n"
             "填错 group 的表现是「字段不见了」，而不是报错——服务器敏感字段别漏到 c。",
     "source": SRC_CONF},
    {"id": "concept_read_schema", "cat": "概念", "title": "read_schema_from_file：要不要从数据表头读字段",
     "tags": "read_schema_from_file define_from_file 从文件读定义 重复定义",
     "body": "true：不用在 xml 里定义行结构，Luban 直接从数据 Excel 的表头推断字段。新手推荐。\n"
             "false：字段在 xml（表结构 bean）里显式定义，适合结构稳定、需要继承/多态的表。\n"
             "红线：为 true 时，绝不能再在 xml 里定义同名 bean，否则报「定义重复」。",
     "source": SRC_ADD},
    {"id": "concept_poly", "cat": "概念", "title": "多态 bean：技能/AI 这类配置的正确表达",
     "tags": "多态 继承 子类 type 抽象 bean Shape",
     "body": "一个抽象基类 + 若干子类：字段类型写基类，填数据时先写具体子类型名（或 alias），再按该子类字段顺序填。\n"
             "分列填写时用 $type 列标明类型；单单元格写法下类型名就是第一个值，如 Circle,1.5。\n"
             "常见坑：直接写抽象类名 Shape，或类型名/别名拼错。",
     "source": SRC_POLY},
    {"id": "concept_sep", "cat": "概念", "title": "sep：一个单元格装下复合数据",
     "tags": "sep 分隔符 流式 单单元格",
     "body": "sep 会把一个字符串拆成数据流再按字段顺序读入。可写在 Excel 字段名（x#sep=,）、bean 定义、或类型标签上。\n"
             "sep 含多个字符时表示「任一字符都是分隔符」，不是整体作为一个分隔符。\n"
             "容器上：(list#sep=|),int 表示元素用 | 分隔；元素自身也可带 sep：(list#sep=|),(Vec#sep=,)。\n"
             "重要：分隔符由类型定义决定，不能因为表里看到中文逗号就全局替换。",
     "source": SRC_SEP},
    {"id": "concept_stream", "cat": "概念", "title": "流式读取 vs 列限定 vs 多行",
     "tags": "流式 列限定 多行 四种读法 空白",
     "body": "复合数据在 Excel 里有四种读法：\n"
             "① 流式·多单元格：占多列，从左到右读，空白常被跳过；\n"
             "② 流式·单单元格 + sep：一格内拆分；\n"
             "③ 列限定·多单元格：用多级 ##var 把子字段钉到具体列，原子类型留空可取默认值；\n"
             "④ 多行（仅容器）：字段名 *name，每行一个元素。\n"
             "关键差异：流式/sep 单元格里留空不等于默认值——bool 要写 false，数字要写 0，空串要写 \"\"。",
     "source": SRC_SEP},
    {"id": "concept_vertical", "cat": "概念", "title": "纵表：单例配置更好读的写法",
     "tags": "纵表 竖表 column vertical 单例",
     "body": "横表一行一条记录；纵表（A1 写 ##column 或 ##vertical）一行一个字段，适合全局单例配置。\n"
             "典型场景：公会开放等级、背包初始大小/上限这类「全表只有一份」的配置。",
     "source": SRC_SEP},

    # ---------------- 规则 ----------------
    {"id": "rule_a1", "cat": "规则", "title": "A1 必须以 ## 开头，否则整表被忽略",
     "tags": "A1 忽略 不导出 表头",
     "body": "xlsx 会读入所有 sheet，但 A1 不以 ## 开头的 sheet 会被整体忽略（可以用来放策划说明）。\n"
             "表现是「我明明建了表，却完全没导出」。\n"
             "另外：注释行如果排在第一位，必须写 ##comment 而不是 ##。",
     "source": SRC_HEADER},
    {"id": "rule_header", "cat": "规则", "title": "表头几行各自的含义",
     "tags": "var type group comment 表头",
     "body": "##var（或首格为 ##）：字段名行；\n"
             "##type：类型行，不要写成中文或「看起来像」的单词；\n"
             "##group：导出分组（c/s/e，留空=全部）；\n"
             "## / ##comment：中文说明行，不进包，多个 ## 行时默认第一行作为代码注释。\n"
             "以 ## 开头的行顺序可以调换，不影响结果。",
     "source": SRC_HEADER},
    {"id": "rule_comment_col", "cat": "规则", "title": "注释列与注释行：别让备注进游戏",
     "tags": "注释列 备注 注释行",
     "body": "字段名以 # 开头（如 #备注）或为空的列 = 注释列，填了也不导出。\n"
             "数据行第一格以 ## 开头 = 注释行，整行忽略。\n"
             "常见事故：备注列没加 #，内容进了包；或者某行改了不生效，因为行首有个 ##。",
     "source": SRC_HEADER},
    {"id": "rule_nullable", "cat": "规则", "title": "可空 T?：与填 0 不是一回事",
     "tags": "可空 null 默认值",
     "body": "原子类型、enum、非多态 bean 都可以写 T? 表示「没有值」，填 null 或留空即可。\n"
             "int? 留空 ≠ 填 0；string? 留空 = null，想要空串要写 \"\"。\n"
             "容器不支持元素可空：不能写 list,int? 或 map,int?,string。",
     "source": SRC_TYPES},
    {"id": "rule_container", "cat": "规则", "title": "容器：list / array / set / map 的硬约束",
     "tags": "list array set map 容器 不可空",
     "body": "list,T / array,T / set,T / map,K,V。\n"
             "元素、key、value 都不可为空。map 的 key 一般用基础类型或 enum。\n"
             "带标签时加在容器上并用括号：(list#sep=|),int、(list#size=4),int。",
     "source": SRC_TYPES},
    {"id": "rule_fieldname", "cat": "规则", "title": "字段名命名：xx_yy_zz",
     "tags": "字段名 命名 Pascal camel 规范",
     "body": "推荐小写下划线风格；生成代码时 Luban 会按目标语言转成 PascalCase / camelCase。\n"
             "别用中文、别用空格、别以数字开头。",
     "source": SRC_HEADER},
    {"id": "rule_basic_fill", "cat": "规则", "title": "基础类型的合法写法",
     "tags": "bool int float string datetime 填法",
     "body": "bool：true/false、0/1、是/否，大小写不敏感，其它值报错。\n"
             "整数/浮点：数字，列限定模式下留空当 0。\n"
             "string：空格子=空串；流式(sep)模式下空串必须写 \"\"。\n"
             "datetime：yyyy-mm-dd hh:mm:ss，导出为 UTC 秒(long)，一般不要留空。\n"
             "枚举：可填项名、别名或数字。",
     "source": SRC_HEADER},
    {"id": "rule_tags", "cat": "规则", "title": "常用校验标签",
     "tags": "ref range path text size 标签 校验",
     "body": "#ref=item.TbItem：引用校验（加 ? 表示可空引用）；\n"
             "#range：数值区间；#path：资源路径存在性；\n"
             "#text=1：按本地化文本处理并校验 key（text 类型等价于 string#text=1）；\n"
             "类型后加 !：禁止默认值（如 int! 不允许为 0）；\n"
             "#size=4：容器长度必须为 4。",
     "source": SRC_TYPES},

    # ---------------- 常见错误 ----------------
    {"id": "err_ignored", "cat": "常见错误", "title": "整张表没导出",
     "tags": "没导出 忽略 A1 未登记",
     "body": "排查两步：① sheet 左上角是不是 ##；② 这张表有没有在 __tables__.xlsx 里登记（input 路径对不对）。",
     "source": SRC_ADD},
    {"id": "err_missing_field", "cat": "常见错误", "title": "某列客户端没有 / 字段不见了",
     "tags": "字段缺失 group 分组 客户端",
     "body": "看 ##group 行：写成 s 的字段不会进客户端包。改成 c 或 c,s，或留空（按项目规范）。\n"
             "同时确认 luban.conf 里导出目标的 groups 是否包含该分组。",
     "source": SRC_HEADER},
    {"id": "err_type", "cat": "常见错误", "title": "类型报错 / 枚举名拼错",
     "tags": "类型报错 枚举 拼错",
     "body": "对照 ##type 行与枚举名单；文字填进了 int 列也会报错。\n"
             "自定义类型要写全名或本模块简名，如 Quality、item.Cost。",
     "source": SRC_HEADER},
    {"id": "err_ref", "cat": "常见错误", "title": "引用 / id 校验失败",
     "tags": "引用 ref id 不存在",
     "body": "填了不存在的道具 id 等。先到目标表确认该 id 存在，再回来填。",
     "source": SRC_TYPES},
    {"id": "err_dup", "cat": "常见错误", "title": "定义重复",
     "tags": "重复定义 duplicate bean",
     "body": "read_schema_from_file=true 的表，其 value_type 不能再在 xml 里定义一次。二选一。",
     "source": SRC_ADD},
    {"id": "err_sep_count", "cat": "常见错误", "title": "sep 段数与字段数对不上",
     "tags": "sep 段数 字段数 错位 流式",
     "body": "一个单元格被 sep 拆出的段数必须和要填的字段数一致。\n"
             "另外流式单元格里把 0/false/空串留空，会导致后面所有字段错位——必须显式写值。",
     "source": SRC_SEP},

    # ---------------- 补充（官方文档要点） ----------------
    {"id": "concept_compact", "cat": "概念", "title": "紧凑格式 compact：省列又省格子",
     "tags": "compact 紧凑 分隔 数组",
     "body": "给容器或结构加 #compact 标签后，元素用默认分隔符紧凑排布，"
             "不用逐列铺开，适合「字段不多、条数多」的列表（如奖励列表）。",
     "source": SRC_TYPES},
    {"id": "concept_parent", "cat": "概念", "title": "bean 继承 parent：复用字段",
     "tags": "parent 继承 复用字段 bean",
     "body": "bean 可写 parent 指向另一个 bean，自动继承它的所有字段。\n"
             "子 bean 只需写新增字段；多态 + 继承组合起来能表达很复杂的配置。",
     "source": SRC_POLY},
    {"id": "concept_datetime", "cat": "概念", "title": "datetime：本地时间转 UTC 秒",
     "tags": "datetime 时间 时区 UTC 秒",
     "body": "datetime 写成 yyyy-mm-dd hh:mm:ss，导出时转成 UTC 秒(long)。\n"
             "一般不要留空；跨时区数据要注意转换。",
     "source": SRC_HEADER},
    {"id": "err_nested_header", "cat": "常见错误", "title": "多级表头不生效：二级 ##var 没写",
     "tags": "多级 表头 嵌套 二级 var",
     "body": "要二级（嵌套）表头，必须在字段名行下面**再写一行 ##var**。\n"
             "只写一行 ## + 中文列名是行说明，不是二级表头——这俩容易混。",
     "source": SRC_HEADER},
    {"id": "err_flags", "cat": "常见错误", "title": "flags 枚举填不对",
     "tags": "flags 枚举 多选 位标记",
     "body": "flags 枚举可多选，填 A|B|C（用 | 连接）。\n"
             "普通枚举只能填一项；把 flags 枚举当普通枚举填会解析失败。",
     "source": SRC_TYPES},
]

_INDEX: Dict[str, List[Dict[str, str]]] = {}


def _build_index():
    if _INDEX:
        return
    for k in KNOWLEDGE:
        for tok in (k["title"] + " " + k["tags"] + " " + k["body"]).replace("\n", " ").split():
            tok = tok.strip().lower()
            if len(tok) < 2:
                continue
            _INDEX.setdefault(tok, []).append(k)
        for tok in k["tags"].split():
            _INDEX.setdefault(tok.lower(), []).append(k)


def search(q: str) -> List[Dict[str, str]]:
    _build_index()
    q = (q or "").strip().lower()
    if not q:
        return list(KNOWLEDGE)
    hits: List[Dict[str, str]] = []
    seen = set()
    for tok in q.split():
        for k in _INDEX.get(tok, []):
            if k["id"] not in seen:
                seen.add(k["id"])
                hits.append(k)
    if not hits:
        for k in KNOWLEDGE:
            if q in (k["title"] + k["body"] + k["tags"]).lower():
                hits.append(k)
    return hits


def by_id(kid: str) -> Optional[Dict[str, str]]:
    for k in KNOWLEDGE:
        if k["id"] == kid:
            return k
    return None


def by_category(cat: str) -> List[Dict[str, str]]:
    return [k for k in KNOWLEDGE if k["cat"] == cat]


CATEGORIES = ["概念", "规则", "常见错误"]

# 知识点列表的两栏：核心知识（概念 + 规则）/ 常见问题（常见错误）
CATEGORY_GROUPS = [
    ("核心知识", ["概念", "规则"]),
    ("常见问题", ["常见错误"]),
]
