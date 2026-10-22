"""LubanEditor 数据模型：与 Luban 定义文件无关的「中间模型」。

界面只操作本模块的对象，写回 xlsx 由 exceldef 负责。
这样撤销/重做、diff、导出都只需要处理一层结构。
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional


@dataclass
class FieldDef:
    name: str = ""
    type: str = "int"
    group: str = ""
    comment: str = ""
    tags: str = ""
    # 配置方式（决定模板怎么展开这一列）：auto/stream/sep/column/multiline/
    # nullable/bang/flags/poly
    fill_style: str = "auto"
    # 单元格格式：空(=stream 默认) / lite / json / lua。写在数据表 ##var 的**字段名**上
    # （`pos#format=lite`），只对结构 / 容器这类非原子字段有意义 —— 见官方文档「紧凑格式」。
    fmt: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "FieldDef":
        return FieldDef(
            name=d.get("name", "") or "",
            type=d.get("type", "int") or "int",
            group=d.get("group", "") or "",
            comment=d.get("comment", "") or "",
            tags=d.get("tags", "") or "",
            fill_style=d.get("fill_style", "auto") or "auto",
            fmt=d.get("fmt", "") or "",
        )


@dataclass
class TableDef:
    full_name: str = ""
    value_type: str = ""
    read_schema_from_file: str = "true"   # true / false / 空
    input: str = ""
    index: str = ""
    mode: str = ""          # 空(=map) / map / list / one
    group: str = ""
    comment: str = ""
    tags: str = ""
    output: str = ""
    _row: Optional[int] = None
    _sheet: str = ""
    _file: str = ""              # 从哪个定义文件读出来的（xlsx 或 xml，保存时写回原处）

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["_row"] = self._row
        d["_sheet"] = self._sheet
        d["_file"] = self._file
        return d

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "TableDef":
        t = TableDef(
            full_name=d.get("full_name", "") or "",
            value_type=d.get("value_type", "") or "",
            read_schema_from_file=d.get("read_schema_from_file", "") or "",
            input=d.get("input", "") or "",
            index=d.get("index", "") or "",
            mode=d.get("mode", "") or "",
            group=d.get("group", "") or "",
            comment=d.get("comment", "") or "",
            tags=d.get("tags", "") or "",
            output=d.get("output", "") or "",
        )
        t._row = d.get("_row")
        t._sheet = d.get("_sheet", "") or ""
        t._file = d.get("_file", "") or ""
        return t

    @property
    def real_mode(self) -> str:
        m = (self.mode or "").strip()
        if m in ("one", "singleton"):
            return "one"
        if m == "list":
            return "list"
        if m == "map":
            return "map"
        if self.index and ("+" in self.index or "," in self.index):
            return "list"
        return "map"

    @property
    def module(self) -> str:
        """表所属模块名，用于自定义类型归属。TbReward -> reward；item.TbEquip -> item"""
        base = self.full_name.split(".")[-1]
        if base.startswith("Tb") and len(base) > 2:
            base = base[2:]
        if "." in self.full_name:
            return self.full_name.split(".")[0]
        return base[:1].lower() + base[1:]


@dataclass
class BeanDef:
    full_name: str = ""
    parent: str = ""
    value_type: str = ""
    sep: str = ""
    alias: str = ""
    group: str = ""
    comment: str = ""
    fields: List[FieldDef] = field(default_factory=list)
    _row: Optional[int] = None
    _sheet: str = ""
    _layout: str = ""
    _file: str = ""              # 从哪个定义文件读出来的（表专用类型在独立文件里）
    scope: str = "common"        # common 或某张表的 full_name（仅界面用，不写盘）

    def to_dict(self) -> Dict[str, Any]:
        return {
            "full_name": self.full_name,
            "parent": self.parent,
            "value_type": self.value_type,
            "sep": self.sep,
            "alias": self.alias,
            "group": self.group,
            "comment": self.comment,
            "fields": [f.to_dict() for f in self.fields],
            "_row": self._row,
            "_sheet": self._sheet,
            "_layout": self._layout,
            "_file": self._file,
            "scope": self.scope,
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "BeanDef":
        b = BeanDef(
            full_name=d.get("full_name", "") or "",
            parent=d.get("parent", "") or "",
            value_type=d.get("value_type", "") or "",
            sep=d.get("sep", "") or "",
            alias=d.get("alias", "") or "",
            group=d.get("group", "") or "",
            comment=d.get("comment", "") or "",
            fields=[FieldDef.from_dict(x) for x in d.get("fields", [])],
            scope=d.get("scope", "common") or "common",
        )
        b._row = d.get("_row")
        b._sheet = d.get("_sheet", "") or ""
        b._layout = d.get("_layout", "") or ""
        b._file = d.get("_file", "") or ""
        return b

    @property
    def short_name(self) -> str:
        return self.full_name.split(".")[-1]


@dataclass
class EnumItem:
    name: str = ""
    alias: str = ""
    value: str = ""
    comment: str = ""


@dataclass
class EnumDef:
    full_name: str = ""
    flags: str = ""
    unique: str = ""
    comment: str = ""
    items: List[EnumItem] = field(default_factory=list)
    _file: str = ""
    scope: str = "common"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "full_name": self.full_name,
            "flags": self.flags,
            "unique": self.unique,
            "comment": self.comment,
            "items": [asdict(i) for i in self.items],
            "_file": self._file,
            "scope": self.scope,
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "EnumDef":
        return EnumDef(
            full_name=d.get("full_name", "") or "",
            flags=d.get("flags", "") or "",
            unique=d.get("unique", "") or "",
            comment=d.get("comment", "") or "",
            items=[EnumItem(**i) for i in d.get("items", [])],
            _file=d.get("_file", "") or "",
            scope=d.get("scope", "common") or "common",
        )

    @property
    def short_name(self) -> str:
        return self.full_name.split(".")[-1]


@dataclass
class ProjectModel:
    root: str = ""
    conf_path: str = ""
    data_dir: str = ""
    groups: List[str] = field(default_factory=list)
    targets: List[str] = field(default_factory=list)
    table_file: str = ""
    bean_file: str = ""
    enum_file: str = ""
    # 表定义可能分散在好几个文件里（__tables__.xlsx + Defines/*.xml，
    # 或 luban.conf 直接登记了整个 Defines 目录）。全部都要读，否则
    # 「声明在 xml 里的表」会被当成没声明 —— 点开数据文件只能只读预览。
    table_files: List[str] = field(default_factory=list)
    # 本次会话从定义文件里读出来的表名。保存时用来区分
    # 「用户真的把某文件的表删光了」和「表头没认出来导致一张表都没读进来」——
    # 后者如果照常写回，会把整份定义文件抹空。
    loaded_tables: List[str] = field(default_factory=list)
    # 除主定义文件外，luban.conf 里另外登记的 bean / enum 定义文件。
    # 配置表专用的自定义类型写进这些文件（例：Defines/TbEquip.xml）；
    # 老工程里可能还残留旧版生成的 TbEquip.beans.xlsx，也能照常读取。
    extra_bean_files: List[str] = field(default_factory=list)
    extra_enum_files: List[str] = field(default_factory=list)
    table_columns: Dict[str, str] = field(default_factory=dict)
    bean_columns: Dict[str, str] = field(default_factory=dict)
    enum_columns: Dict[str, str] = field(default_factory=dict)
    tables: List[TableDef] = field(default_factory=list)
    beans: List[BeanDef] = field(default_factory=list)
    enums: List[EnumDef] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    # ---- 是否像一个 Luban 工程 ----
    @property
    def initialized(self) -> bool:
        # 唯一标志：文件夹里有没有 luban.conf（刷新时重新探测，这个属性实时反映）
        return bool(self.conf_path and os.path.exists(self.conf_path))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "root": self.root,
            "conf_path": self.conf_path,
            "data_dir": self.data_dir,
            "groups": self.groups,
            "targets": self.targets,
            "table_file": self.table_file,
            "bean_file": self.bean_file,
            "enum_file": self.enum_file,
            "table_files": self.table_files,
            "loaded_tables": self.loaded_tables,
            "extra_bean_files": self.extra_bean_files,
            "extra_enum_files": self.extra_enum_files,
            "table_columns": self.table_columns,
            "bean_columns": self.bean_columns,
            "enum_columns": self.enum_columns,
            "tables": [t.to_dict() for t in self.tables],
            "beans": [b.to_dict() for b in self.beans],
            "enums": [e.to_dict() for e in self.enums],
            "warnings": self.warnings,
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "ProjectModel":
        m = ProjectModel(
            root=d.get("root", ""),
            conf_path=d.get("conf_path", ""),
            data_dir=d.get("data_dir", ""),
            groups=d.get("groups", []),
            targets=d.get("targets", []),
            table_file=d.get("table_file", ""),
            bean_file=d.get("bean_file", ""),
            enum_file=d.get("enum_file", ""),
            table_files=d.get("table_files", []),
            loaded_tables=d.get("loaded_tables", []),
            extra_bean_files=d.get("extra_bean_files", []),
            extra_enum_files=d.get("extra_enum_files", []),
            table_columns=d.get("table_columns", {}),
            bean_columns=d.get("bean_columns", {}),
            enum_columns=d.get("enum_columns", {}),
        )
        m.tables = [TableDef.from_dict(x) for x in d.get("tables", [])]
        m.beans = [BeanDef.from_dict(x) for x in d.get("beans", [])]
        m.enums = [EnumDef.from_dict(x) for x in d.get("enums", [])]
        m.warnings = d.get("warnings", [])
        return m

    # ---- 查询 ----
    def find_bean(self, name: str) -> Optional[BeanDef]:
        if not name:
            return None
        for b in self.beans:
            if b.full_name == name:
                return b
        for b in self.beans:
            if b.short_name == name:
                return b
        return None

    def find_enum(self, name: str) -> Optional[EnumDef]:
        if not name:
            return None
        for e in self.enums:
            if e.full_name == name:
                return e
        for e in self.enums:
            if e.short_name == name:
                return e
        return None

    def find_table(self, name: str) -> Optional[TableDef]:
        for t in self.tables:
            if t.full_name == name:
                return t
        return None

    def subclasses_of(self, bean_name: str) -> List[BeanDef]:
        """多态：返回直接/间接以 bean_name 为父类的子结构。"""
        out = []
        for b in self.beans:
            p = b.parent
            if not p:
                continue
            if p == bean_name or p.split(".")[-1] == bean_name.split(".")[-1]:
                out.append(b)
        return out

    def is_polymorphic(self, base: str) -> bool:
        b = self.find_bean(base)
        if not b:
            return False
        return bool(self.subclasses_of(b.full_name)) or bool(b.parent and not b.fields)
