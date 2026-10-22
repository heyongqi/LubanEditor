"""新建 / 初始化 Luban 工程：创建目录结构与定义文件。

产物：
    luban.conf
    gen.bat / gen.sh
    Schemas/        Schema（表 / 结构 / 枚举的定义文件）目录
    Data/           数据表目录（luban.conf 的 dataDir）
    Defines/        XML 定义目录（勾选 xml 时生成 *.xml）
    Schemas/__tables__.xlsx / __beans__.xlsx / __enums__.xlsx  （勾选 xlsx）
    Defines/tables.xml / beans.xml / enums.xml                  （勾选 xml）

为什么 Schema 要单独一个 Schemas 目录：定义文件和数据文件混在 Data 里，
一来找定义要在几十张数据表里翻，二来 Luban 会把 Data 下的 xlsx 当数据表扫。
分开之后 Schema 一眼可见，schemaFiles 里写的也是 Schemas/xxx.xlsx。
"""
from __future__ import annotations

import json
import os
import shutil
import stat
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

DATA_DIRNAME = "Data"
SCHEMAS_DIRNAME = "Schemas"
DEFINES_DIRNAME = "Defines"
TOOLS_DIRNAME = "Tools"
OUTPUT_DIRNAME = "Output"
OUTPUT_DATA_DIRNAME = "Data"
OUTPUT_CODE_DIRNAME = "Code"

CONF_TEMPLATE = {
    "groups": [
        {"names": ["c"], "default": True},
        {"names": ["s"], "default": True},
        {"names": ["e"], "default": True},
    ],
    "schemaFiles": [],
    "dataDir": DATA_DIRNAME,
    "outputDataDir": f"{OUTPUT_DIRNAME}/{OUTPUT_DATA_DIRNAME}",
    "outputCodeDir": f"{OUTPUT_DIRNAME}/{OUTPUT_CODE_DIRNAME}",
    "targets": [
        {"name": "client", "manager": "Tables", "groups": ["c"], "topModule": "cfg"},
        {"name": "server", "manager": "Tables", "groups": ["s"], "topModule": "cfg"},
        {"name": "all", "manager": "Tables", "groups": ["c", "s", "e"], "topModule": "cfg"},
    ],
}

GEN_BAT = r'''@echo off
rem ---------------------------------------------------------------
rem  Luban 导表脚本（dotnet 形式）
rem  默认用 Tools/Luban/Luban.dll；若你的 Luban 在别处，改下面 LUBAN
rem ---------------------------------------------------------------
set "LUBAN=%~dp0Tools\Luban\Luban.dll"

dotnet "%LUBAN%" --conf "%~dp0luban.conf" -t all
if errorlevel 1 (
    echo.
    echo 导表失败，请检查上面的报错信息。
)
pause
'''

GEN_SH = r'''#!/bin/bash
# Luban 导表脚本（dotnet 形式）
# 默认用 Tools/Luban/Luban.dll；若你的 Luban 在别处，改下面 LUBAN
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
LUBAN="$DIR/Tools/Luban/Luban.dll"

dotnet "$LUBAN" --conf "$DIR/luban.conf" -t all
echo "导表完成"
'''

_HEAD_FILL = PatternFill("solid", fgColor="D9E7F7")
_MARK_FILL = PatternFill("solid", fgColor="FBE3D6")


def _style(wb_ws, row, ncol, mark=False):
    for c in range(1, ncol + 1):
        cell = wb_ws.cell(row=row, column=c)
        if isinstance(cell, MergedCell):
            continue    # 合并区域的非左上角单元格只读，跳过
        cell.font = Font(bold=True)
        cell.fill = _MARK_FILL if (mark and c == 1) else _HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_tables_xlsx(path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "tables"
    ws.append(["##var", "full_name", "value_type", "read_schema_from_file", "input",
               "index", "mode", "group", "comment", "tags", "output"])
    ws.append(["##", "全名(包含模块和名字)", "记录类名", "从excel读取定义", "文件列表",
               "表id字段", "模式", "分组", "注释", None, "输出文件名"])
    ws.append(["##", None, None,
               "false时取已有定义，true为从excel标题头和属性栏读取定义",
               "可以多个，以逗号','分隔",
               "为空的话自动取value_type中第一个字段,多主键联合索引为key1+key2,"
               "多主键独立索引为\"key1,key2\"",
               "取值one|map|list，为空自动为map",
               "取值c|s|e，可以有多个，以逗号','分隔。空则表示属于所有分组",
               None, None, "默认为 <module>_<name>.<suffix>"])
    for r in (1, 2, 3):
        _style(ws, r, 11, mark=True)
    for i, w in enumerate([10, 20, 18, 24, 22, 20, 12, 12, 20, 12, 22], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)


def _write_beans_xlsx(path: str):
    """空的结构定义模板（最新版列结构）。

    虽然不推荐在 xlsx 里声明 bean（统一写 xml），但保留这个文件并登记进
    luban.conf 保持标准工程结构。字段列用 `*fields` 前缀（Luban 的重复列
    约定），子列为 name / alias / type / group / comment / tags / variants，
    不再生成重复的 `fields` 列。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "beans"
    fixed = ["##var", "full_name", "parent", "valueType", "sep", "alias",
             "comment", "group", "tags"]
    sub = ["name", "alias", "type", "group", "comment", "tags", "variants"]
    n_fixed = len(fixed)
    n_sub = len(sub)
    # 行 1：固定列 + *fields（父列名）+ 子列占位
    ws.append(fixed + ["*fields"] + [None] * (n_sub - 1))
    # 行 2：##var + 固定列占位 + 子列名
    ws.append(["##var"] + [None] * (n_fixed - 1) + sub)
    # 行 3：## 说明
    ws.append(["##", "全名(包含模块和名字)"] + [None] * (n_fixed - 2) +
              ["字段名", "字段别名", "类型", "分组", "注释", None, "字段变体"])
    # 合并 *fields 父列（第 n_fixed+1 到 n_fixed+n_sub 列）
    ws.merge_cells(start_row=1, start_column=n_fixed + 1,
                   end_row=1, end_column=n_fixed + n_sub)
    for r in (1, 2, 3):
        _style(ws, r, n_fixed + n_sub, mark=True)
    for i, w in enumerate([10, 20, 14, 12, 8, 12, 18, 12, 12, 10, 10, 12, 12, 14, 12, 12],
                          start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)


def _write_enums_xlsx(path: str):
    """空的枚举定义模板（最新版列结构）。

    同 bean：不推荐在 xlsx 里声明，但保留文件结构。枚举项用 `*items` 前缀
    （Luban 的重复列约定），子列为 name / alias / value / comment / tags。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "enums"
    fixed = ["##var", "full_name", "flags", "unique", "group", "comment", "tags"]
    sub = ["name", "alias", "value", "comment", "tags"]
    n_fixed = len(fixed)
    n_sub = len(sub)
    # 行 1：固定列 + *items（父列名）+ 子列占位
    ws.append(fixed + ["*items"] + [None] * (n_sub - 1))
    # 行 2：##var + 固定列占位 + 子列名
    ws.append(["##var"] + [None] * (n_fixed - 1) + sub)
    # 行 3：## 说明
    ws.append(["##", "全名(包含模块和名字)",
               "是否为位标记枚举（即每个枚举项为位标记数据，例如System.IO.FileMode,"
               "填数据时可以为READ|WRITE这样）",
               "枚举项是否唯一", None, None, None,
               "枚举名", "别名", "值", "注释", None])
    # 合并 *items 父列（第 n_fixed+1 到 n_fixed+n_sub 列）
    ws.merge_cells(start_row=1, start_column=n_fixed + 1,
                   end_row=1, end_column=n_fixed + n_sub)
    for r in (1, 2, 3):
        _style(ws, r, n_fixed + n_sub, mark=True)
    for i, w in enumerate([10, 20, 10, 10, 12, 18, 12, 10, 10, 10, 14, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)


_XML_TABLES = '''<?xml version="1.0" encoding="UTF-8"?>
<!-- 表定义：每张表一行 -->
<module name="">
  <!-- 示例
  <table name="TbReward" value="Reward" input="reward.xlsx" index="id" mode="map" group="c,s" comment="奖励表"/>
  -->
</module>
'''


def _dump_conf(conf: dict) -> str:
    """luban.conf 密集型：顶层字段一行一个，list of dict 的每个 dict 压成一行。"""
    def enc(v):
        return json.dumps(v, ensure_ascii=False)

    def inline_dict(d):
        return "{" + ", ".join(f'"{k}": {enc(v)}' for k, v in d.items()) + "}"

    lines = ["{"]
    keys = list(conf.keys())
    for i, k in enumerate(keys):
        v = conf[k]
        comma = "," if i < len(keys) - 1 else ""
        if isinstance(v, list) and v and all(isinstance(x, dict) for x in v):
            # 数组里每个元素单独一行
            lines.append(f'  "{k}": [')
            for j, item in enumerate(v):
                item_comma = "," if j < len(v) - 1 else ""
                lines.append(f"    {inline_dict(item)}{item_comma}")
            lines.append(f"  ]{comma}")
        else:
            lines.append(f'  "{k}": {enc(v)}{comma}')
    lines.append("}")
    return "\n".join(lines) + "\n"


def _copy_tools(tools_source: str, tools_dir: str) -> str:
    """把用户指定的本地 Tools 目录内容复制到工程 Tools 目录。返回说明文字。"""
    if not tools_source:
        return "（未指定，请手动放入 Luban 工具）"
    if not os.path.isdir(tools_source):
        return "（指定的 Tools 位置不存在，请检查）"
    try:
        shutil.copytree(tools_source, tools_dir, dirs_exist_ok=True)
        return "已从本地复制 Luban 工具"
    except Exception:
        return "（复制失败，请手动复制）"


def create_project(root: str, schema_types: List[str], tools_source: str = "") -> tuple:
    """创建工程。schema_types 可含 'xlsx' / 'xml'；tools_source 为本地 Luban 工具目录。"""
    root = os.path.abspath(root)
    if not schema_types:
        return False, "请至少选择一种 Schema 文件类型（xlsx / xml）"
    try:
        os.makedirs(root, exist_ok=True)
        data_dir = os.path.join(root, DATA_DIRNAME)
        schemas_dir = os.path.join(root, SCHEMAS_DIRNAME)
        defines_dir = os.path.join(root, DEFINES_DIRNAME)
        tools_dir = os.path.join(root, TOOLS_DIRNAME)
        output_dir = os.path.join(root, OUTPUT_DIRNAME)
        os.makedirs(data_dir, exist_ok=True)
        os.makedirs(schemas_dir, exist_ok=True)
        os.makedirs(defines_dir, exist_ok=True)
        os.makedirs(tools_dir, exist_ok=True)
        os.makedirs(os.path.join(output_dir, OUTPUT_DATA_DIRNAME), exist_ok=True)
        os.makedirs(os.path.join(output_dir, OUTPUT_CODE_DIRNAME), exist_ok=True)

        conf = json.loads(json.dumps(CONF_TEMPLATE))
        schemas, files = [], []
        # xlsx：表声明 + 空的 bean / enum 模板（不推荐在这里声明类型，但保留标准结构）
        if "xlsx" in schema_types:
            write_tables_xlsx(os.path.join(schemas_dir, "__tables__.xlsx"))
            _write_beans_xlsx(os.path.join(schemas_dir, "__beans__.xlsx"))
            _write_enums_xlsx(os.path.join(schemas_dir, "__enums__.xlsx"))
            conf["schemaFiles"] += [
                {"fileName": f"{SCHEMAS_DIRNAME}/__tables__.xlsx", "type": "table"},
                {"fileName": f"{SCHEMAS_DIRNAME}/__beans__.xlsx", "type": "bean"},
                {"fileName": f"{SCHEMAS_DIRNAME}/__enums__.xlsx", "type": "enum"},
            ]
            schemas += ["__tables__.xlsx", "__beans__.xlsx", "__enums__.xlsx"]
        if "xml" in schema_types:
            with open(os.path.join(defines_dir, "tables.xml"), "w", encoding="utf-8") as f:
                f.write(_XML_TABLES)
            if not any(s.get("fileName") == DEFINES_DIRNAME for s in conf["schemaFiles"]):
                conf["schemaFiles"].append({"fileName": DEFINES_DIRNAME, "type": ""})
            files += ["tables.xml"]

        # luban.conf 密集型
        with open(os.path.join(root, "luban.conf"), "w", encoding="utf-8") as f:
            f.write(_dump_conf(conf))

        # gen.bat / gen.sh（dotnet 形式）
        with open(os.path.join(root, "gen.bat"), "w", encoding="gbk", newline="\r\n") as f:
            f.write(GEN_BAT)
        p = os.path.join(root, "gen.sh")
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(GEN_SH)
        try:
            os.chmod(p, os.stat(p).st_mode | stat.S_IEXEC)
        except Exception:
            pass

        # 把用户指定的本地 Tools 内容复制到工程 Tools 目录
        tools_note = _copy_tools(tools_source, tools_dir)
    except Exception as e:
        return False, f"创建失败：{e}"

    return True, (f"已在 {root} 创建 Luban 工程\n"
                  f"  luban.conf / gen.bat / gen.sh\n"
                  f"  {DATA_DIRNAME}/（数据表）\n"
                  f"  {SCHEMAS_DIRNAME}/  " + "、".join(schemas[:3]) +
                  (f"\n  {DEFINES_DIRNAME}/  " + "、".join(files[:3]) if files else "") +
                  f"\n  {TOOLS_DIRNAME}/  {tools_note}" +
                  f"\n  {OUTPUT_DIRNAME}/{OUTPUT_DATA_DIRNAME}、{OUTPUT_DIRNAME}/{OUTPUT_CODE_DIRNAME}"
                  f"（默认输出目录）")
