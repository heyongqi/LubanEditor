"""调用 Luban.exe 并解析报错；以及保存前的静态自检。

命令不做硬编码：默认模板 + 用户可编辑，适配任何工程（dotnet tool / Luban.exe 都行）。
"""
from __future__ import annotations

import os
import re
import subprocess
from typing import Dict, List, Optional, Any, Tuple

from .model import ProjectModel, TableDef

DEFAULT_CMD = '{exe} --conf "{conf}" -t {target} {extra}'

# 脚本末尾的 `pause` 用这个环境变量跳过：双击运行时照常暂停看结果，
# 工具里跑没人按键，必须跳过，否则会一直卡到超时。
PAUSE_ENV = "LUBANKIT_NO_PAUSE"

ERR_KEYWORDS = ("error", "错误", "exception", "异常", "failed", "失败", "无效", "不合法",
                "not found", "找不到", "重复", "duplicate")

# 已知报错模式 -> 提示（与 knowledge.py 的条目 id 对应）
PATTERN_HINTS: List[Tuple[str, str, str]] = [
    (r"not\s+start\s+with\s+##|不以\s*##\s*开头|##\s*开头", "a1_marker",
     "该 sheet 的 A1 不是以 ## 开头，整张表会被 Luban 忽略"),
    (r"duplicate|重复定义|already\s+define", "dup_define",
     "同一类型被定义了两次：检查 read_schema_from_file=true 的表是否又在 __beans__ 里定义了同名 bean"),
    (r"unknown\s+type|未知类型|type\s+.{0,20}not\s+found", "unknown_type",
     "类型名拼错，或该 enum/bean 还没在定义文件里登记"),
    (r"index|索引|主键", "index_mode",
     "index 与 mode 不匹配：多个主键要用 list，联合主键写 a+b、独立索引写 a,b"),
    (r"sep|分隔", "sep_from_schema",
     "分隔符来自类型/schema 定义，不要凭单元格里的中文逗号猜测"),
    (r"reference|ref|引用", "ref_check",
     "引用校验失败：被引用的 id 在目标表中不存在"),
    (r"empty|为空|null", "null_vs_zero",
     "可空要写 T?；流式(sep)单元格里不能靠留空表示默认值"),
]


def build_command(settings: Dict[str, str], model: ProjectModel) -> str:
    tpl = (settings.get("cmd_template") or "").strip() or DEFAULT_CMD
    exe = (settings.get("luban_exe") or "").strip().strip('"')
    # dotnet 版： Luban.dll 前面自动补 dotnet（模板里不再额外加引号，
    # 由这里统一保证「有空格的部分才加引号」）
    if exe.lower().endswith(".dll") and not exe.lower().startswith("dotnet"):
        exe = 'dotnet "%s"' % exe
    elif " " in exe and not exe.startswith('"'):
        exe = '"%s"' % exe
    conf = settings.get("conf_path") or model.conf_path or ""
    target = settings.get("target", "") or (model.targets[0] if model.targets else "")
    extra = settings.get("extra_args", "")
    return tpl.format(exe=exe, conf=conf, target=target, extra=extra,
                      data_dir=model.data_dir, root=model.root).strip()


def _exec(cmd: str, cwd: str, timeout: int, extra_env: Optional[Dict[str, str]] = None
          ) -> Tuple[Any, str]:
    """跑一条命令，返回 (proc 或 None, 已解码输出)。异常用 code=-2/-3 表示。"""
    # 冻结成 GUI 程序（--noconsole）后，调 Luban.exe 若不隐藏窗口会闪一个黑框
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0  # SW_HIDE
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    # Luban 报错多为 UTF-8 / GBK 混排，强制子进程用 UTF-8 输出减少乱码
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    if extra_env:
        env.update(extra_env)
    try:
        # stdin 接到 DEVNULL：脚本里万一还有 pause / read，读不到输入就立刻返回，
        # 不会把人卡在「请按任意键继续」上直到超时
        proc = subprocess.run(cmd, cwd=cwd or None, shell=isinstance(cmd, str),
                              capture_output=True, timeout=timeout,
                              stdin=subprocess.DEVNULL,
                              startupinfo=startupinfo, creationflags=creationflags,
                              env=env)
    except subprocess.TimeoutExpired:
        return -2, "执行超时"
    except Exception as e:
        return -3, f"启动失败：{e}"
    out = ""
    for b in (proc.stdout, proc.stderr):
        if b:
            try:
                out += b.decode("utf-8") + "\n"
            except UnicodeDecodeError:
                out += b.decode("gbk", errors="replace") + "\n"
    return proc.returncode, out.strip()


def run_luban(settings: Dict[str, str], model: ProjectModel,
              timeout: int = 300) -> Dict[str, Any]:
    cmd = build_command(settings, model)
    cwd = settings.get("work_dir") or (os.path.dirname(model.conf_path) if model.conf_path else model.root)
    if not settings.get("luban_exe"):
        return {"ok": False, "cmd": cmd, "code": -1,
                "output": "尚未设置 Luban.exe 路径（菜单：设置 → Luban 可执行文件）",
                "issues": []}
    code, out = _exec(cmd, cwd, timeout)
    issues = parse_issues(out, model)
    return {"ok": code == 0, "cmd": cmd, "code": code, "output": out,
            "issues": issues}


def script_command(path: str) -> List[str]:
    """执行脚本的**系统命令**：Windows 用 `cmd /c`，其它用 `/bin/sh`。

    不用 shell=True 的字符串，而是直接给出「系统命令 + 参数」的形式 —— 相当于
    在命令行里敲 `cmd /c gen.bat`，路径里有空格也不用额外加引号。
    """
    if os.name == "nt":
        return [os.environ.get("COMSPEC") or "cmd.exe", "/c", path]
    return ["/bin/sh", path]


def run_script(path: str, work_dir: str = "", timeout: int = 600
               ) -> Dict[str, Any]:
    """直接调用系统命令执行一个 gen.bat / gen.sh，返回与 run_luban 一样的结构。

    Windows 上就是 `cmd /c 脚本路径`，等价于在资源管理器里双击它；
    末尾的 `pause` 用环境变量 `LUBANKIT_NO_PAUSE` 跳过（双击运行时照常暂停）。
    """
    if not path or not os.path.exists(path):
        return {"ok": False, "cmd": path or "", "code": -1,
                "output": f"脚本不存在：{path}", "issues": []}
    if not path.lower().endswith(".bat"):
        try:
            os.chmod(path, os.stat(path).st_mode | 0o111)
        except Exception:
            pass
    cmd = script_command(path)
    code, out = _exec(cmd, work_dir, timeout, extra_env={PAUSE_ENV: "1"})
    return {"ok": code == 0, "cmd": " ".join(cmd), "code": code,
            "output": out, "issues": []}


def parse_issues(output: str, model: ProjectModel) -> List[Dict[str, str]]:
    issues: List[Dict[str, str]] = []
    names = []
    for t in model.tables:
        names += [t.full_name, t.value_type, t.input]
    names = [n for n in names if n]
    for line in output.splitlines():
        low = line.lower()
        if not any(k in low for k in ERR_KEYWORDS):
            continue
        table = ""
        for n in sorted(set(names), key=len, reverse=True):
            if n and n in line:
                table = n
                break
        hint, kid = "", ""
        for pat, k, msg in PATTERN_HINTS:
            if re.search(pat, low):
                hint, kid = msg, k
                break
        issues.append({"level": "error", "msg": line.strip(),
                       "table": table, "hint": hint, "knowledge": kid})
        if len(issues) >= 200:
            break
    return issues


# ------------------------------------------------------------------ 保存前自检
def validate_project(model: ProjectModel,
                     fields_of: Optional[Dict[str, List[Any]]] = None) -> List[Dict[str, str]]:
    """静态自检：不跑 Luban 也能拦住大部分低级错误。"""
    res: List[Dict[str, str]] = []
    fields_of = fields_of or {}
    seen = set()
    bean_names = {b.full_name for b in model.beans}
    for t in model.tables:
        tag = t.full_name or "(未命名表)"
        if not t.full_name:
            res.append({"level": "error", "table": tag, "msg": "存在没有 full_name 的表"})
            continue
        if t.full_name in seen:
            res.append({"level": "error", "table": tag, "msg": "表名重复"})
        seen.add(t.full_name)
        if not t.value_type:
            res.append({"level": "error", "table": tag, "msg": "value_type 为空"})
        if not t.input:
            res.append({"level": "error", "table": tag, "msg": "input（数据文件）为空"})
        elif model.data_dir:
            p = os.path.join(model.data_dir, t.input.split("@")[-1])
            if not os.path.exists(p):
                res.append({"level": "warn", "table": tag,
                            "msg": f"数据文件不存在：{p}（还没生成模板？）"})
        rsf = (t.read_schema_from_file or "").strip().lower()
        if rsf in ("true", "1", "yes") and t.value_type in bean_names:
            res.append({"level": "error", "table": tag,
                        "msg": f"read_schema_from_file=true 时，{t.value_type} 不能再在 "
                               f"__beans__.xlsx 里定义，否则报定义重复"})
        mode = t.real_mode
        if mode == "map" and not t.index:
            res.append({"level": "warn", "table": tag,
                        "msg": "mode=map 但未填 index，Luban 会取 value_type 的第一个字段当主键"})
        if t.index and mode == "map" and ("+" in t.index or "," in t.index):
            res.append({"level": "error", "table": tag,
                        "msg": "多个主键时 mode 必须是 list；联合主键写 a+b、独立索引写 a,b"})
        if t.group:
            for g in [x for x in re.split(r"[,，]", t.group) if x]:
                if model.groups and g not in model.groups:
                    res.append({"level": "warn", "table": tag,
                                "msg": f"分组 {g} 不在 luban.conf 的 groups 中"})
        fs = fields_of.get(t.full_name, [])
        if fs:
            fnames = [f.name for f in fs]
            dup = {x for x in fnames if fnames.count(x) > 1}
            for d in dup:
                res.append({"level": "error", "table": tag, "msg": f"字段名重复：{d}"})
            for f in fs:
                if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", f.name or ""):
                    res.append({"level": "error", "table": tag,
                                "msg": f"字段名 {f.name} 不符合命名规范（建议 xx_yy_zz）"})
            if t.index and mode in ("map", "list"):
                keys = [k for k in re.split(r"[+,]", t.index) if k]
                for k in keys:
                    if k not in fnames:
                        res.append({"level": "error", "table": tag,
                                    "msg": f"index 字段 {k} 不在字段列表中"})
    return res
