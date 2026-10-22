# -*- coding: utf-8 -*-
"""发布工程：把 gen.bat / gen.sh 的参数摊开成一张可编辑的表。

为什么单开一个模块：脚本参数这件事有两头 —— 一头是**读**（工程里已经有
gen.bat 了，得把它在用的参数解析出来），一头是**写**（改完参数要能还原成一个
能跑的脚本）。把参数清单、解析、生成放一处，界面只负责画控件。

参数分两组：

- **基础参数**：照 Luban 官方模板脚本（gen_template.bat）里的那几个 ——
  可执行文件、`--conf`、`-t`、`-c`、`-d`、两个输出目录。日常导表只碰这些。
- **高级参数**：官方 CLI 全参数里的其余开关，加常用的 `-x` 扩展项。默认折叠，
  点「更多」才展开。

每个参数都带一句「干什么用的」备注，鼠标悬停还能看完整说明。
"""
from __future__ import annotations

import os
import re
import shlex
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .model import ProjectModel
from .runner import PAUSE_ENV       # 脚本末尾 pause 的跳过开关，两边必须用同一个名字

# ---------------------------------------------------------------- 取值参考
CODE_TARGETS: List[str] = [
    "cs-simple-json", "cs-bin", "cs-dotnet-json", "cs-newtonsoft-json",
    "java-json", "java-bin", "go-json", "go-bin", "lua",
    "cpp-json", "cpp-bin", "python-json", "typescript-json",
    "rust-json", "php-json", "schema-json",
]

DATA_TARGETS: List[str] = [
    "json", "bin", "lua", "bson", "msgpack", "xml", "yaml",
    "text-list", "erlang",
]

# 执行方式：dll = dotnet 版（前面要补 dotnet），exe = Windows 独立版
EXE_MODE_DLL = "dll"
EXE_MODE_EXE = "exe"
EXE_MODES = (EXE_MODE_DLL, EXE_MODE_EXE)

# 在工程目录里找 Luban 时认这些文件名（大小写都认），前面的优先
_EXE_NAMES: Dict[str, Tuple[str, ...]] = {
    EXE_MODE_DLL: ("Luban.dll", "luban.dll", "Luban.exe", "luban.exe"),
    EXE_MODE_EXE: ("Luban.exe", "luban.exe", "Luban.dll", "luban.dll"),
}
_CONF_NAMES: Tuple[str, ...] = ("luban.conf",)


@dataclass
class ParamDef:
    """一个可编辑的发布参数。"""

    key: str                    # 内部键
    label: str                  # 界面显示名
    meaning: str                # 备注：这个参数是干什么的
    section: str = "basic"      # basic / advanced
    group: str = ""             # 高级参数里的小标题
    flag: str = ""              # 命令行开关（空 = 程序本身 / 只做 -x 的 key）
    kind: str = "text"          # text / combo / file / dir / flag
    choices: List[str] = field(default_factory=list)
    placeholder: str = ""
    default: str = ""
    xarg: bool = False          # 走 -x key=value 形式
    multiple: bool = False      # 可重复写多次（逗号分隔）


PARAMS: List[ParamDef] = [
    # ---------------- 基础参数（官方模板脚本里的那几个） ----------------
    ParamDef("exeMode", "执行方式",
             "Luban 有两种发行形式：dll＝dotnet 版（命令前面自动补 dotnet，跨平台），"
             "exe＝Windows 独立版（直接运行）。默认 dll。"
             "选好后下面「执行文件」就按这个类型去找 / 去挑。",
             kind="mode", default=EXE_MODE_DLL),
    ParamDef("exe", "执行文件",
             "Luban 本体：Luban.dll（执行方式选 dll）或 Luban.exe（选 exe）。"
             "它就是命令行最前面那个程序，后面跟的都是参数；"
             "在当前工程里时显示相对路径 / 文件名，在别处则显示全路径。"
             "打开面板时会先去工程目录里找，找到了自动填上。",
             kind="file", placeholder="Luban.dll / Luban.exe"),
    ParamDef("conf", "配置文件 --conf",
             "要加载的 luban.conf。数据目录、定义文件、导出目标、分组全写在这里面，"
             "是导表的总入口，几乎每个脚本都要有。",
             flag="--conf", kind="file", placeholder="luban.conf"),
    ParamDef("target", "导出目标 -t",
             "这一批按哪些分组导出：client 只出 c 组、server 只出 s 组、all 全出。"
             "取值来自 luban.conf 里 targets[].name。",
             flag="-t", kind="combo"),
    ParamDef("codeTarget", "代码目标 -c",
             "生成哪种语言 / 风格的代码：cs-bin（C#读二进制）、cs-simple-json（C#读JSON）、"
             "java-json、go-bin、lua… 可填多个（逗号分隔＝重复写多次 -c）。不填就只出数据。",
             flag="-c", kind="combo", choices=CODE_TARGETS, multiple=True,
             default="cs-simple-json"),
    ParamDef("dataTarget", "数据格式 -d",
             "导出的数据文件格式：json / bin / lua / bson / msgpack… 可填多个（逗号分隔）。"
             "不填就只出代码。",
             flag="-d", kind="combo", choices=DATA_TARGETS, multiple=True,
             default="json"),
    ParamDef("outputCodeDir", "代码输出目录",
             "生成的代码写到哪个目录（-x outputCodeDir=）。留空就用 luban.conf 里的默认值。",
             flag="outputCodeDir", kind="dir", xarg=True),
    ParamDef("outputDataDir", "数据输出目录",
             "导出的数据文件写到哪个目录（-x outputDataDir=）。留空就用 luban.conf 里的默认值。",
             flag="outputDataDir", kind="dir", xarg=True),

    # ---------------- 高级：标准开关 ----------------
    ParamDef("schemaCollector", "schema 收集器 -s",
             "用哪个收集器去解析定义文件（__tables__.xlsx 等）。一般是 default，"
             "只有自己写了收集器插件才需要改。",
             section="advanced", group="标准开关", flag="-s", placeholder="default"),
    ParamDef("pipeline", "生成管线 -p",
             "走哪条生成管线。一般是 default，自定义 Pipeline 插件时才需要改。",
             section="advanced", group="标准开关", flag="-p", placeholder="default"),
    ParamDef("forceLoadTableDatas", "强制加载数据 -f",
             "即使没给 -d 也把数据读进来做校验。想「只校验不出文件」就勾它，"
             "再配 outputSaver=null。",
             section="advanced", group="标准开关", flag="-f", kind="flag"),
    ParamDef("validationFailAsError", "校验失败即失败",
             "--validationFailAsError（也写作 --strict）：数据校验不过就让整次生成失败，"
             "发布前建议勾上。",
             section="advanced", group="标准开关",
             flag="--validationFailAsError", kind="flag"),
    ParamDef("verbose", "详细日志 -v",
             "输出更详细的调试信息。导表报错看不懂时打开。",
             section="advanced", group="标准开关", flag="-v", kind="flag"),

    # ---------------- 高级：过滤 ----------------
    ParamDef("includeTag", "只导出这些 tag -i",
             "只输出带指定 tag 的记录（多个用逗号分隔）。与「排除这些 tag」互斥。",
             section="advanced", group="记录 / 表过滤", flag="-i"),
    ParamDef("excludeTag", "排除这些 tag -e",
             "跳过带指定 tag 的记录（多个用逗号分隔）。与「只导出这些 tag」互斥。",
             section="advanced", group="记录 / 表过滤", flag="-e"),
    ParamDef("outputTable", "只导出这些表 -o",
             "只导出指定的表（表名，多个用逗号分隔）。调单张表时很快。",
             section="advanced", group="记录 / 表过滤", flag="-o"),
    ParamDef("variant", "字段变体 --variant",
             "格式 {结构}.{字段}={变体名}，可填多个（逗号分隔）。一般不用。",
             section="advanced", group="记录 / 表过滤", flag="--variant", multiple=True),
    ParamDef("timeZone", "时区 --timeZone",
             "影响 datetime 类型的解析与输出。留空用本机时区。",
             section="advanced", group="记录 / 表过滤", flag="--timeZone"),

    # ---------------- 高级：其它常用 -x ----------------
    ParamDef("outputSaver", "文件保存器 outputSaver",
             "local＝写文件（默认）；null＝一个文件都不写，只做校验（常配合「强制加载数据」）。",
             section="advanced", group="其它常用 -x", flag="outputSaver",
             kind="combo", choices=["local", "null"], xarg=True),
    ParamDef("cleanUpOutputDir", "清理旧生成物",
             "写出前先清掉输出目录里的旧文件（-x outputSaver.cleanUpOutputDir=）。默认 true —— "
             "所以输出目录**别指向有手写代码的文件夹**。",
             section="advanced", group="其它常用 -x", flag="outputSaver.cleanUpOutputDir",
             kind="combo", choices=["true", "1", "false", "0"], xarg=True),
    ParamDef("fileExt", "数据文件后缀",
             "改导出数据文件的后缀（-x {数据格式}.fileExt=），例如 bytes。留空用 target 默认后缀。",
             section="advanced", group="其它常用 -x", flag="fileExt", xarg=True),
    ParamDef("jsonCompact", "紧凑 JSON",
             "1＝输出无缩进的紧凑 JSON（-x json.compact=1），默认 0 带缩进。",
             section="advanced", group="其它常用 -x", flag="json.compact",
             kind="combo", choices=["0", "1"], xarg=True),
    ParamDef("lineEnding", "行尾符",
             "文本类输出的行尾（-x data.lineEnding=）：CR / LF / CRLF，二进制格式无效。",
             section="advanced", group="其它常用 -x", flag="data.lineEnding",
             kind="combo", choices=["CR", "LF", "CRLF"], xarg=True),
    ParamDef("l10nProvider", "本地化 Provider",
             "填 default 才启用本地化：之后才会做 text 校验，并可做静态替换。",
             section="advanced", group="其它常用 -x", flag="l10n.provider",
             kind="combo", choices=["default"], xarg=True),
    ParamDef("l10nTextFile", "本地化文本文件",
             "本地化文本数据文件（-x l10n.textFile.path=），配合上面 Provider 一起填。",
             section="advanced", group="其它常用 -x", flag="l10n.textFile.path",
             kind="file", xarg=True),
    ParamDef("logConfig", "日志配置 -l",
             "nlog 配置文件路径，默认 nlog.xml。一般不用改。",
             section="advanced", group="其它常用 -x", flag="-l", kind="file"),
    ParamDef("watchDir", "监视目录 -w",
             "监视目录变化并自动重新生成（会常驻不退出的进程，别在 CI / 发布时用）。",
             section="advanced", group="其它常用 -x", flag="-w", kind="dir"),
    ParamDef("extra", "其它 -x 参数",
             "原样追加到命令末尾，写成 key=value，多个用空格分隔。"
             "例如 pathValidator.rootDir=Assets。",
             section="advanced", group="其它常用 -x"),
]

PARAM_BY_KEY: Dict[str, ParamDef] = {p.key: p for p in PARAMS}
BASIC_KEYS = [p.key for p in PARAMS if p.section == "basic"]
ADV_KEYS = [p.key for p in PARAMS if p.section == "advanced"]


# ---------------------------------------------------------------- 命令行解析
_OPTS: Dict[str, str] = {
    "--conf": "conf",
    "-t": "target", "--target": "target",
    "-c": "codeTarget", "--codetarget": "codeTarget",
    "-d": "dataTarget", "--datatarget": "dataTarget",
    "-s": "schemaCollector", "--schemacollector": "schemaCollector",
    "-p": "pipeline", "--pipeline": "pipeline",
    "-o": "outputTable", "--outputtable": "outputTable",
    "-i": "includeTag", "--includetag": "includeTag",
    "-e": "excludeTag", "--excludetag": "excludeTag",
    "--variant": "variant", "--timezone": "timeZone",
    "--customtemplatedir": "customTemplateDir",
    "-l": "logConfig", "--logconfig": "logConfig",
    "-w": "watchDir", "--watchdir": "watchDir",
}
_FLAGS: Dict[str, str] = {
    "-f": "forceLoadTableDatas", "--forceloadtabledatas": "forceLoadTableDatas",
    "--validationfailaserror": "validationFailAsError",
    "--strict": "validationFailAsError",
    "-v": "verbose", "--verbose": "verbose",
}
_XARGS: Dict[str, str] = {
    "outputCodeDir": "outputCodeDir",
    "outputDataDir": "outputDataDir",
    "outputSaver": "outputSaver",
    "outputSaver.cleanUpOutputDir": "cleanUpOutputDir",
    "dataExporter": "dataExporter",
    "json.compact": "jsonCompact",
    "data.lineEnding": "lineEnding",
    "l10n.provider": "l10nProvider",
    "l10n.textFile.path": "l10nTextFile",
}
_FILE_EXT_RE = re.compile(r"^(?:[A-Za-z0-9_]+\.)?fileExt$")


def _unquote(t: str) -> str:
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'":
        return t[1:-1]
    return t


def _tokens(text: str) -> List[str]:
    try:
        toks = shlex.split(text or "", posix=(os.name != "nt"))
    except ValueError:
        toks = (text or "").split()
    return [_unquote(t) for t in toks]


def _expand(text: str, env: Dict[str, str]) -> str:
    """把 `%VAR%` / `$VAR` / `${VAR}` 展开（脚本里常见写法）。"""

    def sub(m):
        return env.get(m.group(1) or m.group(2), m.group(0))

    for _ in range(3):
        new = re.sub(r"%([A-Za-z_][A-Za-z0-9_]*)%|\$\{([A-Za-z_][A-Za-z0-9_]*)\}",
                     lambda m: env.get(m.group(1) or m.group(2), m.group(0)), text)
        if new == text:
            break
        text = new
    for k, v in env.items():
        text = text.replace("$" + k, v)
    return text


_SHELL_NOISE = ("rem ", "#", "@echo", "echo ", "set ", "setlocal", "endlocal",
                "exit", "pause", "cd ", "chcp", "title", "cls", "goto", "read ",
                "if ", "then", "fi", "done", "for ", "dir=", "local ", "shift")


def _is_shell_noise(s: str) -> bool:
    low = s.lower()
    return any(low.startswith(w) for w in _SHELL_NOISE)


def _known_flag(toks: List[str]) -> bool:
    for t in toks:
        low = t.lower()
        if t in _FLAGS or low in _FLAGS or t in _OPTS or low in _OPTS \
                or t in ("-x", "--xargs"):
            return True
    return False


def parse_command_line(line: str, env: Optional[Dict[str, str]] = None
                       ) -> Tuple[Dict[str, str], List[str], List[str]]:
    """解析一行 Luban 命令。

    返回 (已识别参数, 未识别的 -x 参数, 程序部分 token)。
    """
    env = env or {}
    toks = _tokens(line)
    out: Dict[str, str] = {}
    multi: Dict[str, List[str]] = {}
    unknown: List[str] = []
    prog: List[str] = []
    i = 0
    started = False

    def put(key: str, val: str):
        pd = PARAM_BY_KEY.get(key)
        if pd and pd.multiple:
            lst = multi.setdefault(key, [])
            if val and val not in lst:
                lst.append(val)
        else:
            out[key] = val

    while i < len(toks):
        t = toks[i]
        low = t.lower()
        if t in _FLAGS or low in _FLAGS:
            out[_FLAGS.get(t) or _FLAGS[low]] = "1"
            i += 1
            started = True
            continue
        key = _OPTS.get(t) or _OPTS.get(low)
        if key:
            if i + 1 < len(toks):
                put(key, _expand(_unquote(toks[i + 1]), env))
                i += 2
            else:
                put(key, "")
                i += 1
            started = True
            continue
        if t in ("-x", "--xargs"):
            if i + 1 < len(toks):
                arg = toks[i + 1]
                k, _, v = arg.partition("=")
                mapped = _XARGS.get(k)
                if mapped is None and _FILE_EXT_RE.match(k or ""):
                    mapped = "fileExt"
                if mapped:
                    put(mapped, _expand(_unquote(v), env))
                else:
                    unknown.append(arg)
                i += 2
            else:
                i += 1
            started = True
            continue
        if not started and not t.startswith("-"):
            prog.append(_expand(t, env))
            i += 1
            continue
        i += 1
    for k, vals in multi.items():
        out[k] = ",".join(vals)
    return out, unknown, prog


def _join_lines(text: str, kind: str) -> List[str]:
    """把 `^` / `\\` 续行拼回一行，便于解析。"""
    mark = "^" if kind == "bat" else "\\"
    out: List[str] = []
    buf = ""
    for raw in (text or "").splitlines():
        s = raw.rstrip()
        if s.endswith(mark):
            buf += s[:-1].rstrip() + " "
        else:
            out.append(buf + s)
            buf = ""
    if buf:
        out.append(buf)
    return out


def parse_script(text: str, kind: str) -> Dict[str, str]:
    """从已有 gen 脚本里把参数解析出来（用于打开界面时回填）。

    只认「带 `--conf` 的那行」为 Luban 命令：脚本里的 `cd` / `read` / `pause`
    都含 `-p`、`-r` 之类容易被误认的片段，不加这条会解析出一堆假参数。
    """
    env: Dict[str, str] = {}
    set_quoted = re.compile(r'^\s*set\s+"([^"]+)=(.*)"\s*$', re.I)
    set_plain = re.compile(r'^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$', re.I)
    joint = _join_lines(text, kind)
    for line in joint:
        s = line.strip()
        m = set_quoted.match(s) or set_plain.match(s)
        if m:
            env[m.group(1)] = _unquote(m.group(2).strip())

    lines = [l for l in joint if "--conf" in l] or [
        l for l in joint
        if l.strip() and not _is_shell_noise(l.strip()) and _known_flag(_tokens(l))
    ]
    out: Dict[str, str] = {}
    prog: List[str] = []
    unknown: List[str] = []
    for line in lines:
        o, u, p = parse_command_line(line, env)
        if not p:
            continue
        for k, v in o.items():
            if not v:
                continue
            cur = out.get(k, "")
            if not cur:
                out[k] = v
            elif v not in [x.strip() for x in cur.split(",")]:
                out[k] = cur + "," + v
        if u:
            unknown += [x for x in u if x not in unknown]
        prog = prog or p
    if prog:
        out["_exe"] = " ".join(prog)
    if unknown:
        out["extra"] = " ".join(unknown)
    return out


# ---------------------------------------------------------------- 参数默认值
def blank_params() -> Dict[str, str]:
    return {p.key: p.default for p in PARAMS}


def exe_mode_of(exe: str) -> str:
    """从可执行文件路径反推执行方式；看不出来返回空。"""
    last = (exe or "").strip().strip('"').split(" ")[-1].lower()
    for m in EXE_MODES:
        if last.endswith("." + m):
            return m
    return ""


def find_exe(root: str, mode: str = EXE_MODE_DLL) -> str:
    """在工程目录里找 Luban.dll / Luban.exe，找到就返回**相对路径**。

    先找工程根目录，再往下找一层子目录（Luban 常被放进 Tools / bin 之类）。
    返回相对路径是为了让生成的 gen 脚本可移植（脚本自己会 cd 到工程目录）。
    """
    if not root or not os.path.isdir(root):
        return ""
    names = _EXE_NAMES.get(mode) or _EXE_NAMES[EXE_MODE_DLL]
    for n in names:
        if os.path.isfile(os.path.join(root, n)):
            return n
    try:
        subs = sorted(os.listdir(root))
    except Exception:
        subs = []
    for d in subs:
        sd = os.path.join(root, d)
        if d.startswith(".") or not os.path.isdir(sd):
            continue
        for n in names:
            if os.path.isfile(os.path.join(sd, n)):
                return (d + "/" + n).replace("\\", "/")
    return ""


def find_conf(root: str) -> str:
    """在工程目录里找 luban.conf，找到就返回相对路径，没找到返回空（由用户自己填）。"""
    if not root or not os.path.isdir(root):
        return ""
    for n in _CONF_NAMES:
        if os.path.isfile(os.path.join(root, n)):
            return n
    return ""


def _clean_path(v: str) -> str:
    """脚本里常写 `%~dp0luban.conf` / `$DIR/luban.conf`，界面上要显示成相对路径。"""
    s = (v or "").strip().strip('"')
    for pre in ("%~dp0", "%CD%", "$DIR", "${DIR}"):
        if s.startswith(pre):
            s = s[len(pre):]
            break
    return s.lstrip("\\/")


def defaults(model: ProjectModel, settings: Optional[Dict[str, str]] = None,
             script_text: str = "", kind: str = "bat") -> Dict[str, str]:
    """界面打开时的初始值。

    取值优先级：设置里存过的执行方式 → 现有脚本 → 工程目录里自动查找 → 内置默认值
    （执行方式 dll、代码目标 cs-simple-json、数据格式 json）。
    """
    settings = settings or {}
    root = model.root or ""
    p = blank_params()

    # 执行方式默认 dll；只有现有脚本里明确写了 exe 才跟着用 exe（不记住上次选择）
    mode = EXE_MODE_DLL

    if script_text.strip():
        got = parse_script(script_text, kind)
        exe = (got.pop("_exe", "") or "").strip()
        for k, v in got.items():
            if k in p and v:
                p[k] = _clean_path(v) if k == "conf" else v
        if exe:
            p["exe"] = _clean_path(exe)
            m2 = exe_mode_of(p["exe"])
            if m2:
                mode = m2
    p["exeMode"] = mode

    # 可执行文件：脚本没给 / 给了但工程里根本没这个文件 → 去工程目录找；
    # 工程里也没有 → 用设置里记录过的最新路径；再没有就保持为空（让用户自己填）
    if not p["exe"] or not os.path.isfile(os.path.join(root, _clean_path(p["exe"]))):
        found = find_exe(root, mode) or find_exe(
            root, EXE_MODE_EXE if mode == EXE_MODE_DLL else EXE_MODE_DLL)
        if found:
            p["exe"] = found
            m2 = exe_mode_of(found)
            if m2:
                p["exeMode"] = m2
        else:
            remembered = (settings.get("luban_exe") or "").strip()
            if remembered:
                p["exe"] = remembered
                m2 = exe_mode_of(remembered)
                if m2:
                    p["exeMode"] = m2
            else:
                p["exe"] = ""

    # 配置文件：优先工程目录里的 luban.conf，没有就留空让用户自己填
    conf = _clean_path(p.get("conf", ""))
    if not (conf and os.path.isfile(os.path.join(root, conf))):
        conf = find_conf(root)
    p["conf"] = conf

    if not p.get("target"):
        p["target"] = (settings.get("target") or "").strip() \
            or (model.targets[0] if model.targets else "all")
    extra = (settings.get("extra_args") or "").strip()
    if extra and not script_text.strip():
        o, u, _p = parse_command_line(extra, {})
        for k, v in o.items():
            if k in p and v:
                p[k] = v
        if u:
            p["extra"] = " ".join(u)

    # 输出目录：没指定就默认工程 Output/Code、Output/Data（与 luban.conf 一致）
    for key, sub in (("outputCodeDir", "Code"), ("outputDataDir", "Data")):
        if not (p.get(key) or "").strip():
            p[key] = "Output/" + sub
    return p


# ---------------------------------------------------------------- 命令生成
def _q(v: str) -> str:
    v = (v or "").strip()
    if not v:
        return ""
    if " " in v and not (v.startswith('"') and v.endswith('"')):
        return '"%s"' % v
    return v


def _multi(v: str) -> List[str]:
    return [x.strip() for x in re.split(r"[,，]+", (v or "").strip()) if x.strip()]


def _exe_part(exe: str, mode: str = "") -> str:
    """命令行最前面那一段：dll 要写成 `dotnet "xx.dll"`，exe 直接跑。"""
    raw = (exe or "").strip()
    if not raw:
        # 还没指定执行文件 → 留空，不塞占位名（否则看起来像有个默认值）
        return ""
    if raw.lower().startswith("dotnet"):
        return raw
    # 只有 .dll 才需要 dotnet —— 别因为「执行方式选了 dll」就把 exe 也塞给 dotnet
    if raw.lower().endswith(".dll"):
        return "dotnet " + _q(raw)
    return _q(raw)


def _rel2dp0(s: str) -> str:
    """相对路径补 `%~dp0`（脚本目录）前缀；绝对路径原样。"""
    s = (s or "").strip()
    if not s:
        return s
    if s.startswith('"'):
        inner = s[1:]
        if inner.endswith('"'):
            inner = inner[:-1]
        if len(inner) >= 2 and inner[1] == ":":
            return s
        return '"%%~dp0%s"' % inner
    if (len(s) >= 2 and s[1] == ":") or s.startswith("%~dp0"):
        return s
    return "%%~dp0%s" % s


def _bat_exe(exe: str) -> str:
    """bat 脚本里的 exe 段：相对路径补 `%~dp0`。

    Windows 的 NoDefaultCurrentDirectoryInExePath 会让 cmd 不搜当前目录，
    直接写 `Luban.exe` 会报「不是内部或外部命令」；写成 `%~dp0Luban.exe` 才稳。
    """
    exe = (exe or "").strip()
    if not exe:
        return exe
    if exe.lower().startswith("dotnet "):
        return "dotnet " + _rel2dp0(exe[7:].strip())
    return _rel2dp0(exe)


def _materialize(text: str, root: str) -> str:
    """把脚本里的 `%~dp0` / `%CD%` 换成工程绝对路径。

    正式脚本要可移植，所以保留 `%~dp0`；但临时脚本生成在系统临时目录，
    那里的 `%~dp0` 指的是临时目录，必须换成工程目录才能读到 luban.conf。
    """
    r = root.rstrip("\\/")
    out = text.replace("%~dp0", r + os.sep).replace("%CD%", r)
    return out


def build_command(p: Dict[str, str]) -> str:
    """参数 -> 一条完整的 Luban 命令行。"""
    parts = [_exe_part(p.get("exe", ""), p.get("exeMode", ""))]
    conf = (p.get("conf") or "").strip()
    if conf:
        parts.append('--conf "%s"' % conf)
    tgt = (p.get("target") or "").strip()
    if tgt:
        parts.append("-t %s" % tgt)
    for v in _multi(p.get("codeTarget", "")):
        parts.append("-c %s" % v)
    for v in _multi(p.get("dataTarget", "")):
        parts.append("-d %s" % v)
    for key in ("outputCodeDir", "outputDataDir"):
        v = (p.get(key) or "").strip()
        if v:
            parts.append('-x %s="%s"' % (PARAM_BY_KEY[key].flag, v))
    # 高级：普通开关（带值）
    for key in ("schemaCollector", "pipeline", "outputTable", "includeTag",
                "excludeTag", "variant", "timeZone", "logConfig", "watchDir"):
        v = (p.get(key) or "").strip()
        if v:
            parts.append("%s %s" % (PARAM_BY_KEY[key].flag, _q(v)))
    # 高级：布尔开关
    for key in ("forceLoadTableDatas", "validationFailAsError", "verbose"):
        if _truthy(p.get(key)):
            parts.append(PARAM_BY_KEY[key].flag)
    # 高级：-x 扩展项
    for key in ("outputSaver", "cleanUpOutputDir", "fileExt", "jsonCompact",
                "lineEnding", "l10nProvider", "l10nTextFile"):
        v = (p.get(key) or "").strip()
        if not v:
            continue
        flag = PARAM_BY_KEY[key].flag
        if key == "fileExt":
            dts = _multi(p.get("dataTarget", ""))
            if not dts:
                continue
            flag = "%s.fileExt" % dts[0]
        parts.append('-x %s="%s"' % (flag, v))
    extra = (p.get("extra") or "").strip()
    if extra:
        for tk in extra.split():
            parts.append("-x %s" % tk)
    return " ".join(x for x in parts if x).strip()


def _truthy(v) -> bool:
    return str(v or "").strip().lower() in ("1", "true", "yes", "y", "on")


def _bat_lines(p: Dict[str, str], root: str, portable: bool) -> List[str]:
    exe = _bat_exe(_exe_part(p.get("exe", ""), p.get("exeMode", "")))
    conf = (p.get("conf") or "").strip() or "luban.conf"
    tgt = (p.get("target") or "").strip() or "all"
    body: List[str] = []
    body.append('--conf "%s"' % conf)
    body.append("-t %s" % tgt)
    for v in (_multi(p.get("codeTarget", "")) or ["cs-simple-json"]):
        body.append("-c %s" % v)
    for v in (_multi(p.get("dataTarget", "")) or ["json"]):
        body.append("-d %s" % v)
    for key in ("outputCodeDir", "outputDataDir"):
        v = (p.get(key) or "").strip()
        if v:
            body.append('-x %s="%s"' % (PARAM_BY_KEY[key].flag, v))
    for key in ("schemaCollector", "pipeline", "outputTable", "includeTag",
                "excludeTag", "variant", "timeZone", "logConfig", "watchDir"):
        v = (p.get(key) or "").strip()
        if v:
            body.append("%s %s" % (PARAM_BY_KEY[key].flag, _q(v)))
    for key in ("forceLoadTableDatas", "validationFailAsError", "verbose"):
        if _truthy(p.get(key)):
            body.append(PARAM_BY_KEY[key].flag)
    for key in ("outputSaver", "cleanUpOutputDir", "fileExt", "jsonCompact",
                "lineEnding", "l10nProvider", "l10nTextFile"):
        v = (p.get(key) or "").strip()
        if not v:
            continue
        flag = PARAM_BY_KEY[key].flag
        if key == "fileExt":
            dts = _multi(p.get("dataTarget", ""))
            if not dts:
                continue
            flag = "%s.fileExt" % dts[0]
        body.append('-x %s="%s"' % (flag, v))
    for tk in (p.get("extra") or "").split():
        body.append("-x %s" % tk)

    lines = [
        "@echo off",
        "rem ---------------------------------------------------------------",
        "rem  Luban 导表脚本（由 LubanEditor 生成）",
        "rem  可以直接双击运行；LubanEditor 点「发布」运行时不会暂停等待按键。",
        "rem ---------------------------------------------------------------",
        'cd /d "%~dp0"' if portable else 'cd /d "%s"' % root,
        "",
        f'{exe} ^',
    ]
    for j, b in enumerate(body):
        cont = " ^" if j < len(body) - 1 else ""
        lines.append("    " + b + cont)
    lines += [
        "",
        "if errorlevel 1 (",
        "    echo.",
        "    echo 导表失败，请检查上面的报错信息。",
        "    exit /b 1",
        ")",
        "echo 导表完成",
        "if not defined %s pause" % PAUSE_ENV,
    ]
    return lines


def _sh_lines(p: Dict[str, str], root: str, portable: bool) -> List[str]:
    exe = _exe_part(p.get("exe", ""), p.get("exeMode", ""))
    conf = (p.get("conf") or "").strip() or "luban.conf"
    tgt = (p.get("target") or "").strip() or "all"
    body: List[str] = []
    body.append('--conf "%s"' % conf)
    body.append("-t %s" % tgt)
    for v in (_multi(p.get("codeTarget", "")) or ["cs-simple-json"]):
        body.append("-c %s" % v)
    for v in (_multi(p.get("dataTarget", "")) or ["json"]):
        body.append("-d %s" % v)
    for key in ("outputCodeDir", "outputDataDir"):
        v = (p.get(key) or "").strip()
        if v:
            body.append('-x %s="%s"' % (PARAM_BY_KEY[key].flag, v))
    for key in ("schemaCollector", "pipeline", "outputTable", "includeTag",
                "excludeTag", "variant", "timeZone", "logConfig", "watchDir"):
        v = (p.get(key) or "").strip()
        if v:
            body.append("%s %s" % (PARAM_BY_KEY[key].flag, _q(v)))
    for key in ("forceLoadTableDatas", "validationFailAsError", "verbose"):
        if _truthy(p.get(key)):
            body.append(PARAM_BY_KEY[key].flag)
    for key in ("outputSaver", "cleanUpOutputDir", "fileExt", "jsonCompact",
                "lineEnding", "l10nProvider", "l10nTextFile"):
        v = (p.get(key) or "").strip()
        if not v:
            continue
        flag = PARAM_BY_KEY[key].flag
        if key == "fileExt":
            dts = _multi(p.get("dataTarget", ""))
            if not dts:
                continue
            flag = "%s.fileExt" % dts[0]
        body.append('-x %s="%s"' % (flag, v))
    for tk in (p.get("extra") or "").split():
        body.append("-x %s" % tk)

    lines = [
        "#!/bin/bash",
        "# Luban 导表脚本（由 LubanEditor 生成）",
        "# 执行：./gen.sh ；LubanEditor 点「发布」运行时不会等待按键。",
        'DIR="$(cd "$(dirname "$0")" && pwd)"' if portable
        else 'DIR="%s"' % root.replace("\\", "/"),
        'cd "$DIR" || exit 1',
        "",
        exe + " \\",
    ]
    for j, b in enumerate(body):
        cont = " \\" if j < len(body) - 1 else ""
        lines.append("    " + b + cont)
    lines += [
        "",
        'if [ $? -ne 0 ]; then',
        "    echo 导表失败，请检查上面的报错信息。",
        "    exit 1",
        "fi",
        "echo 导表完成",
        'if [ -z "$%s" ]; then read -r -p "按回车键退出..." _; fi' % PAUSE_ENV,
    ]
    return lines


def build_script(p: Dict[str, str], kind: str, root: str = "",
                 portable: bool = True) -> str:
    """参数 -> gen.bat / gen.sh 全文。portable=False 时把工程绝对路径写进 cd。"""
    lines = _bat_lines(p, root, portable) if kind == "bat" \
        else _sh_lines(p, root, portable)
    text = "\n".join(lines) + "\n"
    return text if portable else _materialize(text, root)


def sync_settings(p: Dict[str, str], settings: Dict[str, str]) -> None:
    """把界面上的参数回写到工具设置，让「生成并校验」用同一套参数。"""
    settings["luban_exe"] = (p.get("exe") or "").strip()
    settings["exe_mode"] = (p.get("exeMode") or EXE_MODE_DLL).strip().lower()
    settings["conf_path"] = (p.get("conf") or "").strip()
    settings["target"] = (p.get("target") or "").strip()
    bits: List[str] = []
    for v in _multi(p.get("codeTarget", "")):
        bits.append("-c " + v)
    for v in _multi(p.get("dataTarget", "")):
        bits.append("-d " + v)
    for key in ("schemaCollector", "pipeline", "outputTable", "includeTag",
                "excludeTag", "variant", "timeZone", "logConfig", "watchDir"):
        v = (p.get(key) or "").strip()
        if v:
            bits.append("%s %s" % (PARAM_BY_KEY[key].flag, _q(v)))
    for key in ("forceLoadTableDatas", "validationFailAsError", "verbose"):
        if _truthy(p.get(key)):
            bits.append(PARAM_BY_KEY[key].flag)
    for key in ("outputSaver", "cleanUpOutputDir", "fileExt", "jsonCompact",
                "lineEnding", "l10nProvider", "l10nTextFile"):
        v = (p.get(key) or "").strip()
        if not v:
            continue
        flag = PARAM_BY_KEY[key].flag
        if key == "fileExt":
            dts = _multi(p.get("dataTarget", ""))
            if not dts:
                continue
            flag = "%s.fileExt" % dts[0]
        bits.append('-x %s="%s"' % (flag, v))
    for tk in (p.get("extra") or "").split():
        bits.append("-x " + tk)
    settings["extra_args"] = " ".join(bits).strip()


def script_ext(kind: str) -> str:
    return "bat" if kind == "bat" else "sh"


def preferred_kind() -> str:
    """本机该用哪个脚本：Windows 用 gen.bat，其它用 gen.sh。"""
    return "bat" if os.name == "nt" else "sh"


def read_script(root: str, kind: str) -> str:
    p = os.path.join(root, "gen." + script_ext(kind))
    if not os.path.exists(p):
        return ""
    try:
        with open(p, "r", encoding="gbk" if kind == "bat" else "utf-8",
                  errors="replace") as f:
            return f.read()
    except Exception:
        return ""
