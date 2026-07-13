"""
=============================================================
  动态工具管理器 —— 运行时创建/注册/执行新 Function
=============================================================
核心能力：
  1. 运行时注册新工具（名称、描述、参数 schema、Python 实现）
  2. 安全沙箱执行（受限 builtins + 禁止危险模块）
  3. 工具缺口检测——分析对话历史，提示 LLM 是否需要造新工具
  4. 持久化到 session_state，跨轮次可用
=============================================================
"""

import json
import re
import traceback
import streamlit as st

# ── 安全沙箱：允许的内置函数 ─────────────────
SAFE_BUILTINS = {
    "abs", "all", "any", "bin", "bool", "bytes", "chr", "complex",
    "dict", "divmod", "enumerate", "filter", "float", "format",
    "frozenset", "getattr", "hasattr", "hash", "hex", "int",
    "isinstance", "issubclass", "iter", "len", "list", "map",
    "max", "min", "next", "oct", "ord", "pow", "print", "range",
    "repr", "reversed", "round", "set", "slice", "sorted",
    "str", "sum", "tuple", "type", "zip",
    # 常用标准库（安全）
    "json", "re", "math", "datetime", "collections", "itertools",
    "functools", "statistics", "textwrap", "string", "decimal",
    "fractions", "random", "uuid", "base64", "hashlib",
    "csv", "io", "pathlib", "os",  # os 允许但 path 操作无害
    "time",
    # requests 用于 HTTP 调用
    "requests",
    # 数据科学
    "pd", "np",
    "True", "False", "None", "Exception", "ValueError", "TypeError",
    "KeyError", "IndexError", "StopIteration", "RuntimeError",
}

# ── 危险模块黑名单（即使 LLM 试图 import 也会被拦截） ──
DANGEROUS_IMPORTS = {
    "subprocess", "os.system", "os.popen", "os.exec", "os.spawn",
    "shutil", "socket", "ctypes", "multiprocessing", "threading",
    "signal", "pty", "fcntl", "posix", "pwd", "grp", "crypt",
    "sys",  # sys 太危险，但某些场景可能需要 → 白名单替代
    "builtins", "importlib", "imp", "marshal", "code", "codeop",
    "compile", "eval", "exec", "__builtins__", "__import__",
    "open",  # 禁止直接 open，用 io.StringIO 等
}

# ── 持久化 key ────────────────────────────────
DYNAMIC_TOOLS_KEY = "_dynamic_tools"

# 【关键】模块级动态工具存储 —— 跨线程共享，替代 session_state
# 原因同 excel_tool.py 的 _FILE_CACHE：st.session_state 用了 threading.local
_DYNAMIC_TOOLS_STORE: dict = {}  # {name: {definition, handler_code, ...}}


def _init_store():
    """初始化动态工具存储（模块级 + session_state 双写）。"""
    if DYNAMIC_TOOLS_KEY not in st.session_state:
        st.session_state[DYNAMIC_TOOLS_KEY] = {}


def _get_store():
    """获取动态工具存储。优先读模块级（跨线程可见），回退 session_state。"""
    if _DYNAMIC_TOOLS_STORE:
        return _DYNAMIC_TOOLS_STORE
    _init_store()
    return st.session_state[DYNAMIC_TOOLS_KEY]


def _safe_import(name, globals_dict):
    """安全的 import 钩子：拦截危险模块。"""
    if name in DANGEROUS_IMPORTS:
        raise ImportError(f"模块 '{name}' 被安全策略禁止导入。")
    # 对于 os，只允许 os.path 等无害操作
    if name == "os":
        import os as _os_mod
        # 移除危险函数
        safe_os = type("SafeOS", (), {})()
        for attr in dir(_os_mod):
            if attr in ("system", "popen", "execv", "execve", "spawnl", "spawnle",
                         "spawnlp", "spawnlpe", "spawnv", "spawnve", "spawnvp", "spawnvpe",
                         "fork", "kill", "setuid", "setgid", "chroot", "chmod", "chown"):
                continue
            try:
                setattr(safe_os, attr, getattr(_os_mod, attr))
            except Exception:
                pass
        return safe_os
    return __import__(name)


def _build_safe_globals():
    """构建安全的全局命名空间。"""
    import math
    import datetime
    import collections
    import itertools
    import functools
    import statistics
    import textwrap
    import string
    import decimal
    import fractions
    import random
    import uuid
    import base64
    import hashlib
    import csv
    import io
    from pathlib import Path
    import time
    import json as _json
    import re as _re

    safe_globals = {
        # 安全内置
        **{k: getattr(__builtins__, k) for k in SAFE_BUILTINS if k in dir(__builtins__)},
        # 安全模块
        "json": _json,
        "re": _re,
        "math": math,
        "datetime": datetime,
        "collections": collections,
        "itertools": itertools,
        "functools": functools,
        "statistics": statistics,
        "textwrap": textwrap,
        "string": string,
        "decimal": decimal,
        "fractions": fractions,
        "random": random,
        "uuid": uuid,
        "base64": base64,
        "hashlib": hashlib,
        "csv": csv,
        "io": io,
        "Path": Path,
        "time": time,
        # pandas / numpy（如果在环境中可用）
    }
    # 安全注入 __import__
    safe_globals["__builtins__"] = {
        k: safe_globals.get(k, getattr(__builtins__, k, None))
        for k in SAFE_BUILTINS
        if k in dir(__builtins__)
    }
    safe_globals["__builtins__"]["__import__"] = lambda name, *args, **kw: _safe_import(name, safe_globals)

    # 尝试加载 pandas 和 numpy
    try:
        import pandas as pd
        safe_globals["pd"] = pd
    except ImportError:
        pass
    try:
        import numpy as np
        safe_globals["np"] = np
    except ImportError:
        pass
    try:
        import requests
        safe_globals["requests"] = requests
    except ImportError:
        pass

    # ── matplotlib ──────────────────────────────────────
    # 允许动态工具创建自定义图表，保存到 session_state 供 build_report 使用
    try:
        import matplotlib as _mpl
        _mpl.use("Agg")
        _mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]
        _mpl.rcParams["axes.unicode_minus"] = False
        import matplotlib.pyplot as _plt
        safe_globals["matplotlib"] = _mpl
        safe_globals["plt"] = _plt
        safe_globals["io"] = io
        safe_globals["_session_state"] = st.session_state

        # 便捷函数：动态工具调用此函数将 matplotlib 图表存入报告
        def _sandbox_save_chart(fig, title="图表", description=""):
            try:
                buf = io.BytesIO()
                fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
                buf.seek(0)
                img_bytes = buf.getvalue()
                _plt.close(fig)
                entry = {"title": title, "description": description, "image": img_bytes}
                charts = list(st.session_state.get("generated_charts", []))
                charts = [c for c in charts if c.get("title") != title]
                charts.append(entry)
                st.session_state["generated_charts"] = charts
                return f"✅ 图表「{title}」已保存，将出现在最终报告中。"
            except Exception as _e:
                return f"⚠️ 图表保存失败: {_e}"

        safe_globals["save_chart"] = _sandbox_save_chart
    except ImportError:
        pass

    return safe_globals


def _validate_code(code: str) -> tuple:
    """验证代码安全性，返回 (is_safe, reason)。"""
    if not code or not code.strip():
        return False, "代码为空"

    # 检测危险导入
    for dangerous in DANGEROUS_IMPORTS:
        # 检测 import dangerous 或 from dangerous import ...
        if re.search(rf'\bimport\s+{re.escape(dangerous)}\b', code):
            return False, f"代码尝试导入危险模块: {dangerous}"
        if re.search(rf'\bfrom\s+{re.escape(dangerous)}\s+import\b', code):
            return False, f"代码尝试从危险模块导入: {dangerous}"

    # 检测危险函数调用（os.system, subprocess.call 等）
    for dangerous in ("os.system", "os.popen", "os.exec", "os.spawn",
                       "subprocess.call", "subprocess.run", "subprocess.Popen"):
        if re.search(rf'\b{re.escape(dangerous)}\s*\(', code):
            return False, f"代码调用危险函数: {dangerous}"

    # 检测 eval/exec/compile 调用
    if re.search(r'\b(eval|exec|compile|__import__)\s*\(', code):
        return False, "代码包含 eval/exec/compile/__import__ 调用"

    # 检测文件写入操作
    if re.search(r'\bopen\s*\([^)]*[\'\"][wa]\b', code):
        return False, "代码包含文件写入操作"

    return True, "OK"


def create_tool(name: str, description: str, parameters: str, code: str) -> str:
    """
    运行时创建并注册一个新工具。

    参数:
        name: 工具名称（英文标识符，如 fetch_weather）
        description: 工具用途描述（告知 LLM 何时调用）
        parameters: JSON Schema 格式的参数定义（字符串）
        code: Python 函数体代码（会包装为 handler(**kwargs) 执行）
              handler 接收 parameters 中定义的参数作为 kwargs。
              必须 return 一个字符串结果。

    返回:
        成功或失败信息
    """
    _init_store()
    store = _get_store()

    # ── 1. 验证工具名称 ──
    if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', name):
        return f"❌ 工具名称 '{name}' 不合法。只允许字母、数字、下划线，且不能以数字开头。"

    if name in store:
        return f"⚠️ 工具 '{name}' 已存在。如需覆盖，请先调用 delete_tool。"

    # ── 2. 验证参数 schema ──
    try:
        params_schema = json.loads(parameters) if isinstance(parameters, str) else parameters
    except json.JSONDecodeError as e:
        return f"❌ 参数 schema JSON 解析失败: {e}"

    # ── 3. 验证代码安全性 ──
    is_safe, reason = _validate_code(code)
    if not is_safe:
        return f"❌ 代码安全验证失败: {reason}"

    # ── 4. 构建完整的函数定义 ──
    full_code = f"""
def handler(**kwargs):
{chr(10).join('    ' + line for line in code.strip().split(chr(10)))}
"""
    # ── 5. 试执行以验证代码语法 ──
    safe_globals = _build_safe_globals()
    local_vars = {}
    try:
        exec(full_code, safe_globals, local_vars)
    except Exception as e:
        return f"❌ 代码编译/执行失败: {type(e).__name__}: {e}"

    handler_fn = local_vars.get("handler")
    if not callable(handler_fn):
        return "❌ 代码必须定义一个名为 handler 的可调用函数。"

    # ── 6. 存储工具定义（模块级 + session_state 双写）──
    tool_entry = {
        "definition": {
            "type": "function",
            "function": {
                "name": name,
                "description": description,
                "parameters": params_schema,
            },
        },
        "handler_code": full_code,
        "created_at": str(__import__("datetime").datetime.now()),
    }
    store[name] = tool_entry
    _DYNAMIC_TOOLS_STORE[name] = tool_entry  # 模块级存储，跨线程可见

    chart_hint = (
        "\n"
        "   - 💡 如需生成图表：可用 `plt` (matplotlib) 绘图，"
        "用 `save_chart(fig, title, description)` 存入报告。"
        "沙箱已提供 pd/np/plt/io 等模块。"
    )

    return (
        f"✅ 动态工具 '{name}' 创建成功！\n"
        f"   - 描述: {description[:80]}...\n"
        f"   - 参数: {list(params_schema.get('properties', {}).keys())}\n"
        f"   - 现在你可以在后续对话中直接调用它。"
        + (chart_hint if "图" in description or "chart" in description.lower() or "plot" in description.lower() else "")
    )


def delete_tool(name: str) -> str:
    """删除一个动态创建的工具。"""
    _init_store()
    store = _get_store()
    if name not in store and name not in _DYNAMIC_TOOLS_STORE:
        return f"⚠️ 工具 '{name}' 不存在。"
    store.pop(name, None)
    _DYNAMIC_TOOLS_STORE.pop(name, None)
    return f"🗑️ 动态工具 '{name}' 已删除。"


def list_dynamic_tools() -> str:
    """列出所有已注册的动态工具。"""
    _init_store()
    store = _get_store()
    if not store:
        return "📭 当前没有动态创建的工具。"

    lines = ["📦 **当前动态工具**：\n"]
    for name, info in store.items():
        defn = info["definition"]["function"]
        params = list(defn.get("parameters", {}).get("properties", {}).keys())
        lines.append(f"| `{name}` | {defn['description'][:60]}... | 参数: {', '.join(params)} | 创建于 {info.get('created_at', '?')} |")
    return "\n".join(lines)


def get_dynamic_tool_definitions() -> list:
    """返回所有动态工具的 OpenAI function calling 格式定义列表。"""
    _init_store()
    store = _get_store()
    return [info["definition"] for info in store.values()]


def execute_dynamic_tool(name: str, args: dict) -> str:
    """执行一个动态工具并返回结果。优先从模块级存储获取（跨线程可见）。"""
    # 先查模块级存储（跨线程可见）
    if name in _DYNAMIC_TOOLS_STORE:
        handler_code = _DYNAMIC_TOOLS_STORE[name]["handler_code"]
    else:
        # 回退 session_state
        _init_store()
        store = st.session_state.get(DYNAMIC_TOOLS_KEY, {})
        if name not in store:
            return f"❌ 动态工具 '{name}' 未注册。"
        handler_code = store[name]["handler_code"]
    safe_globals = _build_safe_globals()
    local_vars = {}

    try:
        exec(handler_code, safe_globals, local_vars)
        handler = local_vars["handler"]
        result = handler(**args)
        if result is None:
            return "(工具执行完毕，无返回值)"
        # Streamlit 是 HTML 渲染，原生支持 Unicode emoji，无需转码
        return str(result)
    except Exception as e:
        tb = traceback.format_exc()[-500:]
        # 识别网络相关错误，给出更清晰提示
        if any(k in str(e).lower() for k in ("timeout", "connection", "resolve", "getaddrinfo", "ssl", "proxy")):
            return (
                f"⚠️ 网络请求失败: {type(e).__name__}: {e}\n"
                f"提示: 请检查网络连接、代理设置或目标 URL 是否可访问。"
            )
        return f"❌ 动态工具 '{name}' 执行失败: {type(e).__name__}: {e}\n{tb}"


def detect_tool_gap(failed_action: str, conversation_context: str = "") -> str:
    """
    分析工具缺口：当 LLM 无法完成某个任务时，分析需要什么新工具。

    返回给 LLM 的引导信息，让它自己调用 create_tool。
    """
    return (
        "💡 **工具缺口检测**：当前功能调用无法满足需求。\n\n"
        f"需求: {failed_action}\n\n"
        "如需创建新工具，请调用 `create_tool` 并指定：\n"
        "1. `name`: 工具英文名（如 fetch_data、send_email、custom_chart）\n"
        "2. `description`: 何时调用、做什么\n"
        "3. `parameters`: JSON Schema 格式的参数定义\n"
        "4. `code`: Python 代码，定义 `def handler(**kwargs):` 并 return 结果\n\n"
        "💡 **创建图表工具时**：沙箱已内置 `pd`(pandas)、`np`(numpy)、`plt`(matplotlib)、`io`、"
        "`save_chart(fig, title, description)`（将图表存入报告）。\n"
        "  示例：用 `plt.figure()` 创建自定义图表 → `plt.savefig` / `save_chart(fig, ...)` → 图表自动进入最终报告。\n\n"
        "工具创建后即可在当前对话中调用。"
    )
