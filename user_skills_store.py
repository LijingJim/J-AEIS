"""
=============================================================
  用户自定义 Skill 存储（JSON 持久化）
=============================================================
- 用户通过 Sidebar UI 或 LLM 工具创建 Skill
- 存储在 user_skills.json 中，与内置 Skill 合并使用
- 用户 Skill 可删除，内置 Skill 不可删除
=============================================================
"""
import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

STORE_PATH = Path(__file__).parent / "user_skills.json"
_LOCK = threading.Lock()

# ── 内置 Skill key 集合（不可删除）────────────────
BUILTIN_SKILL_KEYS = {
    "data_analyst",
    "quick_overview",
    "deep_dive",
    "executive_brief",
    "contract_reviewer",
}


def _load() -> dict:
    """从 JSON 文件加载用户 Skill。"""
    if not STORE_PATH.exists():
        return {}
    try:
        with open(str(STORE_PATH), "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {}


def _save(data: dict):
    """写入 JSON 文件。"""
    with _LOCK:
        with open(str(STORE_PATH), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


def list_user_skills() -> List[dict]:
    """列出所有用户创建的 Skill（返回简要信息列表）。"""
    data = _load()
    result = []
    for key, cfg in data.items():
        result.append({
            "key": key,
            "name": cfg.get("name", key),
            "keywords": cfg.get("keywords", []),
            "created_at": cfg.get("created_at", ""),
            "allowed_tools": cfg.get("allowed_tools", []),
        })
    result.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return result


def load_user_skills_as_dict() -> dict:
    """
    加载用户 Skill，返回与内置 SKILLS 相同格式的字典。
    用于 skills.py 合并到 SKILLS 注册表中。
    """
    data = _load()
    result = {}
    for key, cfg in data.items():
        result[key] = {
            "name": cfg.get("name", key),
            "keywords": cfg.get("keywords", []),
            "system_prompt": cfg.get("system_prompt", ""),
            "entry_tool_for_table": cfg.get("entry_tool_for_table", None),
            "entry_tool_for_generic": cfg.get("entry_tool_for_generic", None),
            "allowed_tools": set(cfg.get("allowed_tools", [])),
            "forbidden_tools": set(cfg.get("forbidden_tools", [])),
            "_user_skill": True,  # 标记为用户创建
        }
    return result


def create_user_skill(
    name: str,
    display_name: str,
    keywords: str,
    system_prompt: str,
    allowed_tools: str = "",
    forbidden_tools: str = "",
    entry_tool_for_table: str = "",
    entry_tool_for_generic: str = "",
) -> str:
    """
    创建或更新一个用户 Skill。

    参数：
        name: Skill 唯一标识（英文 key，如 my_finance_review）
        display_name: 显示名称（中文，如"财务对账"）
        keywords: 触发关键词，逗号分隔
        system_prompt: 系统提示词（LLM 角色与工作流）
        allowed_tools: 允许的工具名，逗号分隔
        forbidden_tools: 禁止的工具名，逗号分隔
        entry_tool_for_table: 表格文件的入口工具
        entry_tool_for_generic: 通用文件的入口工具

    返回：成功/失败信息
    """
    import re

    # 验证 key
    if not re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', name):
        return f"❌ Skill 标识 '{name}' 不合法。只允许字母、数字、下划线，且不能以数字开头。"

    if name in BUILTIN_SKILL_KEYS:
        return f"❌ '{name}' 是内置 Skill，不能覆盖。请使用其他名称。"

    if not display_name.strip():
        return "❌ 显示名称不能为空。"

    if not system_prompt.strip():
        return "❌ 系统提示词不能为空。"

    # 解析 keywords
    kw_list = [kw.strip() for kw in keywords.replace("，", ",").split(",") if kw.strip()]
    if not kw_list:
        return "❌ 至少需要一个触发关键词。"

    # 解析工具列表
    at_list = [t.strip() for t in allowed_tools.replace("，", ",").split(",") if t.strip()] if allowed_tools.strip() else []
    ft_list = [t.strip() for t in forbidden_tools.replace("，", ",").split(",") if t.strip()] if forbidden_tools.strip() else []

    # 加载现有数据
    data = _load()
    is_update = name in data

    data[name] = {
        "name": display_name.strip(),
        "keywords": kw_list,
        "system_prompt": system_prompt.strip(),
        "allowed_tools": at_list,
        "forbidden_tools": ft_list,
        "entry_tool_for_table": entry_tool_for_table.strip() or None,
        "entry_tool_for_generic": entry_tool_for_generic.strip() or None,
        "created_at": data.get(name, {}).get("created_at", datetime.now().strftime("%Y-%m-%d %H:%M")),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    _save(data)

    action = "更新" if is_update else "创建"
    return (
        f"✅ 用户 Skill '{display_name}' {action}成功！\n"
        f"   - 标识: {name}\n"
        f"   - 触发词: {', '.join(kw_list)}\n"
        f"   - 可用工具: {', '.join(at_list) if at_list else '(全部)'}\n"
        f"   - 禁止工具: {', '.join(ft_list) if ft_list else '(无)'}\n"
    )


def delete_user_skill(name: str) -> str:
    """删除一个用户创建的 Skill。"""
    if name in BUILTIN_SKILL_KEYS:
        return f"❌ '{name}' 是内置 Skill，不可删除。"

    data = _load()
    if name not in data:
        return f"⚠️ 用户 Skill '{name}' 不存在。"

    display_name = data[name].get("name", name)
    del data[name]
    _save(data)
    return f"🗑️ 用户 Skill '{display_name}'（{name}）已删除。"
