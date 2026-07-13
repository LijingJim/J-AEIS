"""
=============================================================
  会话持久化 + 分析快照 + 上下文压缩
=============================================================
三个能力一层层叠上去：
  1. SQLite 持久化 → 刷新页面对话不丢
  2. 分析快照     → 知道"分析到哪了、产出是什么"
  3. 滑动窗口压缩  → 长对话不超 token 窗口
=============================================================
"""
import sqlite3, json, uuid, re, os
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "sessions.db"

# ── 数据库初始化 ────────────────────────────────
def init_db():
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS conversations (
            session_id  TEXT,
            msg_index   INTEGER,
            msg_role    TEXT,
            msg_content TEXT,
            created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (session_id, msg_index)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS snapshots (
            session_id   TEXT PRIMARY KEY,
            phase        TEXT DEFAULT 'idle',
            main_question TEXT DEFAULT '',
            findings     TEXT DEFAULT '[]',
            charts       TEXT DEFAULT '[]',
            draft_full   TEXT DEFAULT '',
            draft_sections TEXT DEFAULT '[]',
            version      INTEGER DEFAULT 0,
            updated_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # v2.1: 会话命名
    conn.execute("""
        CREATE TABLE IF NOT EXISTS session_meta (
            session_id   TEXT PRIMARY KEY,
            session_name TEXT DEFAULT '',
            auto_topic   TEXT DEFAULT '',
            created_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

# ── 1. 会话管理 ─────────────────────────────────
def get_or_create_session():
    """从 URL 参数取 session_id，没有就生成新的并写入 URL。
    同时为新会话生成时间戳名称。"""
    import streamlit as st
    sid = st.query_params.get("session")
    if not sid:
        sid = str(uuid.uuid4())[:8]
        st.query_params["session"] = sid
        # 自动生成名称：日期_时间
        now = datetime.now()
        auto_name = now.strftime("%m%d_%H%M")
        set_session_name(sid, auto_name)
    return sid

def load_messages(session_id: str):
    """从 SQLite 恢复对话历史。返回 None 表示新会话。"""
    conn = sqlite3.connect(str(DB_PATH))
    rows = conn.execute(
        "SELECT msg_role, msg_content FROM conversations WHERE session_id=? ORDER BY msg_index",
        (session_id,)
    ).fetchall()
    conn.close()
    if not rows:
        return None
    messages = []
    for role, content in rows:
        try:
            msg = json.loads(content)
        except json.JSONDecodeError:
            msg = {"role": role, "content": content}
        messages.append(msg)
    return messages

def save_messages(session_id: str, messages: list):
    """保存完整对话历史到 SQLite（先删后插）。"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("DELETE FROM conversations WHERE session_id=?", (session_id,))
    for i, msg in enumerate(messages):
        msg_copy = dict(msg)
        # 不存 tool_calls 和 reasoning_content（避免过大且无意义）
        msg_copy.pop("tool_calls", None)
        msg_copy.pop("reasoning_content", None)
        conn.execute(
            "INSERT INTO conversations (session_id, msg_index, msg_role, msg_content) VALUES (?,?,?,?)",
            (session_id, i, msg.get("role", "unknown"), json.dumps(msg_copy, ensure_ascii=False, default=str))
        )
    conn.commit()
    conn.close()

def list_sessions(limit=20):
    """列出最近的会话，供侧边栏选择。"""
    conn = sqlite3.connect(str(DB_PATH))
    rows = conn.execute(
        "SELECT c.session_id, COUNT(*) as msg_count, MAX(c.created_at) as last_active, "
        "COALESCE(m.session_name, '') as session_name, COALESCE(m.auto_topic, '') as auto_topic "
        "FROM conversations c LEFT JOIN session_meta m ON c.session_id = m.session_id "
        "GROUP BY c.session_id ORDER BY last_active DESC LIMIT ?",
        (limit,)
    ).fetchall()
    conn.close()
    return [
        {"id": r[0], "msgs": r[1], "last": r[2],
         "name": r[3], "topic": r[4]}
        for r in rows
    ]

def delete_session(session_id: str):
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("DELETE FROM conversations WHERE session_id=?", (session_id,))
    conn.execute("DELETE FROM snapshots WHERE session_id=?", (session_id,))
    conn.execute("DELETE FROM session_meta WHERE session_id=?", (session_id,))
    conn.commit()
    conn.close()

# ── 1.5 会话命名 ───────────────────────────────
def get_session_name(session_id: str) -> str:
    """获取会话名称。返回 '' 表示未命名。"""
    conn = sqlite3.connect(str(DB_PATH))
    row = conn.execute(
        "SELECT session_name, auto_topic FROM session_meta WHERE session_id=?",
        (session_id,)
    ).fetchone()
    conn.close()
    if not row:
        return ""
    # 优先返回用户自定义名，否则返回自动名
    return row[0] or row[1] or ""


def set_session_name(session_id: str, name: str):
    """设置会话名称（用户手动命名）。"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute(
        "INSERT INTO session_meta (session_id, session_name) VALUES (?,?) "
        "ON CONFLICT(session_id) DO UPDATE SET session_name=?",
        (session_id, name, name)
    )
    conn.commit()
    conn.close()


def set_session_topic(session_id: str, topic: str):
    """自动检测会话主题（取首条用户消息或 Skill 名）。"""
    if not topic:
        return
    # 截断过长主题
    topic = topic.strip()[:20]
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute(
        "INSERT INTO session_meta (session_id, auto_topic) VALUES (?,?) "
        "ON CONFLICT(session_id) DO UPDATE SET auto_topic=? "
        "WHERE auto_topic = '' OR auto_topic IS NULL",
        (session_id, topic, topic)
    )
    conn.commit()
    conn.close()

# ── 2. 分析快照 ─────────────────────────────────
def new_snapshot():
    return {
        "phase": "idle",
        "main_question": "",
        "findings": [],
        "charts": [],
        "draft_full": "",
        "draft_sections": [],
        "version": 0,
    }

def get_snapshot(session_id: str) -> dict:
    """读取分析快照。"""
    conn = sqlite3.connect(str(DB_PATH))
    row = conn.execute(
        "SELECT phase, main_question, findings, charts, draft_full, draft_sections, version "
        "FROM snapshots WHERE session_id=?",
        (session_id,)
    ).fetchone()
    conn.close()
    if not row:
        return new_snapshot()
    return {
        "phase": row[0],
        "main_question": row[1],
        "findings": json.loads(row[2]),
        "charts": json.loads(row[3]),
        "draft_full": row[4],
        "draft_sections": json.loads(row[5]),
        "version": row[6],
    }

def save_snapshot(session_id: str, snap: dict):
    """写入分析快照。"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("""
        INSERT OR REPLACE INTO snapshots
        (session_id, phase, main_question, findings, charts, draft_full, draft_sections, version, updated_at)
        VALUES (?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)
    """, (
        session_id,
        snap.get("phase", "idle"),
        snap.get("main_question", ""),
        json.dumps(snap.get("findings", []), ensure_ascii=False),
        json.dumps(snap.get("charts", []), ensure_ascii=False),
        snap.get("draft_full", ""),
        json.dumps(snap.get("draft_sections", []), ensure_ascii=False),
        snap.get("version", 0),
    ))
    conn.commit()
    conn.close()

def auto_extract_snapshot(session_id: str, assistant_text: str, generated_charts: list = None):
    """
    从 LLM 的最终回答中自动提取分析快照。
    识别标题章节、发现列表，拆分为结构化 draft_sections。
    """
    snap = get_snapshot(session_id)
    snap["version"] = snap.get("version", 0) + 1
    snap["draft_full"] = assistant_text

    # 记录图表
    if generated_charts:
        snap["charts"] = [
            {"id": f"chart_{i+1}", "title": c.get("title", f"图表{i+1}")}
            for i, c in enumerate(generated_charts)
        ]

    # 按 ## 标题拆分段落
    sections = []
    parts = re.split(r'\n(?=#{1,3}\s)', assistant_text)
    for i, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue
        # 提取标题
        title_match = re.match(r'#{1,3}\s+(.+)', part)
        title = title_match.group(1).strip() if title_match else f"段落{i+1}"
        sections.append({
            "id": f"s{i+1}",
            "title": title,
            "content": part,
        })
    snap["draft_sections"] = sections

    # 更新阶段
    if any(kw in assistant_text[:200] for kw in ["核心发现", "关键发现", "主要发现"]):
        snap["phase"] = "findings_presented"
    elif sections:
        snap["phase"] = "report_drafted"

    save_snapshot(session_id, snap)
    return snap

def snapshot_to_prompt(snap: dict) -> str:
    """
    将分析快照转为一段 system prompt 注入文字。
    让 LLM 精确知道当前工作状态，支持"改第三段"等指令。
    """
    if snap.get("phase") == "idle":
        return ""

    parts = ["📋 **当前分析状态**"]
    if snap.get("main_question"):
        parts.append(f"- 主问题：{snap['main_question']}")

    if snap.get("findings"):
        parts.append(f"- 已有 {len(snap['findings'])} 个发现")
        for f in snap["findings"]:
            parts.append(f'  - [{f.get("id","?")}] {f.get("title","无标题")}')

    if snap.get("charts"):
        parts.append(f"- 已有 {len(snap['charts'])} 张图表")
        for c in snap["charts"]:
            parts.append(f'  - [{c.get("id","?")}] {c.get("title","无标题")}')

    sections = snap.get("draft_sections", [])
    if sections:
        parts.append(f"- 当前报告共 {len(sections)} 个段落：")
        for s in sections:
            preview = s.get("content", "")[:80].replace("\n", " ")
            parts.append(f'  - **第{s["id"][1:]}段** [{s.get("title","")}]：{preview}...')
        parts.append("用户可能用「改第N段」「第N段换成XX」来指令，请根据上方段落编号精准修改。")

    if snap.get("phase") == "report_drafted" and sections:
        parts.append("用户上次已完成报告初稿。若用户要求修改，直接找到对应段落修改即可，无需重新摸底。")

    parts.append("")
    return "\n".join(parts)


# ── 3. 上下文压缩 ───────────────────────────────
# 策略：滑动窗口 + 摘要
#   - 最近 KEEP_RECENT 条消息保留原文
#   - 更早的消息用 LLM 压缩成一条摘要注入 system
#   - 总条数 < COMPRESS_THRESHOLD 时不压缩

KEEP_RECENT = 10       # 保留最近 N 条原文
COMPRESS_THRESHOLD = 20  # 超过此条数才触发压缩

def _build_compress_prompt(old_messages: list) -> str:
    """构造压缩提示词。"""
    text = ""
    for m in old_messages:
        role = m.get("role", "?")
        content = m.get("content", "") or ""
        if role == "tool":
            content = content[:300]
        elif role == "assistant":
            content = content[:500]
        text += f"[{role}]: {content}\n"
    return (
        "请用 150 字以内的中文，总结以下对话的核心信息："
        "用户是谁、上传了什么文件、做了哪些分析、得出了什么关键结论。"
        "只写事实，不加评论。\n\n"
        f"{text[:4000]}"
    )

def compress_messages(messages: list, llm_client, model: str) -> list:
    """
    滑动窗口压缩：
    总条数 <= COMPRESS_THRESHOLD → 原样返回
    否则：旧消息 → LLM 摘要，插入 system；近 KEEP_RECENT 条保留
    """
    # 只数 user/assistant/tool 消息，忽略 system
    non_system = [m for m in messages if m.get("role") != "system"]
    if len(non_system) <= COMPRESS_THRESHOLD:
        return messages

    # 分离 system 消息和非 system 消息
    system_msgs = [m for m in messages if m.get("role") == "system"]
    old = non_system[:-KEEP_RECENT]
    recent = non_system[-KEEP_RECENT:]

    # 调 LLM 压缩旧消息
    try:
        resp = llm_client.chat.completions.create(
            model=model,
            messages=[
                {"role": "user", "content": _build_compress_prompt(old)}
            ],
            max_tokens=300,
            temperature=0.1,
        )
        summary = resp.choices[0].message.content.strip()
    except Exception:
        summary = f"（历史对话：共 {len(old)} 条，涉及数据分析与报告生成）"

    # 修复消息序列完整性：确保每个 tool 消息前面都有对应的 assistant(tool_calls) 消息
    recent = _ensure_message_sequence_integrity(recent)

    # 重组：system + 摘要 + 最近消息
    compressed = list(system_msgs)
    compressed.append({
        "role": "system",
        "content": f"📝 **历史对话摘要**：{summary}\n（以下为最近 {len(recent)} 条对话原文）"
    })
    compressed.extend(recent)
    return compressed


def _ensure_message_sequence_integrity(messages: list) -> list:
    """确保消息序列对 OpenAI API 合法：
    1. 每个 tool 消息前面必须有 assistant 消息且该 assistant 消息包含对应的 tool_calls
    2. 不能有连续的 tool 消息没有中间的 assistant
    3. assistant 有 tool_calls 但缺少对应 tool 响应时，清理 tool_calls
    """
    if not messages:
        return messages

    result = []
    i = 0
    while i < len(messages):
        msg = messages[i]
        role = msg.get("role", "")

        if role == "tool":
            # tool 消息需要前面有 assistant 消息带 tool_calls
            tool_call_id = msg.get("tool_call_id", "")
            has_valid_parent = False
            if result and result[-1].get("role") == "assistant":
                parent_calls = result[-1].get("tool_calls", [])
                if any(tc.get("id") == tool_call_id for tc in parent_calls):
                    has_valid_parent = True

            if not has_valid_parent:
                # 跳过这个孤立的 tool 消息
                i += 1
                continue

        elif role == "assistant":
            tool_calls = msg.get("tool_calls", [])
            if tool_calls:
                # 收集后面所有连续的 tool 消息
                following_tools = []
                j = i + 1
                while j < len(messages) and messages[j].get("role") == "tool":
                    following_tools.append(messages[j])
                    j += 1

                # 检查哪些 tool_call_id 有对应的 tool 消息
                available_tool_ids = {t.get("tool_call_id") for t in following_tools}
                valid_calls = [tc for tc in tool_calls if tc.get("id") in available_tool_ids]

                if not valid_calls:
                    # 没有有效的 tool 响应，移除 tool_calls
                    msg = dict(msg)
                    msg.pop("tool_calls", None)
                elif len(valid_calls) < len(tool_calls):
                    # 部分 tool 响应缺失，只保留有效的
                    msg = dict(msg)
                    msg["tool_calls"] = valid_calls

                result.append(msg)
                # 添加有对应父消息的 tool 消息
                valid_call_ids = {tc.get("id") for tc in (msg.get("tool_calls", []) or [])}
                for t in following_tools:
                    if t.get("tool_call_id") in valid_call_ids:
                        result.append(t)
                i = j
                continue

        result.append(msg)
        i += 1

    return result

# ── 4. 一键构建发送给 LLM 的上下文 ──────────────
def build_context(session_id: str, raw_messages: list, llm_client, model: str) -> list:
    """
    组装最终发给 LLM 的消息列表：
      1. 注入分析快照到 system prompt
      2. 检查是否需要压缩
    """
    messages = list(raw_messages)

    # 注入快照
    snap = get_snapshot(session_id)
    snap_text = snapshot_to_prompt(snap)
    if snap_text and messages and messages[0].get("role") == "system":
        messages[0] = dict(messages[0])
        messages[0]["content"] = messages[0]["content"] + "\n\n" + snap_text

    # 压缩
    messages = compress_messages(messages, llm_client, model)

    return messages
