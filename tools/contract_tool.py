"""
=============================================================
  合同智能审查引擎 (Contract Review Engine)
=============================================================
双层审查架构：
  第一层 — 关键词快扫（review_contract）：毫秒级，识别明显缺失
  第二层 — LLM 语义深审（review_contract_deep）：理解条款含义，
           应对中英文合同、否定句式、附件引用、阴阳条款等复杂场景
=============================================================
"""
import re, json, os
from typing import Optional, Tuple

import streamlit as st

# ================================================================
#  15 点审查清单（来自 research_plan_llm_agent.tex）
# ================================================================
CHECKLIST = [
    # (类别, 检查项, 关键词, 加分逻辑说明)
    ("基础信息", "合同主体名称完整", ["甲方", "乙方", "单位名称", "公司"], "双方主体信息是否完整可识别"),
    ("基础信息", "合同编号/日期完整", ["合同编号", "签订日期", "签署日期", "合同号"], "合同唯一标识是否存在"),
    ("金额条款", "金额数字明确", ["金额", "费用", "价款", "总价", "元", "万"], "金额是否以明确数字标注"),
    ("金额条款", "大写金额与小写一致", ["大写", "元整", "零壹贰叁肆伍陆柒捌玖拾佰仟"], "大小写金额是否同时存在且一致"),
    ("金额条款", "付款节点清晰", ["付款", "结算", "分期", "首付", "尾款", "预付款"], "是否有明确的付款时间节点"),
    ("履约条款", "交付标准明确", ["交付", "验收", "合格", "标准", "质量"], "交付物/验收标准是否清晰可衡量"),
    ("履约条款", "工期/服务期明确", ["工期", "期限", "服务期", "起止", "日历天", "工作日"], "履约时间范围是否明确"),
    ("履约条款", "违约责任条款存在", ["违约", "赔偿", "罚则", "违约金", "责任"], "是否有违约责任约定"),
    ("法律条款", "争议解决方式明确", ["争议", "仲裁", "诉讼", "管辖", "法院"], "争议解决机制是否约定"),
    ("法律条款", "保密条款存在", ["保密", "商业秘密", "机密", "不得泄露"], "是否包含保密义务"),
    ("法律条款", "知识产权归属明确", ["知识产权", "著作权", "专利", "商标", "归属"], "知识产权归属是否约定清晰"),
    ("合规审查", "盖章/签字要求明确", ["盖章", "签字", "签章", "公章", "合同专用章"], "生效条件（盖章/签字）是否明确"),
    ("合规审查", "合同份数/有效期", ["份数", "有效期", "生效", "终止", "续约"], "合同份数及有效期是否明确"),
    ("风险条款", "单方解约权条款", ["解除", "终止", "解约", "任意解除"], "是否存在单方任意解除权（高风险）"),
    ("风险条款", "无限责任/兜底条款", ["无限", "连带", "兜底", "全部损失", "间接损失"], "是否有无限责任条款（高风险）"),
]

BILLING_KEYWORDS = {
    "按需付费": ["按量", "按需", "按次", "实际使用", "流量", "调用", "请求数", "后付费", "消费", "弹性"],
    "包年包月": ["包年", "包月", "年付", "月付", "预付", "固定费用", "套餐", "订阅"],
    "混合计费": ["混合", "保底", "阶梯", "封顶", "组合"],
}


def _extract_text_from_upload() -> Optional[str]:
    """从 Streamlit 上传区提取合同文本。优先取 .txt，其次解析 PDF/DOCX。"""
    uf = st.session_state.get("uploaded_file")
    if uf is None:
        return None

    fname = getattr(uf, "name", "")

    if fname.lower().endswith((".txt", ".md")):
        uf.seek(0)
        return uf.read().decode("utf-8", errors="replace")

    if fname.lower().endswith(".pdf"):
        try:
            from tools.file_tool import _read_pdf_text
            return _read_pdf_text(uf)
        except Exception:
            return None

    if fname.lower().endswith(".docx"):
        try:
            import docx
            uf.seek(0)
            doc = docx.Document(uf)
            return "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            return None

    # 兜底：尝试按文本读取
    try:
        uf.seek(0)
        return uf.read().decode("utf-8", errors="replace")
    except Exception:
        return None


def _score_item(text: str, keywords: list) -> int:
    """根据关键词命中数打分（0-2 分）。"""
    if not text:
        return 0
    hits = sum(1 for kw in keywords if kw in text)
    if hits >= 3:
        return 2
    if hits >= 1:
        return 1
    return 0


def _detect_billing_type(text: str) -> str:
    """基于关键词推断计费类型。"""
    scores = {}
    for btype, kws in BILLING_KEYWORDS.items():
        scores[btype] = sum(1 for kw in kws if kw in text)
    if all(v == 0 for v in scores.values()):
        return "无法判断（合同未明确计费方式）"
    return max(scores, key=scores.get)


def review_contract() -> str:
    """
    智能审查上传的合同文件（PDF/DOCX/TXT），基于 15 项审查清单逐项打分，
    识别风险条款，推荐计费类型，输出结构化 Markdown 审查报告。
    """
    text = _extract_text_from_upload()
    if not text:
        return (
            "❌ 未找到可解析的合同文件。请先通过侧边栏上传 PDF、DOCX 或 TXT 格式的合同文件。\n\n"
            "支持的文件类型：\n"
            "- .txt / .md — 纯文本合同\n"
            "- .pdf — PDF 合同（自动提取文字）\n"
            "- .docx — Word 合同"
        )

    # 基础统计
    total_chars = len(text)
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    total_lines = len(lines)

    # ---- 逐项审查 ----
    results = []
    risk_items = []
    score_total = 0
    score_max = len(CHECKLIST) * 2  # 满分 30

    for category, item, keywords, note in CHECKLIST:
        s = _score_item(text, keywords)
        score_total += s
        level = {0: "⚠️ 缺失", 1: "⚡ 部分", 2: "✅ 完整"}[s]
        entry = {"类别": category, "检查项": item, "状态": level, "说明": note}
        results.append(entry)
        if s == 0 and category in ("风险条款", "法律条款"):
            risk_items.append(f"- **{item}**：{note}")

    # ---- 计费类型推断 ----
    billing = _detect_billing_type(text)

    # ---- 基本要素提取 ----
    amount_patterns = [
        r'(?:金额|总价|价款|费用)[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)\s*(?:万|元)',
        r'[¥￥]\s*([\d,]+\.?\d*)\s*(?:万|元)?',
    ]
    amounts = []
    for pat in amount_patterns:
        amounts.extend(re.findall(pat, text))
    amounts = list(set(amounts))[:5]

    # ---- 生成报告 ----
    score_pct = round(score_total / score_max * 100, 1)
    grade = (
        "A 优秀" if score_pct >= 85 else
        "B 良好" if score_pct >= 70 else
        "C 一般" if score_pct >= 55 else
        "D 需重点关注"
    )

    report_lines = [
        "# 📋 合同智能审查报告",
        "",
        f"**合同长度**：{total_chars} 字符 / {total_lines} 行",
        f"**综合评分**：{score_total}/{score_max}（{score_pct}%）— **{grade}**",
        f"**推荐计费类型**：{billing}",
        "",
        "---",
        "",
        "## 一、逐项审查清单",
        "",
        "| 类别 | 检查项 | 状态 | 说明 |",
        "|------|--------|------|------|",
    ]

    for r in results:
        report_lines.append(f"| {r['类别']} | {r['检查项']} | {r['状态']} | {r['说明']} |")

    # 风险汇总
    if risk_items:
        report_lines.append("")
        report_lines.append("---")
        report_lines.append("")
        report_lines.append("## 二、⚠️ 风险事项")
        report_lines.extend(risk_items)
    else:
        report_lines.append("")
        report_lines.append("## 二、✅ 未发现重大风险事项")

    # 金额提取
    if amounts:
        report_lines.append("")
        report_lines.append("---")
        report_lines.append("")
        report_lines.append("## 三、💰 识别到的金额信息")
        for a in amounts:
            report_lines.append(f"- {a}")

    # 计费分析
    report_lines.append("")
    report_lines.append("---")
    report_lines.append("")
    report_lines.append("## 四、📊 计费类型分析")
    report_lines.append(f"**推荐**：{billing}")
    report_lines.append("")
    report_lines.append("**关键词命中**：")
    for btype, kws in BILLING_KEYWORDS.items():
        hits = [kw for kw in kws if kw in text]
        if hits:
            report_lines.append(f"- {btype}：{', '.join(hits[:8])}")
    report_lines.append("")
    report_lines.append("💡 *提示：计费类型推荐基于关键词匹配，最终以合同约定条款为准。*")

    # 审查建议
    report_lines.append("")
    report_lines.append("---")
    report_lines.append("")
    report_lines.append("## 五、📝 审查建议")
    if score_pct >= 85:
        report_lines.append("合同要素较为完整，建议重点关注金额一致性及签章生效条件。")
    elif score_pct >= 70:
        report_lines.append("合同基本要素齐备，仍有部分条款需补充完善。建议重点核对缺失项。")
    elif score_pct >= 55:
        report_lines.append("合同存在较多缺失项，建议在签署前补充关键条款（违约、争议解决等）。")
    else:
        report_lines.append("⚠️ 合同存在重大缺失项，不建议直接签署。请与法务/商务协同修订后再审。")

    return "\n".join(report_lines)


# ================================================================
#  第二层：LLM 语义深审引擎
# ================================================================

_DEEP_REVIEW_SYSTEM_PROMPT = """你是一位资深法务合同审查专家，拥有 15 年以上企业法务经验。
你的审查覆盖中国大陆《民法典》合同编、英美普通法系合同惯例，以及 SaaS/IT 行业合同特有问题。

## 审查原则

### 1. 语义理解而非关键词匹配
- 读懂条款的真实法律含义，而非仅检查"关键词是否出现"
- 识别否定句式："不承担任何赔偿责任" ≠ 有赔偿条款
- 识别条件限定："经双方书面同意后可解除" ≠ 任意解除权

### 2. 多语言支持
- 中英文合同均可审查，英文条款用英文回复
- 中英混杂合同逐条按原文语言审查

### 3. 结构性风险识别
- 引用型缺失：条款写了"详见附件三"但附件未提供 → 标记为"⚠️ 附件缺失"
- 框架型缺失：条款写了"双方另行约定" → 标记为"⚡ 待补充"
- 不对等条款：仅约束一方（如"乙方违约金为合同额30%，甲方违约金为已付款项"）
- 兜底转嫁："其他未尽事宜由乙方负责" → 标记为"🔴 不对等转嫁"

### 4. 输出格式
对每项审查输出：
```
| 序号 | 类别 | 检查项 | 评分 | 风险 | 原文引用 | 分析 |
```
- 评分：2=完整 / 1=部分 / 0=缺失
- 风险：🔴高危 / 🟡注意 / 🟢正常 / ⚠️附件缺失 / ⚡待补充
- 原文引用：合同中相关条款的原文摘录（15字以上），未找到写"未找到"
- 分析：一句话法律判断

然后按风险汇总，最后给出整体评估与行动建议。"""


def _build_checklist_for_llm() -> str:
    """将 15 项审查清单格式化为 LLM 可读的文本。"""
    lines = ["## 审查清单（共 15 项）\n"]
    for i, (category, item, keywords, note) in enumerate(CHECKLIST, 1):
        lines.append(f"{i}. **[{category}] {item}** — {note}")
    return "\n".join(lines)


def _get_llm_client():
    """获取 LLM 客户端，优先使用 session_state 配置。"""
    from openai import OpenAI
    from config import API_KEY, BASE_URL, MODEL
    api_key = st.session_state.get("k", API_KEY)
    base_url = st.session_state.get("u", BASE_URL)
    model = st.session_state.get("m", MODEL)
    return OpenAI(api_key=api_key, base_url=base_url), model


def _smart_contract_truncate(text: str, max_chars: int = 12000) -> Tuple[str, bool]:
    """智能截断超长合同：保留首尾 + 中间扫描关键条款区域。
    返回 (截断后文本, 是否发生了截断)。"""
    if len(text) <= max_chars:
        return text, False

    # 尝试按条款编号拆分
    clause_pattern = re.compile(
        r'(?:第[一二三四五六七八九十百千\d]+[条章节款]|'
        r'^\d+[\.\、\)]|'
        r'^[\(（]\d+[\)）]|'
        r'^[A-Z]\d+[\.\、]|'
        r'^Article\s+\d+|^Section\s+\d+|^Clause\s+\d+)',
        re.MULTILINE | re.IGNORECASE,
    )
    parts = clause_pattern.split(text)
    if len(parts) <= 1:
        # 无法按条款拆分，保留首尾
        head = int(max_chars * 0.55)
        tail = int(max_chars * 0.30)
        return text[:head] + f"\n\n…(中间省略 {len(text) - head - tail} 字符)…\n\n" + text[-tail:], True

    # 均匀采样：保留首部条款 + 关键条款 + 尾部条款
    n = len(parts)
    key_indices = {0, 1}  # 开头
    mid = n // 2
    key_indices.update({mid - 1, mid, mid + 1})  # 中间
    key_indices.update({n - 2, n - 1})  # 结尾
    key_indices = {i for i in key_indices if 0 <= i < n}

    # 在关键部位搜索风险相关条款
    risk_keywords = ["违约", "责任", "赔偿", "解除", "终止", "保密", "知识产权",
                     "争议", "仲裁", "诉讼", "liability", "termination", "indemnity",
                     "confidential", "intellectual property", "dispute"]
    for i, part in enumerate(parts):
        if any(kw in part for kw in risk_keywords):
            key_indices.add(i)

    kept = []
    total = 0
    for i in sorted(key_indices):
        snippet = parts[i][:1500]  # 每段最多 1500 字符
        if total + len(snippet) > max_chars:
            break
        kept.append(snippet)
        total += len(snippet)

    truncated = "\n\n---[条款省略]---\n\n".join(kept)
    return truncated[:max_chars], True


def _check_structural_issues(text: str) -> list:
    """检测合同的结构性问题（无需 LLM）。"""
    issues = []

    # 引用附件但未提供
    attachment_refs = re.findall(
        r'(?:详见|参见|见|参照|如|附件|附录|appendix|exhibit|attachment|schedule)',
        text, re.IGNORECASE,
    )
    if len(attachment_refs) >= 3:
        issues.append("⚠️ 合同多处引用附件/附录，请确认附件是否已提供。若未提供，相关条款无法完成审查。")

    # 大量"另行约定"类兜底
    pending_refs = re.findall(r'另行(?:约定|协商|签署|通知)|另行(?:agree|negotiate)', text)
    if len(pending_refs) >= 3:
        issues.append(f"⚡ 发现 {len(pending_refs)} 处\'另行约定\'类条款，关键事项未在本文本中明确。")

    # 单方约束检测（中文）
    if re.search(r'乙方.*(?:不得|禁止|无权|应当|必须)', text) and \
       not re.search(r'甲方.*(?:不得|禁止|无权|应当|必须)', text):
        issues.append("🔴 疑似单方约束：合同大量义务条款仅约束乙方，未约束甲方。")

    # 纯英文合同
    if len(re.findall(r'[a-zA-Z]', text)) > len(text) * 0.7:
        issues.append("💡 检测到英文合同，已切换到英文审查模式。")

    return issues


def review_contract_deep() -> str:
    """
    LLM 驱动的深度合同审查（双层：快扫 + 语义深审）。

    适用场景：
    - 结构各异的合同（框架协议、SaaS、采购、NDA 等）
    - 中英文合同
    - 含否定句式/引用条款/不对等条款的复杂合同
    - 超长合同（自动智能截断）
    """
    text = _extract_text_from_upload()
    if not text:
        return (
            "❌ 未找到可解析的合同文件。请先通过侧边栏上传 PDF、DOCX 或 TXT 格式的合同文件。\n\n"
            "支持的文件类型：\n"
            "- .txt / .md — 纯文本合同\n"
            "- .pdf — PDF 合同（自动提取文字）\n"
            "- .docx — Word 合同"
        )

    # ── 第一层：关键词快扫 ──
    quick_report = review_contract()

    # ── 第二层：LLM 语义深审 ──
    # 智能截断超长合同
    truncated_text, was_truncated = _smart_contract_truncate(text)

    # 结构性预检
    structural_issues = _check_structural_issues(text)

    # 构建审查提示词
    checklist_text = _build_checklist_for_llm()
    truncation_note = (
        f"\n\n⚠️ **注意**：合同原文约 {len(text):,} 字符，"
        f"已智能截断为 {len(truncated_text):,} 字符（保留了首尾及关键条款区域）。"
        f"请在分析中标注哪些判断可能因截断而不完整。"
    ) if was_truncated else ""

    user_prompt = f"""请对以下合同进行逐项语义审查。

{checklist_text}

---

## 结构性预检结果（代码自动检测）
{chr(10).join(f'- {i}' for i in structural_issues) if structural_issues else '- 未发现明显结构性问题'}

---

## 合同原文
{truncated_text}
{truncation_note}

---

## 审查要求

请按以下格式输出审查报告：

### 一、逐项审查表

| # | 类别 | 检查项 | 评分 | 风险 | 原文引用 | 法律分析 |
|---|------|--------|------|------|----------|----------|
| 1 | ... | ... | 2/1/0 | 🔴/🟡/🟢 | "..." | 一句话 |

### 二、风险汇总

按风险等级从高到低排列，每项包含：
- 风险点描述
- 潜在后果
- 修改建议

### 三、结构性风险（如有）

### 四、计费类型建议

### 五、整体评估与行动建议

- 综合评分（满分 30）
- 是否建议签署（可签署 / 有条件签署 / 不建议签署）
- 优先修改清单（Top 3）"""

    # 调用 LLM
    try:
        client, model = _get_llm_client()
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": _DEEP_REVIEW_SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.3,  # 低温度确保一致性
            stream=False,
        )
        deep_report = resp.choices[0].message.content or ""
    except Exception as e:
        deep_report = (
            f"\n\n---\n⚠️ **LLM 深审失败**：{e}\n"
            f"以下仅展示关键词快扫结果。\n---\n"
        )

    # ── 合并报告 ──
    header = (
        "# 📋 合同智能审查报告（双层审查）\n\n"
        f"> 合同长度：{len(text):,} 字符\n"
        f"> 审查模式：关键词快扫 + LLM 语义深审\n"
        f"> 结构性预检：{'发现 ' + str(len(structural_issues)) + ' 个问题' if structural_issues else '通过'}\n"
        + (f"> ⚠️ 合同过长已智能截断（{len(text):,} → {len(truncated_text):,} 字符），深层分析可能不完整\n" if was_truncated else "")
        + "\n---\n\n"
    )

    separator = "\n\n---\n\n## 🔬 附：关键词快扫结果（仅供参考）\n\n"

    return header + deep_report + separator + quick_report
