"""
=============================================================
  公司内部智能 Agent 开发框架
=============================================================
【目标】打造公司内部通用 Agent，聚焦日常事务处理
【方案】OpenAI SDK + Streamlit + Function Calling

【架构】
  用户输入 -> Streamlit UI -> Agent循环 -> 工具层
  工具层: read_excel / send_wecom / query_db (待扩展)

【路线图】
  第1周: 基础对话 + 流式输出
  第2周: Excel读取工具
  第3周: 企业微信通知
  第4周: 部署内网，全员使用

【安装依赖】
  pip install openai streamlit pandas openpyxl tabulate

【运行】
  streamlit run agent.py
=============================================================
"""

import concurrent.futures
import json, os, re
from pathlib import Path
import pandas as pd
import streamlit as st
from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from config import API_KEY, BASE_URL, MODEL, MODEL_PRICING
from tools.excel_tool import (
    read_excel, summarize_excel, filter_excel,
    calc_excel, write_excel, plot_excel, generate_report,
    pivot_excel, top_n_excel,
    get_analysis_data, build_report_doc, build_report_pdf, build_report,
    detect_anomalies, reconcile_sheets,
    _set_cached_file_data, _clear_cached_file_data, _get_cached_file_data,
)
from tools.file_tool import (
    inspect_uploaded_file, list_pdf_info, read_pdf_pages,
    search_pdf, extract_pdf_toc,
)
from tools.contract_tool import review_contract, review_contract_deep
from user_skills_store import (
    create_user_skill as _create_user_skill,
    delete_user_skill as _delete_user_skill,
    list_user_skills as _list_user_skills,
)
from tools.dynamic_tool import (
    create_tool as _create_tool,
    delete_tool as _delete_tool,
    list_dynamic_tools as _list_dynamic_tools,
    get_dynamic_tool_definitions,
    execute_dynamic_tool,
    detect_tool_gap,
)
from skills import SKILLS, detect_skill, DEFAULT_SKILL, reload_skills, is_user_skill
from session_store import (
    init_db, get_or_create_session, load_messages, save_messages,
    get_snapshot, save_snapshot, auto_extract_snapshot,
    snapshot_to_prompt, compress_messages, list_sessions, delete_session,
    get_session_name, set_session_name, set_session_topic,
)
from rag_store import (
    index_file as _rag_index_file,
    search_knowledge_base as _rag_search,
    clear_knowledge_base as _rag_clear,
    get_indexed_files as _rag_list_files,
)

# ── 1. 工具定义（告诉LLM有哪些工具） ─────────────
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "inspect_uploaded_file",
            "description": "通用文件读取/检查工具。支持Excel/CSV、文本、JSON、PDF、Word、PPT、HTML表格、图片、ZIP等。用户上传非表格文件或不确定文件类型时优先调用。",
            "parameters": {
                "type": "object",
                "properties": {
                    "max_chars": {"type": "integer", "description": "文本最多返回字符数，默认6000"},
                    "max_pages": {"type": "integer", "description": "PDF最多读取页数，默认5"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_pdf_info",
            "description": "返回 PDF 元信息：总页数、标题、作者、目录结构等。长文档分析第一步——了解文档规模与结构，决定后续用 search_pdf 定位还是 read_pdf_pages 分段阅读。",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_pdf",
            "description": "在 PDF 中搜索关键词，返回匹配页面及上下文（默认前后各 250 字符）。用于长文档快速定位关键条款、金额、人名、日期等。找到匹配后可用 read_pdf_pages 展开阅读。",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "搜索关键词（中文/英文均可）"},
                    "context_chars": {"type": "integer", "description": "每处匹配返回的上下文字符数，默认 500"},
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_pdf_pages",
            "description": "读取 PDF 指定页码范围。用于长文档分段精读——先用 list_pdf_info 或 search_pdf 定位目标页码范围，再逐段读取。",
            "parameters": {
                "type": "object",
                "properties": {
                    "start_page": {"type": "integer", "description": "起始页码（从 1 开始），默认 1"},
                    "end_page": {"type": "integer", "description": "结束页码（含），默认等于 start_page，即只读一页"},
                    "max_chars": {"type": "integer", "description": "最大返回字符数，默认 6000"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "extract_pdf_toc",
            "description": "提取 PDF 内嵌目录/书签结构。如果 PDF 有书签可快速了解章节分布，比 list_pdf_info 更详细地展示层级结构。",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_excel",
            "description": "读取上传的Excel或CSV文件的行，返回Markdown表格。用于查看数据内容。支持skiprows跳过前N行，用于读取文件中部或底部。",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string",  "description": "工作表名，默认'0'表示第一个sheet"},
                    "nrows":      {"type": "integer", "description": "读取行数，默认50"},
                    "skiprows":   {"type": "integer", "description": "跳过前N行，默认0（从头读），设为50读第51行起"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "summarize_excel",
            "description": "返回Excel文件的结构概览：行列数、每列数据类型、空值情况、数值列统计量。用于快速了解数据。",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "filter_excel",
            "description": "按条件筛选Excel数据行。支持大于、小于、等于、包含等操作。",
            "parameters": {
                "type": "object",
                "properties": {
                    "column":     {"type": "string", "description": "要筛选的列名"},
                    "operator":   {"type": "string", "description": "运算符：> < == != contains"},
                    "value":      {"type": "string", "description": "筛选值"},
                    "sheet_name": {"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": ["column", "operator", "value"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calc_excel",
            "description": "对Excel某列做统计计算，如求和、平均值、最大值、最小值、计数、中位数、标准差。",
            "parameters": {
                "type": "object",
                "properties": {
                    "column":    {"type": "string", "description": "要计算的列名"},
                    "operation": {"type": "string", "description": "操作：sum / mean / max / min / count / median / std"},
                    "sheet_name":{"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": ["column", "operation"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "write_excel",
            "description": "修改Excel中指定行列的单元格值，保存后可下载新文件。",
            "parameters": {
                "type": "object",
                "properties": {
                    "row_index":  {"type": "integer", "description": "行索引，从0开始"},
                    "column":     {"type": "string",  "description": "列名"},
                    "new_value":  {"type": "string",  "description": "新的单元格值"},
                    "sheet_name": {"type": "string",  "description": "工作表名，默认'0'"},
                },
                "required": ["row_index", "column", "new_value"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "plot_excel",
            "description": "生成基础图表（柱状图、折线图、饼图、散点图）。如需更复杂的图表类型（如箱线图、热力图、双轴图、堆积图等），请用 create_tool 创建自定义图表工具（沙箱已内置 plt/save_chart）。",
            "parameters": {
                "type": "object",
                "properties": {
                    "x_column":   {"type": "string", "description": "X轴列名"},
                    "y_column":   {"type": "string", "description": "Y轴列名"},
                    "chart_type": {"type": "string", "description": "图表类型：bar / line / pie / scatter，默认bar"},
                    "sheet_name": {"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": ["x_column", "y_column"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_report",
            "description": "⚠️ 已废弃！自动生成固定模板报告。请改用完整的分析流程：get_analysis_data → pivot/top_n/calc → plot_excel/create_tool 绘图 → 撰写报告 → build_report。",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "pivot_excel",
            "description": "按维度列分组聚合指标列，返回 Markdown 表格。用于分析各类别的总量/均值/计数等。例如：按省份汇总收入、按产品类型统计数量。",
            "parameters": {
                "type": "object",
                "properties": {
                    "group_column":  {"type": "string",  "description": "分组维度列名（如'省份'、'产品类型'）"},
                    "value_column":  {"type": "string",  "description": "聚合指标列名（如'收入'、'数量'）"},
                    "agg_op":        {"type": "string",  "description": "聚合方式：sum/mean/count/max/min，默认sum"},
                    "sheet_name":    {"type": "string",  "description": "工作表名，默认'0'"},
                    "top_n":         {"type": "integer", "description": "返回前N组，默认20"},
                },
                "required": ["group_column", "value_column"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "top_n_excel",
            "description": "按指定列排序，返回数值最大（或最小）的前N行。用于找出TOP贡献者、最高/最低异常值等。",
            "parameters": {
                "type": "object",
                "properties": {
                    "sort_column": {"type": "string",  "description": "排序依据的列名"},
                    "n":           {"type": "integer", "description": "返回行数，默认10"},
                    "ascending":   {"type": "boolean", "description": "true=升序(最小)，false=降序(最大)，默认false"},
                    "sheet_name":  {"type": "string",  "description": "工作表名，默认'0'"},
                },
                "required": ["sort_column"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_analysis_data",
            "description": "返回数据集的完整结构信息（字段列表、数据类型、空值情况、数值统计、维度列、指标列、样本行）。分析任何数据集时第一步必须调用此工具，取代 summarize_excel 作为标准入口。LLM 读完结果后自行决定分析方向。",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string", "description": "工作表名，默认'0'"},
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "build_report_doc",
            "description": "将 LLM 撰写的完整 Markdown 报告文字与图表打包为 Word (.docx) 文件供下载。LLM 负责写内容，此工具只做排版打包。在完成所有分析并写好报告文字后调用。",
            "parameters": {
                "type": "object",
                "properties": {
                    "report_markdown": {"type": "string", "description": "LLM 撰写的完整报告 Markdown 文本，含标题、发现、建议等所有内容"},
                },
                "required": ["report_markdown"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "build_report_pdf",
            "description": "将 LLM 撰写的完整 Markdown 报告文字与图表编译为 PDF 文件（通过 LaTeX/xelatex）供下载。LLM 负责写内容，此工具只做排版打包。在完成所有分析并写好报告文字后调用（可与 build_report_doc 同时调用）。",
            "parameters": {
                "type": "object",
                "properties": {
                    "report_markdown": {"type": "string", "description": "LLM 撰写的完整报告 Markdown 文本，含标题、发现、建议等所有内容"},
                },
                "required": ["report_markdown"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "build_report",
            "description": "一键同时生成 Word (.docx) 和 PDF 两种格式的报告，自动读取侧边栏的格式偏好和表格/图表开关。LLM 负责写内容，此工具负责排版打包和双格式输出。在完成所有分析并写好报告文字后调用（推荐优先使用此工具替代单独调用 build_report_doc 或 build_report_pdf）。",
            "parameters": {
                "type": "object",
                "properties": {
                    "report_markdown": {"type": "string", "description": "LLM 撰写的完整报告 Markdown 文本，含标题、发现、建议等所有内容"},
                },
                "required": ["report_markdown"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "detect_anomalies",
            "description": "对指定数值列进行异常值检测，支持 IQR（四分位距法）、Z-Score、Isolation Forest 三种方法。用于发现数据中的离群点。",
            "parameters": {
                "type": "object",
                "properties": {
                    "column":        {"type": "string",  "description": "要检测的数值列名"},
                    "method":        {"type": "string",  "description": "检测方法：auto（自动全用）/ iqr / zscore / isolation_forest，默认auto"},
                    "sheet_name":    {"type": "string",  "description": "工作表名，默认'0'"},
                    "threshold":     {"type": "number",  "description": "Z-Score 阈值，默认 2.0"},
                    "contamination": {"type": "number",  "description": "Isolation Forest 异常比例，默认 0.05"},
                },
                "required": ["column"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "reconcile_sheets",
            "description": "对两个 Sheet 进行多口径对账：按关键列匹配后比对数值列差异，找出仅单方存在的记录和数值不一致的记录。",
            "parameters": {
                "type": "object",
                "properties": {
                    "key_column":      {"type": "string",  "description": "两Sheet之间匹配用的关键列名（如'计费ID'、'合同号'）"},
                    "compare_columns": {"type": "array",   "items": {"type": "string"}, "description": "要比对的数值列名列表，留空则自动选公有数值列"},
                    "sheet_a":         {"type": "string",  "description": "基准Sheet名称，默认'0'"},
                    "sheet_b":         {"type": "string",  "description": "对比Sheet名称，默认'1'"},
                    "tolerance":       {"type": "number",  "description": "差异容忍度（绝对值），小于此值不报告，默认 0.01"},
                },
                "required": ["key_column"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "review_contract",
            "description": "智能审查上传的合同文件（PDF/DOCX/TXT），基于15项审查清单逐项打分（基础信息/金额/履约/法律/风险），识别风险条款，推荐计费类型（按需/包年包月/混合），输出结构化 Markdown 审查报告。上传合同后即可调用，无需额外参数。",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "review_contract_deep",
            "description": "【推荐】LLM 驱动的深度合同语义审查。先关键词快扫定位明显缺失，再调 LLM 逐项理解条款真实法律含义。能应对：中英文合同、否定句式（'不承担'）、引用型缺失（'详见附件'）、不对等条款（仅约束一方）、框架型兜底（'另行约定'）。输出详细逐项审查表（原文引用+法律分析+风险评级）。比 review_contract 更准确但更慢（需调 LLM）。",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_tool",
            "description": "【元工具】运行时动态创建新 Function。当现有工具无法满足需求时（如 plot_excel 不支持复杂图表），用此工具在线造一个新工具。创建后立即可用。沙箱已内置 pd/np/plt/save_chart(fig,title,description)/io 等模块。",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "新工具名称（英文标识符，如 custom_bar_chart、draw_waterfall、fetch_data）"},
                    "description": {"type": "string", "description": "工具用途描述（告知 LLM 何时调用、做什么），例如：'生成自定义横向柱状图对比各产品线收入'"},
                    "parameters": {"type": "string", "description": "JSON Schema 格式的参数定义（JSON 字符串），包含 type/properties/required 字段"},
                    "code": {"type": "string", "description": "Python 函数体代码（不含 def 行），接收 kwargs 参数，必须 return 字符串结果。沙箱已内置 pd/np/plt/io/save_chart(fig,title,description)，可直接使用。"},
                },
                "required": ["name", "description", "parameters", "code"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_tool",
            "description": "删除一个之前通过 create_tool 动态创建的工具。",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要删除的工具名称"},
                },
                "required": ["name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_dynamic_tools",
            "description": "列出所有已动态创建的工具及其描述和参数。用于查看当前有哪些自定义工具可用。",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_user_skills",
            "description": "列出所有用户自定义的 Skill（技能/工作流）。用户可通过 create_user_skill 创建专属的分析角色。",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_user_skill",
            "description": "创建一个用户自定义 Skill（技能/角色）。定义一个新角色的触发关键词、系统提示词和可用工具白名单。创建后立即可用，后续对话输入关键词即可自动激活。",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Skill 唯一标识（英文key），如 finance_review、hr_screening"},
                    "display_name": {"type": "string", "description": "显示名称（中文），如'财务对账'、'简历筛选'"},
                    "keywords": {"type": "string", "description": "触发关键词，逗号分隔。如'对账,账单,发票,应收账款'"},
                    "system_prompt": {"type": "string", "description": "系统提示词，定义角色、工作流、输出规范"},
                    "allowed_tools": {"type": "string", "description": "允许使用的工具名，逗号分隔。如'reconcile_sheets,read_excel,calc_excel'。留空表示使用全部可用工具"},
                    "forbidden_tools": {"type": "string", "description": "禁止使用的工具名，逗号分隔。如'write_excel,plot_excel'"},
                    "entry_tool_for_table": {"type": "string", "description": "表格文件的第一步入口工具，如'reconcile_sheets'"},
                    "entry_tool_for_generic": {"type": "string", "description": "通用文件的第一步入口工具，如'inspect_uploaded_file'"},
                },
                "required": ["name", "display_name", "keywords", "system_prompt"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_user_skill",
            "description": "删除一个用户自定义的 Skill。只能删除用户自己创建的，不能删除内置 Skill。",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "要删除的 Skill 标识（英文key）"},
                },
                "required": ["name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_knowledge_base",
            "description": "语义检索已上传并索引过的文档内容。当用户询问'之前那份合同里的XX条款'、'文档里有没有提到YY'、'根据上传的资料回答ZZ'等需要查找文档内容的问题时调用。返回最相关的文档片段及其来源。",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "自然语言检索问题，如'违约责任条款'、'2024年销售目标'"},
                    "top_k": {"type": "integer", "description": "返回结果数量，默认5"},
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_knowledge_base",
            "description": "查看已索引的知识库文件列表，了解有哪些文档可供检索。",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }
    },
]

# ── 动态工具处理函数（必须在 TOOL_REGISTRY 之前定义） ──
def _handle_create_tool(name: str, description: str, parameters: str, code: str) -> str:
    """create_tool 的注册表处理函数。"""
    return _create_tool(name=name, description=description, parameters=parameters, code=code)


def _handle_delete_tool(name: str) -> str:
    """delete_tool 的注册表处理函数。"""
    return _delete_tool(name=name)


def _handle_list_dynamic_tools() -> str:
    """list_dynamic_tools 的注册表处理函数。"""
    return _list_dynamic_tools()


def _handle_list_user_skills() -> str:
    """list_user_skills 的注册表处理函数。"""
    skills = _list_user_skills()
    if not skills:
        return "📭 当前没有用户自定义的 Skill。\n\n💡 你可以通过侧边栏「自定义技能」创建，或让我用 create_user_skill 帮你创建。"
    lines = ["📦 **用户自定义 Skill**：\n"]
    for s in skills:
        lines.append(f"| `{s['key']}` | {s['name']} | 触发词: {', '.join(s['keywords'])} | 工具: {', '.join(s['allowed_tools']) or '全部'} |")
    return "\n".join(lines)


def _handle_create_user_skill(name: str, display_name: str, keywords: str, system_prompt: str,
                               allowed_tools: str = "", forbidden_tools: str = "",
                               entry_tool_for_table: str = "", entry_tool_for_generic: str = "") -> str:
    """create_user_skill 的注册表处理函数。"""
    result = _create_user_skill(
        name=name, display_name=display_name, keywords=keywords,
        system_prompt=system_prompt, allowed_tools=allowed_tools,
        forbidden_tools=forbidden_tools,
        entry_tool_for_table=entry_tool_for_table,
        entry_tool_for_generic=entry_tool_for_generic,
    )
    reload_skills()
    return result


def _handle_delete_user_skill(name: str) -> str:
    """delete_user_skill 的注册表处理函数。"""
    result = _delete_user_skill(name=name)
    reload_skills()
    return result


def _handle_rag_search(query: str, top_k: int = 5) -> str:
    """search_knowledge_base 的注册表处理函数（会话隔离）。"""
    sid = st.session_state.get("rag_session_id", "")
    results = _rag_search(query, session_id=sid, top_k=top_k)
    if not results:
        return "知识库中未找到相关内容。请先上传文件（系统会自动索引），或尝试更换检索词。"
    lines = [f"🔍 检索「{query}」— 共 {len(results)} 条结果：\n"]
    for i, r in enumerate(results):
        lines.append(
            f"**[{i + 1}]** 📄 {r['filename']}（片段 {r['chunk_index']}，相似度 {r['score']:.2f}）\n"
            f"{r['content'][:800]}\n"
        )
    return "\n".join(lines)


def _handle_rag_list() -> str:
    """list_knowledge_base 的注册表处理函数（会话隔离）。"""
    sid = st.session_state.get("rag_session_id", "")
    files = _rag_list_files(session_id=sid)
    if not files:
        return "知识库为空。上传文件后系统会自动索引。"
    lines = ["📚 已索引文件："]
    for f in files:
        lines.append(f"- {f['filename']}（{f['chunks']} 个片段）")
    return "\n".join(lines)


def _ai_generate_skill(user_description: str) -> dict | None:
    """调用 LLM 根据用户自然语言描述生成 Skill 配置。
    返回 {"key":..., "name":..., "keywords":..., "prompt":..., "tools":...} 或 None。"""
    available_tools = [t["function"]["name"] for t in TOOLS]

    prompt = f"""你是一个 Skill 配置生成器。根据用户的需求描述，生成一个完整的 Skill 配置。

## 可用工具列表
{', '.join(available_tools)}

## 重要提示
- `create_tool` 是元工具，**不要**列在 tools 字段中（它始终可用）
- 但 **必须** 在 prompt 的末尾追加一句提醒，告知该角色：当现有工具不够用时可使用 `create_tool` 创建自定义工具
- tools 字段只填该角色最常用的 3-5 个业务工具

## 用户需求
{user_description}

## 输出要求
返回一个严格的 JSON 对象（不要 Markdown 代码块，只要纯 JSON），包含以下字段：
{{
    "key": "英文标识（如 finance_review，小写+下划线）",
    "name": "中文名称（如 财务对账，不超过8字）",
    "keywords": "触发关键词，逗号分隔（如 对账,账单,发票,核对）",
    "prompt": "系统提示词，定义角色和工作流。\\n要求：\\n1. 第一句定义角色\\n2. 列出3-5步工作流\\n3. 说明输出格式\\n4. 末尾加一句：当内置工具不满足需求时，可用 create_tool 在线创建自定义工具\\n5. 100-300字",
    "tools": "推荐使用的工具名，逗号分隔（从可用工具列表中选择最相关的3-5个，不要包含create_tool）"
}}

只返回 JSON，不要任何解释。"""

    try:
        from openai import OpenAI
        from config import API_KEY, BASE_URL, MODEL
        api_key = st.session_state.get("k", API_KEY)
        base_url = st.session_state.get("u", BASE_URL)
        model = st.session_state.get("m", MODEL)
        client = OpenAI(api_key=api_key, base_url=base_url)

        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            response_format={"type": "json_object"},
        )
        text = resp.choices[0].message.content or ""

        # 清理可能的 Markdown 代码块 + 多余前缀
        text = text.strip()
        if text.startswith("```"):
            text = re.sub(r'^```(?:json)?\s*', '', text)
            text = re.sub(r'\s*```$', '', text)
        # 从文本中提取 JSON 对象
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            text = match.group()

        config = json.loads(text)
        required = ["key", "name", "keywords", "prompt"]
        if all(k in config for k in required):
            return {
                "key": config["key"],
                "name": config["name"],
                "keywords": config["keywords"],
                "prompt": config["prompt"],
                "tools": config.get("tools", ""),
            }
        return None
    except Exception:
        return None


# ── 2. 工具注册表 ─────────────────────────────────
TOOL_REGISTRY = {
    "inspect_uploaded_file": inspect_uploaded_file,
    "list_pdf_info":     list_pdf_info,
    "search_pdf":        search_pdf,
    "read_pdf_pages":    read_pdf_pages,
    "extract_pdf_toc":   extract_pdf_toc,
    "read_excel":        read_excel,
    "summarize_excel":   summarize_excel,
    "filter_excel":      filter_excel,
    "calc_excel":        calc_excel,
    "write_excel":       write_excel,
    "plot_excel":        plot_excel,
    "generate_report":   generate_report,
    "pivot_excel":       pivot_excel,
    "top_n_excel":       top_n_excel,
    "get_analysis_data": get_analysis_data,
    "build_report_doc":  build_report_doc,
    "build_report_pdf":  build_report_pdf,
    "build_report":      build_report,
    "detect_anomalies":  detect_anomalies,
    "reconcile_sheets":  reconcile_sheets,
    "review_contract":   review_contract,
    "review_contract_deep": review_contract_deep,
    "create_tool":       _handle_create_tool,
    "delete_tool":       _handle_delete_tool,
    "list_dynamic_tools": _handle_list_dynamic_tools,
    "list_user_skills":  _handle_list_user_skills,
    "create_user_skill": _handle_create_user_skill,
    "delete_user_skill": _handle_delete_user_skill,
    "search_knowledge_base": _handle_rag_search,
    "list_knowledge_base":   _handle_rag_list,
}

BRIEF_TOOL_NAMES = {
    "inspect_uploaded_file",
    "read_excel",
    "filter_excel",
    "calc_excel",
    "plot_excel",
    "pivot_excel",
    "top_n_excel",
    "get_analysis_data",
}

EXPORT_KEYWORDS = (
    "pdf", "word", "docx", "导出", "下载", "生成报告", "分析报告", "latex",
)

def _parse_chart_count(user_msg: str) -> int:
    """从用户消息中解析要求的图表数量。"""
    text = user_msg
    cn_map = {'一': 1, '两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6}
    m = re.search(r'([一两二三四五六\d]+)\s*张\s*图', text)
    if m:
        n = m.group(1)
        return cn_map.get(n, int(n) if n.isdigit() else 2)
    # 提到图/图表但没说几张，默认2张
    if re.search(r'图表|画图|作图|出图|可视化|chart|plot', text, re.IGNORECASE):
        return 2
    return 0

TABLE_FILE_SUFFIXES = (".xlsx", ".xls", ".csv")
EVIDENCE_TOOL_NAMES = {
    "read_excel",
    "filter_excel",
    "calc_excel",
    "pivot_excel",
    "top_n_excel",
    "inspect_uploaded_file",
    "list_pdf_info",
    "search_pdf",
    "read_pdf_pages",
    "extract_pdf_toc",
    "get_analysis_data",
    "detect_anomalies",
    "reconcile_sheets",
    "review_contract",
    "review_contract_deep",
}

def dispatch_tool(name, args):
    fn = TOOL_REGISTRY.get(name)
    if fn:
        return fn(**args)
    # 检查动态工具
    dyn_result = execute_dynamic_tool(name, args)
    if not dyn_result.startswith("❌ 动态工具"):
        return dyn_result
    return f"未知工具：{name}"


# ── Token 计数 & 成本追踪 ───────────────────────
def _init_usage_tracker():
    """初始化当前会话的用量追踪器。"""
    import streamlit as st
    if "usage" not in st.session_state:
        st.session_state.usage = {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0,
            "cost_usd": 0.0,
            "calls": 0,
        }


def _track_usage(usage_obj, model: str = None):
    """记录一次 LLM 调用的 token 用量。usage_obj 为 openai.types.CompletionUsage。"""
    import streamlit as st
    _init_usage_tracker()
    u = st.session_state.usage
    pt = getattr(usage_obj, "prompt_tokens", 0) or 0
    ct = getattr(usage_obj, "completion_tokens", 0) or 0
    tt = getattr(usage_obj, "total_tokens", 0) or 0
    u["prompt_tokens"] += pt
    u["completion_tokens"] += ct
    u["total_tokens"] += tt
    u["calls"] += 1

    # 计算成本
    m = model or st.session_state.get("m", MODEL)
    pricing = MODEL_PRICING.get(m)
    if pricing:
        cost = (pt / 1_000_000) * pricing["input"] + (ct / 1_000_000) * pricing["output"]
        u["cost_usd"] += cost


def _render_usage_sidebar():
    """在侧边栏渲染用量面板。"""
    import streamlit as st
    _init_usage_tracker()
    u = st.session_state.usage
    st.caption("📊 **本次会话用量**")
    if u["calls"] == 0:
        st.caption("⏳ 等待首次 LLM 调用…")
        return
    c1, c2 = st.columns(2)
    c1.metric("LLM 调用次数", u["calls"])
    c2.metric("Token 总数", f"{u['total_tokens']:,}")
    c1.metric("输入 Token", f"{u['prompt_tokens']:,}")
    c2.metric("输出 Token", f"{u['completion_tokens']:,}")
    if u["cost_usd"] > 0:
        st.caption(f"💰 估算成本：**${u['cost_usd']:.4f}**")
    with st.expander("🔍 详情"):
        st.write(f"总 Token：{u['total_tokens']:,}")
        st.write(f"输入：{u['prompt_tokens']:,} ｜ 输出：{u['completion_tokens']:,}")
        st.write(f"API 调用：{u['calls']} 次")
        st.write(f"模型：{st.session_state.get('m', MODEL)}")


def _show_usage_inline():
    """在聊天区域底部内联显示用量（Agent 完成后可见）。"""
    import streamlit as st
    _init_usage_tracker()
    u = st.session_state.usage
    if u["calls"] == 0:
        return
    model_name = st.session_state.get("m", MODEL)
    lines = [
        f"📊 累计用量：{u['total_tokens']:,} tokens（输入 {u['prompt_tokens']:,} / 输出 {u['completion_tokens']:,}）"
    ]
    if u["cost_usd"] > 0:
        lines.append(f"💰 估算成本：${u['cost_usd']:.4f} ｜ API 调用 {u['calls']} 次 ｜ 模型 {model_name}")
    else:
        lines.append(f"🔧 API 调用 {u['calls']} 次 ｜ 模型 {model_name}")
    st.caption("  \n".join(lines))


def normalize_message(msg):
    if isinstance(msg, dict):
        return msg
    if hasattr(msg, "model_dump"):
        return msg.model_dump(exclude_none=True)
    role = getattr(msg, "role", None)
    content = getattr(msg, "content", None)
    if role:
        normalized = {"role": role}
        if content is not None:
            normalized["content"] = content
        return normalized
    return {"role": "assistant", "content": str(msg)}

# ── 3. Agent循环（核心） ──────────────────────────
# 原理：LLM不直接执行工具，只返回"我要调用哪个工具+参数"
# 你的代码执行工具，把结果喂回LLM，循环直到LLM给出最终回答

SYSTEM_PROMPT = """你是一位资深数据分析师，对标麦肯锡/BCG 咨询报告的分析深度与表达标准。

## 分析深度三阶框架

每个发现必须穿透三层，缺一不可：
1. **是什么**：数据事实，必须带具体数字和对比基准。
   例："Q3 收入 1,240 万，环比下降 12%，为近 6 个季度首次下滑"（不是"收入有所下降"）。
2. **为什么**：归因推断，给出 1-2 个有依据的假设，不用等确证才写。
   例："与 9 月华东区大客户集中到期未续约高度相关，该区贡献了 Q2 收入的 35%"。
3. **所以呢**：业务含义 + 可执行建议，落到具体动作和时间尺度。
   例："若趋势延续 Q4 缺口约 150 万。建议本周内由销售总监带队逐一拜访 TOP5 到期客户。"

以下是对比示例，展示"平庸回答"和"深度回答"的差距——每个回答都是你的前车之鉴：

【平庸】"本月销售额为 500 万元，较上月有所增长。"
【深度】"8 月销售额 536 万元，环比增长 8.7%（7 月 493 万），创年内单月新高。增长主因是华南区新签 3 个政府云项目集中开票（合计 62 万），剔除该因素后自然增长仅 2.3%。这意味着当前增长对个别大单依赖度过高——建议：①将政府云成单经验标准化为行业方案加速复制；②重点监测 Q4 是否有等量级新单填补缺口。"

【平庸】"利润下降，建议加强成本控制。"
【深度】"Q2 净利润率从 18.3% 降至 14.7%，主因是服务器采购成本同比上升 23%（英伟达 H200 集群扩容），而人均产出仅增 5%。即使算力采购是战略性投入，人效增速远不足以消化成本增速。建议：①冻结非核心岗位招聘至年底；②将算力成本按 BU 用量内部核算分摊，倒逼各 BU 优化模型调用效率。"

【平庸】"应收账款有所增加，需要关注回款。"
【深度】"截至 9 月底应收账款 4,280 万元，同比增 37%，其中账龄 > 90 天的占比已从年初的 8% 升至 19%。前两大客户 A 公司和 B 集团合计欠款 1,560 万，均超 120 天。按当前坏账计提比例，若这 1,560 万无法回收，将直接侵蚀全年净利润约 8%。建议：①A 公司启动法务催收并暂停新项目供货；②B 集团由 VP 级以上出面沟通付款计划；③财务部本月内完成全量应收账龄排查并输出红黄绿灯清单。"

## 分析工作流

**第一步：摸底**
- 表格文件必须先调用 get_analysis_data，其他文件先调用 inspect_uploaded_file
- 读懂数据结构：哪些是维度列（分类分组），哪些是指标列（数值可计算），数据质量如何

**第二步：立题**
- 用一句话说清核心业务问题。例："这份销售数据的关键问题是：哪些客户/产品贡献了主要利润，利润结构是否可持续？"——而不是"分析这份数据"。

**第三步：取证**
- 优先顺序：pivot_excel（分组洞察）> top_n_excel（排名异常）> calc_excel（汇总验证）
- 每次调工具前想清：这个工具结果将如何支撑或否定我的假设？
- 通常 2-4 个关键工具调用即可形成结论，不必穷举所有统计

**第四步：大纲规划（写作前必须执行）**

在动笔之前，先输出报告大纲。大纲格式如下：

```
## 📋 报告大纲

**报告标题**：[一句话概括核心结论]
**报告结构**：[诊断式/叙事式/对比式/预警式/机会式/金字塔式/时间线式]

### 大纲
1. **[章节标题]** — 核心观点一句话
   - 证据来源：[哪个工具的结果]
   - 是否需要图表：是/否 → [图表类型 + 数据维度]
2. **[章节标题]** — 核心观点一句话
   - 证据来源：[哪个工具的结果]
   - 是否需要图表：是/否 → [图表类型 + 数据维度]
3. **[章节标题]** — 核心观点一句话
   - 证据来源：[哪个工具的结果]
   - 是否需要图表：是/否 → [图表类型 + 数据维度]
...
N. **总结与建议** — 行动清单
```

大纲规划要点：
- 每个章节必须有明确的证据来源（不能凭空写）
- 判断每个章节是否需要图表支撑，如果需要，标注图表类型
- 先完成所有需要的图表（调用 plot_excel 或 create_tool），再开始写正文
- 大纲章节数根据数据复杂度灵活调整，2-6 个均可

**第五步：逐段扩写**

按大纲顺序，逐段展开写作。每段遵循以下结构：

```
### [章节编号]. [章节标题]

[核心判断句——一句话给出结论]

[展开论述 3-6 句，按"是什么→为什么→所以呢"三阶框架]
- 是什么：数据事实，带具体数字和对比基准
- 为什么：归因推断，给出有依据的假设
- 所以呢：业务含义 + 可执行建议

[如有图表，在此处插入图表引用，如"见图1"]
```

扩写要点：
- 每段一个核心观点，不要一段塞多个发现
- 宁可写少写深，2-4 个核心发现即可
- 图表紧跟在支撑它的段落后面，不要集中堆在文末
- 段落之间用空行分隔，保持呼吸感

**第六步：排版校对**

写完正文后，按以下排版规范检查并修正：

### 排版规范（必须严格遵守）

**标题层级**：
- 报告总标题用 `#`（一级标题）
- 章节标题用 `##`（二级标题），格式：`## 一、[标题]` 或 `## 1. [标题]`
- 子章节用 `###`（三级标题），格式：`### 1.1 [标题]`
- 禁止跳级（如一级直接到三级）

**段落格式**：
- 每个段落 3-6 句，不超过 8 句
- 段落之间必须有空行
- 禁止一整段超过 200 字不分行

**数字与表格**：
- 关键数字用 **加粗** 突出
- 对比数据用表格呈现，不要堆在段落里
- 表格必须有标题行，列对齐
- 表格前后各留一个空行

**图表位置**：
- 图表紧跟在引用它的段落之后
- 每张图必须有标题和一句话说明
- 图表标题格式：`图N：[具体业务含义]`

**列表与要点**：
- 行动建议用编号列表（1. 2. 3.）
- 并列要素用无序列表（- ）
- 列表项之间不要空行，列表前后各留一个空行

**禁止事项**：
- 禁止大段文字不分行
- 禁止图表和文字之间没有空行
- 禁止标题和正文之间没有空行
- 禁止连续三个以上短段落（合并或展开）
- 禁止在段落中间插入图表引用（图表应独立成行）

## 写作规范
- 数字必须有对比基准和变化幅度："环比增长 8.7%"而非"有所增长"
- 主动句式："Q3 收入下滑 12%"而非"可以看出 Q3 收入有所下降"
- 禁用套话：不写"从数据可以看出"、"综上所述"、"值得注意的是"、"建议进一步分析"
- 每个发现必须给出可执行的行动建议，不能停在"应该关注"

## 图表规范
- 每张图必须有明确的结论指向，能支撑一个发现
- 标题要带具体业务含义，如"各省份月度销售额趋势——华南区持续领跑"
- **用户要求 N 张图时必须实际调用 plot_excel 或动态创建工具 N 次，禁止用文字"建议出图方向"代替真实图表**

## 报告多样性（关键！每次报告必须不同）
每次生成的报告在结构、重点、图表类型上都必须不同，不允许套用固定模板。

### 报告结构多样化示例
根据数据特征选择不同结构，以下供参考（不限于此）：
- **诊断式**：先给整体健康度评分 → 分维度展开问题 → 根本原因 → 行动方案
- **叙事式**：以一个核心业务问题开场 → 逐层深入排查 → 高潮发现 → 收束建议
- **对比式**：AB对照 → 差异分析 → 归因 → 哪方做法值得推广
- **预警式**：先抛风险信号 → 量化影响 → 传导路径 → 止损方案
- **机会式**：先找增长亮点 → 放大分析 → 可复制条件 → 规模化建议
- **金字塔式**：顶部1个核心结论 → 3个支撑论据 → 每个论据带数据和图表
- **时间线式**：按时间顺序展现变化 → 关键拐点标注 → 每个拐点归因 → 预测下一阶段

### 图表类型多样化
禁止总是用简单的柱状图。根据数据特征和叙事需要选择：
- 趋势类：折线图、面积图、堆积面积图
- 对比类：柱状图、分组柱状图、横向柱状图、雷达图
- 构成类：饼图、环形图、堆积柱状图、瀑布图
- 关系类：散点图、气泡图、热力图
- 分布类：箱线图、直方图、密度图
- 高级：双轴图（柱+线组合）、子弹图、帕累托图

当 `plot_excel` 不支持你想要的图表类型时，**必须用 `create_tool` 创建自定义图表工具**：
1. 调用 `create_tool` 创建新工具，code 中使用 `plt` (matplotlib) 绘图
2. 用 `save_chart(fig, "标题", "描述")` 将图表存入报告
3. 然后调用该工具生成图表
4. 沙箱已内置 pd/np/plt/save_chart/io 等模块

### 报告篇幅多样化
- 有时2-3个深度发现就够了（穿透式）
- 有时需要5-7个发现覆盖全局（全景式）
- 不要每次都输出同样的章节数量

## 导出规则
- 只有用户明确要求"导出"/"下载"/"生成报告"/"Word"/"PDF"时才调用 build_report_doc / build_report_pdf
- 导出前必须已完成完整分析正文和所有图表
- **严禁调用 generate_report**（它是固定模板，不符合多样性要求）

## 动态工具创建（元能力）
- 当现有工具（如 plot_excel 只有4种基础图）无法满足需求时，使用 `create_tool` 创建新工具
- 创建前先用 `list_dynamic_tools` 检查是否已有类似工具
- 创建自定义图表工具时，沙箱已内置：`pd`(pandas)、`np`(numpy)、`plt`(matplotlib)、`io`、`save_chart(fig, title, description)`
- `save_chart` 会自动将图表存入报告，被 build_report_doc/build_report_pdf 打包
- 创建后立即在同一轮对话中调用
- 新工具命名规则：动词_对象（如 custom_bar_chart、draw_waterfall）
- 示例场景：plot_excel 不支持横向柱状图 → 调用 create_tool 创建一个 custom_horizontal_bar 工具"""
def _wants_export_report(user_msg: str) -> bool:
    text = (user_msg or "").lower()
    return any(keyword in text for keyword in EXPORT_KEYWORDS)


def _get_active_skill(user_msg: str) -> dict:
    """根据用户消息匹配 Skill，返回 skill 配置字典。
    优先检查手动覆盖，其次自动关键词匹配。"""
    # 检查手动覆盖
    override = st.session_state.get("_skill_override", "__auto__")
    if override != "__auto__" and override in SKILLS:
        skill = SKILLS[override]
    else:
        skill_key = detect_skill(user_msg)
        skill = SKILLS.get(skill_key, SKILLS[DEFAULT_SKILL])

    # 如果用户明确要求导出，允许覆盖 forbid 中的导出工具
    if _wants_export_report(user_msg):
        skill = dict(skill)  # 浅拷贝
        skill["allowed_tools"] = set(skill["allowed_tools"]) | {"build_report_doc", "build_report_pdf", "build_report"}
        skill["forbidden_tools"] = set(skill.get("forbidden_tools", set())) - {"build_report_doc", "build_report_pdf", "build_report"}
    return skill


TABLE_ONLY_TOOLS = {
    "read_excel", "summarize_excel", "filter_excel",
    "calc_excel", "write_excel", "plot_excel",
    "generate_report", "pivot_excel", "top_n_excel",
    "get_analysis_data",
}


def _active_tools_for(skill: dict):
    """根据 Skill 配置过滤可用工具白名单，非表格文件自动排除 Excel 专属工具。
    元工具（create_tool / delete_tool / list_dynamic_tools）始终可用，
    不受 allowed_tools / forbidden_tools 限制——因为它们是 LLM 扩展自身能力的生命线。"""
    allowed = set(skill.get("allowed_tools", set()))
    forbidden = set(skill.get("forbidden_tools", set()))
    # 非表格文件排除 Excel 专属工具
    if _uploaded_file_kind() != "table":
        forbidden = forbidden | TABLE_ONLY_TOOLS
    # 元工具：永远不被拦截（LLM 造工具的能力不可剥夺）
    meta_tools = {"create_tool", "delete_tool", "list_dynamic_tools"}
    # 从 forbidden 中移除元工具，确保它们无论如何都可用
    forbidden = forbidden - meta_tools
    base = [
        tool for tool in TOOLS
        if (tool["function"]["name"] in allowed or tool["function"]["name"] in meta_tools)
        and tool["function"]["name"] not in forbidden
    ]
    # 追加所有已注册的动态工具
    base.extend(get_dynamic_tool_definitions())
    return base


def _mode_instruction(skill: dict) -> str:
    """当前 Skill 模式说明，告知 LLM 当前激活的分析角色。"""
    name = skill.get("name", "数据分析")
    return f"🧠 当前技能：{name}"


def _uploaded_file_kind():
    uploaded = st.session_state.get("uploaded_file")
    if not uploaded or not hasattr(uploaded, "name"):
        return None
    suffix = os.path.splitext(uploaded.name)[1].lower()
    if suffix in TABLE_FILE_SUFFIXES:
        return "table"
    return "generic"


def _entry_tool_for_current_file(skill: dict):
    """根据 Skill 配置和当前文件类型返回建议的入口工具。"""
    kind = _uploaded_file_kind()
    if kind == "table":
        return skill.get("entry_tool_for_table", "get_analysis_data")
    if kind == "generic":
        return skill.get("entry_tool_for_generic", "inspect_uploaded_file")
    return None


def _runtime_instruction(required_entry_tool: str, preflight_done: bool, evidence_calls: int,
                         plot_calls: int, export_mode: bool,
                         tool_calls_used: int = 0, max_tool_calls: int = 14,
                         reserved_plot_slots: int = 0,
                         messages: list = None) -> str:
    """仅在最关键的 2-3 个转折点注入温和引导，避免打断 LLM 的自然推理流。"""
    notes = []
    remaining = max_tool_calls - tool_calls_used
    plots_needed = max(0, reserved_plot_slots - plot_calls)

    # 转折点 0：检测工具缺口（LLM 表达了无法完成某任务）
    if messages:
        last_asst = ""
        for m in reversed(messages):
            if m.get("role") == "assistant" and m.get("content"):
                last_asst = str(m["content"]).lower()
                break
        gap_signals = [
            "无法", "不支持", "没有.*工具", "缺少.*功能", "无法满足",
            "can't", "cannot", "unable to", "not supported",
            "no tool", "missing tool", "no function",
        ]
        if any(re.search(p, last_asst) for p in gap_signals):
            notes.append(
                "💡 **检测到工具缺口**：当前工具似乎无法满足需求。你可以使用 `create_tool` 在线创建新工具！"
                "只需指定名称、描述、参数 schema 和 Python 实现代码（支持 requests/json/re/math 等安全模块），"
                "创建后立即可在当前对话中调用。"
            )

    # 转折点 1：必须完成摸底
    if required_entry_tool and not preflight_done:
        notes.append(f"💡 第一步建议先用 `{required_entry_tool}` 了解数据结构。")

    # 转折点 2：证据充分 + 图表齐备 → 先出大纲再写作
    elif evidence_calls >= 2 and plot_calls >= reserved_plot_slots:
        notes.append(
            "💡 证据和图表已齐备。请按以下流程写作：\n"
            "① **先输出报告大纲**：列出每个章节标题、核心观点、证据来源、是否需要图表。\n"
            "② **按大纲逐段扩写**：每段一个核心观点，按'是什么→为什么→所以呢'展开。\n"
            "③ **排版校对**：检查标题层级、段落间距、数字加粗、图表位置。\n"
            "宁可写 2-3 个深度发现，不要面面俱到但浮于表面。"
        )

    # 转折点 3：证据充分但图表未达标
    elif evidence_calls >= 2 and plots_needed > 0 and remaining <= plots_needed * 2:
        notes.append(f"💡 证据已充分，还需 {plots_needed} 张图。建议优先完成剩余图表，然后进入深度写作。")

    # 转折点 4：资源即将耗尽提醒
    elif remaining <= 2 and evidence_calls >= 1:
        notes.append("💡 工具预算即将用尽，请基于现有证据进入深度写作阶段。")

    # 转折点 5：证据过度积累提醒
    elif evidence_calls >= 5:
        notes.append("💡 证据已非常充分，建议转入作图和写作阶段。")

    if export_mode:
        fmt_parts = []
        if st.session_state.get("export_word", True):
            fmt_parts.append("Word")
        if st.session_state.get("export_pdf", True):
            fmt_parts.append("PDF")
        fmt_str = "+".join(fmt_parts) if fmt_parts else "Word+PDF"
        tbl_note = "不含表格" if not st.session_state.get("include_tables", True) else ""
        fig_note = "不含图表" if not st.session_state.get("include_figures", True) else ""
        extra = "、".join(filter(None, [tbl_note, fig_note]))
        notes.append(f"📄 导出模式：完整写好正文后调用 build_report（推荐）或单独调用 build_report_doc/build_report_pdf。输出格式：{fmt_str}。" + (f" 注意：{extra}。" if extra else ""))

    return "\n".join(notes)


def _smart_truncate(result: str, max_chars: int = 6000) -> str:
    """智能压缩工具返回结果：保留结构而非粗暴切断。
    - 含 Markdown 表格时：保留文字说明 + 表头 + 前 18 行 + 后 5 行
    - 纯文本时：保留前 65% 和后 12%，中间省略"""
    if len(result) <= max_chars:
        return result

    lines = result.split('\n')

    # 检测是否含 Markdown 表格
    table_start = -1
    for i, line in enumerate(lines):
        if '|' in line and line.strip().startswith('|'):
            table_start = i
            break

    if table_start >= 0 and table_start < len(lines):
        header_lines = lines[:table_start]
        table_lines = lines[table_start:]
        n_table = len(table_lines)
        if n_table > 25:
            kept = (
                header_lines +
                table_lines[:3] +
                [f'| ... | *(共 {n_table - 8} 行，此处省略)* | ... |'] +
                table_lines[3:18] +
                [f'| ... | *(尾部省略，原共 {n_table} 行)* | ... |'] +
                table_lines[-5:]
            )
            compressed = '\n'.join(kept)
            if len(compressed) > max_chars:
                compressed = compressed[:max_chars] + f"\n\n*(已压缩，原文 {len(lines)} 行)*"
            return compressed

    # 纯文本：保留首尾
    head = int(max_chars * 0.65)
    tail = int(max_chars * 0.12)
    omitted = len(result) - head - tail
    return result[:head] + f"\n\n…(中间省略 {omitted} 字符)…\n\n" + result[-tail:]


def _self_critique(client, model, messages: list, draft: str) -> str | None:
    """对初稿进行自我审查并输出修订版（Priority 4: Self-Critique Multi-Pass）。

    审查五项：数据溯源 / 三阶穿透 / 套话清理 / 建议可行性 / 空泛判断。
    若原稿已优秀则微调后输出，不重复审查过程。
    """
    critique_prompt = """你是一位严苛的编辑。请审查以下分析报告初稿，逐项检查并直接输出修订版：

## 审查清单（逐项打勾，不允许跳过）
1. □ 每个发现是否穿透三层（是什么→为什么→所以呢）？
2. □ 每个数字是否有对比基准和变化幅度（环比/同比/占比）？
3. □ 是否出现套话（"从数据可以看出"、"综上所述"、"值得注意的是"、"建议进一步分析"）？
4. □ 每个建议是否可执行（具体动作 + 时间 + 负责人）？
5. □ 是否有无证据支撑的空泛判断（如"表现良好"、"需要关注"）？

## 排版审查（同样重要！）
6. □ 标题层级是否正确（# → ## → ###，无跳级）？
7. □ 段落之间是否有空行分隔？
8. □ 关键数字是否用 **加粗** 突出？
9. □ 图表是否紧跟在引用它的段落之后（而非堆在文末）？
10. □ 表格前后是否有空行？表格是否有标题行？
11. □ 是否存在大段文字不分行（>200字）的情况？
12. □ 列表格式是否统一（编号列表/无序列表）？

## 修订要求
- 直接输出修订后的完整报告，不要输出审查过程或打勾结果
- 保留原有结构和核心发现，只强化薄弱环节
- 删掉所有套话和空泛判断
- 为缺对比基准的数字补上可推算的对比基准
- **修正所有排版问题**：标题层级、段落间距、数字加粗、图表位置、表格格式
- 如果原报告已经很好，微调后直接输出
- 保持 Markdown 格式"""

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=messages + [
                {"role": "assistant", "content": draft},
                {"role": "user", "content": critique_prompt},
            ],
            stream=False,
        )
        return resp.choices[0].message.content
    except Exception:
        return None  # 审查失败则回退原稿


# ── LLM 调用重试包装 ────────────────────────────
def _call_llm_stream(client, model: str, messages: list, tools: list, max_retries: int = 3):
    """
    带自动重试的 LLM 流式调用。
    仅在以下瞬时错误时重试：超时、限流、连接中断、服务端 500。
    参数错误（400）、权限错误（401/403）等不重试，直接抛出。
    """
    from openai import (
        APITimeoutError, RateLimitError, APIConnectionError,
        InternalServerError,
    )
    RETRYABLE = (APITimeoutError, RateLimitError, APIConnectionError, InternalServerError)

    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return client.chat.completions.create(
                model=model,
                messages=messages,
                tools=tools,
                tool_choice="auto",
                stream=True,
                stream_options={"include_usage": True},
            )
        except RETRYABLE as e:
            last_exc = e
            if attempt < max_retries:
                wait = 2 ** (attempt - 1)  # 1s → 2s → 4s
                import time
                time.sleep(wait)
        # 不可重试的错误直接抛出
    raise last_exc


def run_agent(user_msg: str, session_id: str = None):
    st.session_state.messages = [normalize_message(msg) for msg in st.session_state.messages]
    st.session_state["generated_charts"] = []
    st.session_state["reasoning_steps"] = []
    st.session_state.pop("last_plot", None)
    st.session_state.pop("last_plot_title", None)

    # ── 预读文件字节到模块级缓存（跨线程可见，解决 threading.local 隔离问题）──
    uploaded = st.session_state.get("uploaded_file")
    if uploaded is not None:
        try:
            if hasattr(uploaded, "getvalue"):
                _set_cached_file_data(uploaded.getvalue(), uploaded.name)
            else:
                uploaded.seek(0)
                _set_cached_file_data(uploaded.read(), uploaded.name)
                uploaded.seek(0)
        except Exception:
            import sys, traceback
            print("[文件预读失败]", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
            _clear_cached_file_data()
    else:
        _clear_cached_file_data()

    # ── 检测并激活 Skill ────────────────────
    skill = _get_active_skill(user_msg)
    st.session_state["active_skill"] = skill.get("name", "数据分析")

    # 动态更新 system prompt，告知 LLM 当前文件状态
    f = st.session_state.get("uploaded_file")
    cached_bytes, cached_name = _get_cached_file_data()
    if f and cached_bytes:
        file_info = (
            f"✅ 当前已上传文件：{cached_name}（{len(cached_bytes):,} bytes 已预读缓存）。"
            f"可直接调用 get_analysis_data / read_excel / inspect_uploaded_file 等工具分析。"
        )
    elif f:
        file_info = (
            f"⚠️ 当前已上传文件：{f.name}，但预读缓存未命中。"
            f"调用工具时会自动重试读取。"
        )
    else:
        file_info = "当前未上传文件。如需分析数据，请先在左侧上传 Excel/CSV 文件。"
    mode_note = _mode_instruction(skill)
    # ── 规划层 + 结构化输出协议 ──
    planning_note = """
## 🎯 执行规划协议

**重要：你的工具列表中有 `create_tool`——无论当前 Skill 限制了多少工具，你永远可以用它在线造新工具！**
- 需要饼图但只有柱状图？→ `create_tool` 造一个饼图工具
- 需要雷达图/瀑布图/双轴图？→ `create_tool` 造
- 需要调外部 API 取数据？→ `create_tool` 造
- 造完立刻就能用，不要因为现有工具不够而妥协质量

**每次工具调用前**在心中明确：
1. 本轮要验证什么假设 / 获取什么信息？
2. 如果多个工具互不依赖，**一次并行调用多个工具**以节省轮次。
3. 工具返回后，下一步做什么？

**报告写作流程**（重要！严格按以下三阶段执行）：

### 阶段一：分析取证
- 完成全部数据分析和取证（get_analysis_data → pivot/top_n/calc → 异常检测等）
- 根据数据特征规划报告结构：诊断式 / 叙事式 / 对比式 / 预警式 / 机会式 / 金字塔式 / 时间线式

### 阶段二：大纲规划
- 取证完成后，**先输出报告大纲**（不要直接写正文！）
- 大纲包含：每个章节标题、核心观点一句话、证据来源、是否需要图表及图表类型
- 根据大纲中需要的图表，调用 plot_excel 或 create_tool 生成所有图表
- 图表全部生成完毕后，再进入阶段三

### 阶段三：逐段扩写 + 排版
- 按大纲顺序逐段展开，每段一个核心观点
- 每段按"是什么→为什么→所以呢"三阶框架组织
- 写完后按排版规范校对：标题层级、段落间距、数字加粗、图表位置、表格格式
- 最后调用 `build_report` / `build_report_doc` / `build_report_pdf` 导出
"""
    base_prompt = skill["system_prompt"] + "\n\n" + planning_note + "\n" + mode_note + "\n" + file_info

    # ── 注入分析快照（让 LLM 知道分析到哪了）──
    if session_id:
        snap = get_snapshot(session_id)
        snap_text = snapshot_to_prompt(snap)
        if snap_text:
            base_prompt += "\n\n" + snap_text

    # ── 注入 RAG 知识库上下文 ──
    rag_files = _rag_list_files(session_id=session_id)
    if rag_files:
        rag_note = "\n\n## 📚 知识库（已索引文档）\n"
        rag_note += "以下文档已向量化索引，可以用 `search_knowledge_base` 语义检索其中内容：\n"
        for f in rag_files:
            rag_note += f"- {f['filename']}（{f['chunks']} 个片段）\n"
        rag_note += "用户询问文档相关问题时，**请主动调用 `search_knowledge_base`** 检索相关片段后再回答。"
        base_prompt += rag_note

    st.session_state.messages[0] = {"role": "system", "content": base_prompt}
    st.session_state.messages.append({"role": "user", "content": user_msg})

    export_mode = _wants_export_report(user_msg)
    active_tools = _active_tools_for(skill)
    required_entry_tool = _entry_tool_for_current_file(skill)
    preflight_done = required_entry_tool is None
    reserved_plot_slots = _parse_chart_count(user_msg)
    max_rounds = 24 if export_mode else 12
    max_tool_calls = 36 if export_mode else (14 + reserved_plot_slots)
    max_plot_calls = 8 if export_mode else max(6, reserved_plot_slots)
    tool_calls_used = 0
    plot_calls_used = 0
    evidence_calls_used = 0

    for _round in range(max_rounds):
        api_key = st.session_state.get("k", API_KEY)
        base_url = st.session_state.get("u", BASE_URL)
        model = st.session_state.get("m", MODEL)
        if not api_key:
            text = "请先在侧边栏填写 API Key，或设置 OPENAI_API_KEY 环境变量。"
            st.session_state.messages.append({"role": "assistant", "content": text})
            for char in text:
                yield char
            return
        llm_client = OpenAI(api_key=api_key, base_url=base_url)

        # ── 上下文压缩：消息太长时自动压缩旧消息 ──
        if session_id:
            st.session_state.messages = compress_messages(
                st.session_state.messages, llm_client, model
            )

        runtime_note = _runtime_instruction(
            required_entry_tool=required_entry_tool,
            preflight_done=preflight_done,
            evidence_calls=evidence_calls_used,
            plot_calls=plot_calls_used,
            export_mode=export_mode,
            tool_calls_used=tool_calls_used,
            max_tool_calls=max_tool_calls,
            reserved_plot_slots=reserved_plot_slots,
            messages=st.session_state.messages,
        )
        call_messages = st.session_state.messages + ([{"role": "system", "content": runtime_note}] if runtime_note else [])

        # 仅在图表严重落后时温和提示模型优先选图，不再强制 tool_choice
        plots_still_needed = max(0, reserved_plot_slots - plot_calls_used)
        remaining_before_call = max_tool_calls - tool_calls_used
        if plots_still_needed > 0 and remaining_before_call <= plots_still_needed * 2:
            call_messages = call_messages + [{"role": "system", "content": f"💡 还需 {plots_still_needed} 张图，预算紧张，建议本轮优先调用 plot_excel。"}]
        tool_choice_param = "auto"

        resp = _call_llm_stream(llm_client, model, call_messages, active_tools)

        # ── 流式累积：边收 token 边记录，同时检测 tool_calls ──
        text_parts = []
        tool_call_chunks = {}  # index → {id, name, args}
        reasoning_parts = []

        for chunk in resp:
            delta = chunk.choices[0].delta

            if delta.content:
                text_parts.append(delta.content)
                yield delta.content  # ← 真流式，token 级别实时输出

            # 推理内容（DeepSeek 等推理模型）
            r = getattr(delta, "reasoning_content", None)
            if r:
                reasoning_parts.append(r)

            # 工具调用增量（多 chunk 拼接）
            for tc_delta in (delta.tool_calls or []):
                idx = tc_delta.index
                if idx not in tool_call_chunks:
                    tool_call_chunks[idx] = {"id": "", "name": "", "args": ""}
                d = tool_call_chunks[idx]
                if tc_delta.id:
                    d["id"] = tc_delta.id
                if tc_delta.function:
                    if tc_delta.function.name:
                        d["name"] += tc_delta.function.name
                    if tc_delta.function.arguments:
                        d["args"] += tc_delta.function.arguments

            # 最终 chunk 带 usage
            if chunk.usage:
                _track_usage(chunk.usage, model)

        full_text = "".join(text_parts).strip()
        reasoning_text = "".join(reasoning_parts).strip()

        # ── 重建 tool_calls 对象（兼容后续 .id / .function.name / .function.arguments 访问）──
        class _TC:
            def __init__(self, d):
                self.id = d["id"]
                self.function = type("_F", (), {"name": d["name"], "arguments": d["args"]})()

        tool_calls = [_TC(tool_call_chunks[i]) for i in sorted(tool_call_chunks.keys())]

        if tool_calls:
            # ── 工具调用分支 ──
            if reasoning_text:
                st.session_state.setdefault("reasoning_steps", []).append({
                    "round": _round + 1, "content": reasoning_text,
                })
            tool_names = " → ".join(tc.function.name for tc in tool_calls)
            yield f"\n\n> 🔍 **调用工具**：`{tool_names}`\n\n"
            asst_msg = {
                "role": "assistant",
                "content": full_text,
                "tool_calls": [
                    {"id": tc.id, "type": "function", "function": {"name": tc.function.name, "arguments": tc.function.arguments}}
                    for tc in tool_calls
                ],
            }
            if reasoning_text:
                asst_msg["reasoning_content"] = reasoning_text
            st.session_state.messages.append(asst_msg)

            preflight_completed_this_round = False

            # ── 并行工具调度（Priority 2）──
            dispatch_items = []  # (tc, tool_name, args)
            for tc in tool_calls:
                tool_name = tc.function.name
                if not preflight_done and tool_name != required_entry_tool:
                    st.session_state.messages.append({
                        "role": "tool", "tool_call_id": tc.id,
                        "content": f"💡 请先用 `{required_entry_tool}` 了解数据结构，再决定分析方向。",
                    })
                elif tool_calls_used >= max_tool_calls:
                    st.session_state.messages.append({
                        "role": "tool", "tool_call_id": tc.id,
                        "content": "工具预算已用尽。请基于现有全部证据，进入深度写作阶段——按三阶框架组织每个发现。",
                    })
                elif tool_name == "plot_excel" and plot_calls_used >= max_plot_calls:
                    st.session_state.messages.append({
                        "role": "tool", "tool_call_id": tc.id,
                        "content": "图表配额已满。请基于已有图表和证据，直接输出深度分析。",
                    })
                elif tool_name in {"build_report_doc", "build_report_pdf", "build_report"} and evidence_calls_used == 0:
                    st.session_state.messages.append({
                        "role": "tool", "tool_call_id": tc.id,
                        "content": "请先完成摸底并形成至少 1-2 个关键证据，再导出报告。",
                    })
                else:
                    try:
                        args = json.loads(tc.function.arguments or "{}")
                    except json.JSONDecodeError as e:
                        st.session_state.messages.append({
                            "role": "tool", "tool_call_id": tc.id,
                            "content": f"工具参数解析失败：{e}",
                        })
                    else:
                        dispatch_items.append((tc, tool_name, args))

            # 第二遍：并行执行所有可调度工具
            if dispatch_items:
                with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(dispatch_items))) as executor:
                    future_map = {}
                    for tc, tool_name, args in dispatch_items:
                        future = executor.submit(dispatch_tool, tool_name, args)
                        future_map[future] = (tc, tool_name)

                    for future in concurrent.futures.as_completed(future_map):
                        tc, tool_name = future_map[future]
                        try:
                            result = future.result()
                        except Exception as e:
                            result = f"工具执行异常：{e}"

                        if tool_name == "plot_excel":
                            plot_calls_used += 1
                        else:
                            tool_calls_used += 1
                        if tool_name == required_entry_tool and not preflight_done:
                            preflight_done = True
                            preflight_completed_this_round = True
                            if isinstance(result, str):
                                result += "\n\n请先基于这份摸底结果确定主问题，再决定下一步是否需要补充证据。"
                        if tool_name in EVIDENCE_TOOL_NAMES:
                            evidence_calls_used += 1
                        if isinstance(result, str) and len(result) > 6000:
                            result = _smart_truncate(result, max_chars=6000)

                        st.session_state.messages.append({
                            "role": "tool",
                            "tool_call_id": tc.id,
                            "content": result,
                        })
        else:
            # ── 文本回答分支（最终输出）──
            text = full_text
            if reasoning_text:
                st.session_state.setdefault("reasoning_steps", []).append({
                    "round": _round + 1, "content": reasoning_text,
                })

            # ── 自 Critique 多轮生成（Priority 4）──
            if evidence_calls_used >= 2 and len(text) > 200:
                yield "\n\n> 🔍 **自审查中**：逐项核验数据溯源、逻辑完整性、建议可行性…\n\n"
                revised = _self_critique(llm_client, model, st.session_state.messages, text)
                if revised and len(revised) > len(text) * 0.5:
                    text = revised
                    yield "> ✅ 审查完成，已输出修订版。\n\n"

            asst_msg = {"role": "assistant", "content": text}
            if reasoning_text:
                asst_msg["reasoning_content"] = reasoning_text
            st.session_state.messages.append(asst_msg)

            # 文本已在流式循环中逐 token 输出，此处仅记录
            # 但自审查后的修订版文本尚未输出，需补输出增量
            if text != full_text:
                yield text

            return

    # 超出最大轮次
    warn = f"\n\n⚠️ 已达最大推理轮次（{max_rounds}轮），请重新提问或拆分任务。"
    st.session_state.messages.append({"role": "assistant", "content": warn})
    for char in warn:
        yield char

# ── 4. 内联渲染辅助 ───────────────────────────────
_CHART_REF_PAT = re.compile(
    r'图表\s*[一二三四五六七八九十\d]|如图|下图|图\s*[一二三四五六七八九十\d]|chart\s*\d'
    r'|以下图表|如下图|见图|参见图|见下图',
    re.IGNORECASE,
)
_FINDING_PAT = re.compile(
    r'^#{1,3}\s*(发现|核心发现|关键发现|结论|图表|分析|趋势|对比)',
    re.MULTILINE,
)

def _render_text_with_charts(text: str, charts: list):
    """按段落渲染文字，遇到图表引用或发现章节标题时插入图片。
    若文字完全没有图表参考标记，则将图表均匀分散在各个发现章节后面。"""
    if not charts:
        st.markdown(text)
        return

    sections = re.split(r'(\n\n+)', text)
    # 合并段落与分隔符
    paragraphs = []
    buf = ""
    for part in sections:
        if re.match(r'\n\n+', part):
            if buf.strip():
                paragraphs.append(buf.strip())
            buf = ""
        else:
            buf += part
    if buf.strip():
        paragraphs.append(buf.strip())

    # 判断文字中是否有明确图表引用
    has_explicit_refs = any(_CHART_REF_PAT.search(p) for p in paragraphs)

    chart_idx = 0
    finding_idx = 0  # 已遇到的"发现/结论"标题数
    n_findings = len(_FINDING_PAT.findall(text)) or 1
    # 每隔几个发现章节放一张图
    charts_per_finding = max(1, len(charts))

    for para in paragraphs:
        if not para:
            continue
        st.markdown(para)

        if chart_idx >= len(charts):
            continue

        if has_explicit_refs:
            # 只在段落含图表引用时插入
            if _CHART_REF_PAT.search(para):
                c = charts[chart_idx]
                st.image(c["image"], caption=c["title"], use_container_width=True)
                if c.get("description"):
                    st.caption(c["description"])
                chart_idx += 1
        else:
            # 无明确引用：在每个"发现/结论"标题段落后插入一张图
            if _FINDING_PAT.search(para) or re.match(r'^#{1,3}\s', para):
                finding_idx += 1
                # 按比例分配图表
                target = round(finding_idx / n_findings * len(charts))
                while chart_idx < target and chart_idx < len(charts):
                    c = charts[chart_idx]
                    st.image(c["image"], caption=c["title"], use_container_width=True)
                    if c.get("description"):
                        st.caption(c["description"])
                    chart_idx += 1

    # 剩余图表追加末尾
    for c in charts[chart_idx:]:
        st.image(c["image"], caption=c["title"], use_container_width=True)
        if c.get("description"):
            st.caption(c["description"])

# ── 5. Streamlit UI ───────────────────────────────
def main():
    st.set_page_config(page_title="公司智能助手", page_icon="🤖", layout="wide")
    st.title("🤖 公司智能助手")
    st.caption("OpenAI SDK + Streamlit | 数据留在本地")

    # ── 会话持久化：从 SQLite 恢复对话 ──
    init_db()
    sid = get_or_create_session()
    st.session_state["rag_session_id"] = sid  # RAG 按会话隔离

    if "messages" not in st.session_state:
        loaded = load_messages(sid)
        if loaded:
            st.session_state.messages = loaded
        else:
            st.session_state.messages = [{"role": "system", "content": SKILLS[DEFAULT_SKILL]["system_prompt"]}]
    if "uploaded_file" not in st.session_state:
        st.session_state.uploaded_file = None

    with st.sidebar:
        st.header("📁 上传文件")
        up = st.file_uploader(
            "上传文件",
            type=[
                "xlsx", "xls", "csv",
                "txt", "md", "log", "json", "xml", "html", "htm",
                "pdf", "docx", "pptx",
                "png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff",
                "zip",
            ],
        )
        if up:
            previous_name = getattr(st.session_state.get("uploaded_file"), "name", None)
            if previous_name != up.name:
                for key in ("generated_charts", "generated_report_docx", "generated_report_name", "generated_report_pdf", "generated_report_pdf_name", "last_plot", "last_plot_title", "uploaded_image_preview", "edited_excel", "edited_excel_name"):
                    st.session_state.pop(key, None)
                _clear_cached_file_data()
            st.session_state.uploaded_file = up
            st.success(f"已上传：{up.name}")

            # ── 自动向量化索引（RAG 知识库，按会话隔离，后台线程不阻塞）──
            try:
                data = up.read()
                up.seek(0)
                sid_rag = st.session_state.get("rag_session_id", "")
                import threading
                threading.Thread(
                    target=lambda d, n, s: _rag_index_file(d, n, session_id=s),
                    args=(data, up.name, sid_rag),
                    daemon=True,
                ).start()
                st.info(f"🧠 正在后台索引「{up.name}」…")
            except Exception:
                pass

        # 缓存状态指示（检查模块级缓存）
        cached_bytes, cached_name = _get_cached_file_data()
        if cached_name:
            st.caption(f"📦 缓存就绪：{cached_name} ({len(cached_bytes):,} bytes)")
        st.divider()

        # ── 技能选择器 ──
        st.header("🧠 当前技能")
        current_skill_name = st.session_state.get("active_skill", "数据分析")
        # 构建所有可用 Skill 的选项
        all_skill_options = {"🤖 自动检测": "__auto__"}
        for key, cfg in SKILLS.items():
            label = f"{'👤 ' if cfg.get('_user_skill') else '📌 '}{cfg['name']}"
            all_skill_options[label] = key
        # 默认选中当前技能或自动
        st.session_state.setdefault("_skill_override", "__auto__")
        selected = st.selectbox(
            "选择技能模式",
            options=list(all_skill_options.keys()),
            index=list(all_skill_options.values()).index(
                st.session_state["_skill_override"]
            ) if st.session_state["_skill_override"] in all_skill_options.values() else 0,
            key="_skill_selector",
            label_visibility="collapsed",
        )
        override_key = all_skill_options[selected]
        if override_key != st.session_state.get("_skill_override"):
            st.session_state["_skill_override"] = override_key
            st.rerun()
        # 显示当前状态
        if override_key == "__auto__":
            st.caption(f"🔍 自动检测 → 当前：**{current_skill_name}**")
        else:
            st.caption(f"🔒 手动锁定 → **{SKILLS[override_key]['name']}**")
        st.divider()

        # ── 自定义 Skill 管理 ──
        st.header("🧩 自定义技能")
        with st.expander("管理自定义 Skill", expanded=False):
            user_skills = _list_user_skills()
            if user_skills:
                st.caption(f"已创建 {len(user_skills)} 个自定义 Skill：")
                for s in user_skills:
                    col_del, col_info = st.columns([1, 6])
                    with col_info:
                        st.caption(f"**{s['name']}** (`{s['key']}`) — 触发词：{', '.join(s['keywords'])}")
                    with col_del:
                        if st.button("🗑️", key=f"del_skill_{s['key']}", help=f"删除 {s['name']}"):
                            _delete_user_skill(s['key'])
                            reload_skills()
                            st.rerun()
            else:
                st.caption("暂无自定义 Skill")

            # 创建方式选择
            create_mode = st.radio("创建方式", ["🤖 AI 辅助创建", "✍️ 手动填写"], key="skill_create_mode", horizontal=True, label_visibility="collapsed")

            if create_mode == "🤖 AI 辅助创建":
                st.caption("用自然语言描述你想要的技能，AI 自动生成配置")
                ai_desc = st.text_area(
                    "描述你需要的技能",
                    placeholder="例如：我想要一个财务对账的技能，用来比对两个 Sheet 的数据差异，找出不一致的记录并给出处理建议。触发词用'对账、账单、核对'。",
                    key="ai_skill_desc",
                    height=80,
                )
                if st.button("🪄 生成 Skill 配置", key="gen_skill_btn", use_container_width=True):
                    if ai_desc.strip():
                        with st.spinner("AI 正在生成 Skill 配置..."):
                            generated = _ai_generate_skill(ai_desc)
                        if generated:
                            st.session_state["_generated_skill"] = generated
                            st.rerun()
                    else:
                        st.error("请先描述你需要的技能。")

                # 显示 AI 生成结果，允许编辑后确认
                gen = st.session_state.get("_generated_skill", {})
                if gen:
                    st.success("✅ AI 已生成配置，请检查并确认创建：")
                    st.caption("💡 可以修改任意字段后再创建")
                    final_key = st.text_input("标识 (英文)", value=gen.get("key", ""), key="ai_key")
                    final_name = st.text_input("名称 (中文)", value=gen.get("name", ""), key="ai_name")
                    final_kw = st.text_input("触发关键词", value=gen.get("keywords", ""), key="ai_kw")
                    final_prompt = st.text_area("系统提示词", value=gen.get("prompt", ""), key="ai_prompt", height=150)
                    final_tools = st.text_input("可用工具 (逗号分隔，留空=全部)", value=gen.get("tools", ""), key="ai_tools")
                    col_confirm, col_cancel = st.columns(2)
                    with col_confirm:
                        if st.button("✅ 确认创建", key="ai_confirm", use_container_width=True):
                            result = _create_user_skill(
                                name=final_key, display_name=final_name,
                                keywords=final_kw, system_prompt=final_prompt,
                                allowed_tools=final_tools,
                            )
                            reload_skills()
                            st.session_state.pop("_generated_skill", None)
                            st.success(result)
                            st.rerun()
                    with col_cancel:
                        if st.button("❌ 放弃", key="ai_cancel", use_container_width=True):
                            st.session_state.pop("_generated_skill", None)
                            st.rerun()
            else:
                with st.form("create_skill_form", clear_on_submit=True):
                    st.markdown("**新建 Skill**")
                    skill_key = st.text_input("标识 (英文)", placeholder="如 finance_review", key="sk_key")
                    skill_name = st.text_input("名称 (中文)", placeholder="如 财务对账", key="sk_name")
                    skill_kw = st.text_input("触发关键词 (逗号分隔)", placeholder="如 对账, 账单, 发票", key="sk_kw")
                    skill_prompt = st.text_area("系统提示词", placeholder="你是资深财务审计师...", key="sk_prompt", height=120)
                    skill_tools = st.text_input("可用工具 (逗号分隔，留空=全部)", placeholder="如 reconcile_sheets, read_excel", key="sk_tools")
                    if st.form_submit_button("💾 创建 Skill"):
                        if skill_key and skill_name and skill_kw and skill_prompt:
                            result = _create_user_skill(
                                name=skill_key, display_name=skill_name,
                                keywords=skill_kw, system_prompt=skill_prompt,
                                allowed_tools=skill_tools,
                            )
                            reload_skills()
                            st.success(result)
                            st.rerun()
                        else:
                            st.error("请填写标识、名称、关键词和提示词。")

        st.divider()
        st.header("⚙️ 设置")
        st.text_input("API Key",  value=API_KEY,  key="k", type="password")
        st.text_input("Base URL", value=BASE_URL, key="u")
        st.text_input("Model",    value=MODEL,    key="m")
        st.text_input("报告保存目录", value=str(Path.cwd()), key="save_dir",
                      help="生成的 Word/PDF 报告将自动保存到此目录")
        st.divider()
        st.caption("📄 报告导出偏好")
        col1, col2 = st.columns(2)
        with col1:
            st.checkbox("Word (.docx)", value=True, key="export_word")
        with col2:
            st.checkbox("PDF", value=True, key="export_pdf")
        col3, col4 = st.columns(2)
        with col3:
            st.checkbox("包含表格", value=True, key="include_tables")
        with col4:
            st.checkbox("包含图表", value=True, key="include_figures")
        if st.button("🗑️ 清空对话"):
            st.session_state.messages = [{"role":"system","content":SKILLS[DEFAULT_SKILL]["system_prompt"]}]
            for key in ("generated_charts", "generated_report_docx", "generated_report_name", "generated_report_pdf", "generated_report_pdf_name", "last_plot", "last_plot_title", "uploaded_image_preview", "edited_excel", "edited_excel_name", "_skill_override", "usage"):
                st.session_state.pop(key, None)
                st.session_state.pop(key, None)
            _clear_cached_file_data()
            _rag_clear(session_id=sid)
            delete_session(sid)
            st.rerun()
        st.caption(f"📎 会话 ID：`{sid}`")

        # ── 用量追踪面板 ──
        st.divider()
        _render_usage_sidebar()

        # ── 会话重命名 ──
        current_name = get_session_name(sid)
        with st.expander("✏️ 命名当前会话", expanded=False):
            new_name = st.text_input(
                "会话名称", value=current_name,
                placeholder="如：Q3销售分析、合同审查_客户A",
                key="rename_input",
                label_visibility="collapsed",
            )
            if new_name and new_name != current_name:
                set_session_name(sid, new_name)
                st.rerun()
            if st.button("🤖 自动提取主题", key="auto_topic_btn", use_container_width=True):
                msgs = st.session_state.get("messages", [])
                for m in msgs:
                    if m.get("role") == "user":
                        topic = m["content"].strip()[:20]
                        set_session_topic(sid, topic)
                        st.rerun()
                        break

        # 会话列表（带删除按钮 + 名称显示）
        sessions = list_sessions()
        if len(sessions) > 1:
            st.divider()
            st.caption(f"📜 历史会话（共 {len(sessions) - 1} 条）")
            for s in sessions:
                if s["id"] == sid:
                    continue
                # 显示名称：用户命名 > 自动主题 > 默认名 > ID
                display = s.get("name") or s.get("topic") or s.get("id", "???")
                col_btn, col_del = st.columns([5, 1])
                with col_btn:
                    lbl = f"{display}（{s['msgs']}条）"
                    if st.button(lbl, key=f"hist_{s['id']}", use_container_width=True, help=f"ID: {s['id']}"):
                        loaded = load_messages(s["id"])
                        if loaded:
                            st.session_state.messages = loaded
                        st.session_state["_skill_override"] = "__auto__"  # 切会话重置技能
                        st.rerun()
                with col_del:
                    if st.button("🗑️", key=f"del_{s['id']}", help=f"删除 {display}"):
                        _rag_clear(session_id=s["id"])
                        delete_session(s["id"])
                        st.rerun()
            # 一键清空
            other_sessions = [s for s in sessions if s["id"] != sid]
            if len(other_sessions) >= 5:
                if st.button("🗑️ 清空全部历史会话", key="clear_all_sessions", use_container_width=True):
                    for s in other_sessions:
                        _rag_clear(session_id=s["id"])
                        delete_session(s["id"])
                    st.rerun()

    st.session_state.messages = [normalize_message(msg) for msg in st.session_state.messages]
    for msg in st.session_state.messages:
        role = msg.get("role")
        if role in ("system", "tool") or not role:
            continue
        content = msg.get("content")
        if not isinstance(content, str) or not content:
            continue
        with st.chat_message(role):
            content = msg.get("content")
            st.markdown(content)

    if "uploaded_image_preview" in st.session_state:
        st.image(st.session_state["uploaded_image_preview"], caption="上传图片预览")

    if "generated_report_docx" in st.session_state:
        st.sidebar.download_button(
            "📄 下载分析报告（Word）",
            data=st.session_state["generated_report_docx"],
            file_name=st.session_state.get("generated_report_name", "分析报告.docx"),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    if "generated_report_pdf" in st.session_state:
        st.sidebar.download_button(
            "📑 下载分析报告（PDF）",
            data=st.session_state["generated_report_pdf"],
            file_name=st.session_state.get("generated_report_pdf_name", "分析报告.pdf"),
            mime="application/pdf",
        )

    # 提供编辑后文件下载
    if "edited_excel" in st.session_state:
        st.sidebar.download_button(
            "📥 下载编辑后的文件",
            data=st.session_state["edited_excel"],
            file_name=st.session_state.get("edited_excel_name", "edited.xlsx"),
        )

    if prompt := st.chat_input("有什么可以帮你？"):
        with st.chat_message("user"):
            st.markdown(prompt)

        # ── 自动提取会话主题（首条消息）──
        user_msgs = [m for m in st.session_state.messages if m.get("role") == "user"]
        if len(user_msgs) == 0:  # 这是第一条用户消息
            set_session_topic(sid, prompt)

        with st.chat_message("assistant"):
            # run_agent 开始时会重置 generated_charts，所以固定从0开始取新图表
            placeholder = st.empty()

            placeholder.write_stream(run_agent(prompt, session_id=sid))

            # ── 用量追踪：Agent 完成后显示用量（侧边栏在脚本早期渲染不会自动刷新）──
            _show_usage_inline()

            # ── 自动保存：对话持久化 + 分析快照提取 ──
            save_messages(sid, st.session_state.messages)
            clean_msgs = [m for m in st.session_state.messages if m.get("role") == "assistant" and isinstance(m.get("content"), str) and m["content"]]
            full_text = clean_msgs[-1]["content"] if clean_msgs else ""
            if full_text:
                auto_extract_snapshot(sid, full_text, st.session_state.get("generated_charts", []))

            # 有新图表 → 清空占位符，重新内联渲染（文字+图交织）
            new_charts = st.session_state.get("generated_charts", [])
            if new_charts:
                placeholder.empty()
                _render_text_with_charts(full_text, new_charts)

            # 显示思考过程折叠区
            reasoning_steps = st.session_state.get("reasoning_steps", [])
            if reasoning_steps:
                with st.expander(f"🧠 查看思考过程（{len(reasoning_steps)} 个推理步骤）", expanded=False):
                    for idx, step in enumerate(reasoning_steps):
                        if idx > 0:
                            st.divider()
                        st.markdown(f"**第 {step['round']} 轮推理**")
                        st.markdown(step["content"])

            # 如果刚生成了报告，重跑一次让侧边栏显示下载按钮
            if st.session_state.get("generated_report_docx") or st.session_state.get("generated_report_pdf"):
                st.rerun()

if __name__ == "__main__":
    main()
