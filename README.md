# 🧠 CompanyAgent — 公司内部智能 Agent

基于 **OpenAI SDK + Streamlit + Function Calling** 的企业级智能助手。上传数据、描述需求，自动分析并生成报告。

## 🚀 为什么选它？

| 能力 | 通用 AI（ChatGPT/Kimi 等） | CompanyAgent |
|------|:---:|:---:|
| 🛠️ **LLM 自己造工具** | ❌ 只能调预设工具 | ✅ `create_tool` 运行时在线造新函数（安全沙箱） |
| 🧩 **用户自定义角色** | ❌ 固定人设 | ✅ Skill 系统：AI 辅助生成 + 关键词自动切换 + 手动锁定 |
| 📦 **开箱即用** | ❌ 需要配置/订阅 | ✅ 双击 `启动Agent.bat` / `安装依赖.bat` 即用 |
| 💾 **会话不丢失** | ⚠️ 刷新可能丢失 | ✅ SQLite 持久化 + 分析快照 + 链路可分享 |
| 📊 **报告深度** | 泛泛而谈 | ✅ 麦肯锡三阶框架：是什么→为什么→所以呢 |
| 🔍 **自审查** | ❌ 一次输出不改 | ✅ 12 项质量核验后自动修订 |
| ⚡ **并行调度** | ❌ 逐个调用 | ✅ ThreadPoolExecutor 最多 8 线程并行 |
| 📋 **合同审查** | ❌ 无 | ✅ 双层审查：15 项快扫 + LLM 语义深审 |

---

## 🏗️ 架构

```
用户输入 → Streamlit UI → Agent 循环 → 工具层
                              │
              ┌───────────────┼───────────────┐
              ▼               ▼               ▼
         excel_tool      file_tool      contract_tool
         (数据分析)      (文件解析)      (合同审查)
                              │
                   ┌──────────┴──────────┐
                   ▼                     ▼
             dynamic_tool          user_skills
             (运行时造工具)         (自定义角色)
```

---

## ✨ 核心能力

- **数据深度分析** — 麦肯锡/BCG 三阶框架，透视、排名、异常检测、多 Sheet 对账
- **智能报告生成** — 自动大纲规划 → 逐段扩写 → 排版校对 → Word + PDF 双格式导出
- **合同双层审查** — 15 项关键词快扫 → LLM 语义深审（中英文/否定句式/附件引用/不对等条款）
- **通用文件解析** — PDF 全文搜索/分段精读、Word、PPT、图片、HTML、ZIP
- **🆕 动态工具创建** — LLM 在线造新工具（安全沙箱内置 pd/np/plt），报告质量不降级
- **🆕 用户自定义 Skill** — AI 辅助生成 + Sidebar UI + 关键词自动匹配 + 手动锁定
- **会话三层记忆** — SQLite 持久化 + 分析快照 + 滑动窗口压缩，刷新不丢、链接可分享
- **自审查机制** — 初稿 12 项质量核验（数据溯源/三阶穿透/套话清理/排版规范）
- **并行工具调度** — ThreadPoolExecutor 最多 8 线程并行，响应时间大幅缩短

## 📦 安装

双击 `安装依赖.bat`（自动检测 Python → 安装依赖 → 配置引导），或手动：

```bash
pip install -r requirements.txt
```

## 🚀 启动

双击 `启动Agent.bat`，浏览器访问 `http://localhost:8501`。
首次在侧边栏填写 API Key 即可。

## 📁 项目结构

```
├── agent.py              # 主入口：Streamlit UI + Agent 循环
├── config.py             # LLM 与企业微信配置
├── session_store.py      # 会话持久化 + 快照 + 上下文压缩
├── skills.py             # 内置技能定义 + 用户 Skill 合并
├── user_skills_store.py  # 用户自定义 Skill 存储（JSON 持久化）
├── requirements.txt      # Python 依赖
├── 安装依赖.bat           # 首次安装脚本
├── 启动Agent.bat          # 日常启动脚本
├── .gitignore
├── tools/
│   ├── excel_tool.py     # Excel/CSV 全生命周期工具
│   ├── file_tool.py      # 通用文件解析（PDF/Word/PPT/图片等）
│   ├── contract_tool.py  # 合同双层审查引擎
│   └── dynamic_tool.py   # 运行时动态工具管理器
└── docs/                 # 用户手册 + 开发手册 + 图表源文件
```

## 🔧 配置

编辑 `config.py` 或通过环境变量：

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OPENAI_API_KEY` | LLM API 密钥 | — |
| `OPENAI_BASE_URL` | API 地址 | `your-api-url` |
| `OPENAI_MODEL` | 模型名称 | `your-model` |

## 📄 License

内部项目 · 公司内部使用

```bash
pip install -r requirements.txt
```

## 🚀 启动

```bash
streamlit run agent.py
```

或双击 `启动Agent.bat`（默认端口 8501）。

启动后在侧边栏填写 API Key，或设置环境变量：

```bash
set OPENAI_API_KEY=your-key
set OPENAI_BASE_URL=your-api-url
set OPENAI_MODEL=your-model
```

## 📁 项目结构

```
├── agent.py              # 主入口：Streamlit UI + Agent 循环
├── config.py             # LLM 与企业微信配置
├── session_store.py      # 会话持久化 + 快照 + 上下文压缩
├── skills.py             # 技能定义与路由
├── requirements.txt      # Python 依赖
├── 启动Agent.bat          # Windows 启动脚本
├── .gitignore
├── tools/
│   ├── excel_tool.py     # Excel/CSV 全生命周期工具
│   ├── file_tool.py      # 通用文件解析（PDF/Word/PPT/图片等）
│   ├── contract_tool.py  # 合同智能审查引擎
│   └── dynamic_tool.py   # 运行时动态工具管理器
└── docs/                 # 开发文档与手册（不上传）
```

## 🔧 配置

编辑 `config.py` 或通过环境变量配置：

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OPENAI_API_KEY` | LLM API 密钥 | — |
| `OPENAI_BASE_URL` | API 地址 | `your-api-url` |
| `OPENAI_MODEL` | 模型名称 | `your-model` |
| `WECOM_CORP_ID` | 企业微信 CorpID | — |
| `WECOM_AGENT_ID` | 企业微信 AgentID | — |
| `WECOM_SECRET` | 企业微信 Secret | — |