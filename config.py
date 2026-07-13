import os

# LLM 配置（兼容 OpenAI 格式，DeepSeek/Moonshot/智谱等均可）
API_KEY  = os.getenv("OPENAI_API_KEY",  "your-api-key-here")
BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com")
MODEL    = os.getenv("OPENAI_MODEL",    "your-model")

# Token 成本追踪（单位：美元/百万 token）
# DeepSeek: https://api-docs.deepseek.com/quick_start/pricing
# 其他模型填入对应价格即可自动计算
MODEL_PRICING = {
    "deepseek-v4-pro":    {"input": 0.14, "output": 0.28},  # 每百万 token 美元
    "deepseek-chat":      {"input": 0.14, "output": 0.28},
    "gpt-4o":             {"input": 2.50, "output": 10.00},
    "gpt-4o-mini":        {"input": 0.15, "output": 0.60},
    "moonshot-v1":        {"input": 0.50, "output": 0.50},
    "qwen-plus":          {"input": 0.50, "output": 2.00},
}

# 企业微信配置（第3周使用）
WECOM_CORP_ID  = os.getenv("WECOM_CORP_ID",  "")
WECOM_AGENT_ID = os.getenv("WECOM_AGENT_ID", "")
WECOM_SECRET   = os.getenv("WECOM_SECRET",   "")
