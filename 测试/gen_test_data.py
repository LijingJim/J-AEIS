# 测试-销售数据（含异常）
# 这是一个模拟的季度销售报表，故意植入了多种数据异常用于测试Agent的分析和异常检测能力。

import pandas as pd
import numpy as np
from pathlib import Path

np.random.seed(42)

# ── 正常数据 ──
regions = ["华东", "华南", "华北", "华中", "西南", "西北", "东北"]
products = ["云服务", "数据中台", "AI平台", "安全产品", "运维托管"]
months = pd.date_range("2025-01-01", "2025-06-30", freq="MS").strftime("%Y-%m")

records = []
for month in months:
    for region in regions:
        for product in products:
            base = {
                "华东-云服务": 580, "华南-云服务": 420, "华北-云服务": 350,
                "华东-数据中台": 310, "华南-数据中台": 280, "华北-数据中台": 220,
                "华东-AI平台": 490, "华南-AI平台": 340, "华北-AI平台": 280,
            }.get(f"{region}-{product}", np.random.randint(80, 200))
            revenue = base + np.random.randint(-30, 50)
            records.append({
                "月份": month,
                "区域": region,
                "产品线": product,
                "销售额_万元": revenue,
                "成本_万元": int(revenue * np.random.uniform(0.4, 0.7)),
                "客户数": np.random.randint(5, 40),
            })

df = pd.DataFrame(records)

# ── 植入异常 ──

# 异常1：缺失值 —— 华中-数据中台 5月成本为空
mask1 = (df["月份"] == "2025-05") & (df["区域"] == "华中") & (df["产品线"] == "数据中台")
df.loc[mask1, "成本_万元"] = np.nan

# 异常2：极端高值 —— 华南-云服务 3月销售额暴涨（正常420→2900）
mask2 = (df["月份"] == "2025-03") & (df["区域"] == "华南") & (df["产品线"] == "云服务")
df.loc[mask2, "销售额_万元"] = 2900

# 异常3：零值 —— 西北-AI平台 2月客户数为0
mask3 = (df["月份"] == "2025-02") & (df["区域"] == "西北") & (df["产品线"] == "AI平台")
df.loc[mask3, "客户数"] = 0

# 异常4：负值 —— 东北-运维托管 4月销售额为负（退款）
mask4 = (df["月份"] == "2025-04") & (df["区域"] == "东北") & (df["产品线"] == "运维托管")
df.loc[mask4, "销售额_万元"] = -45

# 异常5：区域命名不一致 —— "华北" 在6月写成了 "华北区"
mask5 = (df["月份"] == "2025-06") & (df["区域"] == "华北")
df.loc[mask5, "区域"] = "华北区"

# 异常6：华南 6月所有产品成本为0（数据缺失整行）
mask6 = (df["月份"] == "2025-06") & (df["区域"] == "华南")
df.loc[mask6, "成本_万元"] = 0

# 异常7：AI平台 3月销售额骤降（正常~280→15）
mask7 = (df["月份"] == "2025-03") & (df["区域"] == "华北") & (df["产品线"] == "AI平台")
df.loc[mask7, "销售额_万元"] = 15

out_dir = Path(__file__).parent
out_path = out_dir / "测试_异常销售数据.xlsx"
df.to_excel(str(out_path), index=False, engine="openpyxl")
print(f"✅ 已生成: {out_path}")
print(f"   行数: {len(df)}, 列数: {len(df.columns)}")
print(f"   植入异常:")
print(f"     1. 缺失值: 华中-数据中台 5月成本为空")
print(f"     2. 极端高值: 华南-云服务 3月销售额 2900万 (正常~420)")
print(f"     3. 零值: 西北-AI平台 2月客户数为0")
print(f"     4. 负值: 东北-运维托管 4月销售额 -45万")
print(f"     5. 命名不一致: 华北→华北区 (6月)")
print(f"     6. 华南6月成本全为0")
print(f"     7. 骤降: 华北-AI平台 3月销售额→15万 (正常~280)")
