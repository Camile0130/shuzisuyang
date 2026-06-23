import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

random.seed(42)
np.random.seed(42)

# ============================
# 1. 商品信息表
# ============================
product_data = {
    "商品编号": [f"P{i:04d}" for i in range(1, 31)],
    "商品名称": [
        "冰红茶500ml", "绿茶500ml", "乌龙茶500ml", "茉莉花茶500ml", "普洱奶茶450ml",
        "原味薯片80g", "烧烤味薯片80g", "番茄味薯片80g", "虾条70g", "爆米花100g",
        "纯牛奶250ml", "酸奶200g", "乳酸菌饮品350ml", "椰汁330ml", "豆奶250ml",
        "夹心饼干120g", "苏打饼干100g", "曲奇饼干150g", "威化饼干90g", "蛋卷100g",
        "方便面红烧牛肉", "方便面酸菜", "方便面番茄", "杯面鸡汤", "干拌面110g",
        "白砂糖400g", "食盐250g", "酱油500ml", "醋500ml", "料酒500ml",
    ],
    "商品大类": (
        ["饮品"] * 5 + ["零食"] * 5 + ["乳品"] * 5
        + ["饼干糕点"] * 5 + ["方便食品"] * 5 + ["调料"] * 5
    ),
    "商品小类": (
        ["茶饮"] * 5 + ["膨化食品"] * 5 + ["液态奶"] * 3 + ["植物蛋白"] * 2
        + ["饼干"] * 3 + ["糕点"] * 2 + ["泡面"] * 3 + ["杯面/拌面"] * 2
        + ["糖盐"] * 2 + ["酱醋料酒"] * 3
    ),
    "进价": np.random.uniform(1.5, 12.0, 30).round(2),
    "售价": np.random.uniform(3.0, 20.0, 30).round(2),
}
df_info = pd.DataFrame(product_data)
df_info["售价"] = df_info["进价"] * np.random.uniform(1.3, 2.5, 30).round(2)

# ============================
# 2. 销售明细表 (12个月)
# ============================
records = []
start_date = datetime(2025, 1, 1)

for month in range(1, 13):
    month_start = start_date.replace(month=month)
    # 每月 8~20 天有销售记录
    n_days = random.randint(8, 20)
    days = random.sample(range(1, 29), n_days)
    for day in sorted(days):
        date = month_start.replace(day=day)
        n_products = random.randint(5, 15)
        products = random.sample(df_info["商品编号"].tolist(), n_products)
        for pid in products:
            qty = random.randint(5, 80)
            records.append({"商品编号": pid, "销售日期": date, "销售数量": qty})

df_sales = pd.DataFrame(records)
df_sales = df_sales.sort_values("销售日期").reset_index(drop=True)

# ============================
# 3. 写入 Excel
# ============================
os.makedirs("data", exist_ok=True)
with pd.ExcelWriter("data/supermarket_data.xlsx", engine="openpyxl") as writer:
    df_info.to_excel(writer, sheet_name="商品信息", index=False)
    df_sales.to_excel(writer, sheet_name="销售明细", index=False)

print("✅ 模拟数据已生成：data/supermarket_data.xlsx")
print(f"   商品信息：{len(df_info)} 条")
print(f"   销售明细：{len(df_sales)} 条")
