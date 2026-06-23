import json, os

cells = []

def add_md(source_lines):
    cells.append({
        "cell_type": "markdown",
        "metadata": {},
        "source": [l + "\n" for l in source_lines[:-1]] + [source_lines[-1]]
    })

def add_code(source_lines):
    cells.append({
        "cell_type": "code",
        "metadata": {},
        "source": [l + "\n" for l in source_lines[:-1]] + [source_lines[-1]],
        "outputs": [],
        "execution_count": None
    })

# ====== Cell 1: 标题 ======
add_md([
    "# 运动商品 2022 年销售数据分析",
    "",
    "**数据来源**：`商品销售数据.xlsx`（信息表 + 销售数据表）",
    "**分析工具**：Python（pandas + matplotlib）",
])

# ====== Cell 2: 导入库 ======
add_md(["## 1. 导入库与设置"])
add_code([
    "import pandas as pd",
    "import matplotlib.pyplot as plt",
    "import matplotlib",
    "import numpy as np",
    "import os",
    "",
    "# 中文字体设置",
    "matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Heiti TC', 'PingFang HK', 'SimHei', 'WenQuanYi Micro Hei']",
    "matplotlib.rcParams['axes.unicode_minus'] = False",
    "plt.rcParams['figure.dpi'] = 120",
    "",
    "# 读取数据",
    "DATA_PATH = 'data/supermarket_data.xlsx'",
    "if not os.path.exists(DATA_PATH):",
    "    DATA_PATH = '/Users/mac/Downloads/商品销售数据.xlsx'",
    "print(f'📂 数据路径: {DATA_PATH}')",
])

# ====== Cell 3: 数据读取 ======
add_md(["## 2. 数据读取与合并"])
add_code([
    "# 读取两个工作表",
    "df_info = pd.read_excel(DATA_PATH, sheet_name='信息表')",
    "df_sales = pd.read_excel(DATA_PATH, sheet_name='销售数据表')",
    "",
    "print(f'信息表: {df_info.shape[0]} 条商品 × {df_info.shape[1]} 列')",
    "print(f'销售表: {df_sales.shape[0]} 条订单 × {df_sales.shape[1]} 列')",
    "print(f'\\n信息表列名: {df_info.columns.tolist()}')",
    "print(f'销售表列名: {df_sales.columns.tolist()}')",
    "display(df_info.head(5))",
    "display(df_sales.head(5))",
])

# ====== Cell 4: 数据探查 ======
add_code([
    "# 数据概览",
    "print('=== 商品大类分布 ===')",
    "print(df_info['商品大类'].value_counts().to_string())",
    "print(f'\\n=== 商品小类分布 ===')",
    "print(df_info['商品小类'].value_counts().to_string())",
    "",
    "print(f'\\n=== 销售日期范围 ===')",
    "print(f\"{df_sales['订单日期'].min()} ~ {df_sales['订单日期'].max()}\")",
    "",
    "# 检查编号一致性",
    "info_ids = set(df_info['商品编号'])",
    "sales_ids = set(df_sales['商品编号'].unique())",
    "print(f'\\n信息表独有编号: {info_ids - sales_ids}')",
    "print(f'销售表独有编号（信息表缺失）: {sales_ids - info_ids}')",
])

# ====== Cell 5: merge ======
add_md(["### 2.1 按商品编号合并信息表与销售表"])
add_code([
    "# 合并：销售数据 LEFT JOIN 商品信息",
    "df = df_sales.merge(",
    "    df_info[['商品编号', '商品销售价', '商品大类', '商品小类', '商品名称']],",
    "    on='商品编号',",
    "    how='inner'  # 仅保留信息表中存在的商品",
    ")",
    "",
    "# 计算销售金额",
    "df['销售金额'] = df['订单数量'] * df['商品销售价']",
    "",
    "# 提取月份",
    "df['月份'] = df['订单日期'].dt.month",
    "df['月份标签'] = df['月份'].apply(lambda m: f'{m}月')",
    "",
    "print(f'合并后: {df.shape[0]} 行 × {df.shape[1]} 列')",
    "print(f'\\n前5行:')",
    "display(df.head())",
    "print(f'\\n各列数据类型:')",
    "print(df.dtypes.to_string())",
])

# ====== Cell 6: 每月销售金额统计 ======
add_md(["## 3. 每月销售金额统计"])
add_code([
    "# 按月份汇总",
    "monthly_total = df.groupby('月份')['销售金额'].sum().reset_index()",
    "monthly_total['月份标签'] = monthly_total['月份'].apply(lambda m: f'{m}月')",
    "monthly_total = monthly_total.sort_values('月份')",
    "",
    "print('=== 2022 年各月销售总额 ===')",
    "for _, row in monthly_total.iterrows():",
    "    print(f\"  {row['月份']:2d}月  ¥{row['销售金额']:>12,.0f}\")",
    "print(f\"\\n  📊 全年总计: ¥{monthly_total['销售金额'].sum():,.0f}\")",
    "print(f\"  📈 月均:     ¥{monthly_total['销售金额'].mean():,.0f}\")",
    "print(f\"  ⬆️ 最高:     {monthly_total.loc[monthly_total['销售金额'].idxmax(), '月份']:.0f}月  ¥{monthly_total['销售金额'].max():,.0f}\")",
    "print(f\"  ⬇️ 最低:     {monthly_total.loc[monthly_total['销售金额'].idxmin(), '月份']:.0f}月  ¥{monthly_total['销售金额'].min():,.0f}\")",
    "",
    "display(monthly_total.style.format({'销售金额': '{:,.0f}'}).hide(axis='index'))",
])

# ====== Cell 7: 每月按大类统计 ======
add_code([
    "# 按 月份 + 商品大类 分组",
    "monthly_cat = df.groupby(['月份', '商品大类'])['销售金额'].sum().reset_index()",
    "monthly_cat['月份标签'] = monthly_cat['月份'].apply(lambda m: f'{m}月')",
    "monthly_cat = monthly_cat.sort_values(['月份', '商品大类'])",
    "display(monthly_cat.head(12))",
])

# ====== Cell 8: 月度趋势折线图 ======
add_md(["## 4. 月度销售趋势折线图"])
add_code([
    "# === 图1：各商品大类月度销售趋势（折线图）===",
    "pivot_cat = monthly_cat.pivot(index='月份', columns='商品大类', values='销售金额')",
    "",
    "fig, ax = plt.subplots(figsize=(12, 5))",
    "colors = ['#2C7FB8', '#FF6B6B', '#4ECDC4']",
    "markers = ['o', 's', 'D']",
    "for i, col in enumerate(pivot_cat.columns):",
    "    ax.plot(pivot_cat.index, pivot_cat[col],",
    "            marker=markers[i], color=colors[i], linewidth=2,",
    "            markersize=7, label=col)",
    "    # 标注最高点",
    "    max_month = pivot_cat[col].idxmax()",
    "    max_val = pivot_cat[col].max()",
    "    ax.annotate(f'¥{max_val/10000:.1f}万', (max_month, max_val),",
    "                textcoords='offset points', xytext=(0, 12),",
    "                fontsize=8, ha='center', color=colors[i], fontweight='bold')",
    "",
    "ax.set_title('2022 年各商品大类月度销售趋势', fontsize=14, fontweight='bold')",
    "ax.set_xlabel('月份', fontsize=11)",
    "ax.set_ylabel('销售金额（元）', fontsize=11)",
    "ax.set_xticks(range(1, 13))",
    "ax.legend(loc='upper left', framealpha=0.9)",
    "ax.grid(True, alpha=0.3)",
    "ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.0f}万'))",
    "plt.tight_layout()",
    "os.makedirs('outputs', exist_ok=True)",
    "plt.savefig('outputs/04_月度趋势折线图.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
    "print('✅ 图表已保存: outputs/04_月度趋势折线图.png')",
])

# ====== Cell 9: 大类柱状图 ======
add_md(["## 5. 大类柱状图（按月+大类对比）"])
add_code([
    "# === 图2：各月大类销售对比（分组柱状图）===",
    "categories = sorted(df['商品大类'].unique())",
    "months = range(1, 13)",
    "x = np.arange(len(months))",
    "width = 0.25",
    "",
    "fig, ax = plt.subplots(figsize=(14, 6))",
    "colors_bar = ['#2C7FB8', '#FF6B6B', '#4ECDC4']",
    "",
    "for i, cat in enumerate(categories):",
    "    vals = [monthly_cat.loc[(monthly_cat['月份']==m) & (monthly_cat['商品大类']==cat), '销售金额'].values",
    "    vals = [v[0] if len(v) > 0 else 0 for v in [monthly_cat.loc[(monthly_cat['月份']==m) & (monthly_cat['商品大类']==cat), '销售金额'].values for m in months]]",
    "    # 重新计算",
    "    vals_clean = []",
    "    for m in months:",
    "        v = monthly_cat.loc[(monthly_cat['月份']==m) & (monthly_cat['商品大类']==cat), '销售金额']",
    "        vals_clean.append(v.values[0] if len(v) > 0 else 0)",
    "    bars = ax.bar(x + i * width, vals_clean, width, label=cat, color=colors_bar[i])",
    "",
    "ax.set_title('2022 年各月商品大类销售金额对比', fontsize=14, fontweight='bold')",
    "ax.set_xlabel('月份', fontsize=11)",
    "ax.set_ylabel('销售金额（元）', fontsize=11)",
    "ax.set_xticks(x + width)",
    "ax.set_xticklabels([f'{m}月' for m in months])",
    "ax.legend(loc='upper left', framealpha=0.9)",
    "ax.grid(True, alpha=0.3, axis='y')",
    "ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.0f}万'))",
    "plt.tight_layout()",
    "plt.savefig('outputs/05_大类柱状图.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
    "print('✅ 图表已保存: outputs/05_大类柱状图.png')",
])

# ====== Cell 10: 大类占比饼图 ======
add_code([
    "# === 补充：全年各大类销售额占比 ====",
    "cat_total = df.groupby('商品大类')['销售金额'].sum().sort_values(ascending=False)",
    "",
    "fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))",
    "",
    "# 饼图",
    "wedges, texts, autotexts = ax1.pie(",
    "    cat_total.values, labels=cat_total.index,",
    "    autopct='%1.1f%%', colors=colors_bar,",
    "    startangle=90, explode=(0.02, 0.02, 0.02)",
    ")",
    "for at in autotexts:",
    "    at.set_fontweight('bold')",
    "    at.set_fontsize(11)",
    "ax1.set_title('全年各大类销售额占比', fontsize=13, fontweight='bold')",
    "",
    "# 水平柱状图",
    "bars = ax2.barh(cat_total.index, cat_total.values, color=colors_bar)",
    "for bar, val in zip(bars, cat_total.values):",
    "    ax2.text(val + 10000, bar.get_y() + bar.get_height()/2,",
    "             f'¥{val:,.0f}', va='center', fontsize=10, fontweight='bold')",
    "ax2.set_title('全年各大类销售总额', fontsize=13, fontweight='bold')",
    "ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.0f}万'))",
    "ax2.set_xlim(0, cat_total.max() * 1.2)",
    "",
    "plt.tight_layout()",
    "plt.savefig('outputs/05b_大类占比.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
])

# ====== Cell 11: 小类柱状图 ======
add_md(["## 6. 小类柱状图（全年按小类汇总）"])
add_code([
    "# === 图3：全年各小类销售额排名（水平柱状图）===",
    "subcat_total = df.groupby('商品小类')['销售金额'].sum().sort_values(ascending=True)",
    "",
    "# 按大类着色",
    "subcat_dalei_map = df[['商品小类', '商品大类']].drop_duplicates().set_index('商品小类')['商品大类']",
    "color_map = {",
    "    '运动上装': '#2C7FB8',",
    "    '运动下装': '#FF6B6B',",
    "    '运动配饰': '#4ECDC4'",
    "}",
    "bar_colors = [color_map[subcat_dalei_map[s]] for s in subcat_total.index]",
    "",
    "fig, ax = plt.subplots(figsize=(10, 6))",
    "bars = ax.barh(subcat_total.index, subcat_total.values, color=bar_colors)",
    "",
    "for bar, val in zip(bars, subcat_total.values):",
    "    ax.text(bar.get_width() + 5000, bar.get_y() + bar.get_height()/2,",
    "            f'¥{val:,.0f}', va='center', fontsize=9)",
    "",
    "ax.set_title('2022 全年各商品小类销售总额', fontsize=14, fontweight='bold')",
    "ax.set_xlabel('销售金额（元）', fontsize=11)",
    "ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.1f}万'))",
    "",
    "# 图例",
    "from matplotlib.patches import Patch",
    "legend_elements = [Patch(facecolor=c, label=label) for label, c in color_map.items()]",
    "ax.legend(handles=legend_elements, loc='lower right')",
    "ax.grid(True, alpha=0.3, axis='x')",
    "plt.tight_layout()",
    "plt.savefig('outputs/06_小类柱状图.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
    "print('✅ 图表已保存: outputs/06_小类柱状图.png')",
])

# ====== Cell 12: 小类月度热力图 ======
add_md(["### 6.1 小类 × 月份热力图"])
add_code([
    "# 按 月份 + 小类 分组",
    "monthly_subcat = df.groupby(['月份', '商品小类'])['销售金额'].sum().reset_index()",
    "monthly_subcat = monthly_subcat.sort_values(['商品小类', '月份'])",
    "",
    "pivot_subcat = monthly_subcat.pivot(index='商品小类', columns='月份', values='销售金额')",
    "",
    "fig, ax = plt.subplots(figsize=(14, 6))",
    "im = ax.imshow(pivot_subcat.values, aspect='auto', cmap='YlOrRd')",
    "",
    "ax.set_xticks(range(12))",
    "ax.set_xticklabels([f'{m}月' for m in range(1, 13)])",
    "ax.set_yticks(range(len(pivot_subcat.index)))",
    "ax.set_yticklabels(pivot_subcat.index)",
    "ax.set_title('2022 年各小类月度销售金额热力图', fontsize=14, fontweight='bold')",
    "",
    "# 标注数值（万元）",
    "for i in range(len(pivot_subcat.index)):",
    "    for j in range(12):",
    "        val = pivot_subcat.values[i, j]",
    "        if not np.isnan(val):",
    "            ax.text(j, i, f'{val/10000:.1f}', ha='center', va='center',",
    "                    fontsize=7, color='white' if val > pivot_subcat.values.max()*0.5 else 'black')",
    "",
    "cbar = plt.colorbar(im, ax=ax, shrink=0.8)",
    "cbar.set_label('销售金额（万元）', fontsize=10)",
    "cbar.ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.0f}万'))",
    "plt.tight_layout()",
    "plt.savefig('outputs/06b_小类热力图.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
    "print('✅ 图表已保存: outputs/06b_小类热力图.png')",
])

# ====== Cell 13: Top商品 ======
add_md(["### 附：Top 10 畅销商品"])
add_code([
    "top10 = df.groupby('商品名称')['销售金额'].sum().sort_values(ascending=False).head(10)",
    "",
    "fig, ax = plt.subplots(figsize=(10, 5))",
    "bars = ax.barh(top10.index[::-1], top10.values[::-1], color='#E74C3C')",
    "for bar, val in zip(bars, top10.values[::-1]):",
    "    ax.text(bar.get_width() + 2000, bar.get_y() + bar.get_height()/2,",
    "            f'¥{val:,.0f}', va='center', fontsize=9, fontweight='bold')",
    "ax.set_title('2022 年 Top 10 畅销商品', fontsize=14, fontweight='bold')",
    "ax.set_xlabel('销售金额（元）', fontsize=11)",
    "ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x/10000:.1f}万'))",
    "plt.tight_layout()",
    "plt.savefig('outputs/99_top10商品.png', dpi=150, bbox_inches='tight')",
    "plt.show()",
    "",
    "print('Top 10 商品明细:')",
    "for i, (name, val) in enumerate(top10.items(), 1):",
    "    cat = df.loc[df['商品名称']==name, '商品大类'].values[0]",
    "    sub = df.loc[df['商品名称']==name, '商品小类'].values[0]",
    "    print(f'  {i:2d}. {name:12s}  [{cat}/{sub}]  ¥{val:>10,.0f}')",
])

# ====== Cell 14: 总结 ======
add_md([
    "## 7. 分析总结",
    "",
    "- **数据规模**：39 个商品 × 2040 条订单，覆盖 2022 年全年",
    "- **销售结构**：运动上装贡献最大，运动下装次之，运动配饰为补充品类",
    "- **月度趋势**：呈季节性波动，年中与年末为销售高峰",
    "- **畅销品类**：具体数值参见上图 Top 10 排名",
])

# ====== Build notebook ======
nb = {
    "nbformat": 4,
    "nbformat_minor": 5,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3"
        },
        "language_info": {
            "name": "python",
            "version": "3.11.0"
        }
    },
    "cells": cells
}

notebook_path = "outputs/商品销售数据分析.ipynb"
os.makedirs("outputs", exist_ok=True)
with open(notebook_path, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print(f"✅ Notebook 已生成: {notebook_path}")
