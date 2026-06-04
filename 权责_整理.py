import pandas as pd
import warnings
import tkinter as tk
from tkinter import filedialog

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

file_path = r"C:\Users\陈鑫\Desktop\权责收入表.xlsx"
df_raw = pd.read_excel(file_path, header=[0, 1])

# ========== 改进的扁平化列名 ==========
new_columns = []
for col in df_raw.columns:
    level0 = str(col[0]).strip()
    level1 = str(col[1]).strip()

    if level1 in ("nan", "") or level1.startswith("Unnamed"):
        new_columns.append(level0)
    else:
        new_columns.append(f"{level0}_{level1}")

df_raw.columns = new_columns

print("修正后的列名：")
print(df_raw.columns.tolist())

# ========== 去掉末尾的“合计”行 ==========
target_col = "序号"
if target_col in df_raw.columns:
    df_raw = df_raw[~df_raw[target_col].astype(str).str.contains("合计", na=False)]
else:
    print(f"警告：未找到列名「{target_col}」，跳过过滤合计行")

# ========== 需要保留的固定列 ==========
fixed_cols = ["品牌", "铺位/点位号", "费用名称", "楼层", "合同编号"]

# 找出月份相关列
month_cols = [c for c in df_raw.columns if "年" in c and ("_含税" in c or "_不含税" in c)]
months = sorted(set([c.split("_")[0] for c in month_cols]))

# ========== 转换为长表 ==========
rows = []
for _, row in df_raw.iterrows():
    for month in months:
        tax_col = f"{month}_含税"
        net_col = f"{month}_不含税"
        if tax_col not in df_raw.columns or net_col not in df_raw.columns:
            continue
        tax_val = row[tax_col]
        net_val = row[net_col]
        if pd.isna(tax_val) and pd.isna(net_val):
            continue
        rows.append({
            "品牌": row["品牌"],
            "铺位/点位号": row["铺位/点位号"],
            "费用名称": row["费用名称"],
            "月份": month,
            "含税收入": tax_val,
            "不含税收入": net_val,
            "楼层": row["楼层"],
            "合同编号": row["合同编号"]
        })

df_long = pd.DataFrame(rows)

# ========== 弹出保存对话框，让用户选择保存位置 ==========
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

save_excel_path = filedialog.asksaveasfilename(
    title="保存转换后的文件",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    initialfile="权责收入表_精简长表.xlsx"  # 默认文件名
)

if save_excel_path:
    # 保存 Excel
    df_long.to_excel(save_excel_path, index=False)


else:
    print("用户取消了保存，文件未输出。")

