
# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
import datetime  # 引入时间库

# ==================== 1. 配置区域 ====================
print("🔴 第一步：程序启动...")

# 👉 文件夹路径
folder_path = '.'
# 👉 文件名
file_name = "original_data.xlsx"

# 获取今天的日期 (格式: 20251118)
today_str = datetime.datetime.now().strftime('%Y%m%d')

input_excel_path = os.path.join(folder_path, file_name)
# 🔥 修改点：文件名自动加上今天日期
output_filename = f"clean_data_{today_str}.xlsx"
output_path = os.path.join(folder_path, output_filename)

print(f"   📂 读取文件: {input_excel_path}")
print(f"   💾 预定输出: {output_filename}")  # 打印出来给你确认一下


# ==================== 2. 日期强力修复 ====================
def super_fix_date(df, col_name='日期'):
    """统一修复日期格式，解决 1970 和 格式不一致问题"""
    if col_name not in df.columns: return df

    # 替换空值
    df[col_name] = df[col_name].replace([0, '0', 0.0, ''], pd.NA)
    df.dropna(subset=[col_name], inplace=True)

    def convert_date(val):
        try:
            if isinstance(val, pd.Timestamp): return val
            s_val = str(val).strip()
            # 纯数字 (Excel序列号)
            if s_val.replace('.', '', 1).isdigit():
                if float(s_val) < 10000: return pd.NaT
                return pd.to_datetime(float(s_val), unit='D', origin='1899-12-30')
            # 斜杠/横杠日期
            return pd.to_datetime(val, errors='coerce')
        except:
            return pd.NaT

    df[col_name] = df[col_name].apply(convert_date)
    df.dropna(subset=[col_name], inplace=True)
    df = df[df[col_name].dt.year > 2000]  # 过滤异常年份
    df[col_name] = df[col_name].dt.strftime('%Y-%m-%d')  # 转文本
    return df


# ==================== 3. 读取 Excel ====================
print("\n🔴 第三步：读取数据...")
if not os.path.exists(input_excel_path):
    print("❌ 找不到文件！请检查路径。")
    sys.exit()

try:
    xls = pd.ExcelFile(input_excel_path)
    # 自动识别 sheet
    sheet_total = next((s for s in xls.sheet_names if 'total' in s.lower()), None)
    sheet_emp = next((s for s in xls.sheet_names if 'employee' in s.lower()), None)

    # 兜底策略
    if not sheet_total and len(xls.sheet_names) > 1: sheet_total = xls.sheet_names[1]
    if not sheet_emp and len(xls.sheet_names) > 0: sheet_emp = xls.sheet_names[0]

    print(f"   -> 汇总表: {sheet_total}")
    print(f"   -> 员工表: {sheet_emp}")

    df_total = pd.read_excel(input_excel_path, sheet_name=sheet_total, header=1)
    df_emp = pd.read_excel(input_excel_path, sheet_name=sheet_emp, header=1)

except Exception as e:
    print(f"❌ 读取出错: {e}")
    sys.exit()

# ==================== 4. 清洗数据 ====================
print("\n🔴 第四步：清洗并标准化...")

# --- 清洗 Total ---
df_total.columns = df_total.columns.str.strip()
df_total.dropna(how='all', inplace=True)
# 删 Total 里的合计行
if '组别' in df_total.columns:
    df_total = df_total[~df_total['组别'].astype(str).str.contains('合计|Total', case=False, na=False)]
df_total = super_fix_date(df_total)

# --- 清洗 Employee ---
df_emp.columns = df_emp.columns.str.strip()
df_emp.dropna(how='all', inplace=True)
# 删 Employee 里的杂质
mask_junk = (df_emp['业务员'] == df_emp['组别']) | \
            df_emp['业务员'].astype(str).str.contains('合计|Total', case=False, na=False) | \
            df_emp['业务员'].isnull()
df_emp_clean = df_emp[~mask_junk].copy()
df_emp_clean = super_fix_date(df_emp_clean)

print("   ✅ 清洗完成！")

# ==================== 5. 核心：精准提取缺失的【日期、组别、业务员】 ====================
print("\n🔴 第五步：🔍 正在比对 Total 表缺失的人员名单...")

dates_in_total = set(df_total['日期'].unique())
dates_in_emp = set(df_emp_clean['日期'].unique())

# 找出遗漏的日期
missing_dates = dates_in_emp - dates_in_total

df_missing_final = pd.DataFrame()

if missing_dates:
    sorted_dates = sorted(list(missing_dates))
    print(f"   ⚠️ 发现 Total 表缺失以下日期的所有数据: {sorted_dates}")

    # 提取这些日期的详细数据
    df_missing_raw = df_emp_clean[df_emp_clean['日期'].isin(missing_dates)]

    # 🔥 核心：只提取你想要的字段
    target_cols = ['日期', '组别', '业务员']
    valid_cols = [c for c in target_cols if c in df_missing_raw.columns]

    df_missing_final = df_missing_raw[valid_cols].copy()
    df_missing_final.sort_values(by=['日期', '组别'], inplace=True)

    print("\n   👀 缺失名单预览 (前 10 行):")
    print("-" * 30)
    print(df_missing_final.head(10).to_string(index=False))
    print("-" * 30)
    print(f"   👉 共找到 {len(df_missing_final)} 条缺失的人员记录。")

else:
    print("   ✨ 完美：Total 表没有整天缺失的情况。")

# ==================== 6. 保存结果 ====================
print("\n🔴 第六步：保存文件...")
try:
    with pd.ExcelWriter(output_path) as writer:
        df_total.to_excel(writer, sheet_name='Total_Cleaned', index=False)
        df_emp_clean.to_excel(writer, sheet_name='Employee_Cleaned', index=False)

        
    print(f"🎉🎉🎉 成功！文件已生成: {output_filename}")

except PermissionError:
    print("❌ 保存失败：请先关闭 Excel 文件！")