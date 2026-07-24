import pandas as pd

base_path = '业绩数据底表-20260717.csv'
modified_path = '业绩数据20260717.csv'

for enc in ('utf-8','utf-8-sig','gbk','gb18030'):
    try:
        df_base = pd.read_csv(base_path, encoding=enc)
        df_modified = pd.read_csv(modified_path, encoding=enc)
        df_base.columns = [c.strip() for c in df_base.columns]
        df_modified.columns = [c.strip() for c in df_modified.columns]
        break
    except Exception as e:
        print(f"尝试编码 {enc} 失败: {e}")
        continue

print(f"底表行数: {len(df_base)}")
print(f"修改表行数: {len(df_modified)}")

common_cols = [c for c in df_base.columns if c in df_modified.columns]
print(f"\n共有列数: {len(common_cols)}")

df_base = df_base[common_cols]
df_modified = df_modified[common_cols]

diff_mask = df_base.ne(df_modified)
diff_rows = diff_mask.any(axis=1)
print(f"\n有差异的行数: {diff_rows.sum()}")

if diff_rows.sum() > 0:
    print("\n差异列统计:")
    for col in common_cols:
        col_diff = diff_mask[col].sum()
        if col_diff > 0:
            print(f"  {col}: {col_diff} 行不同")
    
    print("\n详细差异示例（前10行）:")
    for i, (idx, row) in enumerate(df_base[diff_rows].head(10).iterrows()):
        print(f"\n--- 行 {idx} ---")
        for col in common_cols:
            if diff_mask.loc[idx, col]:
                print(f"  {col}: 底表='{df_base.loc[idx, col]}' -> 修改表='{df_modified.loc[idx, col]}'")

print("\n=== 检查MGA业务数据 ===")
mga_in_base = df_base[df_base['business_category'].astype(str).str.contains('MGA', na=False)]
print(f"底表中MGA业务行数: {len(mga_in_base)}")
if len(mga_in_base) > 0:
    print("底表中MGA业务的business_category值:", mga_in_base['business_category'].unique())
    print("底表中MGA业务的segment_code值:", mga_in_base['segment_code'].unique())

mga_in_modified = df_modified[df_modified['business_category'].astype(str).str.contains('MGA', na=False)]
print(f"\n修改表中MGA业务行数: {len(mga_in_modified)}")

converted_in_modified = df_modified[df_modified['segment_code'] == '同行经代']
print(f"\n修改表中同行经代行数: {len(converted_in_modified)}")

print("\n=== 验证: 是否所有MGA业务都被转换为同行经代 ===")
mga_segment_in_base = df_base[df_base['business_category'].astype(str).str.contains('MGA', na=False)]['segment_code'].unique()
print(f"底表中MGA业务的segment_code: {mga_segment_in_base}")