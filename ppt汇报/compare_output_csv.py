import pandas as pd
import os

csv_files = ['S1-总览仪表盘.csv', 'S2-业务端视角.csv', 'S3-执行管理端.csv', 'S4-产品端视角.csv']

for csv_name in csv_files:
    new_path = csv_name
    old_path = csv_name.replace('.csv', '_old.csv')
    
    print(f"\n{'='*60}")
    print(f"对比 {csv_name}")
    print(f"{'='*60}")
    
    for enc in ('utf-8-sig', 'utf-8', 'gbk', 'gb18030'):
        try:
            df_new = pd.read_csv(new_path, encoding=enc)
            break
        except Exception:
            continue
    
    if not os.path.exists(old_path):
        print(f"  旧文件不存在: {old_path}")
        continue
    
    for enc in ('utf-8-sig', 'utf-8', 'gbk', 'gb18030'):
        try:
            df_old = pd.read_csv(old_path, encoding=enc)
            break
        except Exception:
            continue
    
    df_new.columns = [c.strip() for c in df_new.columns]
    df_old.columns = [c.strip() for c in df_old.columns]
    
    common_cols = [c for c in df_new.columns if c in df_old.columns]
    
    print(f"  新文件行数: {len(df_new)}")
    print(f"  旧文件行数: {len(df_old)}")
    print(f"  共有列数: {len(common_cols)}")
    
    if len(df_new) != len(df_old):
        print(f"  ⚠️ 行数不一致")
        
    df_new = df_new[common_cols]
    df_old = df_old[common_cols]
    
    diff_mask = df_new.ne(df_old)
    diff_rows = diff_mask.any(axis=1)
    print(f"  有差异的行数: {diff_rows.sum()}")
    
    if diff_rows.sum() > 0:
        print(f"  差异列统计:")
        for col in common_cols:
            col_diff = diff_mask[col].sum()
            if col_diff > 0:
                print(f"    {col}: {col_diff} 行不同")
        
        print(f"\n  详细差异示例（前5行）:")
        count = 0
        for idx, row in df_new[diff_rows].iterrows():
            if count >= 5:
                break
            print(f"\n    --- 行 {idx} ---")
            for col in common_cols:
                if diff_mask.loc[idx, col]:
                    new_val = df_new.loc[idx, col]
                    old_val = df_old.loc[idx, col]
                    print(f"      {col}: 新='{new_val}' -> 旧='{old_val}'")
            count += 1
    else:
        print(f"  ✅ 完全一致")

print(f"\n{'='*60}")