import pandas as pd

csv_files = ['S1-总览仪表盘.csv', 'S2-业务端视角.csv', 'S3-执行管理端.csv', 'S4-产品端视角.csv']

for csv_name in csv_files:
    print(f"\n{'='*60}")
    print(f"检查 {csv_name}")
    print(f"{'='*60}")
    
    df = None
    for enc in ('utf-8-sig', 'utf-8', 'gbk', 'gb18030'):
        try:
            df = pd.read_csv(csv_name, encoding=enc)
            break
        except Exception:
            continue
    
    if df is None:
        print(f"  ❌ 无法读取文件")
        continue
    
    df.columns = [c.strip() for c in df.columns]
    
    print(f"  行数: {len(df)}")
    print(f"  列数: {len(df.columns)}")
    print(f"  列名: {df.columns.tolist()}")
    
    for col in df.columns:
        vals = df[col].dropna().unique()
        if len(vals) > 0 and any('MGA' in str(v) or 'mga' in str(v).lower() for v in vals):
            print(f"\n  发现MGA相关数据在列 '{col}':")
            mga_vals = [v for v in vals if 'MGA' in str(v) or 'mga' in str(v).lower()]
            print(f"    唯一值: {mga_vals}")
    
    if 'segment' in df.columns or '业务细分' in df.columns:
        seg_col = 'segment' if 'segment' in df.columns else '业务细分'
        print(f"\n  业务细分分布:")
        for seg, cnt in df[seg_col].value_counts().items():
            print(f"    {seg}: {cnt} 行")

print(f"\n{'='*60}")