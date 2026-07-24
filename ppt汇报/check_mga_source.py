import pandas as pd

path = '业绩数据底表-20260717.csv'

for enc in ('utf-8-sig', 'utf-8', 'gbk', 'gb18030'):
    try:
        df = pd.read_csv(path, encoding=enc)
        df.columns = [c.strip() for c in df.columns]
        print(f"  ✓ 读取成功，编码: {enc}")
        break
    except Exception as e:
        print(f"  ✗ 编码 {enc} 失败: {e}")
        continue

print(f"\n  总行数: {len(df)}")
print(f"  列名: {df.columns.tolist()}")

if 'business_category' in df.columns:
    print(f"\n  business_category 唯一值:")
    for val, cnt in df['business_category'].value_counts().items():
        print(f"    {val}: {cnt}")

if 'segment_code' in df.columns:
    print(f"\n  segment_code 唯一值:")
    for val, cnt in df['segment_code'].value_counts().items():
        print(f"    {val}: {cnt}")

mga_mask = df['business_category'].astype(str).str.strip() == 'MGA业务'
print(f"\n  MGA业务行数: {mga_mask.sum()}")
if mga_mask.sum() > 0:
    print(f"  MGA业务的segment_code值:")
    for val, cnt in df[mga_mask]['segment_code'].value_counts().items():
        print(f"    {val}: {cnt}")