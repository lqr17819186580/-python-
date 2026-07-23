import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

ref_file = '永明业绩数据-20260717.xlsx'
new_file = '永明业绩数据-20260722.xlsx'

ref_df = pd.read_excel(ref_file)
new_df = pd.read_excel(new_file)

ref_df['保单号码'] = ref_df['保单号码'].fillna('').astype(str).str.strip()
new_df['保单号码'] = new_df['保单号码'].fillna('').astype(str).str.strip()

ref_df = ref_df[ref_df['保单号码'] != '']
new_df = new_df[new_df['保单号码'] != '']

date_cols = ['提交日期', '签单日期', '批核日（年/月/日）']
for col in date_cols:
    if col in ref_df.columns:
        ref_df[col] = pd.to_datetime(ref_df[col], errors='coerce').dt.strftime('%Y/%m/%d')
    if col in new_df.columns:
        new_df[col] = pd.to_datetime(new_df[col], errors='coerce').dt.strftime('%Y/%m/%d')

ref_cols = set(ref_df.columns)
new_cols = set(new_df.columns)

only_in_ref = sorted(ref_cols - new_cols)
only_in_new = sorted(new_cols - ref_cols)
common_cols = sorted(ref_cols & new_cols)

ref_df = ref_df.set_index('保单号码')
new_df = new_df.set_index('保单号码')

ref_index = set(ref_df.index)
new_index = set(new_df.index)

only_in_ref_index = sorted(ref_index - new_index)
only_in_new_index = sorted(new_index - ref_index)
common_index = sorted(ref_index & new_index)

compare_cols = [c for c in common_cols if c != '保单号码']

ref_common = ref_df.loc[common_index, compare_cols]
new_common = new_df.loc[common_index, compare_cols]

ref_empty = ref_common.isna() | (ref_common == '')
new_empty = new_common.isna() | (new_common == '')
both_empty = ref_empty & new_empty

diff_mask = ref_common.ne(new_common) & ~both_empty
diff_summary = diff_mask.sum().sort_values(ascending=False)

total_diff_cells = diff_mask.sum().sum()
total_rows = len(common_index)

diff_details = []
for col in compare_cols:
    col_diff_mask = diff_mask[col]
    if col_diff_mask.any():
        diff_rows = col_diff_mask[col_diff_mask].index.tolist()
        for idx in diff_rows:
            ref_val = ref_common.loc[idx, col]
            new_val = new_common.loc[idx, col]
            ref_val_str = str(ref_val) if pd.notna(ref_val) else ''
            new_val_str = str(new_val) if pd.notna(new_val) else ''
            diff_details.append({
                '保单号码': idx,
                '字段名': col,
                '参考值': ref_val_str,
                '新值': new_val_str
            })

consistent_cols = [col for col in compare_cols if diff_mask[col].sum() == 0]
inconsistent_cols = [col for col in compare_cols if diff_mask[col].sum() > 0]

doc = Document()
style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(11)

doc.add_heading('永明业绩数据 结果验证报告', level=1)
today = datetime.now().strftime('%Y年%m月%d日')
doc.add_paragraph(f'参考文件: {ref_file}')
doc.add_paragraph(f'输出文件: {new_file}')
doc.add_paragraph(f'验证日期: {today}')
doc.add_paragraph('')

doc.add_heading('目录', level=2)
doc.add_paragraph('（请在 Word 中右键此目录 → 更新域，以生成完整目录）')
doc.add_paragraph('')

doc.add_heading('一、数据量验证', level=2)
p = doc.add_paragraph()
p.add_run('验证数据的行数、列数是否与参考文件一致，分析差异原因。')

doc.add_heading('数据量差异原因分析', level=3)

data_diff = [
    ('数据行数', ref_df.shape[0], new_df.shape[0], new_df.shape[0] - ref_df.shape[0]),
    ('数据列数', ref_df.shape[1], new_df.shape[1], new_df.shape[1] - ref_df.shape[1]),
    ('保单数量', len(ref_index), len(new_index), len(new_index) - len(ref_index)),
    ('共同保单', len(common_index), len(common_index), 0),
    ('仅参考表有保单', len(only_in_ref_index), 0, -len(only_in_ref_index)),
    ('仅新表有保单', 0, len(only_in_new_index), len(only_in_new_index)),
]

table = doc.add_table(rows=1, cols=5)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '指标'
hdr[1].text = '参考文件'
hdr[2].text = '输出文件'
hdr[3].text = '差异'
hdr[4].text = '状态'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

for name, ref_val, new_val, diff_val in data_diff:
    cells = table.add_row().cells
    cells[0].text = name
    cells[1].text = str(ref_val)
    cells[2].text = str(new_val)
    cells[3].text = f'{diff_val:+.0f}' if isinstance(diff_val, int) else str(diff_val)
    status = '✅ 一致' if diff_val == 0 else '⚠️ 差异'
    cells[4].text = status

if only_in_ref_index:
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.add_run(f'仅参考表有保单（{len(only_in_ref_index)}个）: ').bold = True
    p.add_run(', '.join(only_in_ref_index[:10]))
    if len(only_in_ref_index) > 10:
        p.add_run(f'...（共{len(only_in_ref_index)}个）')

if only_in_new_index:
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.add_run(f'仅新表有保单（{len(only_in_new_index)}个）: ').bold = True
    p.add_run(', '.join(only_in_new_index[:10]))
    if len(only_in_new_index) > 10:
        p.add_run(f'...（共{len(only_in_new_index)}个）')

doc.add_heading('二、字段一致性验证', level=2)
p = doc.add_paragraph()
p.add_run('逐列验证数据中每个字段的值是否与参考文件一致，包括列名一致性、值一致率、不一致的详细分析。')

doc.add_heading('列名一致性', level=3)

table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '项目'
hdr[1].text = '结果'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

cells = table.add_row().cells
cells[0].text = '共同列'
cells[1].text = f'{len(common_cols)}列'

cells = table.add_row().cells
cells[0].text = '仅参考表有'
cells[1].text = f'{len(only_in_ref)}列' if only_in_ref else '无'

cells = table.add_row().cells
cells[0].text = '仅新表有'
cells[1].text = f'{len(only_in_new)}列' if only_in_new else '无'

if only_in_new:
    doc.add_paragraph('')
    p = doc.add_paragraph()
    p.add_run('仅新表有列详情: ').bold = True
    p.add_run(', '.join(only_in_new))

doc.add_heading('字段值一致率总览', level=3)

p = doc.add_paragraph()
p.add_run(f'共同列共{len(common_cols)}列，其中{len(consistent_cols)}列完全一致，{len(inconsistent_cols)}列有差异。')
p.add_run(f'对比范围: 共{total_rows}个共同保单。')

doc.add_heading('不一致列详细分析', level=3)

for col in inconsistent_cols:
    diff_count = diff_mask[col].sum()
    consistent_count = total_rows - diff_count
    consistency_rate = (consistent_count / total_rows * 100) if total_rows > 0 else 0
    
    doc.add_paragraph(f'')
    p = doc.add_paragraph()
    p.add_run(f'{col}（{diff_count}行差异，{consistency_rate:.1f}%一致）').bold = True
    
    sample_diffs = [d for d in diff_details if d['字段名'] == col][:5]
    if sample_diffs:
        doc.add_paragraph('差异样本（标识列: 保单号码):')
        for sample in sample_diffs:
            p = doc.add_paragraph(f"  • 保单{sample['保单号码']}: 参考值='{sample['参考值']}' → 新值='{sample['新值']}'")

doc.add_heading('三、映射规则验证', level=2)
p = doc.add_paragraph()
p.add_run('列出数据中存在的映射规则，并验证其在数据中的实际执行效果。')

doc.add_heading('差异类型分析', level=3)

diff_types = []
for col in inconsistent_cols:
    diff_count = diff_mask[col].sum()
    if diff_count == total_rows:
        diff_types.append((col, '完全差异（可能为格式转换）'))
    elif diff_count > total_rows * 0.5:
        diff_types.append((col, '大部分差异（可能为规则变更）'))
    elif diff_count > total_rows * 0.1:
        diff_types.append((col, '部分差异（可能为数据更新）'))
    else:
        diff_types.append((col, '少量差异（可能为个别数据修正）'))

table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '字段名'
hdr[1].text = '差异类型'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

for col, diff_type in diff_types:
    cells = table.add_row().cells
    cells[0].text = col
    cells[1].text = diff_type

doc.add_heading('四、格式一致性验证', level=2)
p = doc.add_paragraph()
p.add_run('验证数据的格式是否与参考文件一致。')

doc.add_heading('日期格式验证', level=3)

date_analysis = []
for col in date_cols:
    if col in compare_cols:
        diff_count = diff_mask[col].sum()
        consistency_rate = ((total_rows - diff_count) / total_rows * 100) if total_rows > 0 else 0
        date_analysis.append((col, consistency_rate))

table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '日期字段'
hdr[1].text = '一致率'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

for col, rate in date_analysis:
    cells = table.add_row().cells
    cells[0].text = col
    cells[1].text = f'{rate:.1f}%'

doc.add_heading('数值字段精度验证', level=3)

numeric_cols = ['保费', '保费（港币）', 'APE']
numeric_analysis = []
for col in numeric_cols:
    if col in compare_cols:
        diff_count = diff_mask[col].sum()
        consistency_rate = ((total_rows - diff_count) / total_rows * 100) if total_rows > 0 else 0
        numeric_analysis.append((col, diff_count, consistency_rate))

table = doc.add_table(rows=1, cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '数值字段'
hdr[1].text = '差异行数'
hdr[2].text = '一致率'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

for col, diff_count, rate in numeric_analysis:
    cells = table.add_row().cells
    cells[0].text = col
    cells[1].text = str(diff_count)
    cells[2].text = f'{rate:.1f}%'

doc.add_heading('五、综合验证结论', level=2)

if len(inconsistent_cols) == 0:
    p = doc.add_paragraph()
    p.add_run('✅ 验证结论: 所有数据完全一致，验证通过。').bold = True
else:
    p = doc.add_paragraph()
    p.add_run('⚠️ 验证结论: 存在以下需要关注的差异').bold = True
    
    doc.add_paragraph('')
    doc.add_heading('需要关注的差异项', level=3)
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = '字段名'
    hdr[1].text = '差异行数'
    hdr[2].text = '一致率'
    for c in hdr:
        for p in c.paragraphs:
            p.runs[0].bold = True
    
    for col in inconsistent_cols:
        diff_count = diff_mask[col].sum()
        consistency_rate = ((total_rows - diff_count) / total_rows * 100) if total_rows > 0 else 0
        cells = table.add_row().cells
        cells[0].text = col
        cells[1].text = str(diff_count)
        cells[2].text = f'{consistency_rate:.1f}%'

doc.add_heading('六、差异详情', level=2)

if diff_details:
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = '保单号码'
    hdr[1].text = '字段名'
    hdr[2].text = '参考值'
    hdr[3].text = '新值'
    for c in hdr:
        for p in c.paragraphs:
            p.runs[0].bold = True

    for detail in diff_details:
        cells = table.add_row().cells
        cells[0].text = detail['保单号码']
        cells[1].text = detail['字段名']
        cells[2].text = detail['参考值']
        cells[3].text = detail['新值']
else:
    doc.add_paragraph('无数据差异')

output_file = f'永明业绩数据对比报告-{datetime.now().strftime("%Y%m%d")}.docx'
doc.save(output_file)

print(f"\n对比报告已保存: {output_file}")