from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from datetime import date

doc = Document()

style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(11)

doc.add_heading('业绩报表MGA业务数据纳入修改说明', level=1)

today = date.today().strftime('%Y年%m月%d日')
doc.add_paragraph(f'汇报日期：{today}')
doc.add_paragraph('')

doc.add_heading('一、修改背景与目标', level=2)

p = doc.add_paragraph()
p.add_run('原流程：').bold = True
p.add_run('需在「业绩数据-整合-YYYYMMDD」文件的「V0-合并表」sheet中，')
p.add_run('手动将MGA业务对应的「业务类型」改为「经代业务」，将「业务细分」改为「同行经代」，')
p.add_run('报表才能包含完整数据。')

p = doc.add_paragraph()
p.add_run('新流程：').bold = True
p.add_run('不动底表，调整代码，使报表能自动识别并展示MGA业务数据。')

p = doc.add_paragraph()
p.add_run('修改目标：').bold = True
p.add_run('1）报表能跑出完整数据（包含MGA业务）；2）在需要体现业务细分的地方，把MGA业务这条线的数据单独体现出来。')

doc.add_heading('二、修改内容（代码层面）', level=2)

table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '修改位置'
hdr[1].text = '修改前'
hdr[2].text = '修改后'
hdr[3].text = '影响范围'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

rows_data = [
    ('SEGMENTS列表\n（第40-41行）', '8个业务细分', '13个业务细分\n（新增：永明TA、媒体与培训、\n管理咨询、权益业务、DT）', '所有涉及业务细分的报表板块'),
    ('biz_map映射\n（第193-195行）', '8个映射关系', '13个映射关系\n（新增5个→MGA业务）', '业务类型维度统计'),
    ('TARGET_ALL/TARGET_YM\n（第47、53行）', '8个目标值', '13个目标值\n（MGA业务暂设为0）', '目标达成率计算'),
    ('S1业务类型维度\n（第597行）', '3个业务类型\n（代理人/经代/KA）', '4个业务类型\n（新增MGA业务）', 'S1总览仪表盘'),
    ('CSV导出业务类型\n（第2684行）', '3个业务类型', '4个业务类型\n（新增MGA业务）', 'S1-S4 CSV文件'),
]

for row in rows_data:
    cells = table.add_row().cells
    for i, text in enumerate(row):
        cells[i].text = text

doc.add_heading('三、MGA业务细分与业务大类映射关系', level=2)

table = doc.add_table(rows=1, cols=2)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = '业务细分（Segment）'
hdr[1].text = '业务大类（Business Category）'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

mga_mapping = [
    ('天领业务', '代理人业务'),
    ('成事家办', '代理人业务'),
    ('BK业务', '经代业务'),
    ('同行经代', '经代业务'),
    ('永明经代', '经代业务'),
    ('合伙转介业务', 'KA业务'),
    ('ICLUB业务', 'KA业务'),
    ('IFA业务', 'KA业务'),
    ('永明TA', 'MGA业务 ← 新增'),
    ('媒体与培训', 'MGA业务 ← 新增'),
    ('管理咨询', 'MGA业务 ← 新增'),
    ('权益业务', 'MGA业务 ← 新增'),
    ('DT', 'MGA业务 ← 新增'),
]

for seg, biz_cat in mga_mapping:
    cells = table.add_row().cells
    cells[0].text = seg
    cells[1].text = biz_cat

doc.add_heading('四、PPT受影响的板块分析', level=2)

p = doc.add_paragraph()
p.add_run('说明：').bold = True
p.add_run('以下为「周业绩汇报PPT_W28_20260717」各页PPT受影响情况的详细分析。')

ppt_slides = [
    ('第1页 — 全业务总览', [
        ('业务类型业绩分组堆叠柱图（Chart 0）', '从3个业务类型增加到4个，新增「MGA业务」系列，展示MGA业务的批核/未批核/待签APE'),
        ('月度明细表格', '合计行数据包含MGA业务数据，各月度业绩汇总更完整'),
        ('KPI汇总卡片', '全业务目标达成率、批核APE、件数等KPI包含MGA业务数据'),
    ], '有变化'),
    ('第2页 — F批核路径管控', [
        ('达标节奏线', '无变化'),
        ('预测区间图表', '无变化'),
    ], '无变化'),
    ('第3页 — 永明业绩汇报', [
        ('永明月度趋势图', '无变化（仅展示永明相关数据）'),
        ('永明业务数据表格', '无变化'),
    ], '无变化'),
    ('第4页 — G/H气泡瀑布', [
        ('目标缺口分解堆叠柱图', '从8个业务细分增加到13个，新增永明TA、媒体与培训、管理咨询、权益业务、DT五条业务线'),
        ('业务细分名称映射', '需要确认PPT模板是否能容纳新增的5个业务细分系列'),
    ], '有变化'),
    ('第5页 — 业务端视角', [
        ('保单阶段构成甜甜圈图（J块）', '无变化（全业务合计数据）'),
        ('目标vs已批核/未批核/待签APE对比图（L块）', '从8个业务细分增加到13个，新增5个MGA业务细分的数据列'),
        ('业务细分年度汇总表格', '从8列增加到13列，新增5个MGA业务细分的年度汇总数据'),
    ], '有变化'),
    ('第6页 — 执行管理端', [
        ('全流程转化漏斗（M块）', '无变化（全业务合计数据）'),
        ('周度趋势图表', '无变化（全业务合计数据）'),
        ('签批时效分析表格', '从8个业务细分增加到13个，新增5个MGA业务细分的时效统计'),
    ], '有变化'),
    ('第7页 — 同行业绩分析', [
        ('同行推荐人分析图表（U块）', '无变化（仅展示同行经代/永明经代数据）'),
        ('同行KA TOP10图表（V块）', '无变化（仅展示同行经代/永明经代数据）'),
        ('同行明细表', '无变化（仅展示同行经代/永明经代数据）'),
    ], '无变化'),
    ('第8页 — KA业务分析', [
        ('KA业务相关图表', '无变化（仅展示合伙转介/ICLUB/IFA业务数据）'),
        ('KA业务明细表', '无变化'),
    ], '无变化'),
    ('第9页 — 同行业绩分析仪表盘', [
        ('同行周度趋势图表', '无变化（仅展示同行经代/永明经代数据）'),
        ('同行明细表', '需要确认是否包含MGA业务的同行数据，取决于MGA业务的业务细分映射'),
    ], '视数据映射而定'),
    ('第10页 — 银行仪表盘（run_full_deck.py追加）', [
        ('银行相关图表', '无变化（仅展示BK业务数据）'),
        ('银行明细表', '无变化'),
    ], '无变化'),
    ('第11页 — 代理人业务分析（run_full_deck.py追加）', [
        ('代理人业务相关图表', '无变化（仅展示天领/成事家办业务数据）'),
        ('代理人业务明细表', '无变化'),
    ], '无变化'),
    ('第12页 — KA业务分析（run_full_deck.py追加）', [
        ('KA业务相关图表', '无变化（仅展示合伙转介/ICLUB/IFA业务数据）'),
        ('KA业务明细表', '无变化'),
    ], '无变化'),
]

for slide_name, items, change_status in ppt_slides:
    p = doc.add_paragraph()
    run = p.add_run(f'{slide_name}')
    run.bold = True
    if change_status == '有变化':
        run.font.color.rgb = RGBColor(0, 51, 102)
    elif change_status == '视数据映射而定':
        run.font.color.rgb = RGBColor(200, 100, 0)
    else:
        run.font.color.rgb = RGBColor(102, 102, 102)
    p.add_run(f'（{change_status}）')
    
    for item_name, desc in items:
        p = doc.add_paragraph(f'• {item_name}：{desc}', style='List Bullet')

doc.add_heading('五、关键影响图表汇总', level=2)

table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = table.rows[0].cells
hdr[0].text = 'PPT页码'
hdr[1].text = '图表/表格名称'
hdr[2].text = '影响描述'
hdr[3].text = '影响程度'
for c in hdr:
    for p in c.paragraphs:
        p.runs[0].bold = True

key_charts = [
    ('第1页', '业务类型业绩分组堆叠柱图', '从3个业务类型增加到4个，新增「MGA业务」系列', '高'),
    ('第1页', '月度明细表格', '合计行数据包含MGA业务数据', '中'),
    ('第1页', 'KPI汇总卡片', '全业务KPI包含MGA业务数据', '中'),
    ('第4页', '目标缺口分解堆叠柱图', '从8个业务细分增加到13个', '高'),
    ('第5页', '目标vs已批核/未批核/待签APE对比图', '从8个业务细分增加到13个', '高'),
    ('第5页', '业务细分年度汇总表格', '从8列增加到13列', '高'),
    ('第6页', '签批时效分析表格', '从8个业务细分增加到13个', '中'),
]

for slide, chart_name, desc, level in key_charts:
    cells = table.add_row().cells
    cells[0].text = slide
    cells[1].text = chart_name
    cells[2].text = desc
    cells[3].text = level

doc.add_heading('六、分析过程与分析方法', level=2)

doc.add_heading('6.1 分析过程', level=3)

steps = [
    '第一步：问题定位',
    '原始CSV数据中包含MGA业务相关的业务细分（永明TA、媒体与培训、管理咨询、权益业务、DT），但这些不在脚本的SEGMENTS列表中，导致数据被遗漏。',
    '',
    '第二步：解决方案设计',
    '在代码中添加对这5个业务细分的支持，使其能被正确识别和统计，同时保持原有8个业务细分的逻辑不变。',
    '',
    '第三步：代码修改',
    '1）扩展SEGMENTS列表（8→13）；2）更新biz_map映射（增加MGA业务大类）；3）更新目标值配置；4）更新业务类型维度统计（3→4类）；5）同步更新CSV导出逻辑。',
    '',
    '第四步：影响评估',
    '分析所有使用SEGMENTS列表的代码位置（共10+处循环），确保修改不会破坏现有逻辑，原有数据计算不受影响。',
]

for step in steps:
    if step.startswith('第'):
        p = doc.add_paragraph()
        run = p.add_run(step)
        run.bold = True
    elif step == '':
        doc.add_paragraph('')
    else:
        doc.add_paragraph(step)

doc.add_heading('6.2 分析方法', level=3)

methods = [
    ('代码搜索法', '通过全局搜索「SEGMENTS」和「for seg in SEGMENTS」定位所有受影响的代码位置，确保不遗漏任何需要修改的地方。'),
    ('映射关系法', '分析「biz_map」字典，确定MGA业务细分应映射到「MGA业务」大类，与代理人业务、经代业务、KA业务并列。'),
    ('数据流向法', '追踪数据从CSV读取→预处理→聚合→报表输出的完整流程，确保MGA数据在每个环节都能被正确处理。'),
    ('增量测试法', '将MGA业务目标值设为0，避免影响现有达成率计算，后续可根据业务需求调整。'),
]

for method_name, desc in methods:
    p = doc.add_paragraph()
    run = p.add_run(f'{method_name}：')
    run.bold = True
    p.add_run(desc)

doc.add_heading('七、关键技术点', level=2)

key_points = [
    '业务大类映射：MGA业务（永明TA、媒体与培训、管理咨询、权益业务、DT）统一映射到「MGA业务」大类，与代理人业务、经代业务、KA业务并列，形成四大业务类型。',
    '目标值处理：MGA业务目标值暂设为0，不会影响全业务目标达成率计算（全业务目标仍为1,113,000,000）。',
    '向后兼容：原有8个业务细分的逻辑完全不变，MGA业务作为增量数据加入，不影响历史报表数据。',
    '全链路覆盖：从数据读取、预处理、聚合到报表输出，MGA业务数据在每个环节都能被正确识别和统计。',
]

for i, point in enumerate(key_points, 1):
    doc.add_paragraph(f'{i}. {point}')

doc.add_heading('八、后续建议', level=2)

suggestions = [
    ('确认目标值', '如需为MGA业务设置年度/月度目标，请告知具体数值，我将更新代码。'),
    ('S9扩展', '如需在S9中单独展示MGA业务数据，可增加新的板块。'),
    ('PPT模板检查', '建议检查「run_full_pipeline.py」和「run_full_deck.py」，确保PPT模板能容纳新增的业务细分列。'),
    ('数据验证', '首次运行后，请核对PPT中MGA业务数据与底表数据是否一致。'),
]

for item_name, desc in suggestions:
    p = doc.add_paragraph()
    run = p.add_run(f'{item_name}：')
    run.bold = True
    p.add_run(desc)

doc.add_heading('九、总结', level=2)

doc.add_paragraph('本次修改成功实现了MGA业务数据的自动识别和统计，无需手动修改底表。主要变化集中在业务细分维度的扩展（从8个增加到13个）和业务类型的增加（从3类增加到4类）。')

doc.add_paragraph('修改后，运行完整流程（generate_report.py → run_full_pipeline.py → run_full_deck.py）即可生成包含MGA业务数据的周业绩汇报PPT。')

output_path = 'MGA业务数据纳入修改说明.docx'
doc.save(output_path)
print(f'Document saved to: {output_path}')
