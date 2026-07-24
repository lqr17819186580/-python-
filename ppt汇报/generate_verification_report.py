import docx
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = docx.Document()
doc.styles['Normal'].font.name = '微软雅黑'
doc.styles['Normal'].font.size = docx.shared.Pt(11)

title = doc.add_heading('MGA业务验证报告清单', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_heading('一、PPT页面与板块影响清单', level=1)

doc.add_paragraph('说明：以下清单详细列出MGA业务作为独立业务线后，对PPT各页面及板块的影响，以及对应的脚本修改点。')

pages_info = [
    ('Page 1', '全维度业绩分析仪表盘', 'A-目标达成率', 'MGA业务数据纳入全业务达成率计算', '需要修改', 
     'update_ppt.py第2126行SLIDE1_SUBS，generate_report.py第45行TARGET_ALL、第50行TARGET_YM',
     'generate_report.py第169行APE数值转换修复、第40行SEGMENTS添加MGA业务'),
    ('Page 1', '全维度业绩分析仪表盘', 'B-保单阶段分布', 'MGA业务数据纳入全业务统计', '需要修改',
     'update_ppt.py第2126行SLIDE1_SUBS',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 1', '全维度业绩分析仪表盘', 'D-业务线月度趋势图', 'MGA业务作为独立业务线显示，月度数据包含MGA', '需要修改',
     'update_ppt.py第630行slide1_chart_business_type(S1)、第810行图表修复',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 1', '全维度业绩分析仪表盘', 'K-业务细分甜甜圈', 'MGA业务作为独立扇区显示', '需要修改',
     'chart_updates.py第41行slide1_chart_business_type函数',
     'generate_report.py第193行biz_map添加MGA业务映射'),
    ('Page 2', 'F批核路径管控', 'C-KPI卡片', '全业务KPI包含MGA业务数据', '需要修改',
     'update_ppt.py第2208行SLIDE2_SUBS',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 2', 'F批核路径管控', 'E-月度三线图', '预约/签单/批核数据包含MGA', '需要修改',
     'update_ppt.py第848-850行slide2_appointment/slide2_signed/slide2_approved',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 2', 'F批核路径管控', 'F-预测图', '月度数据包含MGA业务', '需要修改',
     'update_ppt.py第847行slide2_forecast_chart(S1)',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 3', '永明业绩汇报', 'G-牌照表', '牌照数据统计包含MGA业务', '需要修改',
     'update_ppt.py第3392行牌照表写入逻辑',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 3', '永明业绩汇报', 'H-缺口瀑布图', '缺口计算包含MGA业务', '需要修改',
     'update_ppt.py第1524-1526行Waterfall变量_wv_mga',
     'update_ppt.py第221行CH_MGA = _channel_kpis(S2["A"], "MGA业务")'),
    ('Page 3', '永明业绩汇报', 'I-TOP10产品表', '永明产品排名包含MGA业务', '需要修改',
     'update_ppt.py第3721行TOP10产品排名',
     'generate_report.py第169行APE数值转换修复'),
    ('Page 4', '业务端视角仪表盘', 'I-业务线月度趋势', 'MGA业务作为独立业务线显示', '需要修改',
     'update_ppt.py第1318行slide4_channel_trend(S2, seg)、第1296行_I_CHART_MAP',
     'generate_report.py第2745行SEGS(S2 CSV)添加MGA业务'),
    ('Page 4', '业务端视角仪表盘', 'G-气泡图', 'MGA业务作为独立气泡显示', '需要修改',
     'update_ppt.py第1451行_G_SEGS添加("MGA业务", "Shape 81", "Text 82")',
     'generate_report.py第2745行SEGS(S2 CSV)添加MGA业务'),
    ('Page 4', '业务端视角仪表盘', 'H-瀑布图', '瀑布图包含MGA业务数据', '需要修改',
     'update_ppt.py第1524-1526行Waterfall变量_wv_mga及缺口计算',
     'update_ppt.py第221行CH_MGA = _channel_kpis(S2["A"], "MGA业务")'),
    ('Page 5', '业务端视角仪表盘', 'K-甜甜圈', '9个业务细分包含MGA业务', '需要修改',
     'update_ppt.py第1133行slide5_donut(S2)',
     'generate_report.py第2745行SEGS(S2 CSV)添加MGA业务'),
    ('Page 5', '业务端视角仪表盘', 'L-业务细分柱状图', '9个业务细分柱状图包含MGA业务', '需要修改',
     'update_ppt.py第1028行_L_SEGMENTS添加"MGA业务"',
     'generate_report.py第2745行SEGS(S2 CSV)添加MGA业务'),
    ('Page 6', '执行管理端分析仪表盘', 'Q/R/S-热力图', '周度数据包含MGA业务', '需要修改',
     'update_ppt.py第1708行slide6_weekly_trend(S3)',
     'generate_report.py第2974行SEGS(S3 CSV)添加MGA业务'),
    ('Page 7', '执行管理端分析仪表盘', 'Q/R/S-热力矩阵', 'MGA业务的时效数据', '需要修改',
     'update_ppt.py第3165行热力矩阵克隆逻辑',
     'generate_report.py第2974行SEGS(S3 CSV)添加MGA业务'),
    ('Page 7', '执行管理端分析仪表盘', 'T-时效表', 'MGA业务时效统计', '需要修改',
     'update_ppt.py第3165行热力矩阵数据填充',
     'generate_report.py第2974行SEGS(S3 CSV)添加MGA业务'),
    ('Page 8', '同行业绩分析仪表盘', 'V-KA业绩表', '同行业绩分析包含MGA业务', '需要修改',
     'update_ppt.py第1885行_update_slide8_peer_ka_table',
     'generate_report.py第901行mask_tonghang添加MGA业务'),
    ('Page 9', '同行业绩分析仪表盘', '同行周度走势', '无直接影响', '无需修改', '-', '-'),
    ('Page 10', 'BK批核', '银行仪表盘', '无直接影响', '无需修改', '-', '-'),
    ('Page 11', '代理人业务分析仪表盘', '代理人业务分析', '无直接影响', '无需修改', '-', '-'),
    ('Page 12', 'KA业务分析仪表盘', 'KA业务分析', '无直接影响', '无需修改', '-', '-'),
]
table = doc.add_table(rows=len(pages_info)+1, cols=7)
table.style = 'Table Grid'
table.autofit = False
table.allow_autofit = False
headers = ['Page', '页面标题', '板块', '影响说明', '是否修改', 'PPT脚本修改', 'Excel脚本修改']
col_widths = [Cm(1.2), Cm(2.5), Cm(1.5), Cm(4.5), Cm(1.2), Cm(3.0), Cm(3.0)]
for i, h in enumerate(headers):
    table.cell(0, i).text = h
    table.cell(0, i).width = col_widths[i]
for i, (page, title, block, effect, need_modify, ppt_script, excel_script) in enumerate(pages_info, 1):
    table.cell(i, 0).text = page
    table.cell(i, 0).width = col_widths[0]
    table.cell(i, 1).text = title
    table.cell(i, 1).width = col_widths[1]
    table.cell(i, 2).text = block
    table.cell(i, 2).width = col_widths[2]
    table.cell(i, 3).text = effect
    table.cell(i, 3).width = col_widths[3]
    table.cell(i, 4).text = need_modify
    table.cell(i, 4).width = col_widths[4]
    table.cell(i, 5).text = ppt_script
    table.cell(i, 5).width = col_widths[5]
    table.cell(i, 6).text = excel_script
    table.cell(i, 6).width = col_widths[6]

doc.add_heading('二、Excel脚本修改清单 (generate_report.py)', level=1)
excel_changes = [
    ('第40行', 'SEGMENTS', '添加"MGA业务"到业务细分列表（共9个）', 
     'SEGMENTS = [\'天领业务\',\'成事家办\',\'BK业务\',\'同行经代\',\'永明经代\',\'合伙转介业务\',\'ICLUB业务\',\'IFA业务\',\'MGA业务\']'),
    ('第45行', 'TARGET_ALL', '添加"MGA业务":0', '\'MGA业务\':0'),
    ('第50行', 'TARGET_YM', '添加"MGA业务":0', '\'MGA业务\':0'),
    ('第193行', 'biz_map', '添加"MGA业务":"MGA业务"（独立业务类型）', '\'MGA业务\':\'MGA业务\'（独立映射）'),
    ('第169行', 'APE数值转换', '修复ape字段含逗号导致转换失败', 
     'df[col] = pd.to_numeric(df[col].astype(str).str.replace(\',\', \'\'), errors=\'coerce\').fillna(0)'),
    ('第901行', 'mask_tonghang', '同行业绩分析添加"MGA业务"', 'mask_tonghang = df[\'业务细分\'].isin([...,\'MGA业务\'])'),
    ('第2745行', 'SEGS（S2 CSV）', '添加"MGA业务"', 'SEGS列表添加"MGA业务"'),
    ('第2747行', 'TALL（S2 CSV）', '添加"MGA业务":0', '\'MGA业务\':0'),
    ('第2749行', 'TYM（S2 CSV）', '添加"MGA业务":0', '\'MGA业务\':0'),
    ('第2871行', '同行业绩分析（S2 CSV）', 'mask添加"MGA业务"', 'mask包含"MGA业务"'),
    ('第2974行', 'SEGS（S3 CSV）', '添加"MGA业务"', 'SEGS列表添加"MGA业务"'),
]
table = doc.add_table(rows=len(excel_changes)+1, cols=4)
table.style = 'Table Grid'
table.cell(0, 0).text = '位置'
table.cell(0, 1).text = '变量/函数'
table.cell(0, 2).text = '修改内容'
table.cell(0, 3).text = '代码示例'
for i, (pos, var, desc, code) in enumerate(excel_changes, 1):
    table.cell(i, 0).text = pos
    table.cell(i, 1).text = var
    table.cell(i, 2).text = desc
    table.cell(i, 3).text = code

doc.add_heading('三、PPT脚本修改清单', level=1)

doc.add_heading('3.1 update_ppt.py 修改点', level=2)
ppt_changes = [
    ('第221行', 'CH_MGA', '添加MGA业务的channel_kpis调用', 
     'CH_MGA = _channel_kpis(S2["A"], "MGA业务")'),
    ('第1028行', '_L_SEGMENTS', '添加"MGA业务"到业务细分列表（9个）', 
     '_L_SEGMENTS = ["BK业务","永明经代","同行经代","天领业务","ICLUB业务","成事家办","合伙转介业务","IFA业务","MGA业务"]'),
    ('第1284行', 'CHANNEL_ORDER', '添加"MGA业务"到通道顺序列表', 
     'CHANNEL_ORDER = ["永明经代","天领业务","BK业务","合伙转介业务","成事家办","同行经代","ICLUB业务","MGA业务"]'),
    ('第1451行', '_G_SEGS', '添加"MGA业务"气泡图Shape映射', 
     '("MGA业务", "Shape 81", "Text 82")'),
    ('第1524行', 'Waterfall', '添加_wv_mga变量', '_wv_mga = CH_MGA["issued_m"]'),
    ('第1526行', 'Waterfall gap', '缺口计算包含MGA业务', 
     '_wv_gap = _WF_MAX - (_wv_bk+_wv_ym+_wv_th+_wv_tl+_wv_ic+_wv_cs+_wv_hh+_wv_mga+_wv_ub+_wv_pd)'),
]
table = doc.add_table(rows=len(ppt_changes)+1, cols=4)
table.style = 'Table Grid'
table.cell(0, 0).text = '位置'
table.cell(0, 1).text = '变量/函数'
table.cell(0, 2).text = '修改内容'
table.cell(0, 3).text = '代码示例'
for i, (pos, var, desc, code) in enumerate(ppt_changes, 1):
    table.cell(i, 0).text = pos
    table.cell(i, 1).text = var
    table.cell(i, 2).text = desc
    table.cell(i, 3).text = code

doc.add_heading('3.2 chart_updates.py 修改点', level=2)
chart_changes = [
    ('第41行', 'slide1_chart_business_type', '业务类型图表数据包含MGA业务', 
     'cats = ["经代业务", "代理人业务", "KA 业务", "MGA业务"]'),
]
table = doc.add_table(rows=len(chart_changes)+1, cols=4)
table.style = 'Table Grid'
table.cell(0, 0).text = '位置'
table.cell(0, 1).text = '函数'
table.cell(0, 2).text = '修改内容'
table.cell(0, 3).text = '代码示例'
for i, (pos, func, desc, code) in enumerate(chart_changes, 1):
    table.cell(i, 0).text = pos
    table.cell(i, 1).text = func
    table.cell(i, 2).text = desc
    table.cell(i, 3).text = code

doc.add_heading('四、数据源对比', level=1)

doc.add_heading('4.1 文件信息', level=2)
info = [
    ('底表文件', '业绩数据底表-20260717.csv'),
    ('参考文件', '业绩数据20260717.csv'),
    ('总行数', '3,504行（两者一致）'),
]
table = doc.add_table(rows=len(info), cols=2)
table.style = 'Table Grid'
for i, (label, value) in enumerate(info):
    table.cell(i, 0).text = label
    table.cell(i, 1).text = value

doc.add_heading('4.2 业务类型分布对比', level=2)
doc.add_paragraph('底表业务类型：')
base_biz = ['经代业务：1,592条', '代理人业务：884条', 'KA业务：855条', 'MGA业务：172条（新增独立业务线）']
for item in base_biz:
    doc.add_paragraph(f'• {item}', style='List Bullet')

doc.add_paragraph('参考表业务类型（手工处理后）：')
ref_biz = ['经代业务：1,764条（含MGA业务数据）', '代理人业务：884条', 'KA业务：855条']
for item in ref_biz:
    doc.add_paragraph(f'• {item}', style='List Bullet')

doc.add_heading('4.3 MGA业务数据详情', level=2)
mga_info = [
    ('业务类型', 'MGA业务（独立业务线）'),
    ('业务细分', 'MGA业务'),
    ('记录数', '172条'),
    ('生效记录数', '80条'),
    ('批核APE', '19,377,929港元'),
    ('未批核APE', '5,823,000港元'),
]
table = doc.add_table(rows=len(mga_info), cols=2)
table.style = 'Table Grid'
for i, (label, value) in enumerate(mga_info):
    table.cell(i, 0).text = label
    table.cell(i, 1).text = value

doc.add_heading('4.4 关键指标对比', level=2)
metrics = [
    ('全业务达成率', '54.4%', '54.4%', '一致'),
    ('全业务批核APE', '605.2M', '605.2M', '一致'),
    ('永明达成率', '52.4%', '52.4%', '一致'),
    ('PPT页数', '12页', '12页', '一致'),
]
table = doc.add_table(rows=len(metrics)+1, cols=4)
table.style = 'Table Grid'
table.cell(0, 0).text = '指标'
table.cell(0, 1).text = '底表生成'
table.cell(0, 2).text = '参考文件'
table.cell(0, 3).text = '对比结果'
for i, (name, base, ref, result) in enumerate(metrics, 1):
    table.cell(i, 0).text = name
    table.cell(i, 1).text = base
    table.cell(i, 2).text = ref
    table.cell(i, 3).text = result

doc.add_heading('五、完整执行流程', level=1)
doc.add_paragraph('生成最终PPT需要运行以下三个脚本：')
steps = [
    ('Step 1', 'python generate_report.py 业绩数据底表-20260717.csv', '生成Excel报表和S1-S4 CSV附表'),
    ('Step 2', 'python run_full_pipeline.py', '更新PPT图表和文本（生成11页PPT）'),
    ('Step 3', 'python run_full_deck.py', '追加代理人/KA页面（生成12页PPT）'),
]
table = doc.add_table(rows=len(steps)+1, cols=3)
table.style = 'Table Grid'
table.cell(0, 0).text = '步骤'
table.cell(0, 1).text = '命令'
table.cell(0, 2).text = '说明'
for i, (step, cmd, desc) in enumerate(steps, 1):
    table.cell(i, 0).text = step
    table.cell(i, 1).text = cmd
    table.cell(i, 2).text = desc

doc.add_heading('六、最终PPT页面结构（12页）', level=1)
pages = [
    ('Page 1', '全维度业绩分析仪表盘', '全维度业绩分析仪表盘', '一致'),
    ('Page 2', '月度效能深析', '月度效能深析', '一致'),
    ('Page 3', '永明业绩汇报', '永明业绩汇报', '一致'),
    ('Page 4', '业务端视角仪表盘', '业务端视角仪表盘', '一致'),
    ('Page 5', '业务端视角仪表盘', '业务端视角仪表盘', '一致'),
    ('Page 6', '执行管理端分析仪表盘', '执行管理端分析仪表盘', '一致'),
    ('Page 7', '执行管理端分析仪表盘', '执行管理端分析仪表盘', '一致'),
    ('Page 8', '同行业绩分析仪表盘', '同行业绩分析仪表盘', '一致'),
    ('Page 9', '同行业绩分析仪表盘', '同行业绩分析仪表盘', '一致'),
    ('Page 10', 'BK批核', 'BK批核', '一致'),
    ('Page 11', '代理人业务分析仪表盘', '代理人业务分析仪表盘', '一致'),
    ('Page 12', 'KA业务分析仪表盘', 'KA业务分析仪表盘', '一致'),
]
table = doc.add_table(rows=len(pages)+1, cols=4)
table.style = 'Table Grid'
table.autofit = False
table.allow_autofit = False
col_widths = [Cm(1.5), Cm(3.5), Cm(3.5), Cm(2.0)]
table.cell(0, 0).text = '页码'
table.cell(0, 0).width = col_widths[0]
table.cell(0, 1).text = '底表生成页面标题'
table.cell(0, 1).width = col_widths[1]
table.cell(0, 2).text = '参考文件页面标题'
table.cell(0, 2).width = col_widths[2]
table.cell(0, 3).text = '对比结果'
table.cell(0, 3).width = col_widths[3]
for i, (page, base_title, ref_title, result) in enumerate(pages, 1):
    table.cell(i, 0).text = page
    table.cell(i, 0).width = col_widths[0]
    table.cell(i, 1).text = base_title
    table.cell(i, 1).width = col_widths[1]
    table.cell(i, 2).text = ref_title
    table.cell(i, 2).width = col_widths[2]
    table.cell(i, 3).text = result
    table.cell(i, 3).width = col_widths[3]

doc.add_heading('七、结论', level=1)

doc.add_heading('7.1 核心问题', level=2)
doc.add_paragraph('底表中MGA业务（172条记录）在代码中没有被识别为有效业务细分，且ape字段因包含逗号（如"546,000.00"）导致数值转换失败，所有批核APE统计为0。')

doc.add_heading('7.2 修复方案', level=2)
solutions = [
    '将MGA业务作为独立业务线添加到所有业务细分列表中',
    '在biz_map中将MGA业务映射为独立的"MGA业务"类型',
    '修复ape字段数值转换逻辑，支持含逗号的数字格式',
    '在同行业绩分析中包含MGA业务',
    '在PPT生成脚本中添加MGA业务的数据处理和图表显示',
    '运行run_full_deck.py完成最终12页PPT生成',
]
for item in solutions:
    doc.add_paragraph(f'• {item}', style='List Bullet')

doc.add_heading('7.3 预期效果', level=2)
effects = [
    '全业务达成率从0.0%恢复到54.4%（与参考文件一致）',
    'MGA业务批核APE=19.4M，未批核APE=5.8M',
    '所有报表和PPT中MGA业务作为独立业务线呈现',
    '同行业绩分析包含MGA业务数据',
    '最终PPT共12页，与参考文件结构完全一致',
]
for item in effects:
    doc.add_paragraph(f'• {item}', style='List Bullet')

doc.add_heading('7.4 验证结果', level=2)
doc.add_paragraph('生成的周业绩汇报PPT_FINAL.pptx与参考文件周业绩汇报PPT_W28_20260717.pptx数据一致，MGA业务作为独立业务线在所有相关页面正确显示，PPT共12页。')

doc.add_paragraph('报告生成时间：2026年7月24日')

doc.save(r'D:\数据中台支持\业绩整合自动化脚本\ppt汇报\MGA业务验证报告清单.docx')
print('Word文档生成成功！')