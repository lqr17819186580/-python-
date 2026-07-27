import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

data = [
    {
        '截图': '',
        'PPT的P几': 'P1',
        '什么板块': 'C. 业务类型业绩 — 渠道批核/未批核/待签（APE in M）',
        'PPT的影响说明': '参考文件图表0只有3个业务类型（经代业务、代理人业务、KA业务）；测试文件图表0增加了MGA业务，共4个业务类型',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的1-总览仪表盘 sheet的G. 业务类型维度-2026 | Business Type Dimension',
        'excel的影响说明': 'S1-总览仪表盘.csv的G部分已包含MGA业务数据：MGA业务 19,377,930 80 23,967,148 65 9,235,800 21',
        '判断': '已修改',
        '为什么这么判断': 'generate_report.py已将MGA业务加入biz_map映射，S1-G部分数据已包含MGA；chart_updates.py的slide1_chart_business_type函数已包含MGA业务分类',
        '怎么改': '无需修改，已完成'
    },
    {
        '截图': '',
        'PPT的P几': 'P1',
        '什么板块': 'D. 月度业绩走势—按预约时间 | Monthly by Reservation Date',
        'PPT的影响说明': '参考文件和测试文件图表数据一致，均为全业务月度趋势汇总，不按业务类型细分',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的1-总览仪表盘 sheet的C. 月度业绩走势—按预约时间',
        'excel的影响说明': 'S1-C部分为全业务汇总数据，不按业务类型拆分，MGA业务已包含在合计中',
        '判断': '无需改',
        '为什么这么判断': '该图表为全业务汇总趋势，不区分业务类型，MGA业务数据已自然包含在合计中',
        '怎么改': '无需修改'
    },
    {
        '截图': '',
        'PPT的P几': 'P4',
        '什么板块': 'I. 业务线月度三线趋势 — 预约/签单/批核',
        'PPT的影响说明': '参考文件有7个业务线图表（天领业务、成事家办、BK业务、同行经代、永明经代、合伙转介业务、ICLUB业务）；测试文件增加了MGA业务图表C9041，共8个',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的C-APE/D-APE/E-APE. 月度预约/签单/批核业绩—全业务',
        'excel的影响说明': 'S2-C-APE/D-APE/E-APE部分已包含MGA业务数据，2026-01至2026-07各月数据完整',
        '判断': '已修改',
        '为什么这么判断': 'update_ppt.py已添加自动复制C9037图表创建C9041的逻辑，用于展示MGA业务月度趋势',
        '怎么改': '无需修改，已完成'
    },
    {
        '截图': '',
        'PPT的P几': 'P4',
        '什么板块': 'G. 达成率×目标规模×件数 — 业务线战略象限气泡图',
        'PPT的影响说明': '参考文件气泡图包含7个业务线；测试文件已添加MGA业务气泡展示',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的A. 业务细分年度汇总—全业务',
        'excel的影响说明': 'S2-A部分已包含MGA业务数据：目标APE=0，批核APE=19,377,930',
        '判断': '已修改',
        '为什么这么判断': 'update_ppt.py已更新_G_SEGS常量，添加了MGA业务到气泡图分类中',
        '怎么改': '无需修改，已完成'
    },
    {
        '截图': '',
        'PPT的P几': 'P4',
        '什么板块': 'H. 全业务目标缺口分解（瀑布图）',
        'PPT的影响说明': '参考文件瀑布图包含7个业务线；测试文件已添加MGA业务的缺口展示',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的A. 业务细分年度汇总—全业务',
        'excel的影响说明': 'S2-A部分已包含MGA业务数据，目标缺口=0-19,377,930=-19,377,930',
        '判断': '已修改',
        '为什么这么判断': 'update_ppt.py已添加自动复制Text_314标签的逻辑，用于MGA业务的缺口展示',
        '怎么改': '无需修改，已完成'
    },
    {
        '截图': '',
        'PPT的P几': 'P5',
        '什么板块': 'L. 目标 vs 已批核/未批核/待签 APE — 各业务线对比',
        'PPT的影响说明': '参考文件图表1包含5个业务线（永明经代、天领业务、BK业务、合伙转介业务、成事家办）；测试文件数据有变化，但图表中未直接显示MGA业务',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的A. 业务细分年度汇总—全业务',
        'excel的影响说明': 'S2-A部分已包含MGA业务数据，但chart_updates.py的slide5_target_vs_actual函数只筛选了8个业务细分，未包含MGA业务',
        '判断': '改PPT',
        '为什么这么判断': '数据层（generate_report.py）已包含MGA业务，但chart_updates.py的slide5_target_vs_actual函数在筛选业务细分时未包含MGA业务',
        '怎么改': '在chart_updates.py的slide5_target_vs_actual函数中，将MGA业务添加到df筛选列表和order列表中'
    },
    {
        '截图': '',
        'PPT的P几': 'P5',
        '什么板块': 'K. 甜甜圈: 已批核/未批核/待签（全业务合计）',
        'PPT的影响说明': '参考文件和测试文件数据一致，均为全业务合计数据',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的A. 业务细分年度汇总—全业务',
        'excel的影响说明': 'S2-A合计行已包含MGA业务数据（合计批核APE=605,158,199）',
        '判断': '无需改',
        '为什么这么判断': '该图表为全业务合计，MGA业务数据已自然包含在合计中',
        '怎么改': '无需修改'
    },
    {
        '截图': '',
        'PPT的P几': 'P7',
        '什么板块': '执行管理端分析仪表盘',
        'PPT的影响说明': '参考文件和测试文件均无图表，测试文件显示包含MGA文本',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的3-执行管理端 sheet',
        'excel的影响说明': 'S3各部分（A-APE/B-APE/C-APE等）已包含MGA业务数据',
        '判断': '无需改',
        '为什么这么判断': '该页面无图表，文本中包含MGA是正常的，因为数据已更新',
        '怎么改': '无需修改'
    },
    {
        '截图': '',
        'PPT的P几': 'P8',
        '什么板块': 'U. 同行推荐人分析',
        'PPT的影响说明': '参考文件和测试文件数据一致，无MGA业务展示',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的J. 同行推荐人分析',
        'excel的影响说明': 'S2-J部分为同行推荐人数据，与MGA业务无关',
        '判断': '无需改',
        '为什么这么判断': '该板块为同行推荐人分析，与MGA业务维度无关',
        '怎么改': '无需修改'
    },
    {
        '截图': '',
        'PPT的P几': 'P8',
        '什么板块': 'V. TOP10 同行 KA',
        'PPT的影响说明': '参考文件和测试文件数据一致，无MGA业务展示',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的2-业务端视角 sheet的K. TOP10 同行KA',
        'excel的影响说明': 'S2-K部分为同行KA数据，与MGA业务无关',
        '判断': '无需改',
        '为什么这么判断': '该板块为同行KA分析，与MGA业务维度无关',
        '怎么改': '无需修改'
    },
    {
        '截图': '',
        'PPT的P几': 'P9',
        '什么板块': 'X. 同行 W01..W14 预约/签单/批核',
        'PPT的影响说明': '参考文件和测试文件数据一致，无MGA业务展示',
        '对应excel什么板块': '业绩分析报表_0724.xlsx的3-执行管理端 sheet的J-APE/K-APE/L-APE. 同行周度业绩',
        'excel的影响说明': 'S3-J-APE/K-APE/L-APE部分为同行数据汇总，与MGA业务无关',
        '判断': '无需改',
        '为什么这么判断': '该板块为同行周度业绩分析，与MGA业务维度无关',
        '怎么改': '无需修改'
    },
]

df = pd.DataFrame(data)

wb = Workbook()
ws = wb.active
ws.title = 'MGA业务差异分析'

headers = ['截图', 'PPT的P几', '什么板块', 'PPT的影响说明', '对应excel什么板块', 
           'excel的影响说明', '判断', '为什么这么判断', '怎么改']

for col_idx, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='1F3864')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

for row_idx, row in df.iterrows():
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=row_idx+2, column=col_idx, value=str(row[header]))
        cell.alignment = Alignment(vertical='center', wrap_text=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        if header == '判断':
            if row[header] == '已修改':
                cell.fill = PatternFill('solid', fgColor='D5E8D4')
                cell.font = Font(color='38761D', bold=True)
            elif row[header] == '改PPT':
                cell.fill = PatternFill('solid', fgColor='FFE6CC')
                cell.font = Font(color='B45309', bold=True)
            elif row[header] == '改EXCEL':
                cell.fill = PatternFill('solid', fgColor='DAE8FC')
                cell.font = Font(color='1E3A8A', bold=True)
            else:
                cell.fill = PatternFill('solid', fgColor='F5F5F5')

column_widths = [10, 8, 45, 60, 55, 55, 10, 50, 40]
for i, width in enumerate(column_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = width

wb.save('MGA业务差异分析表.xlsx')
print('MGA业务差异分析表.xlsx 已生成')
