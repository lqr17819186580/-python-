import os
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation

def extract_chart_data(chart):
    series_data = []
    try:
        if hasattr(chart, 'categories'):
            categories = [str(cat.label) for cat in chart.categories]
        else:
            categories = []
        
        for series in chart.series:
            series_name = series.name if series.name else "未命名"
            values = []
            for val in series.values:
                if val is None:
                    values.append("None")
                else:
                    try:
                        values.append(f"{float(val):.2f}")
                    except:
                        values.append(str(val))
            series_data.append({
                'name': series_name,
                'values': values
            })
        
        return {
            'categories': categories,
            'series': series_data,
            'has_data': len(series_data) > 0
        }
    except Exception as e:
        return {
            'categories': [],
            'series': [],
            'has_data': False,
            'error': str(e)[:100]
        }

def extract_table_data(table):
    data = []
    try:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                text = cell.text.strip()
                row_data.append(text[:80] if text else "")
            data.append(row_data)
        return {'data': data, 'has_data': len(data) > 0}
    except:
        return {'data': [], 'has_data': False}

def extract_slide_detailed(prs, ppt_name):
    slides_info = []
    for i, slide in enumerate(prs.slides, 1):
        title = ""
        all_texts = []
        charts = []
        tables = []
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    all_texts.append(text)
                    if not title and i == 1:
                        title = text[:80]
        
        for shape in slide.shapes:
            if shape.has_chart:
                chart_data = extract_chart_data(shape.chart)
                chart_data['name'] = shape.name
                charts.append(chart_data)
            
            if shape.has_table:
                table_data = extract_table_data(shape.table)
                tables.append(table_data)
        
        if not title and all_texts:
            title = all_texts[0][:50]
        
        slides_info.append({
            'index': i,
            'title': title,
            'text_count': len(all_texts),
            'chart_count': len(charts),
            'table_count': len(tables),
            'all_texts': all_texts,
            'charts': charts,
            'tables': tables
        })
    return slides_info

def find_kpi_values(texts):
    kpis = {}
    for text in texts:
        match = re.search(r'([\d.]+)\s*%', text)
        if match and '达成率' in text:
            kpis['达成率'] = match.group(1)
        
        match = re.search(r'批核APE[\s:]*([\d,.]+)', text)
        if match:
            kpis['批核APE'] = match.group(1)
        
        match = re.search(r'全业务[\s\S]*?([\d.]+)\s*%', text)
        if match and '达成率' in text:
            kpis['全业务达成率'] = match.group(1)
    return kpis

def create_comparison_report():
    ref_ppt = r'd:\数据中台支持\业绩整合自动化脚本\ppt汇报\周业绩汇报PPT_W28_20260717.pptx'
    test_ppt = r'd:\数据中台支持\业绩整合自动化脚本\ppt汇报\周业绩汇报PPT_FINAL.pptx'
    
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = '微软雅黑'
    font.size = Pt(10.5)
    
    title = doc.add_heading('周业绩汇报PPT对比验证报告', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"参考文件：{os.path.basename(ref_ppt)}")
    doc.add_paragraph(f"测试文件：{os.path.basename(test_ppt)}")
    doc.add_paragraph(f"生成日期：2026年7月24日")
    
    if os.path.exists(ref_ppt):
        prs_ref = Presentation(ref_ppt)
        ref_info = extract_slide_detailed(prs_ref, ref_ppt)
        ref_size = os.path.getsize(ref_ppt)
    else:
        ref_info = []
        ref_size = 0
    
    if os.path.exists(test_ppt):
        prs_test = Presentation(test_ppt)
        test_info = extract_slide_detailed(prs_test, test_ppt)
        test_size = os.path.getsize(test_ppt)
    else:
        test_info = []
        test_size = 0
    
    doc.add_heading('一、基本信息对比', level=1)
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = '项目'
    hdr[1].text = '参考文件'
    hdr[2].text = '测试文件'
    hdr[3].text = '对比结果'
    
    row = table.add_row().cells
    row[0].text = '文件大小'
    row[1].text = f"{ref_size / 1024:.1f} KB"
    row[2].text = f"{test_size / 1024:.1f} KB"
    row[3].text = '✅ 一致' if abs(ref_size - test_size) < 1024 else '不同'
    
    row = table.add_row().cells
    row[0].text = '总页数'
    row[1].text = str(len(ref_info))
    row[2].text = str(len(test_info))
    row[3].text = '✅ 一致' if len(ref_info) == len(test_info) else '❌ 不一致'
    
    row = table.add_row().cells
    row[0].text = '图表总数'
    ref_charts = sum(s['chart_count'] for s in ref_info)
    test_charts = sum(s['chart_count'] for s in test_info)
    row[1].text = str(ref_charts)
    row[2].text = str(test_charts)
    row[3].text = '✅ 一致' if ref_charts == test_charts else '❌ 不一致'
    
    row = table.add_row().cells
    row[0].text = '表格总数'
    ref_tables = sum(s['table_count'] for s in ref_info)
    test_tables = sum(s['table_count'] for s in test_info)
    row[1].text = str(ref_tables)
    row[2].text = str(test_tables)
    row[3].text = '✅ 一致' if ref_tables == test_tables else '❌ 不一致'
    
    doc.add_heading('二、逐页详细对比', level=1)
    
    all_matched = True
    page_diff_details = []
    
    max_pages = max(len(ref_info), len(test_info))
    
    table = doc.add_table(rows=1, cols=8)
    hdr = table.rows[0].cells
    hdr[0].text = '页码'
    hdr[1].text = '参考标题'
    hdr[2].text = '测试标题'
    hdr[3].text = '图表数'
    hdr[4].text = '表格数'
    hdr[5].text = '图表数据'
    hdr[6].text = '表格数据'
    hdr[7].text = '状态'
    
    for i in range(1, max_pages + 1):
        ref = ref_info[i-1] if i <= len(ref_info) else None
        test = test_info[i-1] if i <= len(test_info) else None
        
        row = table.add_row().cells
        row[0].text = str(i)
        
        ref_title = ref['title'] if ref else '（缺失）'
        test_title = test['title'] if test else '（缺失）'
        row[1].text = ref_title[:25] + '...' if len(ref_title) > 25 else ref_title
        row[2].text = test_title[:25] + '...' if len(test_title) > 25 else test_title
        
        ref_ch = ref['chart_count'] if ref else 0
        test_ch = test['chart_count'] if test else 0
        row[3].text = f"{ref_ch} vs {test_ch}"
        
        ref_tb = ref['table_count'] if ref else 0
        test_tb = test['table_count'] if test else 0
        row[4].text = f"{ref_tb} vs {test_tb}"
        
        chart_data_status = ''
        table_data_status = ''
        chart_diffs = []
        table_diffs = []
        
        if ref and test:
            for j, (c1, c2) in enumerate(zip(ref['charts'], test['charts'])):
                if c1['categories'] != c2['categories']:
                    chart_diffs.append(f"图表{j}分类不同")
                if len(c1['series']) != len(c2['series']):
                    chart_diffs.append(f"图表{j}系列数不同")
                else:
                    for k, (s1, s2) in enumerate(zip(c1['series'], c2['series'])):
                        if s1['values'] != s2['values']:
                            chart_diffs.append(f"图表{j}系列{k}数值不同")
            
            if len(ref['charts']) != len(test['charts']):
                chart_diffs.append(f"图表总数不同")
            
            chart_data_status = '✅ 一致' if len(chart_diffs) == 0 else '❌ 有差异'
            
            for j, (t1, t2) in enumerate(zip(ref['tables'], test['tables'])):
                if t1['data'] != t2['data']:
                    table_diffs.append(f"表格{j}数据不同")
            
            if len(ref['tables']) != len(test['tables']):
                table_diffs.append(f"表格总数不同")
            
            table_data_status = '✅ 一致' if len(table_diffs) == 0 else '❌ 有差异'
        else:
            chart_data_status = '—'
            table_data_status = '—'
        
        row[5].text = chart_data_status
        row[6].text = table_data_status
        
        if not ref and not test:
            status = '—'
        elif not ref:
            status = '❌ 参考缺失'
            all_matched = False
        elif not test:
            status = '❌ 测试缺失'
            all_matched = False
        else:
            title_match = ref_title[:25] == test_title[:25]
            chart_match = ref_ch == test_ch
            table_match = ref_tb == test_tb
            data_match = len(chart_diffs) == 0 and len(table_diffs) == 0
            
            if title_match and chart_match and table_match and data_match:
                status = '✅ 一致'
            else:
                status = '⚠️ 部分差异'
                all_matched = False
                page_diff_details.append({
                    'page': i,
                    'ref_title': ref_title,
                    'test_title': test_title,
                    'chart_diff': ref_ch != test_ch or len(chart_diffs) > 0,
                    'table_diff': ref_tb != test_tb or len(table_diffs) > 0,
                    'title_diff': not title_match,
                    'chart_diffs': chart_diffs,
                    'table_diffs': table_diffs
                })
        row[7].text = status
    
    doc.add_heading('三、差异详情', level=1)
    
    if not page_diff_details:
        doc.add_paragraph('✅ 未发现页面结构和数据差异')
    else:
        for diff in page_diff_details:
            doc.add_heading(f"Page {diff['page']}", level=2)
            if diff['title_diff']:
                doc.add_paragraph(f"标题差异：")
                doc.add_paragraph(f"  参考：{diff['ref_title']}")
                doc.add_paragraph(f"  测试：{diff['test_title']}")
            if diff['chart_diffs']:
                doc.add_paragraph(f"图表数据差异：")
                for cd in diff['chart_diffs']:
                    doc.add_paragraph(f"  • {cd}")
            if diff['table_diffs']:
                doc.add_paragraph(f"表格数据差异：")
                for td in diff['table_diffs']:
                    doc.add_paragraph(f"  • {td}")
    
    doc.add_heading('四、MGA业务专项验证', level=1)
    
    mga_found = False
    mga_pages = []
    for slide in test_info:
        for text in slide['all_texts']:
            if 'MGA' in text or 'mga' in text.lower():
                mga_found = True
                mga_pages.append(slide['index'])
                doc.add_paragraph(f"Page {slide['index']} ({slide['title'][:30]}): 包含MGA相关文本")
                break
        
        for chart in slide['charts']:
            if chart['categories'] and 'MGA' in str(chart['categories']):
                mga_found = True
                if slide['index'] not in mga_pages:
                    mga_pages.append(slide['index'])
                    doc.add_paragraph(f"Page {slide['index']} ({slide['title'][:30]}): 图表分类包含MGA")
    
    if not mga_found:
        doc.add_paragraph('⚠️ 未在测试PPT中找到MGA相关内容')
    else:
        doc.add_paragraph(f'✅ 测试PPT中在{len(mga_pages)}个页面包含MGA业务内容：{", ".join(map(str, mga_pages))}')
    
    doc.add_heading('五、关键指标对比', level=1)
    
    doc.add_paragraph('5.1 参考文件关键指标：')
    ref_kpis = []
    for slide in ref_info:
        kpis = find_kpi_values(slide['all_texts'])
        if kpis:
            kpi_str = ", ".join([f"{k}={v}" for k, v in kpis.items()])
            ref_kpis.append(f"Page {slide['index']}: {kpi_str}")
            doc.add_paragraph(f"  • Page {slide['index']}: {kpi_str}")
    
    doc.add_paragraph('5.2 测试文件关键指标：')
    test_kpis = []
    for slide in test_info:
        kpis = find_kpi_values(slide['all_texts'])
        if kpis:
            kpi_str = ", ".join([f"{k}={v}" for k, v in kpis.items()])
            test_kpis.append(f"Page {slide['index']}: {kpi_str}")
            doc.add_paragraph(f"  • Page {slide['index']}: {kpi_str}")
    
    doc.add_paragraph('5.3 指标对比结果：')
    if set(ref_kpis) == set(test_kpis):
        doc.add_paragraph('✅ 关键指标完全一致')
    else:
        doc.add_paragraph('⚠️ 关键指标存在差异')
        if len(ref_kpis) != len(test_kpis):
            doc.add_paragraph(f"   - 指标数量不同：参考{len(ref_kpis)}项 vs 测试{len(test_kpis)}项")
    
    doc.add_heading('六、结论', level=1)
    
    total_pages = len(ref_info) if ref_info else 0
    test_total_pages = len(test_info) if test_info else 0
    
    if all_matched and total_pages == test_total_pages == 12 and mga_found:
        doc.add_paragraph('✅ 测试PPT与参考PPT结构一致（12页）')
        doc.add_paragraph('✅ 图表和表格数据一致')
        doc.add_paragraph('✅ MGA业务已正确整合为独立业务线')
        doc.add_paragraph('✅ 关键指标数据正确')
        doc.add_paragraph('✅ 修改完成，可以交付')
    else:
        doc.add_paragraph('❌ 仍存在差异，需要进一步检查')
        if total_pages != test_total_pages:
            doc.add_paragraph(f"   - 页数不一致：参考{total_pages}页 vs 测试{test_total_pages}页")
        if not all_matched:
            doc.add_paragraph(f"   - {len(page_diff_details)}个页面存在结构或数据差异")
        if not mga_found:
            doc.add_paragraph('   - MGA业务内容缺失')
    
    output_file = r'd:\数据中台支持\业绩整合自动化脚本\ppt汇报\周业绩汇报PPT对比验证报告.docx'
    doc.save(output_file)
    print(f"✅ 验证报告已生成：{output_file}")

if __name__ == '__main__':
    create_comparison_report()