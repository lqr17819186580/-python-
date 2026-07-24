import os
import re
import tempfile
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches

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
        
        match = re.search(r'缺口[\s:]*([\d,.]+)', text)
        if match:
            kpis['缺口'] = match.group(1)
        
        match = re.search(r'目标[\s:]*([\d,.]+)', text)
        if match:
            kpis['目标'] = match.group(1)
        
        match = re.search(r'实际[\s:]*([\d,.]+)', text)
        if match:
            kpis['实际'] = match.group(1)
    return kpis

def export_ppt_slides_to_images(ppt_path, output_dir):
    try:
        os.makedirs(output_dir, exist_ok=True)
        import win32com.client
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        powerpoint.Visible = False
        
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        
        image_paths = {}
        for i, slide in enumerate(presentation.Slides):
            image_path = os.path.join(output_dir, f"slide_{i+1}.png")
            slide.Export(image_path, "PNG", 1920, 1080)
            if os.path.exists(image_path):
                image_paths[i+1] = image_path
        
        presentation.Close()
        powerpoint.Quit()
        
        return image_paths
    except Exception as e:
        print(f"PPT导出图片失败(win32com): {e}")
        try:
            os.makedirs(output_dir, exist_ok=True)
            cmd = f'powershell -Command "$ppt = New-Object -ComObject PowerPoint.Application; $ppt.Visible = $false; $pres = $ppt.Presentations.Open(\'{ppt_path}\', $true, $false, $false); for ($i=1; $i -le $pres.Slides.Count; $i++) {{ $pres.Slides.Item($i).Export(\'{os.path.join(output_dir, "slide_")}\' + $i + \'.png\', \'PNG\', 1920, 1080) }}; $pres.Close(); $ppt.Quit()"'
            os.system(cmd)
            image_paths = {}
            for i in range(1, 13):
                image_path = os.path.join(output_dir, f"slide_{i}.png")
                if os.path.exists(image_path):
                    image_paths[i] = image_path
            return image_paths
        except Exception as e2:
            print(f"PPT导出图片失败(powershell): {e2}")
            return {}

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
    title.runs[0].font.size = Pt(20)
    title.runs[0].font.bold = True
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"参考文件：{os.path.basename(ref_ppt)}")
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    subtitle2 = doc.add_paragraph()
    subtitle2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle2.add_run(f"测试文件：{os.path.basename(test_ppt)}")
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    subtitle3 = doc.add_paragraph()
    subtitle3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle3.add_run(f"生成日期：2026年7月24日")
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    
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
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    
    col_widths = [Cm(4.0), Cm(5.0), Cm(5.0), Cm(3.0)]
    hdr = table.rows[0].cells
    hdr[0].text = '项目'
    hdr[0].width = col_widths[0]
    hdr[1].text = '参考文件'
    hdr[1].width = col_widths[1]
    hdr[2].text = '测试文件'
    hdr[2].width = col_widths[2]
    hdr[3].text = '对比结果'
    hdr[3].width = col_widths[3]
    
    row = table.add_row().cells
    row[0].text = '文件大小'
    row[1].text = f"{ref_size / 1024:.1f} KB"
    row[2].text = f"{test_size / 1024:.1f} KB"
    row[3].text = '✅ 一致' if abs(ref_size - test_size) < 1024 else '❌ 不同'
    
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
    
    doc.add_paragraph()
    
    doc.add_heading('二、逐页结构对比', level=1)
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    table.autofit = False
    table.allow_autofit = False
    
    col_widths = [Cm(1.5), Cm(4.5), Cm(4.5), Cm(2.0), Cm(2.0), Cm(2.5)]
    hdr = table.rows[0].cells
    hdr[0].text = '页码'
    hdr[0].width = col_widths[0]
    hdr[1].text = '参考标题'
    hdr[1].width = col_widths[1]
    hdr[2].text = '测试标题'
    hdr[2].width = col_widths[2]
    hdr[3].text = '图表数'
    hdr[3].width = col_widths[3]
    hdr[4].text = '表格数'
    hdr[4].width = col_widths[4]
    hdr[5].text = '状态'
    hdr[5].width = col_widths[5]
    
    all_matched = True
    page_diff_details = []
    
    max_pages = max(len(ref_info), len(test_info))
    
    for i in range(1, max_pages + 1):
        ref = ref_info[i-1] if i <= len(ref_info) else None
        test = test_info[i-1] if i <= len(test_info) else None
        
        row = table.add_row().cells
        row[0].text = str(i)
        
        ref_title = ref['title'] if ref else '（缺失）'
        test_title = test['title'] if test else '（缺失）'
        row[1].text = ref_title[:30] + '...' if len(ref_title) > 30 else ref_title
        row[2].text = test_title[:30] + '...' if len(test_title) > 30 else test_title
        
        ref_ch = ref['chart_count'] if ref else 0
        test_ch = test['chart_count'] if test else 0
        row[3].text = f"{ref_ch}/{test_ch}"
        
        ref_tb = ref['table_count'] if ref else 0
        test_tb = test['table_count'] if test else 0
        row[4].text = f"{ref_tb}/{test_tb}"
        
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
            
            for j, (t1, t2) in enumerate(zip(ref['tables'], test['tables'])):
                if t1['data'] != t2['data']:
                    table_diffs.append(f"表格{j}数据不同")
            
            if len(ref['tables']) != len(test['tables']):
                table_diffs.append(f"表格总数不同")
        
        if not ref and not test:
            status = '—'
        elif not ref:
            status = '❌ 参考缺失'
            all_matched = False
        elif not test:
            status = '❌ 测试缺失'
            all_matched = False
        else:
            title_match = ref_title[:30] == test_title[:30]
            chart_match = ref_ch == test_ch
            table_match = ref_tb == test_tb
            data_match = len(chart_diffs) == 0 and len(table_diffs) == 0
            
            if title_match and chart_match and table_match and data_match:
                status = '✅ 一致'
            else:
                status = '⚠️ 有差异'
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
        row[5].text = status
    
    doc.add_paragraph()
    
    doc.add_heading('三、关键页面内容对比', level=1)
    
    key_pages = [1, 2, 3, 4, 5]
    
    for page_num in key_pages:
        if page_num <= len(ref_info) and page_num <= len(test_info):
            doc.add_heading(f"Page {page_num}: {ref_info[page_num-1]['title']}", level=2)
            
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            table.autofit = False
            table.allow_autofit = False
            
            ref_cell = table.cell(0, 0)
            ref_cell.width = Cm(8.0)
            ref_paragraph = ref_cell.paragraphs[0]
            ref_run = ref_paragraph.add_run("参考文件")
            ref_run.font.bold = True
            ref_run.font.color.rgb = RGBColor(0, 102, 204)
            
            test_cell = table.cell(0, 1)
            test_cell.width = Cm(8.0)
            test_paragraph = test_cell.paragraphs[0]
            test_run = test_paragraph.add_run("测试文件")
            test_run.font.bold = True
            test_run.font.color.rgb = RGBColor(0, 153, 76)
            
            ref_texts = ref_info[page_num-1]['all_texts'][:5]
            test_texts = test_info[page_num-1]['all_texts'][:5]
            
            for text in ref_texts:
                ref_cell.add_paragraph(text[:80], style='Body Text')
            
            for text in test_texts:
                test_cell.add_paragraph(text[:80], style='Body Text')
            
            if ref_info[page_num-1]['charts']:
                ref_cell.add_paragraph()
                ref_cell.add_paragraph("图表数据:", style='Heading 3')
                for chart in ref_info[page_num-1]['charts'][:2]:
                    if chart['categories']:
                        cat_str = ", ".join(chart['categories'][:6])
                        if len(chart['categories']) > 6:
                            cat_str += "..."
                        ref_cell.add_paragraph(f"  分类: {cat_str}")
                    for s in chart['series'][:2]:
                        val_str = ", ".join(s['values'][:6])
                        if len(s['values']) > 6:
                            val_str += "..."
                        ref_cell.add_paragraph(f"  {s['name']}: {val_str}")
            
            if test_info[page_num-1]['charts']:
                test_cell.add_paragraph()
                test_cell.add_paragraph("图表数据:", style='Heading 3')
                for chart in test_info[page_num-1]['charts'][:2]:
                    if chart['categories']:
                        cat_str = ", ".join(chart['categories'][:6])
                        if len(chart['categories']) > 6:
                            cat_str += "..."
                        test_cell.add_paragraph(f"  分类: {cat_str}")
                    for s in chart['series'][:2]:
                        val_str = ", ".join(s['values'][:6])
                        if len(s['values']) > 6:
                            val_str += "..."
                        test_cell.add_paragraph(f"  {s['name']}: {val_str}")
            
            doc.add_paragraph()
    
    doc.add_heading('四、差异详情', level=1)
    
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
    
    doc.add_paragraph()
    
    doc.add_heading('五、MGA业务专项验证', level=1)
    
    mga_found = False
    mga_pages = []
    for slide in test_info:
        for text in slide['all_texts']:
            if 'MGA' in text or 'mga' in text.lower():
                mga_found = True
                if slide['index'] not in mga_pages:
                    mga_pages.append(slide['index'])
                break
        
        for chart in slide['charts']:
            if chart['categories'] and 'MGA' in str(chart['categories']):
                mga_found = True
                if slide['index'] not in mga_pages:
                    mga_pages.append(slide['index'])
    
    doc.add_paragraph(f"测试PPT中包含MGA业务内容的页面：{', '.join(map(str, mga_pages)) if mga_pages else '无'}")
    
    if mga_found:
        for page_num in mga_pages[:3]:
            if page_num <= len(test_info):
                doc.add_heading(f"Page {page_num} MGA业务内容", level=2)
                
                for chart in test_info[page_num-1]['charts']:
                    if 'MGA' in str(chart['categories']):
                        doc.add_paragraph(f"图表: {chart['name']}")
                        doc.add_paragraph(f"  分类: {', '.join(chart['categories'])}")
                        for s in chart['series']:
                            doc.add_paragraph(f"  系列 {s['name']}: {', '.join(s['values'])}")
                
                for text in test_info[page_num-1]['all_texts']:
                    if 'MGA' in text or 'mga' in text.lower():
                        doc.add_paragraph(f"文本: {text[:100]}")
        
        doc.add_paragraph('')
        doc.add_paragraph('✅ MGA业务已正确整合为独立业务线')
    else:
        doc.add_paragraph('⚠️ 未在测试PPT中找到MGA相关内容')
    
    doc.add_paragraph()
    
    doc.add_heading('六、关键指标对比', level=1)
    
    doc.add_heading('6.1 参考文件关键指标', level=2)
    ref_kpis = {}
    for slide in ref_info:
        kpis = find_kpi_values(slide['all_texts'])
        if kpis:
            ref_kpis[slide['index']] = kpis
            kpi_str = ", ".join([f"{k}={v}" for k, v in kpis.items()])
            doc.add_paragraph(f"  • Page {slide['index']}: {kpi_str}")
    
    doc.add_heading('6.2 测试文件关键指标', level=2)
    test_kpis = {}
    for slide in test_info:
        kpis = find_kpi_values(slide['all_texts'])
        if kpis:
            test_kpis[slide['index']] = kpis
            kpi_str = ", ".join([f"{k}={v}" for k, v in kpis.items()])
            doc.add_paragraph(f"  • Page {slide['index']}: {kpi_str}")
    
    doc.add_heading('6.3 指标对比结果', level=2)
    all_kpis_match = True
    kpi_diff_list = []
    
    for page in set(ref_kpis.keys()) | set(test_kpis.keys()):
        ref_k = ref_kpis.get(page, {})
        test_k = test_kpis.get(page, {})
        if ref_k != test_k:
            all_kpis_match = False
            kpi_diff_list.append({
                'page': page,
                'ref': ref_k,
                'test': test_k
            })
    
    if all_kpis_match:
        doc.add_paragraph('✅ 关键指标完全一致')
    else:
        doc.add_paragraph('⚠️ 关键指标存在差异：')
        for diff in kpi_diff_list:
            doc.add_paragraph(f"  • Page {diff['page']}:")
            if diff['ref']:
                doc.add_paragraph(f"     参考: {', '.join([f'{k}={v}' for k, v in diff['ref'].items()])}")
            if diff['test']:
                doc.add_paragraph(f"     测试: {', '.join([f'{k}={v}' for k, v in diff['test'].items()])}")
    
    doc.add_paragraph()
    
    doc.add_heading('七、结论', level=1)
    
    total_pages = len(ref_info) if ref_info else 0
    test_total_pages = len(test_info) if test_info else 0
    
    p1 = doc.add_paragraph()
    if all_matched and total_pages == test_total_pages == 12:
        run = p1.add_run('✅ 测试PPT与参考PPT结构一致（12页）')
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 153, 0)
    else:
        run = p1.add_run('❌ 页面结构存在差异')
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
    
    p2 = doc.add_paragraph()
    if all_kpis_match:
        run = p2.add_run('✅ 关键指标数据正确')
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 153, 0)
    else:
        run = p2.add_run('❌ 关键指标存在差异')
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
    
    p3 = doc.add_paragraph()
    if mga_found:
        run = p3.add_run('✅ MGA业务已正确整合为独立业务线')
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 153, 0)
    else:
        run = p3.add_run('❌ MGA业务内容缺失')
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
    
    doc.add_paragraph()
    
    if all_matched and total_pages == test_total_pages == 12 and mga_found and all_kpis_match:
        summary = doc.add_paragraph()
        run = summary.add_run('综合评价：修改完成，可以交付')
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 153, 0)
        summary.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        summary = doc.add_paragraph()
        run = summary.add_run('综合评价：仍存在差异，需要进一步检查')
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        summary.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if total_pages != test_total_pages:
            doc.add_paragraph(f"   - 页数不一致：参考{total_pages}页 vs 测试{test_total_pages}页")
        if not all_matched:
            doc.add_paragraph(f"   - {len(page_diff_details)}个页面存在结构或数据差异")
        if not mga_found:
            doc.add_paragraph('   - MGA业务内容缺失')
        if not all_kpis_match:
            doc.add_paragraph(f"   - {len(kpi_diff_list)}个页面关键指标存在差异")
    
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run('报告生成时间：2026年7月24日')
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(150, 150, 150)
    
    output_file = r'd:\数据中台支持\业绩整合自动化脚本\ppt汇报\周业绩汇报PPT对比验证报告_图文版.docx'
    doc.save(output_file)
    print(f"✅ 验证报告已生成：{output_file}")

if __name__ == '__main__':
    create_comparison_report()