from pptx import Presentation
from zipfile import ZipFile
from lxml import etree

def extract_chart_data(ppt_path, chart_name):
    """Extract chart data from PPT using XML parsing."""
    print(f"\nExtracting chart data for '{chart_name}' from: {ppt_path}")
    
    with ZipFile(ppt_path, 'r') as zf:
        # Find chart relationship file
        rels_file = None
        for name in zf.namelist():
            if name.endswith('slide7.xml.rels'):
                rels_file = name
                break
        
        if rels_file is None:
            print("  ERROR: slide7.xml.rels not found")
            return None
        
        with zf.open(rels_file) as f:
            rels_tree = etree.parse(f)
            ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            chart_part = None
            for rel in rels_tree.iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
                    target = rel.get('Target')
                    if chart_name in target or ('0' in chart_name and 'chart1' not in target):
                        chart_part = target
                        break
        
        if chart_part is None:
            print("  ERROR: Chart part not found")
            return None
        
        print(f"  Chart part: {chart_part}")
        
        with zf.open(chart_part) as f:
            chart_tree = etree.parse(f)
            ns_c = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
            
            # Extract categories
            cats = []
            cat_nodes = chart_tree.iter(ns_c + 'cat')
            for cat in cat_nodes:
                for v in cat.iter(ns_c + 'v'):
                    cats.append(v.text)
            
            # Extract series
            series_data = []
            series_names = []
            series_nodes = chart_tree.iter(ns_c + 'ser')
            for ser in series_nodes:
                name = "Unknown"
                for tx in ser.iter(ns_c + 'tx'):
                    for v in tx.iter(ns_c + 'v'):
                        name = v.text
                        break
                    break
                series_names.append(name)
                
                values = []
                for val in ser.iter(ns_c + 'val'):
                    for v in val.iter(ns_c + 'v'):
                        values.append(v.text)
                series_data.append(values)
            
            return {
                'categories': cats,
                'series_names': series_names,
                'series_data': series_data
            }

if __name__ == "__main__":
    template_data = extract_chart_data("template.pptx", "Chart 0")
    final_data = extract_chart_data("周业绩汇报PPT_FINAL_MGA.pptx", "Chart 0")
    
    print("\n" + "="*60)
    print("CHART 0 DATA COMPARISON (U 同行推荐人分析)")
    print("="*60)
    
    print("\nTemplate categories:", template_data['categories'] if template_data else "N/A")
    print("Final categories:", final_data['categories'] if final_data else "N/A")
    
    if template_data and final_data:
        print("\nTemplate series:")
        for i, (name, data) in enumerate(zip(template_data['series_names'], template_data['series_data'])):
            print(f"  [{i}] {name}: {data[:5]}... (total {len(data)})")
        
        print("\nFinal series:")
        for i, (name, data) in enumerate(zip(final_data['series_names'], final_data['series_data'])):
            print(f"  [{i}] {name}: {data[:5]}... (total {len(data)})")
    
    # Chart 1
    template_data1 = extract_chart_data("template.pptx", "Chart 1")
    final_data1 = extract_chart_data("周业绩汇报PPT_FINAL_MGA.pptx", "Chart 1")
    
    print("\n" + "="*60)
    print("CHART 1 DATA COMPARISON (V 同行 TOP 10 KEY ACCOUNT分析)")
    print("="*60)
    
    print("\nTemplate categories:", template_data1['categories'] if template_data1 else "N/A")
    print("Final categories:", final_data1['categories'] if final_data1 else "N/A")
    
    if template_data1 and final_data1:
        print("\nTemplate series:")
        for i, (name, data) in enumerate(zip(template_data1['series_names'], template_data1['series_data'])):
            print(f"  [{i}] {name}: {data[:5]}... (total {len(data)})")
        
        print("\nFinal series:")
        for i, (name, data) in enumerate(zip(final_data1['series_names'], final_data1['series_data'])):
            print(f"  [{i}] {name}: {data[:5]}... (total {len(data)})")