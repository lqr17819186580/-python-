from pptx import Presentation
from zipfile import ZipFile
from lxml import etree

def check_chart_axis_format(ppt_path):
    """Check chart axis number format in PPTX."""
    print(f"\nChecking chart axis format for: {ppt_path}")
    
    with ZipFile(ppt_path, 'r') as zf:
        # Find all chart XML files
        chart_files = [f for f in zf.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
        print(f"\nFound {len(chart_files)} chart files:")
        for cf in chart_files:
            print(f"  {cf}")
        
        # Check each chart for axis format
        ns_c = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
        
        for cf in chart_files:
            with zf.open(cf) as f:
                tree = etree.parse(f)
                
                # Find value axis
                valAx_nodes = tree.iter(ns_c + 'valAx')
                for valAx in valAx_nodes:
                    # Check numFmt
                    numFmt = valAx.find(ns_c + 'numFmt')
                    if numFmt is not None:
                        format_code = numFmt.get('formatCode')
                        sourceLinked = numFmt.get('sourceLinked')
                        print(f"\n{cf} - Value Axis numFmt:")
                        print(f"  formatCode: {format_code}")
                        print(f"  sourceLinked: {sourceLinked}")
                
                # Check series titles
                ser_nodes = tree.iter(ns_c + 'ser')
                for ser in ser_nodes:
                    tx = ser.find(ns_c + 'tx')
                    if tx is not None:
                        v = tx.find('.//' + ns_c + 'v')
                        if v is not None:
                            print(f"  Series title: {v.text}")

if __name__ == "__main__":
    check_chart_axis_format("template.pptx")
    check_chart_axis_format("周业绩汇报PPT_FINAL_MGA.pptx")