"""Step-by-step trace of So What shapes through the pipeline"""
from pptx import Presentation
from copy import deepcopy
from lxml import etree
import sys

def check_sowhat(prs, label):
    """Count and show So What shapes on slide 7"""
    slide7 = prs.slides[6]
    shapes = []
    for i, sh in enumerate(slide7.shapes):
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t.startswith('So What') or (len(t) > 100 and len(t) < 200):
                shapes.append((i, sh, t))
    print(f"\n[{label}] Slide 7: {len(shapes)} So What shapes")
    for idx, sh, t in shapes:
        print(f"  [{idx}] name={sh.name} top={sh.top} left={sh.left} text={t[:60]}...")
    return shapes

# Step 0: Check template
print("=" * 60)
print("STEP 0: Load template.pptx")
prs = Presentation('template.pptx')
check_sowhat(prs, "template")

# Step 1: Run update_ppt.py manually (simulate)
# Actually just check if the intermediate file exists
import os
stage1 = '周业绩汇报PPT_AUTO_UPDATED.pptx'
stage2 = '周业绩汇报PPT_AUTO_UPDATED_updated.pptx'

# Try stage 3 output
print("\n" + "=" * 60)
print("STEP: Final check")
prs_final = Presentation('周业绩汇报PPT_FINAL.pptx')
check_sowhat(prs_final, "FINAL")

# Difference analysis
print("\n" + "=" * 60)
print("ANALYSIS:")
tpl_shapes = [(sh.top, sh.left, sh.text_frame.text.strip()[:60]) 
              for sh in prs.slides[6].shapes 
              if sh.has_text_frame and (sh.text_frame.text.strip().startswith('So What') or len(sh.text_frame.text.strip()) > 100)]
fin_shapes = [(sh.top, sh.left, sh.text_frame.text.strip()[:60]) 
              for sh in prs_final.slides[6].shapes 
              if sh.has_text_frame and (sh.text_frame.text.strip().startswith('So What') or len(sh.text_frame.text.strip()) > 100)]

print(f"  Template: {len(tpl_shapes)} shapes")
print(f"  Final:    {len(fin_shapes)} shapes")

# Find missing
fin_tops = {t[0] for t in fin_shapes}
for top, left, text in tpl_shapes:
    if top not in fin_tops:
        print(f"  MISSING: top={top}, left={left}, text={text}")