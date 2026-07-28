"""Diagnose where So What shapes are lost on slide 7"""
from pptx import Presentation
import copy

def count_sowhat(prs, slide_idx=6):
    slide = prs.slides[slide_idx]
    count = 0
    for sh in slide.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t.startswith('So What') or len(t) > 100:
                count += 1
                print(f'  [{count}] top={sh.top}, left={sh.left}, text={t[:60]}...')
    return count

# Stage 0: template
print("=" * 60)
print("Stage 0: template.pptx")
prs0 = Presentation('template.pptx')
n0 = count_sowhat(prs0)
print(f"  Total: {n0} So What shapes")
del prs0

# Stage 1: After update_ppt.py output
print()
print("=" * 60)
print("Stage 1: After update_ppt.py (周业绩汇报PPT_AUTO_UPDATED.pptx)")
try:
    prs1 = Presentation('周业绩汇报PPT_AUTO_UPDATED.pptx')
    n1 = count_sowhat(prs1)
    print(f"  Total: {n1} So What shapes")
    del prs1
except:
    print("  File not found, trying to run update_ppt.py...")

# Stage 2: After apply_w14_patches.py output
print()
print("=" * 60)
print("Stage 2: After apply_w14_patches.py")
try:
    prs2 = Presentation('周业绩汇报PPT_AUTO_UPDATED_updated.pptx')
    n2 = count_sowhat(prs2)
    print(f"  Total: {n2} So What shapes")
    del prs2
except:
    print("  File not found")

# Stage 3: Final
print()
print("=" * 60)
print("Stage 3: Final (周业绩汇报PPT_FINAL.pptx)")
prs3 = Presentation('周业绩汇报PPT_FINAL.pptx')
n3 = count_sowhat(prs3)
print(f"  Total: {n3} So What shapes")
del prs3

print()
print("=" * 60)
print("Summary:")
print(f"  Template: {n0}")
print(f"  Final: {n3}")
print(f"  Lost: {n0 - n3}")