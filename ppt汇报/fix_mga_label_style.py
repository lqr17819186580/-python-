#!/usr/bin/env python3
"""Fix MGA批核 label style - remove white fill and border"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

prs = Presentation("周业绩汇报PPT_FINAL_fixed.pptx")
slide4 = prs.slides[3]

print("Checking for MGA批核 label...")
found = False
for sh in slide4.shapes:
    if sh.has_text_frame:
        text = sh.text_frame.text.strip()
        if "MGA批核" in text:
            print(f"  Found: name={sh.name!r}, text={text!r}, top={sh.top}, left={sh.left}")
            found = True
            break

if not found:
    print("  MGA批核 label not found, adding...")
    left = Emu(6626860)
    top = Emu(5662930)
    width = Emu(289560)
    height = Emu(182880)
    
    tb = slide4.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = "MGA\n批核"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x20, 0x20, 0x20)
    run.font.name = 'Calibri'
    
    tb.fill.background()
    tb.line.fill.background()
    print("  Added MGA批核 label with transparent style")

prs.save("周业绩汇报PPT_FINAL_fixed.pptx")
print("Saved to 周业绩汇报PPT_FINAL_fixed.pptx")