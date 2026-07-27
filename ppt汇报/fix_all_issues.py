#!/usr/bin/env python3
"""Fix all four issues:
1. Slide 2: "1–4月已批核" → "1–6月已批核"
2. Slide 4: Remove extra white text box
3. Slide 7: Restore T title with correct style from template
4. Slide 9: Compare peer W28 data with reference PPT
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

IN_PPT = "周业绩汇报PPT_FINAL.pptx"
OUT_PPT = "周业绩汇报PPT_FINAL_fixed.pptx"
REF_PPT = "周业绩汇报PPT_W28_20260717.pptx"

def fix_slide2_month_label(slide):
    print("\n[Slide 2] Fixing month label...")
    for sh in slide.shapes:
        if sh.has_text_frame:
            text = sh.text_frame.text.strip()
            if "1–4 月已批核" in text:
                print(f"  Found: {text!r}")
                if sh.text_frame.paragraphs and sh.text_frame.paragraphs[0].runs:
                    old_text = sh.text_frame.paragraphs[0].runs[0].text
                    new_text = old_text.replace("1–4 月已批核", "1–6 月已批核")
                    sh.text_frame.paragraphs[0].runs[0].text = new_text
                    print(f"  Updated: {old_text!r} → {new_text!r}")
                break

def fix_slide4_white_textbox(slide):
    print("\n[Slide 4] Removing extra white text boxes...")
    removed = 0
    for sh in list(slide.shapes):
        if not sh.has_text_frame:
            continue
        text = sh.text_frame.text.strip()
        if not text:
            continue
        try:
            has_white_fill = False
            if sh.fill.type == 1:
                try:
                    if sh.fill.fore_color.rgb == RGBColor(0xFF, 0xFF, 0xFF):
                        has_white_fill = True
                except:
                    pass
            
            has_border = False
            try:
                if sh.line.fill.type != 0:
                    has_border = True
            except:
                pass
            
            if has_white_fill and has_border:
                print(f"  Removing: name={sh.name!r}, text={text!r}, top={sh.top}, left={sh.left}")
                sh._element.getparent().remove(sh._element)
                removed += 1
        except Exception as e:
            pass
    print(f"  Removed {removed} white text boxes")

def fix_slide7_title(slide):
    print("\n[Slide 7] Fixing T title...")
    for sh in list(slide.shapes):
        if sh.has_text_frame:
            text = sh.text_frame.text.strip()
            if "签批时效" in text:
                sh._element.getparent().remove(sh._element)
                print(f"  Removed existing title: {text!r}")
                break

    left = Emu(6254496)
    top = Emu(3694176)
    width = Emu(5724144)
    height = Emu(192024)
    
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    run = para.add_run()
    run.text = "T  签批时效综合分析 — 业务线时效 & 分档分布"
    run.font.size = Pt(14.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.name = 'Calibri'
    
    tb.fill.background()
    try:
        tb.line.fill.background()
    except:
        pass
    
    print("  Added T title with correct style")

def compare_slide9_data():
    print("\n[Slide 9] Comparing peer W28 data...")
    
    prs_final = Presentation(IN_PPT)
    prs_ref = Presentation(REF_PPT)
    
    slide9_final = prs_final.slides[8]
    slide9_ref = prs_ref.slides[8]
    
    def get_peer_data(slide, label):
        data = {}
        for sh in slide.shapes:
            if sh.has_chart:
                for series in sh.chart.series:
                    if len(series.values) >= 28:
                        data[series.name] = series.values[27]
        return data
    
    final_data = get_peer_data(slide9_final, "FINAL")
    ref_data = get_peer_data(slide9_ref, "REF")
    
    print(f"  FINAL W28 data: {final_data}")
    print(f"  REF W28 data: {ref_data}")
    
    if final_data != ref_data:
        print("  ⚠️ DATA MISMATCH!")
        for key in final_data:
            if key in ref_data:
                diff = final_data[key] - ref_data[key]
                print(f"    {key}: FINAL={final_data[key]}, REF={ref_data[key]}, diff={diff}")
    else:
        print("  ✅ Data matches!")

def main():
    print(f"Loading {IN_PPT}...")
    prs = Presentation(IN_PPT)
    
    slide2 = prs.slides[1]
    fix_slide2_month_label(slide2)
    
    slide4 = prs.slides[3]
    fix_slide4_white_textbox(slide4)
    
    slide7 = prs.slides[6]
    fix_slide7_title(slide7)
    
    compare_slide9_data()
    
    prs.save(OUT_PPT)
    print(f"\nSaved to {OUT_PPT}")

if __name__ == "__main__":
    main()