from pptx import Presentation

prs = Presentation('周业绩汇报PPT_FINAL.pptx')
slide = prs.slides[4]

print('Slide 5 J section shapes layout:')
biz_lines = []
for s in slide.shapes:
    if s.has_text_frame and s.name.startswith('Text'):
        text = s.text_frame.text.strip()
        if text in ['BK业务', '永明经代', '同行经代', '天领业务', 'ICLUB', '成事家办', '合伙转介', 'IFA业务', 'MGA业务']:
            biz_lines.append({
                'name': text,
                'shape_name': s.name,
                'top': s.top,
                'height': s.height,
                'bottom': s.top + s.height
            })

biz_lines.sort(key=lambda x: x['top'])

for i, bl in enumerate(biz_lines):
    print(f'{i+1}. {bl["name"]}: top={bl["top"]}, height={bl["height"]}, bottom={bl["bottom"]}')

print('\nJ section header:')
for s in slide.shapes:
    if s.has_text_frame and s.name.startswith('Text'):
        text = s.text_frame.text.strip()
        if '保单阶段构成' in text:
            print(f'  {s.name}: "{text}" - top={s.top}, bottom={s.top + s.height}')

print('\nK section start:')
for s in slide.shapes:
    if s.has_text_frame and s.name.startswith('Text'):
        text = s.text_frame.text.strip()
        if '保单总量构成' in text:
            print(f'  {s.name}: "{text}" - top={s.top}')

print('\nAvailable space for J section:')
j_header_top = None
k_section_top = None
for s in slide.shapes:
    if s.has_text_frame and s.name.startswith('Text'):
        text = s.text_frame.text.strip()
        if '保单阶段构成' in text:
            j_header_top = s.top + s.height + Emu(100000)
        if '保单总量构成' in text:
            k_section_top = s.top

if j_header_top and k_section_top:
    available_height = k_section_top - j_header_top
    print(f'  J section start: {j_header_top}')
    print(f'  K section start: {k_section_top}')
    print(f'  Available height: {available_height}')
    print(f'  Available height (pt): {available_height / 914400}')
    print(f'  Space per business line (9 lines): {available_height / 9}')
    print(f'  Space per business line (9 lines, pt): {(available_height / 9) / 914400}')