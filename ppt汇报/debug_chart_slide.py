import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from update_ppt import _chart_slide, _SL_11, slides

print(f"_SL_11 is _NullSlide: {isinstance(_SL_11, type('')) or 'NullSlide' in str(type(_SL_11))}")
print(f"Total slides: {len(slides)}")

result = _chart_slide("Chart 0", _SL_11)
print(f"_chart_slide('Chart 0', _SL_11) returned slide index: {slides.index(result) if result in slides else 'Not found'}")