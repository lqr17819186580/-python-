import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from chart_updates import slide1_chart_business_type
from update_ppt import S1

cd = slide1_chart_business_type(S1)
print("slide1_chart_business_type(S1) returns:")
print(f"  Categories: {[c.label for c in cd.categories]}")
for s in cd._series:
    print(f"  Series '{s.name}': {list(s.values)}")