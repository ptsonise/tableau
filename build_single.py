"""Joins src/*.py into ONE file for the planner's PC:  python build_single.py  ->  dist/update_inbound.py (git-ignored, generated)."""
import os, glob
here = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(here, 'dist'), exist_ok=True)
out = os.path.join(here, 'dist', 'update_inbound.py')
open(out, 'w', encoding='utf-8', newline='').write(''.join(open(p, encoding='utf-8', newline='').read() for p in sorted(glob.glob(os.path.join(here, 'src', '*.py')))))
print('written', out)
