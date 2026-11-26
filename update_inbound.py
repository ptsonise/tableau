"""Loader for the raw-material planner. The code lives in src/*.py (15 area files, run in file-name order, ONE shared namespace).
Run:  python update_inbound.py            Single file for the planner's PC:  python build_single.py  ->  dist/update_inbound.py
Do NOT read dist/update_inbound.py (generated, 1 MB): edit the file in src/ that owns the area (see CLAUDE.md 'Code map')."""
import os, glob
_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in sorted(glob.glob(os.path.join(_HERE, 'src', '*.py'))):
    exec(compile(open(_p, encoding='utf-8').read(), _p, 'exec'), globals())
