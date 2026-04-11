#!/usr/bin/env python3
"""Orchestrator for the test_features.pptx fixture.

Runs each slide module in order against a shared Presentation, then applies
postprocess patches and saves the fixture. Each 'fixtures/slides_NN_*.py'
module is responsible for its own slide range. To add a new slide, append it
to the matching category module (or create a new one and add it to the import
list below).

Slide categories:
  slides_01_text_basics   — 1-14  text body basics
  slides_02_text_extra    — 15-22 capitalization, clrMapOvr, cs/sym, rot, vertical, hlinks, img bullet
  slides_03_fills         — 23-33 gradient/alpha/blip/pattern fills, stroke/arrow/join
  slides_04_shapes        — 34-40 groups, connectors, preset/custom geom, gears, rects, cxnLst
  slides_05_tables        — 41-42 cell merge, borders, diagonal borders, tblPr
  slides_06_images        — 43-51 crop/alpha, ext refs, lum, duotone, bg patt, line grad, hlinks, effects, 3D
  slides_07_charts        — 52-73 standard charts + 3D/trendline/errBars/composite/stacked
  slides_08_misc          — 74-88 notes, comments, SmartArt, OLE, media, OMML, transition, hidden, WMF, effects
  slides_09_chartex       — 89-94 ChartEx (patched by postprocess)
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from fixtures import _ctx  # noqa: F401 — sets up prs/blank
from fixtures import slides_01_text_basics  # noqa: F401
from fixtures import slides_02_text_extra  # noqa: F401
from fixtures import slides_03_fills  # noqa: F401
from fixtures import slides_04_shapes  # noqa: F401
from fixtures import slides_05_tables  # noqa: F401
from fixtures import slides_06_images  # noqa: F401
from fixtures import slides_07_charts  # noqa: F401
from fixtures import slides_08_misc  # noqa: F401
from fixtures import slides_09_chartex  # noqa: F401

_ctx.prs.save(_ctx.OUTPUT_PATH)

# Postprocess: OMML injection, ChartEx XML, media binaries, etc.
from fixtures import _postprocess  # noqa: F401, E402 — runs patching side-effects on import

print(f"Saved {_ctx.OUTPUT_PATH} with {len(_ctx.prs.slides)} slides")
