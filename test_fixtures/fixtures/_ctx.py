"""Shared mutable state and helpers for the fixture modules.

All slide modules import 'prs' and 'blank' from here and mutate 'prs' in
place at import time. The orchestrator in 'gen_test_features.py' imports the
slide modules in order, then runs the postprocess step.

Cross-module helpers (nsmap, set_fill_xml) live here so any slide module can
reach them via `from ._ctx import ...`.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from lxml import etree

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

OUTPUT_PATH = 'test_fixtures/test_features.pptx'

# Default namespace map used across slide modules for lxml find() calls.
nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def set_fill_xml(shape, fill_xml_str):
    """Replace any existing fill with custom fill XML in the shape's spPr."""
    ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    ns_p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    spPr = shape._element.find(f'{ns_p}spPr')
    if spPr is None:
        spPr = shape._element.find(f'.//{ns_a}spPr')
    if spPr is None:
        return
    for tag in ('solidFill', 'noFill', 'gradFill', 'blipFill', 'pattFill'):
        for el in spPr.findall(f'{ns_a}{tag}'):
            spPr.remove(el)
    prstGeom = spPr.find(f'{ns_a}prstGeom')
    fill_el = etree.fromstring(fill_xml_str)
    if prstGeom is not None:
        prstGeom.addnext(fill_el)
    else:
        spPr.append(fill_el)
