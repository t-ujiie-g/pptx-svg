#!/usr/bin/env python3
"""
Generate a test PPTX that exercises all recently implemented features.

Slides:
  1. Text body vertical alignment (top / center / bottom)
  2. Paragraph spacing (spcBef / spcAft) + indent (marL / indent)
  3. Bullet characters + auto-numbering
  4. Run decorations (underline / strikethrough / superscript / subscript)
  5. Body insets (lIns / tIns / rIns / bIns)
  6. Master/Layout inheritance: slide with NO explicit bg (inherits master bg)
  7. Master/Layout inheritance: placeholder shapes with text style defaults
  8. Master/Layout inheritance: slide WITH explicit bg (overrides master)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from lxml import etree

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]  # blank layout

# ── Slide 1: Vertical alignment ─────────────────────────────────────────────

slide1 = prs.slides.add_slide(blank)

def add_box(slide, left, top, width, height, text, anchor, fill_rgb=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    txBox.text = text
    tf = txBox.text_frame
    tf.word_wrap = True
    # Set anchor (vertical alignment)
    if anchor == "top":
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT
        # anchor is set via xml manipulation below
    p = tf.paragraphs[0]
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(0, 0, 0)
    # Set fill
    if fill_rgb:
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*fill_rgb)
    # Set body anchor
    bodyPr = txBox.text_frame._txBody.find(
        '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
    anchor_map = {"top": "t", "center": "ctr", "bottom": "b"}
    bodyPr.set('anchor', anchor_map.get(anchor, "t"))
    return txBox

# Title
title_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title_box.text_frame.paragraphs[0].text = "Slide 1: Vertical Alignment (top / center / bottom)"
title_box.text_frame.paragraphs[0].font.size = Pt(24)
title_box.text_frame.paragraphs[0].font.bold = True

add_box(slide1, Inches(0.5), Inches(1.2), Inches(2.8), Inches(2.5),
        "TOP aligned text\n(anchor=t)", "top", (220, 230, 240))
add_box(slide1, Inches(3.6), Inches(1.2), Inches(2.8), Inches(2.5),
        "CENTER aligned text\n(anchor=ctr)", "center", (200, 220, 240))
add_box(slide1, Inches(6.7), Inches(1.2), Inches(2.8), Inches(2.5),
        "BOTTOM aligned text\n(anchor=b)", "bottom", (180, 210, 240))

# Also add boxes with explicit insets
inset_box = slide1.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(4), Inches(2.5))
inset_box.text_frame.paragraphs[0].text = "Large insets (lIns=36pt, tIns=36pt)"
inset_box.text_frame.paragraphs[0].font.size = Pt(12)
fill = inset_box.fill
fill.solid()
fill.fore_color.rgb = RGBColor(240, 230, 200)
bodyPr = inset_box.text_frame._txBody.find(
    '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
bodyPr.set('lIns', str(Emu(Pt(36))))
bodyPr.set('tIns', str(Emu(Pt(36))))
bodyPr.set('rIns', str(Emu(Pt(36))))
bodyPr.set('bIns', str(Emu(Pt(36))))

# ── Slide 2: Paragraph spacing + indent ──────────────────────────────────────

slide2 = prs.slides.add_slide(blank)

title2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title2.text_frame.paragraphs[0].text = "Slide 2: Paragraph Spacing & Indent"
title2.text_frame.paragraphs[0].font.size = Pt(24)
title2.text_frame.paragraphs[0].font.bold = True

box2 = slide2.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(5))
tf2 = box2.text_frame
tf2.word_wrap = True
fill2 = box2.fill
fill2.solid()
fill2.fore_color.rgb = RGBColor(245, 245, 250)

# First paragraph - normal
p0 = tf2.paragraphs[0]
p0.text = "Paragraph 1: No special spacing or indent."
p0.font.size = Pt(14)

nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

# Second paragraph - spacing before 24pt
p1 = tf2.add_paragraph()
p1.text = "Paragraph 2: spcBef=24pt (gap above this line)"
p1.font.size = Pt(14)
pPr1 = p1._p.find('a:pPr', nsmap)
if pPr1 is None:
    pPr1 = etree.SubElement(p1._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p1._p.insert(0, pPr1)
spcBef = etree.SubElement(pPr1, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcBef')
spcPts = etree.SubElement(spcBef, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPts')
spcPts.set('val', '2400')  # 24pt in hundredths

# Third paragraph - indent marL
p2 = tf2.add_paragraph()
p2.text = "Paragraph 3: marL=1in (indented from left)"
p2.font.size = Pt(14)
pPr2 = p2._p.find('a:pPr', nsmap)
if pPr2 is None:
    pPr2 = etree.SubElement(p2._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p2._p.insert(0, pPr2)
pPr2.set('marL', str(914400))  # 1 inch in EMU

# Fourth paragraph - spacing after 18pt
p3 = tf2.add_paragraph()
p3.text = "Paragraph 4: spcAft=18pt (gap below this line)"
p3.font.size = Pt(14)
pPr3 = p3._p.find('a:pPr', nsmap)
if pPr3 is None:
    pPr3 = etree.SubElement(p3._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p3._p.insert(0, pPr3)
spcAft = etree.SubElement(pPr3, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcAft')
spcPtsA = etree.SubElement(spcAft, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPts')
spcPtsA.set('val', '1800')  # 18pt

# Fifth paragraph - after the gap
p4 = tf2.add_paragraph()
p4.text = "Paragraph 5: After the 18pt gap. Should be visibly separated from P4."
p4.font.size = Pt(14)

# ── Slide 3: Bullets ────────────────────────────────────────────────────────

slide3 = prs.slides.add_slide(blank)

title3 = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title3.text_frame.paragraphs[0].text = "Slide 3: Bullets & Auto-Numbering"
title3.text_frame.paragraphs[0].font.size = Pt(24)
title3.text_frame.paragraphs[0].font.bold = True

# Character bullets
box3a = slide3.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(2.5))
tf3a = box3a.text_frame
tf3a.word_wrap = True
fill3a = box3a.fill
fill3a.solid()
fill3a.fore_color.rgb = RGBColor(230, 245, 230)

bullet_items = [
    ("First bullet item", "\u2022"),   # •
    ("Second bullet item", "\u2022"),
    ("Third with dash", "\u2013"),      # –
    ("Fourth with arrow", "\u25B6"),    # ▶
]

for i, (text, char) in enumerate(bullet_items):
    if i == 0:
        p = tf3a.paragraphs[0]
    else:
        p = tf3a.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))   # 0.5in
    pPr.set('indent', str(-228600))  # -0.25in (hanging indent)
    buChar = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar.set('char', char)

# Auto-numbering
box3b = slide3.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(2.5))
tf3b = box3b.text_frame
tf3b.word_wrap = True
fill3b = box3b.fill
fill3b.solid()
fill3b.fore_color.rgb = RGBColor(230, 230, 245)

auto_items = ["Arabic period", "Second item", "Third item", "Fourth item"]
for i, text in enumerate(auto_items):
    if i == 0:
        p = tf3b.paragraphs[0]
    else:
        p = tf3b.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buAutoNum = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum')
    buAutoNum.set('type', 'arabicPeriod')

# Roman numerals
box3c = slide3.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(4), Inches(2.5))
tf3c = box3c.text_frame
tf3c.word_wrap = True
fill3c = box3c.fill
fill3c.solid()
fill3c.fore_color.rgb = RGBColor(245, 240, 230)

roman_items = ["Roman upper item", "Another item", "Third one"]
for i, text in enumerate(roman_items):
    if i == 0:
        p = tf3c.paragraphs[0]
    else:
        p = tf3c.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buAutoNum = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum')
    buAutoNum.set('type', 'romanUcPeriod')

# Alpha lowercase
box3d = slide3.shapes.add_textbox(Inches(5), Inches(4.2), Inches(4.5), Inches(2.5))
tf3d = box3d.text_frame
tf3d.word_wrap = True
fill3d = box3d.fill
fill3d.solid()
fill3d.fore_color.rgb = RGBColor(240, 245, 245)

alpha_items = ["Alpha lower a", "Alpha lower b", "Alpha lower c"]
for i, text in enumerate(alpha_items):
    if i == 0:
        p = tf3d.paragraphs[0]
    else:
        p = tf3d.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buAutoNum = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum')
    buAutoNum.set('type', 'alphaLcParenR')

# ── Slide 4: Run decorations ───────────────────────────────────────────────

slide4 = prs.slides.add_slide(blank)

title4 = slide4.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title4.text_frame.paragraphs[0].text = "Slide 4: Run Decorations"
title4.text_frame.paragraphs[0].font.size = Pt(24)
title4.text_frame.paragraphs[0].font.bold = True

box4 = slide4.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(5))
tf4 = box4.text_frame
tf4.word_wrap = True

# Underline (single)
p_u = tf4.paragraphs[0]
run_u = p_u.add_run()
run_u.text = "This text has single underline"
run_u.font.size = Pt(18)
run_u.font.underline = True

# Strikethrough
p_s = tf4.add_paragraph()
run_s = p_s.add_run()
run_s.text = "This text has strikethrough"
run_s.font.size = Pt(18)
# Set strike via XML
rPr = run_s._r.find('a:rPr', nsmap)
if rPr is None:
    rPr = etree.SubElement(run_s._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run_s._r.insert(0, rPr)
rPr.set('strike', 'sngStrike')

# Both underline and strikethrough
p_us = tf4.add_paragraph()
run_us = p_us.add_run()
run_us.text = "Both underline AND strikethrough"
run_us.font.size = Pt(18)
run_us.font.underline = True
rPr_us = run_us._r.find('a:rPr', nsmap)
rPr_us.set('strike', 'sngStrike')

# Superscript
p_sup = tf4.add_paragraph()
run_normal = p_sup.add_run()
run_normal.text = "E = mc"
run_normal.font.size = Pt(18)
run_super = p_sup.add_run()
run_super.text = "2"
run_super.font.size = Pt(18)
rPr_sup = run_super._r.find('a:rPr', nsmap)
if rPr_sup is None:
    rPr_sup = etree.SubElement(run_super._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run_super._r.insert(0, rPr_sup)
rPr_sup.set('baseline', '30000')

# Subscript
p_sub = tf4.add_paragraph()
run_h2 = p_sub.add_run()
run_h2.text = "H"
run_h2.font.size = Pt(18)
run_2 = p_sub.add_run()
run_2.text = "2"
run_2.font.size = Pt(18)
rPr_2 = run_2._r.find('a:rPr', nsmap)
if rPr_2 is None:
    rPr_2 = etree.SubElement(run_2._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run_2._r.insert(0, rPr_2)
rPr_2.set('baseline', '-25000')
run_o = p_sub.add_run()
run_o.text = "O (water)"
run_o.font.size = Pt(18)

# Mixed decorations in one paragraph
p_mix = tf4.add_paragraph()
run_b = p_mix.add_run()
run_b.text = "Bold "
run_b.font.bold = True
run_b.font.size = Pt(18)
run_i = p_mix.add_run()
run_i.text = "Italic "
run_i.font.italic = True
run_i.font.size = Pt(18)
run_bi = p_mix.add_run()
run_bi.text = "BoldItalicUnderline"
run_bi.font.bold = True
run_bi.font.italic = True
run_bi.font.underline = True
run_bi.font.size = Pt(18)

# ── Slide 5: Combined test ──────────────────────────────────────────────────

slide5 = prs.slides.add_slide(blank)
# Dark background
bg = slide5.background
fill_bg = bg.fill
fill_bg.solid()
fill_bg.fore_color.rgb = RGBColor(27, 58, 107)

title5 = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
tf5t = title5.text_frame
tf5t.paragraphs[0].text = "Slide 5: Combined Features"
tf5t.paragraphs[0].font.size = Pt(28)
tf5t.paragraphs[0].font.bold = True
tf5t.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
bodyPr5 = tf5t._txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
bodyPr5.set('anchor', 'ctr')

# Box with center-anchored, spaced, bulleted content
box5 = slide5.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5))
tf5 = box5.text_frame
tf5.word_wrap = True
fill5 = box5.fill
fill5.solid()
fill5.fore_color.rgb = RGBColor(240, 240, 250)

bodyPr5b = tf5._txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
bodyPr5b.set('anchor', 'ctr')
bodyPr5b.set('lIns', str(Emu(Pt(18))))
bodyPr5b.set('tIns', str(Emu(Pt(18))))

items5 = [
    "Center-anchored box with bullets",
    "Second item with spacing",
    "Third item with underline run",
]
for i, text in enumerate(items5):
    if i == 0:
        p = tf5.paragraphs[0]
    else:
        p = tf5.add_paragraph()
    p.text = text
    p.font.size = Pt(16)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buChar = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar.set('char', '\u2022')
    # Add spacing before
    spcBef = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcBef')
    spcPts = etree.SubElement(spcBef, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPts')
    spcPts.set('val', '1200')  # 12pt

# ── Set up slide master for inheritance testing ────────────────────────────
# Modify the slide master to have:
#   - A dark navy background (#1B3A6B)
#   - Title style: 44pt, centered, white text
#   - Body style: 24pt, left-aligned, light gray text

sld_master = prs.slide_masters[0]
nsmap_p = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
           'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
           'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

# Set master background
master_xml = sld_master._element
# Remove existing bg if any
for old_bg in master_xml.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}bg'):
    master_xml.remove(old_bg)
# Find cSld or create it
cSld = master_xml.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
if cSld is None:
    cSld = etree.SubElement(master_xml, '{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
# Remove existing bg from cSld
for old_bg in cSld.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}bg'):
    cSld.remove(old_bg)
# Add background: navy blue
bg_elem = etree.SubElement(cSld, '{http://schemas.openxmlformats.org/presentationml/2006/main}bg')
cSld.insert(0, bg_elem)  # bg should be first child of cSld
bgPr = etree.SubElement(bg_elem, '{http://schemas.openxmlformats.org/presentationml/2006/main}bgPr')
solidFill_bg = etree.SubElement(bgPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
srgbClr_bg = etree.SubElement(solidFill_bg, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr_bg.set('val', '1B3A6B')
effectLst = etree.SubElement(bgPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst')

# Set up txStyles on the master
# Remove existing txStyles if any
for old_ts in master_xml.findall('{http://schemas.openxmlformats.org/presentationml/2006/main}txStyles'):
    master_xml.remove(old_ts)

txStyles = etree.SubElement(master_xml, '{http://schemas.openxmlformats.org/presentationml/2006/main}txStyles')

# Title style: 44pt, centered, white
titleStyle = etree.SubElement(txStyles, '{http://schemas.openxmlformats.org/presentationml/2006/main}titleStyle')
lvl1pPr_t = etree.SubElement(titleStyle, '{http://schemas.openxmlformats.org/drawingml/2006/main}lvl1pPr')
lvl1pPr_t.set('algn', 'ctr')
defRPr_t = etree.SubElement(lvl1pPr_t, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
defRPr_t.set('sz', '4400')  # 44pt
solidFill_t = etree.SubElement(defRPr_t, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
srgbClr_t = etree.SubElement(solidFill_t, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr_t.set('val', 'FFFFFF')

# Body style: lvl1=24pt left white, lvl2=20pt left lightgray
bodyStyle = etree.SubElement(txStyles, '{http://schemas.openxmlformats.org/presentationml/2006/main}bodyStyle')
lvl1pPr_b = etree.SubElement(bodyStyle, '{http://schemas.openxmlformats.org/drawingml/2006/main}lvl1pPr')
lvl1pPr_b.set('algn', 'l')
lvl1pPr_b.set('marL', '342900')
buChar_b1 = etree.SubElement(lvl1pPr_b, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
buChar_b1.set('char', '\u2022')
defRPr_b1 = etree.SubElement(lvl1pPr_b, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
defRPr_b1.set('sz', '2400')  # 24pt
solidFill_b1 = etree.SubElement(defRPr_b1, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
srgbClr_b1 = etree.SubElement(solidFill_b1, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr_b1.set('val', 'FFFFFF')

lvl2pPr_b = etree.SubElement(bodyStyle, '{http://schemas.openxmlformats.org/drawingml/2006/main}lvl2pPr')
lvl2pPr_b.set('algn', 'l')
lvl2pPr_b.set('marL', '685800')
buChar_b2 = etree.SubElement(lvl2pPr_b, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
buChar_b2.set('char', '\u2013')
defRPr_b2 = etree.SubElement(lvl2pPr_b, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
defRPr_b2.set('sz', '2000')  # 20pt
solidFill_b2 = etree.SubElement(defRPr_b2, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
srgbClr_b2 = etree.SubElement(solidFill_b2, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr_b2.set('val', 'CCCCCC')

# Other style: 18pt, gray
otherStyle = etree.SubElement(txStyles, '{http://schemas.openxmlformats.org/presentationml/2006/main}otherStyle')
lvl1pPr_o = etree.SubElement(otherStyle, '{http://schemas.openxmlformats.org/drawingml/2006/main}lvl1pPr')
defRPr_o = etree.SubElement(lvl1pPr_o, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
defRPr_o.set('sz', '1800')
solidFill_o = etree.SubElement(defRPr_o, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
srgbClr_o = etree.SubElement(solidFill_o, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr_o.set('val', '999999')

# ── Slide 6: No explicit background — should inherit navy from master ──────
slide6 = prs.slides.add_slide(blank)
# Do NOT set any background on this slide

title6 = slide6.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
tf6t = title6.text_frame
tf6t.paragraphs[0].text = "Slide 6: Inherits master navy bg"
tf6t.paragraphs[0].font.size = Pt(24)
tf6t.paragraphs[0].font.bold = True
tf6t.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

info6 = slide6.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(3))
tf6i = info6.text_frame
tf6i.word_wrap = True
p6 = tf6i.paragraphs[0]
p6.text = "Navy bg (#1B3A6B) = inheritance OK"
p6.font.size = Pt(18)
p6.font.color.rgb = RGBColor(200, 220, 255)
p6b = tf6i.add_paragraph()
p6b.text = "White bg = inheritance FAILED"
p6b.font.size = Pt(18)
p6b.font.color.rgb = RGBColor(255, 150, 150)

# ── Slide 7: Placeholder shapes with text style defaults ──────────────────
# Use title+content layout (index 1) to get real placeholders
title_content_layout = prs.slide_layouts[1]  # Title and Content layout
slide7 = prs.slides.add_slide(title_content_layout)

# The title placeholder should pick up master titleStyle (44pt ctr white)
title_ph = slide7.placeholders[0]  # Title placeholder
title_ph.text = "Slide 7: Placeholder Title (inherits 44pt ctr white)"

# The content placeholder should pick up master bodyStyle
body_ph = slide7.placeholders[1]  # Content placeholder
body_ph.text_frame.clear()
p7a = body_ph.text_frame.paragraphs[0]
p7a.text = "Level 0: Should inherit 24pt white from master bodyStyle"
# Don't set font size or color — let it inherit!

p7b = body_ph.text_frame.add_paragraph()
p7b.text = "Level 1: Should inherit 20pt lightgray from master bodyStyle lvl2"
p7b.level = 1

p7c = body_ph.text_frame.add_paragraph()
p7c.text = "Level 0 again: back to 24pt white"

# Also add a non-placeholder textbox for comparison
info7 = slide7.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(9), Inches(1.5))
tf7i = info7.text_frame
tf7i.word_wrap = True
p7info = tf7i.paragraphs[0]
p7info.text = "^ Title should be 44pt centered white. Body bullets should be 24pt/20pt white/gray."
p7info.font.size = Pt(12)
p7info.font.color.rgb = RGBColor(180, 200, 255)

# ── Slide 8: Explicit background overrides master ─────────────────────────
slide8 = prs.slides.add_slide(blank)
# Set an explicit green background — should NOT inherit master navy
bg8 = slide8.background
fill_bg8 = bg8.fill
fill_bg8.solid()
fill_bg8.fore_color.rgb = RGBColor(34, 139, 34)  # Forest green

title8 = slide8.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
tf8t = title8.text_frame
tf8t.paragraphs[0].text = "Slide 8: Explicit green bg (#228B22) — overrides master"
tf8t.paragraphs[0].font.size = Pt(20)
tf8t.paragraphs[0].font.bold = True
tf8t.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

info8 = slide8.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(3))
tf8i = info8.text_frame
tf8i.word_wrap = True
p8 = tf8i.paragraphs[0]
p8.text = "This slide has an explicit green background. It should NOT show the master's navy."
p8.font.size = Pt(18)
p8.font.color.rgb = RGBColor(220, 255, 220)
p8b = tf8i.add_paragraph()
p8b.text = "If you see green, explicit bg override is working correctly."
p8b.font.size = Pt(18)
p8b.font.color.rgb = RGBColor(255, 255, 255)

# Save
output_path = 'test_fixtures/test_features.pptx'
prs.save(output_path)
print(f"Saved {output_path} with {len(prs.slides)} slides")
