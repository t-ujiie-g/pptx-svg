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
  9. East Asian fonts + font theme references (+mj-lt/+mn-lt/+mj-ea/+mn-ea)
 10. Line spacing (a:lnSpc) — EMU and percentage
 11. Character spacing (a:rPr spc) + lstStyle
 12. normAutofit (fontScale / lnSpcReduction)
 13. Text wrapping (long text, CJK, wrap="none")
 14. Bullet formatting (a:buFont, a:buSzPct, a:buSzPts, a:buClr)
 15. Capitalization (a:rPr cap="all" / "small")
 16. Color map override (p:clrMapOvr — dark bg + light text via bg1↔dk1 swap)
 17. CS/Sym fonts + kerning (a:cs, a:sym, kern attr)
 18. Text rotation (bodyPr rot) + tab stops (a:tabLst)
 19. Vertical text (bodyPr vert) + text columns (numCol/spcCol)
 20. Hyperlink (a:hlinkClick) + RTL (a:pPr rtl)
 21. Image bullet (a:buBlip)
 22. Hover link + link color (a:hlinkHover)
 23. Linear gradient fills (a:gradFill + a:lin at 0°/90°/45°)
 24. Radial/path gradient fills (a:path circle/rect) + gradient on ellipse
 25. Gradient background (p:bg → p:bgPr → a:gradFill)
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

# ── Set theme EA fonts ────────────────────────────────────────────────────
# Modify theme1.xml to add East Asian font entries
# Access theme via slide master's relationship
ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
theme_part = None
for rel in sld_master.part.rels.values():
    if 'theme' in rel.reltype:
        theme_part = rel.target_part
        break

if theme_part is not None:
    theme_elem = etree.fromstring(theme_part.blob)
    font_scheme = theme_elem.find(f'.//{{{ns_a}}}fontScheme')
    if font_scheme is not None:
        # Add a:ea to majorFont
        major_font_elem = font_scheme.find(f'{{{ns_a}}}majorFont')
        if major_font_elem is not None:
            for old_ea in major_font_elem.findall(f'{{{ns_a}}}ea'):
                major_font_elem.remove(old_ea)
            ea_major = etree.SubElement(major_font_elem, f'{{{ns_a}}}ea')
            ea_major.set('typeface', 'Yu Gothic')
        # Add a:ea to minorFont
        minor_font_elem = font_scheme.find(f'{{{ns_a}}}minorFont')
        if minor_font_elem is not None:
            for old_ea in minor_font_elem.findall(f'{{{ns_a}}}ea'):
                minor_font_elem.remove(old_ea)
            ea_minor = etree.SubElement(minor_font_elem, f'{{{ns_a}}}ea')
            ea_minor.set('typeface', 'Yu Gothic')
    # Write back the modified theme XML
    theme_part._blob = etree.tostring(theme_elem, xml_declaration=True, encoding='UTF-8', standalone=True)
    print("Theme EA fonts set: Yu Gothic")

# ── Slide 9: East Asian fonts + font theme references ───────────────────
slide9 = prs.slides.add_slide(blank)

title9 = slide9.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title9.text_frame.paragraphs[0].text = "Slide 9: East Asian Fonts & Theme References"
title9.text_frame.paragraphs[0].font.size = Pt(24)
title9.text_frame.paragraphs[0].font.bold = True

# Box with explicit EA font (a:ea typeface="MS PGothic")
box9a = slide9.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(2))
tf9a = box9a.text_frame
tf9a.word_wrap = True
fill9a = box9a.fill
fill9a.solid()
fill9a.fore_color.rgb = RGBColor(240, 240, 255)
p9a = tf9a.paragraphs[0]
run9a = p9a.add_run()
run9a.text = "日本語テキスト（MS PGothic）"
run9a.font.size = Pt(18)
# Set a:ea typeface directly in XML
rPr9a = run9a._r.find('a:rPr', nsmap)
if rPr9a is None:
    rPr9a = etree.SubElement(run9a._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run9a._r.insert(0, rPr9a)
ea9a = etree.SubElement(rPr9a, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea9a.set('typeface', 'MS PGothic')
latin9a = etree.SubElement(rPr9a, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
latin9a.set('typeface', 'Arial')

# Second para: theme font reference +mj-ea and +mj-lt
p9b = tf9a.add_paragraph()
run9b = p9b.add_run()
run9b.text = "テーマ参照: +mj-ea / +mj-lt"
run9b.font.size = Pt(16)
rPr9b = run9b._r.find('a:rPr', nsmap)
if rPr9b is None:
    rPr9b = etree.SubElement(run9b._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run9b._r.insert(0, rPr9b)
ea9b = etree.SubElement(rPr9b, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea9b.set('typeface', '+mj-ea')
latin9b = etree.SubElement(rPr9b, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
latin9b.set('typeface', '+mj-lt')

# Third para: minor font theme reference +mn-ea and +mn-lt
p9c = tf9a.add_paragraph()
run9c = p9c.add_run()
run9c.text = "マイナーフォント: +mn-ea / +mn-lt"
run9c.font.size = Pt(14)
rPr9c = run9c._r.find('a:rPr', nsmap)
if rPr9c is None:
    rPr9c = etree.SubElement(run9c._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run9c._r.insert(0, rPr9c)
ea9c = etree.SubElement(rPr9c, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea9c.set('typeface', '+mn-ea')
latin9c = etree.SubElement(rPr9c, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
latin9c.set('typeface', '+mn-lt')

# Box with mixed Latin + EA text in same run
box9b = slide9.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(2))
tf9b = box9b.text_frame
tf9b.word_wrap = True
fill9b = box9b.fill
fill9b.solid()
fill9b.fore_color.rgb = RGBColor(255, 245, 235)
p9d = tf9b.paragraphs[0]
run9d = p9d.add_run()
run9d.text = "Mixed: ABC + あいう + 123"
run9d.font.size = Pt(18)
rPr9d = run9d._r.find('a:rPr', nsmap)
if rPr9d is None:
    rPr9d = etree.SubElement(run9d._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run9d._r.insert(0, rPr9d)
ea9d = etree.SubElement(rPr9d, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea9d.set('typeface', 'Meiryo')
latin9d = etree.SubElement(rPr9d, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
latin9d.set('typeface', 'Segoe UI')

# Info text
info9 = slide9.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(2))
tf9info = info9.text_frame
tf9info.word_wrap = True
p9info = tf9info.paragraphs[0]
p9info.text = "Verify: EA fonts should appear in font-family. Theme refs (+mj-ea/+mn-ea) should resolve to Yu Gothic."
p9info.font.size = Pt(12)
p9info.font.color.rgb = RGBColor(100, 100, 100)

# ── Slide 10: Line spacing ──────────────────────────────────────────────
slide10 = prs.slides.add_slide(blank)

title10 = slide10.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title10.text_frame.paragraphs[0].text = "Slide 10: Line Spacing (a:lnSpc)"
title10.text_frame.paragraphs[0].font.size = Pt(24)
title10.text_frame.paragraphs[0].font.bold = True

# Box with percentage-based line spacing (150%)
box10a = slide10.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(3))
tf10a = box10a.text_frame
tf10a.word_wrap = True
fill10a = box10a.fill
fill10a.solid()
fill10a.fore_color.rgb = RGBColor(230, 245, 255)

for i, text in enumerate(["Line 1: 150% spacing", "Line 2: wider gap above", "Line 3: still 150%"]):
    if i == 0:
        p = tf10a.paragraphs[0]
    else:
        p = tf10a.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    lnSpc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lnSpc')
    spcPct = etree.SubElement(lnSpc, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPct')
    spcPct.set('val', '150000')  # 150%

# Box with point-based line spacing (36pt absolute)
box10b = slide10.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(3))
tf10b = box10b.text_frame
tf10b.word_wrap = True
fill10b = box10b.fill
fill10b.solid()
fill10b.fore_color.rgb = RGBColor(255, 240, 230)

for i, text in enumerate(["Line 1: 36pt absolute", "Line 2: fixed height", "Line 3: still 36pt"]):
    if i == 0:
        p = tf10b.paragraphs[0]
    else:
        p = tf10b.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    lnSpc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lnSpc')
    spcPts = etree.SubElement(lnSpc, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPts')
    spcPts.set('val', '3600')  # 36pt

# Box with tight line spacing (80%)
box10c = slide10.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(4), Inches(2.5))
tf10c = box10c.text_frame
tf10c.word_wrap = True
fill10c = box10c.fill
fill10c.solid()
fill10c.fore_color.rgb = RGBColor(245, 255, 240)

for i, text in enumerate(["Tight: 80% spacing", "Lines should be close together", "Third tight line"]):
    if i == 0:
        p = tf10c.paragraphs[0]
    else:
        p = tf10c.add_paragraph()
    p.text = text
    p.font.size = Pt(14)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    lnSpc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lnSpc')
    spcPct = etree.SubElement(lnSpc, '{http://schemas.openxmlformats.org/drawingml/2006/main}spcPct')
    spcPct.set('val', '80000')  # 80%

# ── Slide 11: Character spacing + lstStyle ──────────────────────────────
slide11 = prs.slides.add_slide(blank)

title11 = slide11.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title11.text_frame.paragraphs[0].text = "Slide 11: Character Spacing & lstStyle"
title11.text_frame.paragraphs[0].font.size = Pt(24)
title11.text_frame.paragraphs[0].font.bold = True

# Box with character spacing variations
box11a = slide11.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(2.5))
tf11a = box11a.text_frame
tf11a.word_wrap = True
fill11a = box11a.fill
fill11a.solid()
fill11a.fore_color.rgb = RGBColor(245, 240, 255)

# Normal spacing
p11_0 = tf11a.paragraphs[0]
run11_0 = p11_0.add_run()
run11_0.text = "Normal spacing (spc=0)"
run11_0.font.size = Pt(16)

# Wide spacing (spc=300 = 3pt)
p11_1 = tf11a.add_paragraph()
run11_1 = p11_1.add_run()
run11_1.text = "Wide spacing (spc=300, 3pt)"
run11_1.font.size = Pt(16)
rPr11_1 = run11_1._r.find('a:rPr', nsmap)
if rPr11_1 is None:
    rPr11_1 = etree.SubElement(run11_1._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run11_1._r.insert(0, rPr11_1)
rPr11_1.set('spc', '300')

# Very wide spacing (spc=1000 = 10pt)
p11_2 = tf11a.add_paragraph()
run11_2 = p11_2.add_run()
run11_2.text = "Very wide (spc=1000, 10pt)"
run11_2.font.size = Pt(16)
rPr11_2 = run11_2._r.find('a:rPr', nsmap)
if rPr11_2 is None:
    rPr11_2 = etree.SubElement(run11_2._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run11_2._r.insert(0, rPr11_2)
rPr11_2.set('spc', '1000')

# Tight spacing (spc=-100 = -1pt)
p11_3 = tf11a.add_paragraph()
run11_3 = p11_3.add_run()
run11_3.text = "Tight spacing (spc=-100, -1pt)"
run11_3.font.size = Pt(16)
rPr11_3 = run11_3._r.find('a:rPr', nsmap)
if rPr11_3 is None:
    rPr11_3 = etree.SubElement(run11_3._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run11_3._r.insert(0, rPr11_3)
rPr11_3.set('spc', '-100')

# Box with lstStyle (shape-specific list style defaults)
box11b = slide11.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(2.5))
tf11b = box11b.text_frame
tf11b.word_wrap = True
fill11b = box11b.fill
fill11b.solid()
fill11b.fore_color.rgb = RGBColor(255, 250, 235)

# Add a:lstStyle to the txBody with level defaults
ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
txBody11b = tf11b._txBody
# Remove existing lstStyle
for old_ls in txBody11b.findall(f'{{{ns_a}}}lstStyle'):
    txBody11b.remove(old_ls)
lstStyle = etree.SubElement(txBody11b, f'{{{ns_a}}}lstStyle')
# Insert lstStyle after bodyPr (should be second child of txBody)
bodyPr11b = txBody11b.find(f'{{{ns_a}}}bodyPr')
if bodyPr11b is not None:
    idx = list(txBody11b).index(bodyPr11b) + 1
    txBody11b.remove(lstStyle)
    txBody11b.insert(idx, lstStyle)

# Level 1: 20pt, bold, dark blue
lvl1pPr_ls = etree.SubElement(lstStyle, f'{{{ns_a}}}lvl1pPr')
defRPr_ls1 = etree.SubElement(lvl1pPr_ls, f'{{{ns_a}}}defRPr')
defRPr_ls1.set('sz', '2000')  # 20pt
defRPr_ls1.set('b', '1')
solidFill_ls1 = etree.SubElement(defRPr_ls1, f'{{{ns_a}}}solidFill')
srgbClr_ls1 = etree.SubElement(solidFill_ls1, f'{{{ns_a}}}srgbClr')
srgbClr_ls1.set('val', '003366')
ea_ls1 = etree.SubElement(defRPr_ls1, f'{{{ns_a}}}ea')
ea_ls1.set('typeface', 'Meiryo')

p11_ls1 = tf11b.paragraphs[0]
p11_ls1.text = "lstStyle lvl1: inherits 20pt bold #003366 Meiryo"
# Don't set font size/color — let lstStyle provide defaults

p11_ls2 = tf11b.add_paragraph()
p11_ls2.text = "Second paragraph: should also inherit lstStyle defaults"

# ── Slide 12: normAutofit (fontScale / lnSpcReduction) ───────────────────
slide12 = prs.slides.add_slide(blank)

title12 = slide12.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title12.text_frame.paragraphs[0].text = "Slide 12: normAutofit (fontScale / lnSpcReduction)"
title12.text_frame.paragraphs[0].font.size = Pt(24)
title12.text_frame.paragraphs[0].font.bold = True

def add_normAutofit_box(slide, left, top, width, height, label, lines,
                         fill_rgb, border_rgb, font_scale=None, ln_spc_reduction=None):
    """Add a textbox with a:normAutofit in bodyPr."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    fill = box.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*fill_rgb)
    # Border
    ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    sp_pr = box._element.find(f'{{{ns_a}}}spPr')
    if sp_pr is None:
        sp_pr = box._element.find(f'{{{ns_p}}}spPr')
    if sp_pr is None:
        # textbox shapes: spPr is direct child of p:sp
        sp_pr = etree.SubElement(box._element, f'{{{ns_a}}}spPr')
    ln_elem = etree.SubElement(sp_pr, f'{{{ns_a}}}ln')
    ln_elem.set('w', '12700')
    ln_fill = etree.SubElement(ln_elem, f'{{{ns_a}}}solidFill')
    ln_clr = etree.SubElement(ln_fill, f'{{{ns_a}}}srgbClr')
    ln_clr.set('val', '%02X%02X%02X' % border_rgb)

    # Title line
    p0 = tf.paragraphs[0]
    run0 = p0.add_run()
    run0.text = label
    run0.font.size = Pt(20)
    run0.font.bold = True

    # Content lines
    for line in lines:
        p = tf.add_paragraph()
        run = p.add_run()
        run.text = line
        run.font.size = Pt(16)

    # Set bodyPr with normAutofit
    bodyPr = tf._txBody.find(f'{{{ns_a}}}bodyPr')
    bodyPr.set('lIns', '91440')
    bodyPr.set('tIns', '45720')
    bodyPr.set('rIns', '91440')
    bodyPr.set('bIns', '45720')

    if font_scale is not None:
        naf = etree.SubElement(bodyPr, f'{{{ns_a}}}normAutofit')
        if font_scale != 100000:
            naf.set('fontScale', str(font_scale))
        if ln_spc_reduction is not None and ln_spc_reduction > 0:
            naf.set('lnSpcReduction', str(ln_spc_reduction))

    return box

# Shape 1: fontScale=80000 (80%)
add_normAutofit_box(slide12, Inches(0.5), Inches(1.2), Inches(4), Inches(1.8),
    "fontScale=80%",
    ["Line 2: scaled down", "Line 3: smaller text", "Line 4: fits in box"],
    (232, 234, 246), (63, 81, 181),
    font_scale=80000)

# Shape 2: fontScale=62500 + lnSpcReduction=20000
add_normAutofit_box(slide12, Inches(5), Inches(1.2), Inches(4), Inches(1.8),
    "62.5% + lnSpc 20%",
    ["Line 2: even smaller", "Line 3: tighter spacing", "Line 4: compact text"],
    (255, 243, 224), (255, 152, 0),
    font_scale=62500, ln_spc_reduction=20000)

# Shape 3: normAutofit default (100%, no fontScale attr)
add_normAutofit_box(slide12, Inches(0.5), Inches(3.5), Inches(4), Inches(1.8),
    "normAutofit default",
    ["Line 2: no scaling", "Line 3: same as original"],
    (232, 245, 233), (76, 175, 80),
    font_scale=100000)

# Shape 4: no autofit (reference)
add_normAutofit_box(slide12, Inches(5), Inches(3.5), Inches(4), Inches(1.8),
    "No autofit (ref)",
    ["Line 2: original size", "Line 3: may overflow"],
    (243, 229, 245), (156, 39, 176),
    font_scale=None)  # no normAutofit

# Info label at bottom
info12 = slide12.shapes.add_textbox(Inches(0.5), Inches(5.8), Inches(9), Inches(1))
tf12info = info12.text_frame
tf12info.word_wrap = True
p12info = tf12info.paragraphs[0]
p12info.text = "Compare: top-left (80%) should be visibly smaller than bottom-right (no autofit). Top-right (62.5%+20%) should be smallest and tightest."
p12info.font.size = Pt(12)
p12info.font.color.rgb = RGBColor(100, 100, 100)

# ── Slide 13: Text wrapping ───────────────────────────────────────────────
slide13 = prs.slides.add_slide(blank)

title13 = slide13.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title13.text_frame.paragraphs[0].text = "Slide 13: Text Wrapping"
title13.text_frame.paragraphs[0].font.size = Pt(24)
title13.text_frame.paragraphs[0].font.bold = True

# Box with long Latin text that should wrap
box13a = slide13.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(2.5))
tf13a = box13a.text_frame
tf13a.word_wrap = True
fill13a = box13a.fill
fill13a.solid()
fill13a.fore_color.rgb = RGBColor(230, 240, 255)
p13a = tf13a.paragraphs[0]
p13a.text = "This is a long paragraph of Latin text that should automatically wrap at the shape boundary. Word-level wrapping should break at spaces."
p13a.font.size = Pt(16)

# Second paragraph: even longer
p13a2 = tf13a.add_paragraph()
p13a2.text = "Another paragraph with a very long sentence that exercises the wrapping algorithm more thoroughly, including multiple lines of content."
p13a2.font.size = Pt(14)

# Box with CJK text that should wrap character by character
box13b = slide13.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(2.5))
tf13b = box13b.text_frame
tf13b.word_wrap = True
fill13b = box13b.fill
fill13b.solid()
fill13b.fore_color.rgb = RGBColor(255, 245, 230)
p13b = tf13b.paragraphs[0]
run13b = p13b.add_run()
run13b.text = "日本語テキストの折り返しテスト。文字単位で折り返されるべきです。漢字とひらがなが混在しています。"
run13b.font.size = Pt(16)
rPr13b = run13b._r.find('a:rPr', nsmap)
if rPr13b is None:
    rPr13b = etree.SubElement(run13b._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run13b._r.insert(0, rPr13b)
ea13b = etree.SubElement(rPr13b, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea13b.set('typeface', 'Yu Gothic')

# Box with mixed Latin + CJK text
p13b2 = tf13b.add_paragraph()
run13b2 = p13b2.add_run()
run13b2.text = "Mixed テキスト: ABC あいう 123 日本語 English words and 漢字が混在"
run13b2.font.size = Pt(14)
rPr13b2 = run13b2._r.find('a:rPr', nsmap)
if rPr13b2 is None:
    rPr13b2 = etree.SubElement(run13b2._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run13b2._r.insert(0, rPr13b2)
ea13b2 = etree.SubElement(rPr13b2, '{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
ea13b2.set('typeface', 'Yu Gothic')

# Box with wrap="none" — text should NOT wrap
box13c = slide13.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(4), Inches(1.5))
tf13c = box13c.text_frame
tf13c.word_wrap = False  # This sets wrap="none"
fill13c = box13c.fill
fill13c.solid()
fill13c.fore_color.rgb = RGBColor(255, 230, 230)
p13c = tf13c.paragraphs[0]
p13c.text = "wrap=none: This text should NOT wrap even though it is long and exceeds the box width"
p13c.font.size = Pt(16)
# Ensure wrap="none" is set in bodyPr
bodyPr13c = tf13c._txBody.find('{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
bodyPr13c.set('wrap', 'none')

# Narrow box to test wrapping with large font
box13d = slide13.shapes.add_textbox(Inches(5), Inches(4.2), Inches(2.5), Inches(2.5))
tf13d = box13d.text_frame
tf13d.word_wrap = True
fill13d = box13d.fill
fill13d.solid()
fill13d.fore_color.rgb = RGBColor(230, 255, 230)
p13d = tf13d.paragraphs[0]
p13d.text = "Large font in narrow box"
p13d.font.size = Pt(28)
p13d2 = tf13d.add_paragraph()
p13d2.text = "Should wrap to multiple lines"
p13d2.font.size = Pt(20)

# ── Slide 14: Bullet formatting ──────────────────────────────────────────
slide14 = prs.slides.add_slide(blank)

title14 = slide14.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title14.text_frame.paragraphs[0].text = "Slide 14: Bullet Formatting (buFont / buSzPct / buSzPts / buClr)"
title14.text_frame.paragraphs[0].font.size = Pt(22)
title14.text_frame.paragraphs[0].font.bold = True

# Box with custom bullet font (Wingdings)
box14a = slide14.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(2.5))
tf14a = box14a.text_frame
tf14a.word_wrap = True
fill14a = box14a.fill
fill14a.solid()
fill14a.fore_color.rgb = RGBColor(240, 235, 255)

bullet_format_items = [
    ("Wingdings bullet font", "Wingdings", "\u006C"),       # Wingdings 'l' = checkmark-like
    ("Symbol bullet font", "Symbol", "\u00B7"),              # Symbol middle dot
    ("Default font bullet", "", "\u2022"),                    # Regular bullet
]

for i, (text, bu_font, bu_char) in enumerate(bullet_format_items):
    if i == 0:
        p = tf14a.paragraphs[0]
    else:
        p = tf14a.add_paragraph()
    p.text = text
    p.font.size = Pt(16)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    if bu_font:
        buFontElem = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buFont')
        buFontElem.set('typeface', bu_font)
    buChar14 = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar14.set('char', bu_char)

# Box with bullet size (percentage + points)
box14b = slide14.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(2.5))
tf14b = box14b.text_frame
tf14b.word_wrap = True
fill14b = box14b.fill
fill14b.solid()
fill14b.fore_color.rgb = RGBColor(230, 255, 240)

bu_size_items = [
    ("buSzPct=150000 (150%)", 150000, None),
    ("buSzPct=75000 (75%)", 75000, None),
    ("buSzPts=3200 (32pt)", None, 3200),
    ("buSzPts=800 (8pt)", None, 800),
]

for i, (text, sz_pct, sz_pts) in enumerate(bu_size_items):
    if i == 0:
        p = tf14b.paragraphs[0]
    else:
        p = tf14b.add_paragraph()
    p.text = text
    p.font.size = Pt(16)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buChar14b = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar14b.set('char', '\u2022')
    if sz_pct is not None:
        buSzPct = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buSzPct')
        buSzPct.set('val', str(sz_pct))
    if sz_pts is not None:
        buSzPts = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buSzPts')
        buSzPts.set('val', str(sz_pts))

# Box with bullet color
box14c = slide14.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(4), Inches(2.5))
tf14c = box14c.text_frame
tf14c.word_wrap = True
fill14c = box14c.fill
fill14c.solid()
fill14c.fore_color.rgb = RGBColor(255, 245, 235)

bu_color_items = [
    ("Red bullet", "FF0000"),
    ("Green bullet", "00AA00"),
    ("Blue bullet", "0000FF"),
    ("No bullet color (inherits)", None),
]

for i, (text, bu_clr) in enumerate(bu_color_items):
    if i == 0:
        p = tf14c.paragraphs[0]
    else:
        p = tf14c.add_paragraph()
    p.text = text
    p.font.size = Pt(16)
    pPr = p._p.find('a:pPr', nsmap)
    if pPr is None:
        pPr = etree.SubElement(p._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        p._p.insert(0, pPr)
    pPr.set('marL', str(457200))
    pPr.set('indent', str(-228600))
    buChar14c = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar14c.set('char', '\u2022')
    if bu_clr is not None:
        buClrElem = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buClr')
        srgbClr14 = etree.SubElement(buClrElem, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
        srgbClr14.set('val', bu_clr)

# Box with combined: buFont + buSzPct + buClr
box14d = slide14.shapes.add_textbox(Inches(5), Inches(4.2), Inches(4.5), Inches(2.5))
tf14d = box14d.text_frame
tf14d.word_wrap = True
fill14d = box14d.fill
fill14d.solid()
fill14d.fore_color.rgb = RGBColor(245, 240, 255)

p14d = tf14d.paragraphs[0]
p14d.text = "Combined: Wingdings + 120% + red"
p14d.font.size = Pt(16)
pPr14d = p14d._p.find('a:pPr', nsmap)
if pPr14d is None:
    pPr14d = etree.SubElement(p14d._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p14d._p.insert(0, pPr14d)
pPr14d.set('marL', str(457200))
pPr14d.set('indent', str(-228600))
buFont14d = etree.SubElement(pPr14d, '{http://schemas.openxmlformats.org/drawingml/2006/main}buFont')
buFont14d.set('typeface', 'Wingdings')
buSzPct14d = etree.SubElement(pPr14d, '{http://schemas.openxmlformats.org/drawingml/2006/main}buSzPct')
buSzPct14d.set('val', '120000')
buClr14d = etree.SubElement(pPr14d, '{http://schemas.openxmlformats.org/drawingml/2006/main}buClr')
srgbClr14d = etree.SubElement(buClr14d, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
srgbClr14d.set('val', 'FF0000')
buChar14d = etree.SubElement(pPr14d, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
buChar14d.set('char', '\u006C')

p14d2 = tf14d.add_paragraph()
p14d2.text = "Normal bullet for comparison"
p14d2.font.size = Pt(16)
pPr14d2 = p14d2._p.find('a:pPr', nsmap)
if pPr14d2 is None:
    pPr14d2 = etree.SubElement(p14d2._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p14d2._p.insert(0, pPr14d2)
pPr14d2.set('marL', str(457200))
pPr14d2.set('indent', str(-228600))
buChar14d2 = etree.SubElement(pPr14d2, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
buChar14d2.set('char', '\u2022')

# ── Slide 15: Capitalization (cap) ──────────────────────────────────────
slide15 = prs.slides.add_slide(blank)

title15 = slide15.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
title15.text_frame.paragraphs[0].text = "Slide 15: Capitalization (a:rPr cap)"
title15.text_frame.paragraphs[0].font.size = Pt(24)
title15.text_frame.paragraphs[0].font.bold = True

# Box with cap="all"
box15a = slide15.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4), Inches(2.5))
tf15a = box15a.text_frame
tf15a.word_wrap = True
fill15a = box15a.fill
fill15a.solid()
fill15a.fore_color.rgb = RGBColor(230, 245, 255)

p15a = tf15a.paragraphs[0]
run15a = p15a.add_run()
run15a.text = "cap=all: This Should Be All Caps"
run15a.font.size = Pt(18)
rPr15a = run15a._r.find('a:rPr', nsmap)
if rPr15a is None:
    rPr15a = etree.SubElement(run15a._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run15a._r.insert(0, rPr15a)
rPr15a.set('cap', 'all')

p15a2 = tf15a.add_paragraph()
run15a2 = p15a2.add_run()
run15a2.text = "Mixed: "
run15a2.font.size = Pt(16)
run15a2b = p15a2.add_run()
run15a2b.text = "all caps part"
run15a2b.font.size = Pt(16)
rPr15a2b = run15a2b._r.find('a:rPr', nsmap)
if rPr15a2b is None:
    rPr15a2b = etree.SubElement(run15a2b._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run15a2b._r.insert(0, rPr15a2b)
rPr15a2b.set('cap', 'all')
run15a2c = p15a2.add_run()
run15a2c.text = " and normal part"
run15a2c.font.size = Pt(16)

p15a3 = tf15a.add_paragraph()
run15a3 = p15a3.add_run()
run15a3.text = "No cap: Regular text for comparison"
run15a3.font.size = Pt(16)

# Box with cap="small"
box15b = slide15.shapes.add_textbox(Inches(5), Inches(1.2), Inches(4.5), Inches(2.5))
tf15b = box15b.text_frame
tf15b.word_wrap = True
fill15b = box15b.fill
fill15b.solid()
fill15b.fore_color.rgb = RGBColor(255, 245, 230)

p15b = tf15b.paragraphs[0]
run15b = p15b.add_run()
run15b.text = "cap=small: Small Caps Text"
run15b.font.size = Pt(18)
rPr15b = run15b._r.find('a:rPr', nsmap)
if rPr15b is None:
    rPr15b = etree.SubElement(run15b._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run15b._r.insert(0, rPr15b)
rPr15b.set('cap', 'small')

p15b2 = tf15b.add_paragraph()
run15b2 = p15b2.add_run()
run15b2.text = "No cap: Regular text"
run15b2.font.size = Pt(16)

p15b3 = tf15b.add_paragraph()
run15b3 = p15b3.add_run()
run15b3.text = "Bold Small Caps"
run15b3.font.size = Pt(18)
run15b3.font.bold = True
rPr15b3 = run15b3._r.find('a:rPr', nsmap)
if rPr15b3 is None:
    rPr15b3 = etree.SubElement(run15b3._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    run15b3._r.insert(0, rPr15b3)
rPr15b3.set('cap', 'small')

# Info text
info15 = slide15.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(2))
tf15info = info15.text_frame
tf15info.word_wrap = True
p15info = tf15info.paragraphs[0]
p15info.text = "cap=all: lowercase a-z should display as UPPERCASE A-Z. cap=small: should display with SVG font-variant=small-caps."
p15info.font.size = Pt(12)
p15info.font.color.rgb = RGBColor(100, 100, 100)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 16: Color map override (p:clrMapOvr)
# ═══════════════════════════════════════════════════════════════════════════════
slide16 = prs.slides.add_slide(prs.slide_layouts[0])

nsmap = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# Remove existing p:clrMapOvr (python-pptx adds a:masterClrMapping by default)
existing_ovr = slide16._element.findall(
    '{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr'
)
for ovr in existing_ovr:
    slide16._element.remove(ovr)

# Add p:clrMapOvr with a:overrideClrMapping that swaps bg1↔tx1 (dk1↔lt1)
# This makes the slide use dark background + light text
clrMapOvr = etree.SubElement(
    slide16._element,
    '{http://schemas.openxmlformats.org/presentationml/2006/main}clrMapOvr'
)
overrideMapping = etree.SubElement(
    clrMapOvr,
    '{http://schemas.openxmlformats.org/drawingml/2006/main}overrideClrMapping',
    attrib={
        'bg1': 'dk1',
        'tx1': 'lt1',
        'bg2': 'lt2',
        'tx2': 'dk2',
        'accent1': 'accent1',
        'accent2': 'accent2',
        'accent3': 'accent3',
        'accent4': 'accent4',
        'accent5': 'accent5',
        'accent6': 'accent6',
        'hlink': 'hlink',
        'folHlink': 'folHlink',
    }
)

# Add a textbox with scheme-colored text (should resolve to light color on dark bg)
tb16 = slide16.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.5))
tf16 = tb16.text_frame
tf16.word_wrap = True
p16 = tf16.paragraphs[0]
p16.text = "Color Map Override: bg1=dk1, tx1=lt1"
p16.font.size = Pt(28)
p16.font.bold = True
# Use scheme color tx1 (which should resolve to lt1 = white on this slide)
rPr16 = p16.runs[0]._r.find('a:rPr', nsmap)
if rPr16 is None:
    rPr16 = etree.SubElement(p16.runs[0]._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    p16.runs[0]._r.insert(0, rPr16)
solidFill16 = etree.SubElement(rPr16, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
schemeClr16 = etree.SubElement(solidFill16, '{http://schemas.openxmlformats.org/drawingml/2006/main}schemeClr', attrib={'val': 'tx1'})

# Add a second textbox with explicit description
tb16b = slide16.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(2))
tf16b = tb16b.text_frame
tf16b.word_wrap = True
p16b = tf16b.paragraphs[0]
p16b.text = "This slide has clrMapOvr swapping bg1/tx1. Text using scheme color tx1 should appear light (white) because tx1 maps to lt1."
p16b.font.size = Pt(16)
# Also use scheme color for this text
rPr16b = p16b.runs[0]._r.find('a:rPr', nsmap)
if rPr16b is None:
    rPr16b = etree.SubElement(p16b.runs[0]._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    p16b.runs[0]._r.insert(0, rPr16b)
solidFill16b = etree.SubElement(rPr16b, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
schemeClr16b = etree.SubElement(solidFill16b, '{http://schemas.openxmlformats.org/drawingml/2006/main}schemeClr', attrib={'val': 'tx1'})

# Add slide background using scheme color bg1 (which maps to dk1 = dark)
bg16 = slide16._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld', nsmap)
if bg16 is None:
    bg16 = slide16._element.find('p:cSld', nsmap)
# Insert p:bg before p:cSld
pBg = etree.Element('{http://schemas.openxmlformats.org/presentationml/2006/main}bg')
bgPr = etree.SubElement(pBg, '{http://schemas.openxmlformats.org/presentationml/2006/main}bgPr')
solidFillBg = etree.SubElement(bgPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
# Use a dark solid color for background to make the slide visually dark
srgbClrBg = etree.SubElement(solidFillBg, '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr', attrib={'val': '1A1A2E'})
# Insert bg element at the right position in slide XML
cSld = slide16._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
if cSld is not None:
    slide16._element.insert(list(slide16._element).index(cSld), pBg)

# ── Slide 17: CS/Sym fonts + kerning ─────────────────────────────────────────

slide17 = prs.slides.add_slide(blank)

# Textbox with Complex Script font (a:cs)
tb17a = slide17.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.5))
tf17a = tb17a.text_frame
tf17a.word_wrap = True
p17a = tf17a.paragraphs[0]
p17a.text = "Complex Script Font (Arabic style)"
p17a.font.size = Pt(24)
p17a.font.bold = True
# Add a:cs element to rPr
rPr17a = p17a.runs[0]._r.find('a:rPr', nsmap)
if rPr17a is None:
    rPr17a = etree.SubElement(p17a.runs[0]._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    p17a.runs[0]._r.insert(0, rPr17a)
etree.SubElement(rPr17a, '{http://schemas.openxmlformats.org/drawingml/2006/main}cs', attrib={'typeface': 'Arial'})

# Textbox with Symbol font (a:sym)
tb17b = slide17.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
tf17b = tb17b.text_frame
p17b = tf17b.paragraphs[0]
p17b.text = "Symbol Font Text"
p17b.font.size = Pt(24)
rPr17b = p17b.runs[0]._r.find('a:rPr', nsmap)
if rPr17b is None:
    rPr17b = etree.SubElement(p17b.runs[0]._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    p17b.runs[0]._r.insert(0, rPr17b)
etree.SubElement(rPr17b, '{http://schemas.openxmlformats.org/drawingml/2006/main}sym', attrib={'typeface': 'Wingdings'})

# Textbox with kerning
tb17c = slide17.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(1.5))
tf17c = tb17c.text_frame
p17c = tf17c.paragraphs[0]
p17c.text = "Kerning enabled at 1200 hundredths-pt"
p17c.font.size = Pt(24)
rPr17c = p17c.runs[0]._r.find('a:rPr', nsmap)
if rPr17c is None:
    rPr17c = etree.SubElement(p17c.runs[0]._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    p17c.runs[0]._r.insert(0, rPr17c)
rPr17c.set('kern', '1200')

# ── Slide 18: Text rotation + tab stops ──────────────────────────────────────

slide18 = prs.slides.add_slide(blank)

# Textbox with rotated text (bodyPr rot)
tb18a = slide18.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(2))
tf18a = tb18a.text_frame
p18a = tf18a.paragraphs[0]
p18a.text = "Rotated text (45 degrees)"
p18a.font.size = Pt(20)
# Set rot on bodyPr (45 degrees = 2700000 in 60000ths)
bodyPr18a = tb18a._element.find('.//a:bodyPr', nsmap)
bodyPr18a.set('rot', '2700000')

# Textbox with tab stops
tb18b = slide18.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(3))
tf18b = tb18b.text_frame
tf18b.word_wrap = True
p18b = tf18b.paragraphs[0]
p18b.text = "Col1\tCol2\tCol3"
p18b.font.size = Pt(18)
# Add tab stops via XML
pPr18b = p18b._pPr
if pPr18b is None:
    pPr18b = etree.SubElement(p18b._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p18b._p.insert(0, pPr18b)
tabLst = etree.SubElement(pPr18b, '{http://schemas.openxmlformats.org/drawingml/2006/main}tabLst')
etree.SubElement(tabLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}tab', attrib={'pos': '2743200', 'algn': 'l'})  # 3 inches
etree.SubElement(tabLst, '{http://schemas.openxmlformats.org/drawingml/2006/main}tab', attrib={'pos': '5486400', 'algn': 'r'})  # 6 inches

# ── Slide 19: Vertical text + text columns ──────────────────────────────────

slide19 = prs.slides.add_slide(blank)

# Textbox with vertical text (vert="vert")
tb19a = slide19.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(2), Inches(5))
tf19a = tb19a.text_frame
p19a = tf19a.paragraphs[0]
p19a.text = "Vertical text"
p19a.font.size = Pt(20)
bodyPr19a = tb19a._element.find('.//a:bodyPr', nsmap)
bodyPr19a.set('vert', 'vert')

# Textbox with eaVert
tb19b = slide19.shapes.add_textbox(Inches(3), Inches(0.5), Inches(2), Inches(5))
tf19b = tb19b.text_frame
p19b = tf19b.paragraphs[0]
p19b.text = "EA Vertical"
p19b.font.size = Pt(20)
bodyPr19b = tb19b._element.find('.//a:bodyPr', nsmap)
bodyPr19b.set('vert', 'eaVert')

# Textbox with text columns (numCol + spcCol)
tb19c = slide19.shapes.add_textbox(Inches(5.5), Inches(0.5), Inches(4), Inches(5))
tf19c = tb19c.text_frame
tf19c.word_wrap = True
p19c = tf19c.paragraphs[0]
p19c.text = "This is column text that spans multiple columns. The text should flow from the first column to the second column when it reaches the bottom."
p19c.font.size = Pt(14)
bodyPr19c = tb19c._element.find('.//a:bodyPr', nsmap)
bodyPr19c.set('numCol', '2')
bodyPr19c.set('spcCol', '457200')  # 0.5 inch spacing

# ── Slide 20: Hyperlink + RTL ────────────────────────────────────────────────

slide20 = prs.slides.add_slide(blank)

# Textbox with hyperlink
tb20a = slide20.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(2))
tf20a = tb20a.text_frame
tf20a.word_wrap = True
p20a = tf20a.paragraphs[0]
r20a = p20a.runs[0] if len(p20a.runs) > 0 else p20a.add_run()
r20a.text = "Click here to visit example.com"
r20a.font.size = Pt(20)
r20a.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
r20a.font.underline = True
# Add hlinkClick via XML — need to add a relationship first
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
rel = slide20.part.relate_to('https://example.com', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
rPr20a = r20a._r.find('a:rPr', nsmap)
if rPr20a is None:
    rPr20a = etree.SubElement(r20a._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    r20a._r.insert(0, rPr20a)
etree.SubElement(rPr20a, '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rel})

# RTL paragraph
tb20b = slide20.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(2))
tf20b = tb20b.text_frame
tf20b.word_wrap = True
p20b = tf20b.paragraphs[0]
p20b.text = "RTL paragraph text (right-to-left)"
p20b.font.size = Pt(20)
pPr20b = p20b._pPr
if pPr20b is None:
    pPr20b = etree.SubElement(p20b._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p20b._p.insert(0, pPr20b)
pPr20b.set('rtl', '1')

# ── Slide 21: Image bullet (a:buBlip) ───────────────────────────────────────

slide21 = prs.slides.add_slide(blank)

# For image bullets, we need a small embedded image
# We'll create a tiny 1x1 red PNG in memory
import struct, zlib, io
def make_tiny_png(r, g, b):
    """Create a minimal 1x1 PNG."""
    # IHDR
    ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
    ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff
    ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
    # IDAT
    raw = bytes([0, r, g, b])
    compressed = zlib.compress(raw)
    idat_crc = zlib.crc32(b'IDAT' + compressed) & 0xffffffff
    idat = struct.pack('>I', len(compressed)) + b'IDAT' + compressed + struct.pack('>I', idat_crc)
    # IEND
    iend_crc = zlib.crc32(b'IEND') & 0xffffffff
    iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
    return b'\x89PNG\r\n\x1a\n' + ihdr + idat + iend

png_data = make_tiny_png(255, 0, 0)  # red dot
png_stream = io.BytesIO(png_data)

# Add image as a relationship to the slide
from pptx.opc.package import Part
from pptx.opc.constants import CONTENT_TYPE as CT
image_part, rId_img = slide21.part.get_or_add_image_part(png_stream)

# Create textbox with bullet paragraphs
tb21 = slide21.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
tf21 = tb21.text_frame
tf21.word_wrap = True
p21 = tf21.paragraphs[0]
p21.text = "Image bullet paragraph 1"
p21.font.size = Pt(20)

# Add image bullet via XML: <a:buBlip><a:blip r:embed="rIdN"/></a:buBlip>
pPr21 = p21._pPr
if pPr21 is None:
    pPr21 = etree.SubElement(p21._p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    p21._p.insert(0, pPr21)
buBlip = etree.SubElement(pPr21, '{http://schemas.openxmlformats.org/drawingml/2006/main}buBlip')
etree.SubElement(buBlip, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed': rId_img})

# Add second paragraph with same image bullet
p21b_elem = etree.SubElement(tf21._txBody, '{http://schemas.openxmlformats.org/drawingml/2006/main}p')
pPr21b = etree.SubElement(p21b_elem, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
buBlip21b = etree.SubElement(pPr21b, '{http://schemas.openxmlformats.org/drawingml/2006/main}buBlip')
etree.SubElement(buBlip21b, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed': rId_img})
r21b = etree.SubElement(p21b_elem, '{http://schemas.openxmlformats.org/drawingml/2006/main}r')
rPr21b = etree.SubElement(r21b, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr', attrib={'lang': 'en-US', 'sz': '2000'})
t21b = etree.SubElement(r21b, '{http://schemas.openxmlformats.org/drawingml/2006/main}t')
t21b.text = "Image bullet paragraph 2"

# ── Slide 22: Hover link + link color ──────────────────────────────────────────
slide22 = prs.slides.add_slide(blank)
tb22 = slide22.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(5))
tf22 = tb22.text_frame
tf22.word_wrap = True

# Paragraph 1: hlinkClick (normal hyperlink)
p22a = tf22.paragraphs[0]
p22a.text = ""
r22a = p22a.add_run()
r22a.text = "Click link"
r22a.font.size = Pt(24)
# Add hlinkClick via XML
rPr22a = r22a._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
if rPr22a is None:
    rPr22a = etree.SubElement(r22a._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    r22a._r.insert(0, rPr22a)
# Add relationship for click link
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
rId_click = slide22.part.relate_to('https://example.com/click', RT.HYPERLINK, is_external=True)
etree.SubElement(rPr22a, '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rId_click})

# Paragraph 2: hlinkHover (mouse-over hyperlink)
p22b = tf22.add_paragraph()
r22b = p22b.add_run()
r22b.text = "Hover link"
r22b.font.size = Pt(24)
rPr22b = r22b._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
if rPr22b is None:
    rPr22b = etree.SubElement(r22b._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    r22b._r.insert(0, rPr22b)
rId_hover = slide22.part.relate_to('https://example.com/hover', RT.HYPERLINK, is_external=True)
etree.SubElement(rPr22b, '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkHover',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rId_hover})

# Paragraph 3: both hlinkClick + hlinkHover
p22c = tf22.add_paragraph()
r22c = p22c.add_run()
r22c.text = "Both links"
r22c.font.size = Pt(24)
rPr22c = r22c._r.find('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
if rPr22c is None:
    rPr22c = etree.SubElement(r22c._r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    r22c._r.insert(0, rPr22c)
rId_click2 = slide22.part.relate_to('https://example.com/both-click', RT.HYPERLINK, is_external=True)
rId_hover2 = slide22.part.relate_to('https://example.com/both-hover', RT.HYPERLINK, is_external=True)
etree.SubElement(rPr22c, '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rId_click2})
etree.SubElement(rPr22c, '{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkHover',
                 attrib={'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rId_hover2})

# ── Slide 23: Linear gradient fills ──────────────────────────────────────────

slide23 = prs.slides.add_slide(blank)

# Helper to inject gradFill XML into a shape's spPr (after prstGeom, replacing solidFill)
def set_gradient_fill(shape, grad_xml_str):
    """Replace any existing fill with a:gradFill in the shape's spPr."""
    ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    ns_p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    # spPr can be under a: or p: namespace depending on context
    spPr = shape._element.find(f'{ns_p}spPr')
    if spPr is None:
        spPr = shape._element.find(f'.//{ns_a}spPr')
    if spPr is None:
        print(f"WARNING: spPr not found on shape")
        return
    # Remove existing fills (always a: namespace children)
    for tag in ('solidFill', 'noFill', 'gradFill'):
        for el in spPr.findall(f'{ns_a}{tag}'):
            spPr.remove(el)
    # Insert after prstGeom (a: namespace child)
    prstGeom = spPr.find(f'{ns_a}prstGeom')
    grad_el = etree.fromstring(grad_xml_str)
    if prstGeom is not None:
        prstGeom.addnext(grad_el)
    else:
        spPr.append(grad_el)

# Shape 1: 3-stop linear gradient (left→right, ang=0)
s23a = slide23.shapes.add_shape(1, Inches(0.5), Inches(0.5), Inches(3), Inches(2))
s23a.text = "Linear 0deg"
s23a.text_frame.paragraphs[0].font.size = Pt(18)
s23a.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
set_gradient_fill(s23a, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
    <a:gs pos="50000"><a:srgbClr val="FFFF00"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
  </a:gsLst>
  <a:lin ang="0" scaled="1"/>
</a:gradFill>''')

# Shape 2: 2-stop linear gradient (top→bottom, ang=5400000 = 90°)
s23b = slide23.shapes.add_shape(1, Inches(4), Inches(0.5), Inches(3), Inches(2))
s23b.text = "Linear 90deg"
s23b.text_frame.paragraphs[0].font.size = Pt(18)
s23b.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
set_gradient_fill(s23b, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="003366"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="66CCFF"/></a:gs>
  </a:gsLst>
  <a:lin ang="5400000" scaled="1"/>
</a:gradFill>''')

# Shape 3: Diagonal linear gradient (45° = 2700000), rotWithShape="0"
s23c = slide23.shapes.add_shape(1, Inches(7.5), Inches(0.5), Inches(2), Inches(2))
s23c.text = "Linear 45deg"
s23c.text_frame.paragraphs[0].font.size = Pt(14)
set_gradient_fill(s23c, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="0">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="006600"/></a:gs>
  </a:gsLst>
  <a:lin ang="2700000" scaled="1"/>
</a:gradFill>''')


# ── Slide 24: Radial/path gradient fills + gradient on ellipse ───────────────

slide24 = prs.slides.add_slide(blank)

# Shape 1: Radial circle gradient
s24a = slide24.shapes.add_shape(1, Inches(0.5), Inches(0.5), Inches(3), Inches(2.5))
s24a.text = "Radial circle"
s24a.text_frame.paragraphs[0].font.size = Pt(18)
s24a.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
set_gradient_fill(s24a, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="FF6600"/></a:gs>
  </a:gsLst>
  <a:path path="circle">
    <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
  </a:path>
</a:gradFill>''')

# Shape 2: Radial rect gradient with off-center fillToRect
s24b = slide24.shapes.add_shape(1, Inches(4), Inches(0.5), Inches(3), Inches(2.5))
s24b.text = "Radial rect"
s24b.text_frame.paragraphs[0].font.size = Pt(18)
set_gradient_fill(s24b, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="FFFF00"/></a:gs>
    <a:gs pos="50000"><a:srgbClr val="FF0000"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="660000"/></a:gs>
  </a:gsLst>
  <a:path path="rect">
    <a:fillToRect l="25000" t="25000" r="75000" b="75000"/>
  </a:path>
</a:gradFill>''')

# Shape 3: Gradient on ellipse (oval)
from pptx.enum.shapes import MSO_SHAPE
s24c = slide24.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.5), Inches(0.5), Inches(2), Inches(2.5))
s24c.text = "Ellipse grad"
s24c.text_frame.paragraphs[0].font.size = Pt(14)
s24c.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
set_gradient_fill(s24c, '''<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="9933FF"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="330066"/></a:gs>
  </a:gsLst>
  <a:lin ang="16200000" scaled="1"/>
</a:gradFill>''')


# ── Slide 25: Gradient background ────────────────────────────────────────────

slide25 = prs.slides.add_slide(blank)

# Inject gradient background via XML
ns_p = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
cSld25 = slide25._element.find(f'{ns_p}cSld')
spTree25 = cSld25.find(f'{ns_p}spTree')
bg_xml = f'''<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:bgPr>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="1B2838"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="2A475E"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="66C0F4"/></a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="0"/>
    </a:gradFill>
    <a:effectLst/>
  </p:bgPr>
</p:bg>'''
bg_el = etree.fromstring(bg_xml)
cSld25.insert(list(cSld25).index(spTree25), bg_el)

# Text shape on gradient background (noFill)
s25a = slide25.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
s25a.text = "Gradient Background"
s25a.text_frame.paragraphs[0].font.size = Pt(36)
s25a.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
s25a.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Save
output_path = 'test_fixtures/test_features.pptx'
prs.save(output_path)
print(f"Saved {output_path} with {len(prs.slides)} slides")
