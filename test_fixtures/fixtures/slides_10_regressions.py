"""Slide 96: regression fixtures.

Slides for behaviors that don't fit neatly into a category module — typically
small reproductions for bugs that have been fixed, kept so regressions can be
detected automatically.

Currently covered:
  - empty custGeom (issue #39): a custGeom with no path commands must NOT
    render as a colored rectangle.
  - p:style/a:fontRef text color: a shape whose <p:style> includes
    <a:fontRef> should apply that color to text runs that have no explicit
    color, matching PowerPoint's behavior (accent-styled rectangles get white
    text via fontRef -> lt1).
"""
from pptx.util import Inches
from lxml import etree

from ._ctx import prs, blank

ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"

# ── Slide 96: empty custGeom + fontRef text color ──────────────────────────
slide96 = prs.slides.add_slide(blank)

# Reference rect with <p:style>/<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>.
# python-pptx's add_shape() emits exactly this style block, so the only thing
# we need to do is add a text run with NO explicit color. PowerPoint renders
# the text in white (via fontRef -> lt1 -> theme background light).
ref = slide96.shapes.add_shape(1, Inches(0.5), Inches(0.5), Inches(2.5), Inches(0.6))
ref.text_frame.text = "Reference rect (prstGeom)"

# Empty custGeom (the regressed shape). Must NOT render as a colored block.
spTree95 = slide96.shapes._spTree
empty_custgeom_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9501" name="EmptyCustGeom"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="3000000" y="1500000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst><a:path w="508000" h="508000"/></a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="00B4D8"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>
"""
spTree95.append(etree.fromstring(empty_custgeom_xml))

# Valid custGeom — used as a control so we can verify normal shapes still
# render after the empty-custGeom fix.
valid_custgeom_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9502" name="ValidCustGeom"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="6000000" y="1500000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>
    <a:custGeom>
      <a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
      <a:rect l="0" t="0" r="0" b="0"/>
      <a:pathLst>
        <a:path w="100" h="100">
          <a:moveTo><a:pt x="50" y="0"/></a:moveTo>
          <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
          <a:lnTo><a:pt x="0" y="100"/></a:lnTo>
          <a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="06A77D"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
</p:sp>
"""
spTree95.append(etree.fromstring(valid_custgeom_xml))

# ── Slide 97: header/footer field placeholders (date / footer / slide number) ──
# Reproduces sample.pptx slide 20: a slide carrying date (<a:fld type="datetime1">),
# footer (ftr) and slide-number (<a:fld type="slidenum"> with a STALE cached "1")
# placeholders. The renderer must (a) show the ACTUAL slide number, not the cached
# "1", (b) fill the date field with the host-provided current date, and (c) keep
# the footer text.
slide97 = prs.slides.add_slide(blank)
spTree97 = slide97.shapes._spTree

footer_ph_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9701" name="Footer Placeholder"/><p:cNvSpPr/>
    <p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2438400" cy="365125"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>moon-pptx footer</a:t></a:r></a:p></p:txBody>
</p:sp>
"""
date_ph_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9702" name="Date Placeholder"/><p:cNvSpPr/>
    <p:nvPr><p:ph type="dt" idx="10"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2438400" cy="365125"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
    <a:fld id="{{1F3B9E54-0000-4000-8000-000000000099}}" type="datetime1"><a:rPr/></a:fld></a:p></p:txBody>
</p:sp>
"""
sldnum_ph_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9703" name="Slide Number Placeholder"/><p:cNvSpPr/>
    <p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2438400" cy="365125"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
    <a:fld id="{{1F3B9E54-0000-4000-8000-000000000099}}" type="slidenum"><a:rPr/><a:t>1</a:t></a:fld></a:p></p:txBody>
</p:sp>
"""
for xml in (footer_ph_xml, date_ph_xml, sldnum_ph_xml):
    spTree97.append(etree.fromstring(xml))

# ── Slide 98: underline styles + underline color (a:uFill) + double strike ────
# Underline color round-trip (R1): the underline value (u=) and a:uFill color are
# parsed, round-tripped (data-ooxml-underline-color) and re-serialized to <a:uFill>.
# (Decoration STYLE/COLOR are not drawn visually — browsers ignore
# text-decoration-style/-color on SVG text — but must survive export.)
slide98 = prs.slides.add_slide(blank)
spTree98 = slide98.shapes._spTree

# Each run: u="<variant>" with a red a:uFill underline color; one dblStrike run.
deco_runs = (
    '<a:r><a:rPr lang="en-US" sz="2000" u="dbl"><a:uFill><a:solidFill>'
    '<a:srgbClr val="FF0000"/></a:solidFill></a:uFill></a:rPr><a:t>double </a:t></a:r>'
    '<a:r><a:rPr lang="en-US" sz="2000" u="wavy"><a:uFill><a:solidFill>'
    '<a:srgbClr val="00B050"/></a:solidFill></a:uFill></a:rPr><a:t>wavy </a:t></a:r>'
    '<a:r><a:rPr lang="en-US" sz="2000" u="dotted"/><a:t>dotted </a:t></a:r>'
    '<a:r><a:rPr lang="en-US" sz="2000" strike="dblStrike"/><a:t>dblstrike</a:t></a:r>'
)
deco_sp_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="9801" name="DecoText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="8229600" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p>{deco_runs}</a:p></p:txBody>
</p:sp>
"""
spTree98.append(etree.fromstring(deco_sp_xml))

# ── Slide 99: flipH/flipV mirroring (issue #55) ───────────────────────────────
# Regression: the renderer emitted rotate() but dropped flipH/flipV, so shapes
# and their (mirrored) text rendered un-flipped vs PowerPoint. Each text box
# below sets a flip on <a:xfrm>; the SVG must carry a negative-scale transform.
slide99 = prs.slides.add_slide(blank)
spTree99 = slide99.shapes._spTree

flip_variants = (
    ("FlipH", 'flipH="1"', 457200),
    ("FlipV", 'flipV="1"', 3886200),
    ("FlipHV", 'flipH="1" flipV="1"', 7315200),
)
for idx, (name, flip_attr, x) in enumerate(flip_variants):
    sp_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="{9900 + idx}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm {flip_attr}><a:off x="{x}" y="2743200"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>{name}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
"""
    spTree99.append(etree.fromstring(sp_xml))

# ── Slide 100: preset geometries from the ECMA-376 definitions (issue #62) ────
# Regression: these presets drew blank (a full-circle arcTo was emitted as one
# SVG arc with coincident endpoints) or distorted (hand-approximated guide
# lists, reversed sin/cos operands, missing hc/vc and +/). The wedgeRectCallout
# variants cover `?:` edge selection: each tip must attach to its nearest edge.
slide100 = prs.slides.add_slide(blank)
spTree100 = slide100.shapes._spTree

issue62_presets = (
    "rect", "donut", "cloudCallout", "star12",
    "circularArrow", "uturnArrow", "quadArrow", "lightningBolt",
    "sun", "moon", "cube", "flowChartDocument",
)
issue62_wedges = (
    ("W_below", 30000, 80000),
    ("W_above", 30000, -80000),
    ("W_left", -80000, 0),
    ("W_right", 80000, 10000),
)


def issue62_sp(sp_id, name, prst, x, y, cx, cy, avlst=""):
    return f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="{prst}"><a:avLst>{avlst}</a:avLst></a:prstGeom>
    <a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>
  </p:spPr>
</p:sp>
"""


for idx, prst in enumerate(issue62_presets):
    x = 228600 + (idx % 6) * 1447800
    y = 228600 + (idx // 6) * 1600200
    spTree100.append(etree.fromstring(
        issue62_sp(10000 + idx, f"P_{prst}", prst, x, y, 1188720, 868680)))
for idx, (name, adj1, adj2) in enumerate(issue62_wedges):
    av = (f'<a:gd name="adj1" fmla="val {adj1}"/>'
          f'<a:gd name="adj2" fmla="val {adj2}"/>')
    spTree100.append(etree.fromstring(issue62_sp(
        10100 + idx, name, "wedgeRectCallout",
        457200 + idx * 2133600, 4114800, 1371600, 914400, av)))

# ── Slide 101: presets generated from presetShapeDefinitions.xml ─────────────
# Each preset is drawn from its ECMA-376 definition, including per-path fill
# modes: arc (filled wedge + stroke-only arc), cube (shaded faces), chartPlus
# (lines drawn over the box), bentConnector3 (stroke-only), actionButtonHome
# (shaded icon), and star7 (absent from the old hand-written table).
slide101 = prs.slides.add_slide(blank)
spTree101 = slide101.shapes._spTree
generated_presets = ("arc", "cube", "chartPlus", "bentConnector3", "actionButtonHome", "star7")
for idx, prst in enumerate(generated_presets):
    sp_xml = issue62_sp(10200 + idx, f"G_{prst}", prst,
                        228600 + idx * 1447800, 1371600, 1188720, 868680)
    sp = etree.fromstring(sp_xml)
    # A light outline so stroke-only paths stand out on the dark background.
    sp.find(f"{{{ns_p}}}spPr").append(etree.fromstring(
        f'<a:ln xmlns:a="{ns_a}" w="19050"><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></a:ln>'))
    spTree101.append(sp)

# ── Slide 102: custGeom per-path fill / stroke / coordinate space ─────────────
# A custom geometry with three a:path elements: a filled, unstroked square
# (w/h 100), a stroke-only diagonal in its own 50×50 space, and a darkened
# inner square. Each path's attributes must render and survive export.
slide102 = prs.slides.add_slide(blank)
cust_sp_xml = f"""
<p:sp xmlns:p="{ns_p}" xmlns:a="{ns_a}">
  <p:nvSpPr><p:cNvPr id="10300" name="CustStyled"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="2286000" y="1371600"/><a:ext cx="2743200" cy="2743200"/></a:xfrm>
    <a:custGeom><a:avLst/><a:gdLst/><a:rect l="l" t="t" r="r" b="b"/>
      <a:pathLst>
        <a:path w="100" h="100" stroke="0">
          <a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo>
          <a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/>
        </a:path>
        <a:path w="50" h="50" fill="none">
          <a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo>
        </a:path>
        <a:path w="100" h="100" fill="darken" stroke="0">
          <a:moveTo><a:pt x="25" y="25"/></a:moveTo><a:lnTo><a:pt x="50" y="25"/></a:lnTo>
          <a:lnTo><a:pt x="50" y="50"/></a:lnTo><a:lnTo><a:pt x="25" y="50"/></a:lnTo><a:close/>
        </a:path>
      </a:pathLst>
    </a:custGeom>
    <a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>
    <a:ln w="28575"><a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></a:ln>
  </p:spPr>
</p:sp>
"""
slide102.shapes._spTree.append(etree.fromstring(cust_sp_xml))

# ── Slides 103-105: slide-own gradient / pattern / image backgrounds ──────────
# Each slide carries its own non-solid <p:bg>; it must render and survive the
# SVG round-trip (data-ooxml-bg-grad-* / -bg-patt-* / -bg-blip-*) on export.
import base64
import io

ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def set_slide_bg(slide, fill_xml):
    """Insert <p:bg><p:bgPr>fill<a:effectLst/></p:bgPr></p:bg> into p:cSld."""
    cSld = slide._element.find(f"{{{ns_p}}}cSld")
    bg = etree.fromstring(
        f'<p:bg xmlns:p="{ns_p}" xmlns:a="{ns_a}" xmlns:r="{ns_r}"><p:bgPr>'
        f'{fill_xml}<a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)


slide103 = prs.slides.add_slide(blank)
set_slide_bg(slide103,
             '<a:gradFill rotWithShape="1"><a:gsLst>'
             '<a:gs pos="0"><a:srgbClr val="1F4E79"/></a:gs>'
             '<a:gs pos="100000"><a:srgbClr val="9DC3E6"/></a:gs>'
             '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>')

slide104 = prs.slides.add_slide(blank)
set_slide_bg(slide104,
             '<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr>'
             '<a:bgClr><a:srgbClr val="FFF2CC"/></a:bgClr></a:pattFill>')

slide105 = prs.slides.add_slide(blank)
_bg_png = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==')
_, bg_rid = slide105.part.get_or_add_image_part(io.BytesIO(_bg_png))
set_slide_bg(slide105,
             f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{bg_rid}"/>'
             '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>')

# ── Slide 106: shape ids, names and alt text survive editing ──────────────────
# A shape with descr/title, a picture with descr, and an appear animation that
# targets the shape by its p:cNvPr id (p:spTgt spid). Editing the slide must keep
# the ids (or the animation loses its target), names and alt text.


slide106 = prs.slides.add_slide(blank)
alt_box = slide106.shapes.add_shape(1, Inches(1), Inches(1), Inches(3), Inches(1.5))
alt_box.name = "Revenue Box"
alt_box.text_frame.text = "Revenue"
cnv = alt_box._element.find(f".//{{{ns_p}}}cNvPr")
cnv.set("descr", "Quarterly revenue, up 12% & rising")
cnv.set("title", "Revenue")

alt_pic = slide106.shapes.add_picture(io.BytesIO(_bg_png), Inches(5), Inches(1), Inches(2), Inches(2))
alt_pic.name = "Logo"
alt_pic._element.find(f".//{{{ns_p}}}cNvPr").set("descr", "Company logo")

_spid = alt_box.shape_id
slide106._element.append(etree.fromstring(
    f'<p:timing xmlns:p="{ns_p}"><p:tnLst><p:par>'
    '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
    '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
    '<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
    '</p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst>'
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="1" '
    'presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst>'
    '<p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" '
    'fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl>'
    f'<p:spTgt spid="{_spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility'
    '</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
    '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
    '</p:par></p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'))

# ── Slide 107: a hidden shape (p:cNvPr hidden="1") is kept but not drawn ──────
slide107 = prs.slides.add_slide(blank)
shown = slide107.shapes.add_shape(1, Inches(1), Inches(1), Inches(3), Inches(1.5))
shown.name = "Shown"
shown.text_frame.text = "Visible"
hidden = slide107.shapes.add_shape(1, Inches(5), Inches(1), Inches(3), Inches(1.5))
hidden.name = "Hidden"
hidden.text_frame.text = "Hidden text"
hidden._element.find(f".//{{{ns_p}}}cNvPr").set("hidden", "1")
