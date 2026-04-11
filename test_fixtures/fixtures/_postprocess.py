"""Post-save patching: OMML injection, media, ChartEx, etc.

Runs after 'prs.save(OUTPUT_PATH)'. Rewrites the ZIP in place to insert
content python-pptx cannot express directly.
"""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from lxml import etree
import base64, io, os, re, struct, tempfile, zipfile, zlib

from ._ctx import prs, OUTPUT_PATH
from .slides_09_chartex import cx_chart_slides, cx_chart_names

output_path = OUTPUT_PATH

# Patch slide78 XML to inject actual OMML
OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
etree.register_namespace('m', OMML_NS)

total_slides = len(prs.slides)  # 94 (88 + 6 cx charts)
# python-pptx numbers slide files sequentially
slide78_path = 'ppt/slides/slide79.xml'

# Read entire ZIP into memory, then close it before writing
zin = zipfile.ZipFile(output_path, 'r')
slide78_xml = zin.read(slide78_path).decode('utf-8')
all_entries = {}
for item in zin.infolist():
    all_entries[item.filename] = (item, zin.read(item.filename))
zin.close()

# Replace the placeholder paragraph with OMML
omml_xml = (
    '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:oMath>'
    '<m:r><m:t>x</m:t></m:r>'
    '<m:r><m:t>=</m:t></m:r>'
    '<m:f>'  # fraction
    '<m:num><m:r><m:t>-b±</m:t></m:r>'
    '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
    '<m:deg/><m:e><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e>'
    '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
    '<m:r><m:t>-4ac</m:t></m:r></m:e></m:rad></m:num>'
    '<m:den><m:r><m:t>2a</m:t></m:r></m:den>'
    '</m:f>'
    '</m:oMath>'
    '</m:oMathPara>'
)
slide78_xml = re.sub(
    r'<a:r>(?:<a:rPr[^/]*/>)?<a:t>MATH_PLACEHOLDER</a:t></a:r>',
    omml_xml,
    slide78_xml
)
# Add m: namespace to root if not present
if 'xmlns:m=' not in slide78_xml:
    slide78_xml = slide78_xml.replace(
        'xmlns:a=',
        f'xmlns:m="{OMML_NS}" xmlns:a='
    )

# ── Patch slide 79: inject transition + timing XML ──
slide79_path = 'ppt/slides/slide80.xml'
slide79_xml = all_entries[slide79_path][1].decode('utf-8') if isinstance(all_entries[slide79_path][1], bytes) else all_entries[slide79_path][1]
# Inject p:transition and p:timing before </p:sld>
transition_xml = '<p:transition spd="med"><p:fade/></p:transition>'
timing_xml = (
    '<p:timing>'
    '<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/></p:par></p:tnLst>'
    '</p:timing>'
)
slide79_xml = slide79_xml.replace('</p:sld>', transition_xml + timing_xml + '</p:sld>')

# ── Patch slide 80: inject show="0" for hidden slide ──
slide80_path = 'ppt/slides/slide81.xml'
slide80_xml = all_entries[slide80_path][1].decode('utf-8') if isinstance(all_entries[slide80_path][1], bytes) else all_entries[slide80_path][1]
# Add show="0" to p:sld element
slide80_xml = slide80_xml.replace('<p:sld ', '<p:sld show="0" ', 1)

# ── Patch slide 82: inject OMML nary (integral) ──
slide82_path = 'ppt/slides/slide83.xml'
slide82_xml = all_entries[slide82_path][1].decode('utf-8') if isinstance(all_entries[slide82_path][1], bytes) else all_entries[slide82_path][1]
nary_omml = (
    '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:oMath>'
    '<m:nary>'
    '<m:naryPr><m:chr m:val="\u222B"/><m:limLoc m:val="subSup"/></m:naryPr>'
    '<m:sub><m:r><m:t>0</m:t></m:r></m:sub>'
    '<m:sup><m:r><m:t>\u221E</m:t></m:r></m:sup>'
    '<m:e>'
    '<m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e>'
    '<m:sup><m:r><m:t>-x</m:t></m:r></m:sup></m:sSup>'
    '<m:r><m:t>dx</m:t></m:r>'
    '</m:e>'
    '</m:nary>'
    '<m:r><m:t>=</m:t></m:r>'
    '<m:r><m:t>1</m:t></m:r>'
    '</m:oMath>'
    '</m:oMathPara>'
)
slide82_xml = re.sub(
    r'<a:r>(?:<a:rPr[^/]*/>)?<a:t>NARY_PLACEHOLDER</a:t></a:r>',
    nary_omml, slide82_xml
)
if 'xmlns:m=' not in slide82_xml:
    slide82_xml = slide82_xml.replace('xmlns:a=', f'xmlns:m="{OMML_NS}" xmlns:a=')

# ── Patch slide 83: inject OMML matrix ──
slide83_path = 'ppt/slides/slide84.xml'
slide83_xml = all_entries[slide83_path][1].decode('utf-8') if isinstance(all_entries[slide83_path][1], bytes) else all_entries[slide83_path][1]
matrix_omml = (
    '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:oMath>'
    '<m:r><m:t>A=</m:t></m:r>'
    '<m:d>'
    '<m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
    '<m:e>'
    '<m:m>'
    '<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>'
    '<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr>'
    '</m:m>'
    '</m:e>'
    '</m:d>'
    '</m:oMath>'
    '</m:oMathPara>'
)
slide83_xml = re.sub(
    r'<a:r>(?:<a:rPr[^/]*/>)?<a:t>MATRIX_PLACEHOLDER</a:t></a:r>',
    matrix_omml, slide83_xml
)
if 'xmlns:m=' not in slide83_xml:
    slide83_xml = slide83_xml.replace('xmlns:a=', f'xmlns:m="{OMML_NS}" xmlns:a=')

# ── Patch slide 84: inject OMML accent + bar + subsup ──
slide84_path = 'ppt/slides/slide85.xml'
slide84_xml = all_entries[slide84_path][1].decode('utf-8') if isinstance(all_entries[slide84_path][1], bytes) else all_entries[slide84_path][1]
acc_bar_omml = (
    '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    '<m:oMath>'
    '<m:acc><m:accPr><m:chr m:val="\u0302"/></m:accPr>'
    '<m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>'
    '<m:r><m:t>=</m:t></m:r>'
    '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>'
    '<m:e><m:r><m:t>y</m:t></m:r></m:e></m:bar>'
    '<m:r><m:t>+</m:t></m:r>'
    '<m:sSubSup>'
    '<m:e><m:r><m:t>z</m:t></m:r></m:e>'
    '<m:sub><m:r><m:t>i</m:t></m:r></m:sub>'
    '<m:sup><m:r><m:t>n</m:t></m:r></m:sup>'
    '</m:sSubSup>'
    '</m:oMath>'
    '</m:oMathPara>'
)
slide84_xml = re.sub(
    r'<a:r>(?:<a:rPr[^/]*/>)?<a:t>ACC_BAR_PLACEHOLDER</a:t></a:r>',
    acc_bar_omml, slide84_xml
)
if 'xmlns:m=' not in slide84_xml:
    slide84_xml = slide84_xml.replace('xmlns:a=', f'xmlns:m="{OMML_NS}" xmlns:a=')

# ── Patch slide 85: inject blur effect shape ──
slide85_path = 'ppt/slides/slide86.xml'
slide85_xml = all_entries[slide85_path][1].decode('utf-8') if isinstance(all_entries[slide85_path][1], bytes) else all_entries[slide85_path][1]

# Add a shape with a:blur effect
blur_shape_xml = '''<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="401" name="BlurShape"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="1828800" y="1828800"/><a:ext cx="5486400" cy="2743200"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
    <a:effectLst>
      <a:blur rad="76200"/>
    </a:effectLst>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1"/><a:t>Blur Effect</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''

# Insert the blur shape into the slide's spTree
slide85_xml = slide85_xml.replace('</p:spTree>', blur_shape_xml + '\n</p:spTree>')

# ── Patch slide 86: inject preset shadow shapes ──
slide86_path = 'ppt/slides/slide87.xml'
slide86_xml = all_entries[slide86_path][1].decode('utf-8') if isinstance(all_entries[slide86_path][1], bytes) else all_entries[slide86_path][1]

# Add white background to slide 86
slide86_bg_xml = '''<p:bg>
  <p:bgPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:effectLst/>
  </p:bgPr>
</p:bg>'''
slide86_xml = slide86_xml.replace('<p:spTree>', slide86_bg_xml + '\n<p:spTree>')

prstshdw_shape1_xml = '''<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="501" name="PrstShdw1"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="1828800"/><a:ext cx="3657600" cy="2286000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
    <a:effectLst>
      <a:prstShdw prst="shdw1" dist="76200" dir="2700000">
        <a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr>
      </a:prstShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1">
      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    </a:rPr><a:t>shdw1 (Bottom Right)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''

prstshdw_shape2_xml = '''<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvSpPr>
    <p:cNvPr id="502" name="PrstShdw2"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="5029200" y="1828800"/><a:ext cx="3657600" cy="2286000"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
    <a:effectLst>
      <a:prstShdw prst="shdw2" dist="76200" dir="5400000">
        <a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr>
      </a:prstShdw>
    </a:effectLst>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1">
      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    </a:rPr><a:t>shdw2 (Bottom)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''

slide86_xml = slide86_xml.replace('</p:spTree>', prstshdw_shape1_xml + '\n' + prstshdw_shape2_xml + '\n</p:spTree>')

# ── Patch slide 87: inject fillOverlay effects ──
slide87_path = 'ppt/slides/slide88.xml'
slide87_xml = all_entries[slide87_path][1].decode('utf-8') if isinstance(all_entries[slide87_path][1], bytes) else all_entries[slide87_path][1]

# Inject fillOverlay into shapes using regex to find each spPr solidFill
blend_modes = ["over", "mult", "screen", "darken", "lighten"]
fo_count = [0]  # mutable counter for replacement callback
def inject_fill_overlay(m):
    idx = fo_count[0]
    fo_count[0] += 1
    if idx >= len(blend_modes):
        return m.group(0)  # skip title shape etc.
    blend = blend_modes[idx]
    fo_xml = (
        f'<a:effectLst>'
        f'<a:fillOverlay blend="{blend}">'
        f'<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>'
        f'</a:fillOverlay>'
        f'</a:effectLst>'
    )
    return m.group(0).replace('</p:spPr>', fo_xml + '</p:spPr>')
# Match each <p:spPr>...<a:solidFill>...</a:solidFill>...</p:spPr> block
slide87_xml = re.sub(
    r'<p:spPr>.*?<a:solidFill>.*?</a:solidFill>.*?</p:spPr>',
    inject_fill_overlay,
    slide87_xml,
    flags=re.DOTALL,
)

# ── Patch slide 81: inject WMF image + picture shape ──
slide81_path = 'ppt/slides/slide82.xml'
slide81_xml = all_entries[slide81_path][1].decode('utf-8') if isinstance(all_entries[slide81_path][1], bytes) else all_entries[slide81_path][1]

# Build a Placeable WMF binary with colored shapes (rectangle + ellipse + polygon)
def build_wmf_binary():
    """Generate a Placeable WMF with a blue rectangle, red ellipse, and green triangle."""
    records = bytearray()

    def add_record(func_num, params=b''):
        # WMF record: size (uint32 words) + function (uint16) + params
        param_words = len(params) // 2
        size = 3 + param_words  # 3 words for header
        records.extend(struct.pack('<IH', size, func_num) + params)

    # SetMapMode (MM_ANISOTROPIC = 8)
    add_record(0x0103, struct.pack('<h', 8))
    # SetWindowOrg (0, 0)
    add_record(0x020B, struct.pack('<hh', 0, 0))
    # SetWindowExt (1000, 800)
    add_record(0x020C, struct.pack('<hh', 800, 1000))

    # ── Blue filled rectangle ──
    # CreateBrushIndirect: style=0(solid), colorR=0x0000FF, hatch=0
    add_record(0x02FC, struct.pack('<h', 0) + struct.pack('<BBBx', 0, 0, 255) + struct.pack('<h', 0))
    # SelectObject (index 0)
    add_record(0x012D, struct.pack('<h', 0))
    # Rectangle (bottom=400, right=450, top=50, left=50) — Y,X order
    add_record(0x041B, struct.pack('<hhhh', 400, 450, 50, 50))

    # ── Red filled ellipse ──
    # CreateBrushIndirect: style=0(solid), color=0xFF0000, hatch=0
    add_record(0x02FC, struct.pack('<h', 0) + struct.pack('<BBBx', 255, 0, 0) + struct.pack('<h', 0))
    # SelectObject (index 1)
    add_record(0x012D, struct.pack('<h', 1))
    # Ellipse (bottom=400, right=950, top=50, left=550)
    add_record(0x0418, struct.pack('<hhhh', 400, 950, 50, 550))

    # ── Green filled triangle (Polygon) ──
    # CreateBrushIndirect: style=0(solid), color=0x00FF00, hatch=0
    add_record(0x02FC, struct.pack('<h', 0) + struct.pack('<BBBx', 0, 255, 0) + struct.pack('<h', 0))
    # SelectObject (index 2)
    add_record(0x012D, struct.pack('<h', 2))
    # Polygon: 3 points — (500,750), (250,500), (750,500) — X,Y order in params
    add_record(0x0324, struct.pack('<h', 3) + struct.pack('<hh', 500, 750) + struct.pack('<hh', 250, 500) + struct.pack('<hh', 750, 500))

    # DeleteObject (0, 1, 2)
    add_record(0x01F0, struct.pack('<h', 0))
    add_record(0x01F0, struct.pack('<h', 1))
    add_record(0x01F0, struct.pack('<h', 2))

    # EOF record
    records.extend(struct.pack('<IH', 3, 0x0000))

    # Build standard WMF header (18 bytes)
    file_type = 1  # memory metafile
    header_size = 9  # words
    version = 0x0300
    file_size = (18 + len(records)) // 2  # in words
    num_objects = 3
    max_record = max(struct.unpack_from('<I', records, i)[0] for i in range(0, len(records), 2) if i + 4 <= len(records) and i % 2 == 0)  # approximate
    # Recalculate max_record properly
    max_rec_size = 0
    pos = 0
    while pos < len(records):
        rec_size = struct.unpack_from('<I', records, pos)[0]
        if rec_size > max_rec_size:
            max_rec_size = rec_size
        pos += rec_size * 2
    std_header = struct.pack('<HHIHIHH', file_type, header_size, version, file_size, num_objects, max_rec_size, 0)

    # Placeable WMF header (22 bytes)
    magic = 0x9AC6CDD7
    hmf = 0
    bbox_left, bbox_top, bbox_right, bbox_bottom = 0, 0, 1000, 800
    inch = 96  # units per inch
    reserved = 0
    placeable = struct.pack('<IHhhhhHI', magic, hmf, bbox_left, bbox_top, bbox_right, bbox_bottom, inch, reserved)
    # Checksum: XOR of first 10 uint16 values
    chk = 0
    for i in range(0, 20, 2):
        chk ^= struct.unpack_from('<H', placeable, i)[0]
    placeable += struct.pack('<H', chk & 0xFFFF)

    return bytes(placeable) + bytes(std_header) + bytes(records)

wmf_data = build_wmf_binary()

# Replace WMF_PLACEHOLDER textbox with a p:pic shape
# Find and remove the placeholder <p:sp> containing "WMF_PLACEHOLDER"
slide81_xml = re.sub(
    r'<p:sp>.*?WMF_PLACEHOLDER.*?</p:sp>',
    '',
    slide81_xml,
    flags=re.DOTALL
)

# Insert p:pic shape referencing the WMF image before </p:spTree>
pic_xml = (
    '<p:pic>'
    '<p:nvPicPr>'
    '<p:cNvPr id="100" name="WMF Picture"/>'
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
    '<p:nvPr/>'
    '</p:nvPicPr>'
    '<p:blipFill>'
    '<a:blip r:embed="rWmf1"/>'
    '<a:stretch><a:fillRect/></a:stretch>'
    '</p:blipFill>'
    '<p:spPr>'
    '<a:xfrm>'
    '<a:off x="914400" y="2286000"/>'
    '<a:ext cx="4572000" cy="3657600"/>'
    '</a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '</p:spPr>'
    '</p:pic>'
)
slide81_xml = slide81_xml.replace('</p:spTree>', pic_xml + '</p:spTree>')

# Add relationship for the WMF image in slide81.xml.rels
slide81_rels_path = 'ppt/slides/_rels/slide82.xml.rels'
if slide81_rels_path in all_entries:
    slide81_rels = all_entries[slide81_rels_path][1].decode('utf-8')
else:
    slide81_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
wmf_rel = '<Relationship Id="rWmf1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image_wmf1.wmf"/>'
slide81_rels = slide81_rels.replace('</Relationships>', wmf_rel + '</Relationships>')

# Add [Content_Types].xml entry for .wmf
content_types_path = '[Content_Types].xml'
content_types_xml = all_entries[content_types_path][1].decode('utf-8')
if 'Extension="wmf"' not in content_types_xml:
    content_types_xml = content_types_xml.replace(
        '</Types>',
        '<Default Extension="wmf" ContentType="image/x-wmf"/></Types>'
    )

# ── Inject cx:chart (ChartEx) data into slides 89-94 ─────────────────────────
# cx:chartSpace XML templates for each chart type
cx_chart_xmls = {
    'waterfall': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="5">
      <cx:pt idx="0">Q1</cx:pt><cx:pt idx="1">Q2</cx:pt><cx:pt idx="2">Q3</cx:pt>
      <cx:pt idx="3">Q4</cx:pt><cx:pt idx="4">Total</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="5" formatCode="General">
      <cx:pt idx="0">100</cx:pt><cx:pt idx="1">-20</cx:pt><cx:pt idx="2">50</cx:pt>
      <cx:pt idx="3">-10</cx:pt><cx:pt idx="4">120</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="waterfall"><cx:dataId val="0"/>
      <cx:layoutPr><cx:subtotals><cx:idx val="4"/></cx:subtotals></cx:layoutPr>
    </cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>''',
    'treemap': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="5">
      <cx:pt idx="0">Product A</cx:pt><cx:pt idx="1">Product B</cx:pt>
      <cx:pt idx="2">Product C</cx:pt><cx:pt idx="3">Product D</cx:pt>
      <cx:pt idx="4">Product E</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="5">
      <cx:pt idx="0">50</cx:pt><cx:pt idx="1">30</cx:pt><cx:pt idx="2">20</cx:pt>
      <cx:pt idx="3">15</cx:pt><cx:pt idx="4">10</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="treemap"><cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>''',
    'sunburst': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="4">
      <cx:pt idx="0">North</cx:pt><cx:pt idx="1">South</cx:pt>
      <cx:pt idx="2">East</cx:pt><cx:pt idx="3">West</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="4">
      <cx:pt idx="0">40</cx:pt><cx:pt idx="1">30</cx:pt>
      <cx:pt idx="2">20</cx:pt><cx:pt idx="3">10</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="sunburst"><cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>''',
    'histogram': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="5">
      <cx:pt idx="0">0-20</cx:pt><cx:pt idx="1">20-40</cx:pt>
      <cx:pt idx="2">40-60</cx:pt><cx:pt idx="3">60-80</cx:pt>
      <cx:pt idx="4">80-100</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="5">
      <cx:pt idx="0">5</cx:pt><cx:pt idx="1">15</cx:pt>
      <cx:pt idx="2">25</cx:pt><cx:pt idx="3">12</cx:pt>
      <cx:pt idx="4">8</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>''',
    'boxWhisker': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="1">
      <cx:pt idx="0">Dataset</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="10">
      <cx:pt idx="0">12</cx:pt><cx:pt idx="1">15</cx:pt>
      <cx:pt idx="2">18</cx:pt><cx:pt idx="3">22</cx:pt>
      <cx:pt idx="4">25</cx:pt><cx:pt idx="5">28</cx:pt>
      <cx:pt idx="6">30</cx:pt><cx:pt idx="7">35</cx:pt>
      <cx:pt idx="8">40</cx:pt><cx:pt idx="9">45</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea></cx:chart>
</cx:chartSpace>''',
    'funnel': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cx:chartData><cx:data id="0">
    <cx:strDim type="cat"><cx:lvl ptCount="4">
      <cx:pt idx="0">Prospects</cx:pt><cx:pt idx="1">Qualified</cx:pt>
      <cx:pt idx="2">Proposals</cx:pt><cx:pt idx="3">Closed</cx:pt>
    </cx:lvl></cx:strDim>
    <cx:numDim type="val"><cx:lvl ptCount="4">
      <cx:pt idx="0">1000</cx:pt><cx:pt idx="1">600</cx:pt>
      <cx:pt idx="2">300</cx:pt><cx:pt idx="3">100</cx:pt>
    </cx:lvl></cx:numDim>
  </cx:data></cx:chartData>
  <cx:chart><cx:plotArea><cx:plotAreaRegion>
    <cx:series layoutId="funnel"><cx:dataId val="0"/></cx:series>
  </cx:plotAreaRegion></cx:plotArea>
    <cx:legend pos="r"/>
  </cx:chart>
</cx:chartSpace>''',
}

# Chart type keys in order matching slides 89-94
cx_types = ['waterfall', 'treemap', 'sunburst', 'histogram', 'boxWhisker', 'funnel']

# Patch each cx:chart slide XML to include a graphicFrame + create chartex files
cx_chartex_files = {}  # path -> xml content
cx_slide_patches = {}  # slide_path -> patched xml

for i, cx_type in enumerate(cx_types):
    slide_num = 89 + i
    # python-pptx slide files are 1-indexed but may start from different offset
    slide_file = f'ppt/slides/slide{slide_num + 1}.xml'  # 0-indexed in files
    chartex_file = f'ppt/charts/chartEx{i + 1}.xml'
    chartex_rid = f'rIdCx{i + 1}'

    # Store chartex XML
    cx_chartex_files[chartex_file] = cx_chart_xmls[cx_type]

    # Patch slide XML to inject graphicFrame
    if slide_file in all_entries:
        slide_xml = all_entries[slide_file][1].decode('utf-8')
        # Build graphicFrame XML
        gf_xml = (
            '<p:graphicFrame>'
            '<p:nvGraphicFramePr>'
            f'<p:cNvPr id="99{i}" name="ChartEx {i+1}"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            '<p:nvPr/>'
            '</p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="7315200" cy="4572000"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex">'
            f'<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
            f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            f' r:id="{chartex_rid}"/>'
            '</a:graphicData></a:graphic>'
            '</p:graphicFrame>'
        )
        # Insert before </p:spTree>
        slide_xml = slide_xml.replace('</p:spTree>', gf_xml + '</p:spTree>')
        cx_slide_patches[slide_file] = slide_xml

        # Patch slide rels to add chartex relationship
        rels_file = f'ppt/slides/_rels/slide{slide_num + 1}.xml.rels'
        if rels_file in all_entries:
            rels_xml = all_entries[rels_file][1].decode('utf-8')
        else:
            rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        cx_rel = f'<Relationship Id="{chartex_rid}" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx{i+1}.xml"/>'
        rels_xml = rels_xml.replace('</Relationships>', cx_rel + '</Relationships>')
        # Store patched rels
        cx_slide_patches[rels_file] = rels_xml

# Add chartex content type overrides
for i in range(len(cx_types)):
    chartex_override = f'<Override PartName="/ppt/charts/chartEx{i+1}.xml" ContentType="application/vnd.ms-office.chartex+xml"/>'
    if chartex_override not in content_types_xml:
        content_types_xml = content_types_xml.replace('</Types>', chartex_override + '</Types>')

# Rewrite ZIP with patched slides + WMF image
with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
    for fname, (item, data) in all_entries.items():
        if fname == slide78_path:
            zout.writestr(item, slide78_xml.encode('utf-8'))
        elif fname == slide79_path:
            zout.writestr(item, slide79_xml.encode('utf-8'))
        elif fname == slide80_path:
            zout.writestr(item, slide80_xml.encode('utf-8'))
        elif fname == slide81_path:
            zout.writestr(item, slide81_xml.encode('utf-8'))
        elif fname == slide81_rels_path:
            zout.writestr(item, slide81_rels.encode('utf-8'))
        elif fname == slide82_path:
            zout.writestr(item, slide82_xml.encode('utf-8'))
        elif fname == slide83_path:
            zout.writestr(item, slide83_xml.encode('utf-8'))
        elif fname == slide84_path:
            zout.writestr(item, slide84_xml.encode('utf-8'))
        elif fname == slide85_path:
            zout.writestr(item, slide85_xml.encode('utf-8'))
        elif fname == slide86_path:
            zout.writestr(item, slide86_xml.encode('utf-8'))
        elif fname == slide87_path:
            zout.writestr(item, slide87_xml.encode('utf-8'))
        elif fname == content_types_path:
            zout.writestr(item, content_types_xml.encode('utf-8'))
        elif fname in cx_slide_patches:
            zout.writestr(item, cx_slide_patches[fname].encode('utf-8'))
        else:
            zout.writestr(item, data)
    # Add WMF binary as new entry
    zout.writestr('ppt/media/image_wmf1.wmf', wmf_data)
    # Add ChartEx XML files
    for cx_path, cx_xml in cx_chartex_files.items():
        zout.writestr(cx_path, cx_xml.encode('utf-8'))
    # Add patched rels that are new (not in original ZIP)
    for rels_path, rels_xml in cx_slide_patches.items():
        if rels_path not in all_entries:
            zout.writestr(rels_path, rels_xml.encode('utf-8'))

print(f"Saved {output_path} with {total_slides} slides")
