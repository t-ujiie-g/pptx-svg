// ── Color constructors ──────────────────────────────────────────────────────

///|
test "Color::white" {
  let c = Color::white()
  assert_eq(c.r, 255)
  assert_eq(c.g, 255)
  assert_eq(c.b, 255)
  assert_eq(c.alpha, 255)
  assert_false(c.is_none())
}

///|
test "Color::black" {
  let c = Color::black()
  assert_eq(c.r, 0)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

///|
test "Color::none" {
  let c = Color::none()
  assert_true(c.is_none())
}

// ── parse_hex_color ─────────────────────────────────────────────────────────

///|
test "parse_hex_color 6-char" {
  let c = parse_hex_color("FF8800")
  assert_eq(c.r, 255)
  assert_eq(c.g, 136)
  assert_eq(c.b, 0)
  assert_eq(c.alpha, 255)
}

///|
test "parse_hex_color 8-char with alpha" {
  let c = parse_hex_color("80FF0000")
  assert_eq(c.alpha, 128)
  assert_eq(c.r, 255)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

///|
test "parse_hex_color short string" {
  // Falls back to black
  let c = parse_hex_color("FFF")
  assert_eq(c.r, 0)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

// ── Color::to_hex / to_hex_rgb / to_css ─────────────────────────────────────

///|
test "Color::to_hex roundtrip" {
  let c = parse_hex_color("4488CC")
  assert_eq(c.to_hex(), "4488CC")
  assert_eq(c.to_hex_rgb(), "4488CC")
}

///|
test "Color::to_hex with alpha" {
  let c = parse_hex_color("80FF0000")
  assert_eq(c.to_hex(), "80FF0000")
  assert_eq(c.to_hex_rgb(), "FF0000")
}

///|
test "Color::to_hex none" {
  assert_eq(Color::none().to_hex(), "none")
  assert_eq(Color::none().to_hex_rgb(), "none")
}

///|
test "Color::to_css opaque" {
  let c = parse_hex_color("FF0000")
  assert_eq(c.to_css(), "rgb(255,0,0)")
}

///|
test "Color::to_css with alpha" {
  let c = parse_hex_color("80FF0000")
  let css = c.to_css()
  assert_true(@xml.str_starts_with(css, "rgba(255,0,0,"))
}

///|
test "Color::to_css none" {
  assert_eq(Color::none().to_css(), "none")
}

// ── alpha_to_css_str ────────────────────────────────────────────────────────

///|
test "alpha_to_css_str" {
  assert_eq(alpha_to_css_str(255), "1")
  assert_eq(alpha_to_css_str(0), "0")
  assert_eq(alpha_to_css_str(128), "0.50")
}

// ── resolve_scheme_color ────────────────────────────────────────────────────

///|
test "resolve_scheme_color basic" {
  let theme = ThemeData::default()
  let dk1 = resolve_scheme_color("dk1", theme)
  assert_false(dk1.is_none())
  // bg1 → lt1 alias
  let bg1 = resolve_scheme_color("bg1", theme)
  let lt1 = resolve_scheme_color("lt1", theme)
  assert_eq(bg1.r, lt1.r)
  assert_eq(bg1.g, lt1.g)
  assert_eq(bg1.b, lt1.b)
  // tx1 → dk1 alias
  let tx1 = resolve_scheme_color("tx1", theme)
  assert_eq(tx1.r, dk1.r)
}

// ── apply_color_modifiers ───────────────────────────────────────────────────

///|
test "apply_color_modifiers empty" {
  let c = parse_hex_color("FF0000")
  let result = apply_color_modifiers(c, [])
  assert_eq(result.r, 255)
  assert_eq(result.g, 0)
  assert_eq(result.b, 0)
}

///|
test "apply_color_modifiers alpha" {
  let c = parse_hex_color("FF0000")
  let nodes = @xml.parse_xml("<a:alpha val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // alpha 50000/100000 * 255 ≈ 127-128
  assert_true(result.alpha < 130)
  assert_true(result.alpha > 125)
}

///|
test "apply_color_modifiers shade" {
  let c = Color::white()
  let nodes = @xml.parse_xml("<a:shade val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // 255 * 50000/100000 = 127
  assert_eq(result.r, 127)
  assert_eq(result.g, 127)
  assert_eq(result.b, 127)
}

///|
test "apply_color_modifiers tint" {
  let c = Color::black()
  let nodes = @xml.parse_xml("<a:tint val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // 255 - (255-0)*50000/100000 = 128 (integer rounding)
  assert_eq(result.r, 128)
  assert_eq(result.g, 128)
  assert_eq(result.b, 128)
}

///|
/// Regression: a:hueOff val is in 60000ths of a degree. The conversion to the
/// internal hue scale (tenths of a degree) must divide by 6000 — not by 60.
/// The previous formula (val / 600 * 10) over-shifted hue by 100x.
test "apply_color_modifiers hueOff is in correct degrees" {
  // Pure red (h=0). Apply hueOff val=60000 (= 1° offset). Result should be
  // *almost* red, not yellow/green. With the old buggy formula val/600*10
  // = 1000, the hue would jump to 100°, turning red into green.
  let c = parse_hex_color("FF0000")
  let nodes = @xml.parse_xml("<a:hueOff val=\"60000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // After 1° hue offset on red, R should still dominate by a large margin.
  assert_true(result.r > 250)
  assert_true(result.g < 10)
  assert_true(result.b < 10)
}

// ── parse_geom / to_prst ────────────────────────────────────────────────────

///|
test "parse_geom basic shapes" {
  assert_eq(parse_geom("rect", []).to_prst(), "rect")
  assert_eq(parse_geom("ellipse", []).to_prst(), "ellipse")
  assert_eq(parse_geom("roundRect", []).to_prst(), "roundRect")
  assert_eq(parse_geom("line", []).to_prst(), "line")
  assert_eq(parse_geom("straightConnector1", []).to_prst(), "line")
}

///|
test "parse_geom other" {
  assert_eq(parse_geom("star5", []).to_prst(), "star5")
}

///|
test "parse_geom roundRect adj" {
  // Empty avlst → default 16667
  match parse_geom("roundRect", []) {
    ShapeGeom::RoundRect(adj) => assert_eq(adj, 16667)
    _ => fail("expected RoundRect")
  }
  // Author-specified adj preserved
  match parse_geom("roundRect", [50000]) {
    ShapeGeom::RoundRect(adj) => assert_eq(adj, 50000)
    _ => fail("expected RoundRect")
  }
}

///|
test "parse_connector_geom" {
  assert_eq(parse_connector_geom("straightConnector1", []).to_prst(), "line")
  assert_eq(
    parse_connector_geom("bentConnector3", [50000]).to_prst(),
    "bentConnector3",
  )
}

// ── parse_rels ──────────────────────────────────────────────────────────────

///|
test "parse_rels basic" {
  let xml = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/></Relationships>"
  let rels = parse_rels(xml)
  assert_eq(rels.length(), 2)
  assert_eq(find_rel_target(rels, "rId1"), "../slideLayouts/slideLayout1.xml")
  assert_eq(find_rel_target(rels, "rId2"), "../media/image1.png")
  assert_eq(find_rel_target(rels, "rId99"), "")
}

// ── parse_slide_size ────────────────────────────────────────────────────────

///|
test "parse_slide_size standard" {
  let xml = "<p:presentation><p:sldSz cx=\"9144000\" cy=\"6858000\"/></p:presentation>"
  let size = parse_slide_size(xml)
  assert_eq(size.cx, 9144000)
  assert_eq(size.cy, 6858000)
}

///|
test "parse_slide_size widescreen" {
  let xml = "<p:presentation><p:sldSz cx=\"12192000\" cy=\"6858000\"/></p:presentation>"
  let size = parse_slide_size(xml)
  assert_eq(size.cx, 12192000)
}

// ── EffectList / Shadow / Glow ──────────────────────────────────────────────

///|
test "EffectList::none" {
  let e = EffectList::none()
  assert_true(e.is_none())
  assert_true(e.outer_shadow.is_none())
  assert_true(e.inner_shadow.is_none())
  assert_true(e.glow.is_none())
  assert_true(e.soft_edge.is_none())
  assert_true(e.reflection.is_none())
  assert_true(e.blur.is_none())
}

///|
test "OuterShadow::none and is_none" {
  let os = OuterShadow::none()
  assert_true(os.is_none())
  assert_eq(os.blur_rad, 0)
  assert_eq(os.dist, 0)
  assert_true(os.color.is_none())
  let os2 : OuterShadow = {
    blur_rad: 152400,
    dist: 114300,
    dir: 5400000,
    color: parse_hex_color("000000"),
    sx: 0,
    sy: 0,
    algn: "ctr",
    rot_with_shape: false,
  }
  assert_true(!os2.is_none())
  assert_eq(os2.blur_rad, 152400)
  assert_eq(os2.dir, 5400000)
}

///|
test "InnerShadow::none and is_none" {
  let ishdw = InnerShadow::none()
  assert_true(ishdw.is_none())
  let ishdw2 : InnerShadow = {
    blur_rad: 114300,
    dist: 76200,
    dir: 2700000,
    color: parse_hex_color("000000"),
  }
  assert_true(!ishdw2.is_none())
}

///|
test "Glow::none and is_none" {
  let gl = Glow::none()
  assert_true(gl.is_none())
  assert_eq(gl.rad, 0)
  assert_true(gl.color.is_none())
  let gl2 : Glow = { rad: 228600, color: parse_hex_color("FFC000") }
  assert_true(!gl2.is_none())
  assert_eq(gl2.rad, 228600)
}

///|
test "SoftEdge::none and is_none" {
  let se = SoftEdge::none()
  assert_true(se.is_none())
  assert_eq(se.rad, 0)
  let se2 : SoftEdge = { rad: 63500 }
  assert_true(!se2.is_none())
}

///|
test "Reflection::none and is_none" {
  let r = Reflection::none()
  assert_true(r.is_none())
  assert_eq(r.st_alpha, 0)
  assert_eq(r.end_alpha, 0)
  let r2 : Reflection = {
    blur_rad: 6350,
    dist: 0,
    dir: 5400000,
    st_alpha: 50000,
    end_alpha: 300,
    fade_dir: 5400000,
    sx: 100000,
    sy: -100000,
    algn: "bl",
    rot_with_shape: false,
  }
  assert_true(!r2.is_none())
}

///|
test "Blur::none and is_none" {
  let b = Blur::none()
  assert_true(b.is_none())
  assert_eq(b.rad, 0)
  let b2 : Blur = { rad: 50800 }
  assert_true(!b2.is_none())
  assert_eq(b2.rad, 50800)
}

///|
test "parse_slide captures endParaRPr into paragraph epr_* fields" {
  // endParaRPr color/size/font should be stored on the paragraph as
  // last-resort fallbacks, NOT pre-merged into runs. A run with no explicit
  // color must still parse with Color::none() so layout/master inheritance
  // can run first; the epr_* values are only applied as the final fallback
  // in main_inherit.mbt::apply_epr_fallbacks().
  let slide_xml =
    #|<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    #|<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    #|  <p:cSld><p:spTree>
    #|    <p:sp>
    #|      <p:nvSpPr><p:cNvPr id="1" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    #|      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    #|      <p:txBody>
    #|        <a:bodyPr/><a:lstStyle/>
    #|        <a:p>
    #|          <a:r><a:rPr lang="ja-JP"/><a:t>0.028</a:t></a:r>
    #|          <a:endParaRPr lang="ja-JP" sz="1800"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/></a:endParaRPr>
    #|        </a:p>
    #|        <a:p>
    #|          <a:r><a:rPr lang="ja-JP"/><a:t>no epr</a:t></a:r>
    #|        </a:p>
    #|      </p:txBody>
    #|    </p:sp>
    #|  </p:spTree></p:cSld>
    #|</p:sld>
  let slide = parse_slide(slide_xml, SlideSize::default(), ThemeData::default())
  let paras = slide.shapes[0].paragraphs
  assert_eq(paras.length(), 2)
  // Paragraph 0 has endParaRPr — epr_* fields should be populated.
  assert_eq(paras[0].epr_font_size, 1800)
  assert_eq(paras[0].epr_color.r, 255)
  assert_eq(paras[0].epr_color.g, 255)
  assert_eq(paras[0].epr_color.b, 255)
  assert_eq(paras[0].epr_font_face, "Arial")
  // The run itself must still have no explicit color/size — those are only
  // applied later in apply_epr_fallbacks, after layout/master inheritance.
  assert_true(paras[0].runs[0].color.is_none())
  assert_eq(paras[0].runs[0].font_size, 0)
  // Paragraph 1 has no endParaRPr — epr_* stays empty.
  assert_eq(paras[1].epr_font_size, 0)
  assert_true(paras[1].epr_color.is_none())
  assert_eq(paras[1].epr_font_face, "")
}

///|
test "BlipFill::is_none considers svg_rid" {
  // Fully empty blip — is_none
  let bf0 = BlipFill::none()
  assert_true(bf0.is_none())
  // Raster-only (r:embed) — not none
  let bf_raster : BlipFill = { ..BlipFill::none(), rid: "rId7" }
  assert_false(bf_raster.is_none())
  // SVG-only (asvg:svgBlip, no r:embed) — not none. This is the regression
  // fixed for cover-slide SVG logos: previously is_none() returned true and
  // the picture shape rendered as a placeholder rectangle.
  let bf_svg : BlipFill = { ..BlipFill::none(), svg_rid: "rId8" }
  assert_false(bf_svg.is_none())
  // Both set — not none
  let bf_both : BlipFill = { ..BlipFill::none(), rid: "rId7", svg_rid: "rId8" }
  assert_false(bf_both.is_none())
}

///|
test "EffectList with single effect is not none" {
  // Only outer shadow set
  let e1 : EffectList = {
    ..EffectList::none(),
    outer_shadow: {
      blur_rad: 100,
      dist: 0,
      dir: 0,
      color: parse_hex_color("000000"),
      sx: 0,
      sy: 0,
      algn: "",
      rot_with_shape: true,
    },
  }
  assert_true(!e1.is_none())
  // Only glow set
  let e2 : EffectList = {
    ..EffectList::none(),
    glow: { rad: 100, color: parse_hex_color("FF0000") },
  }
  assert_true(!e2.is_none())
  // Only soft edge set
  let e3 : EffectList = { ..EffectList::none(), soft_edge: { rad: 100 } }
  assert_true(!e3.is_none())
}

///|
test "parse_theme captures fmtScheme fill/ln style lists" {
  // Minimal theme with one fillStyleLst entry (gradient using phClr) and one
  // lnStyleLst entry. Both should survive into ThemeData as raw XML with the
  // `phClr` placeholder intact, so that resolve_fill_ref / resolve_ln_ref can
  // later substitute the referenced scheme color.
  let theme_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
    #|  <a:themeElements>
    #|    <a:clrScheme name="Test">
    #|      <a:dk1><a:srgbClr val="000000"/></a:dk1>
    #|      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    #|      <a:dk2><a:srgbClr val="112233"/></a:dk2>
    #|      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    #|      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    #|      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    #|      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
    #|      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
    #|      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
    #|      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    #|      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    #|      <a:folHlink><a:srgbClr val="954F63"/></a:folHlink>
    #|    </a:clrScheme>
    #|    <a:fontScheme name="Test">
    #|      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
    #|      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    #|    </a:fontScheme>
    #|    <a:fmtScheme name="Test">
    #|      <a:fillStyleLst>
    #|        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    #|        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
    #|      </a:fillStyleLst>
    #|      <a:lnStyleLst>
    #|        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    #|      </a:lnStyleLst>
    #|      <a:effectStyleLst/>
    #|      <a:bgFillStyleLst/>
    #|    </a:fmtScheme>
    #|  </a:themeElements>
    #|</a:theme>
  let theme = parse_theme(theme_xml)
  // Two fill styles captured, one ln style captured.
  assert_eq(theme.fill_style_xmls.length(), 2)
  assert_eq(theme.ln_style_xmls.length(), 1)
  // phClr placeholder must still be present (substitution happens at resolve time).
  assert_true(@xml.str_starts_with(theme.fill_style_xmls[0], "<a:solidFill"))
  assert_true(@xml.str_starts_with(theme.fill_style_xmls[1], "<a:gradFill"))
  assert_true(@xml.str_starts_with(theme.ln_style_xmls[0], "<a:ln"))
}

///|
test "parse_slide resolves p:style/a:fillRef via theme fmtScheme" {
  // Shape has no explicit fill/line in spPr, only a <p:style>/<a:fillRef idx=2>
  // pointing at the theme's fmtScheme. The gradient fill should be substituted
  // in and applied to the shape.
  let theme_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
    #|  <a:themeElements>
    #|    <a:clrScheme name="Test">
    #|      <a:dk1><a:srgbClr val="000000"/></a:dk1>
    #|      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    #|      <a:dk2><a:srgbClr val="112233"/></a:dk2>
    #|      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    #|      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
    #|      <a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    #|      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
    #|      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
    #|      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
    #|      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    #|      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    #|      <a:folHlink><a:srgbClr val="954F63"/></a:folHlink>
    #|    </a:clrScheme>
    #|    <a:fontScheme name="Test"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    #|    <a:fmtScheme name="Test">
    #|      <a:fillStyleLst>
    #|        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    #|        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
    #|      </a:fillStyleLst>
    #|      <a:lnStyleLst>
    #|        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    #|      </a:lnStyleLst>
    #|      <a:effectStyleLst/>
    #|      <a:bgFillStyleLst/>
    #|    </a:fmtScheme>
    #|  </a:themeElements>
    #|</a:theme>
  let theme = parse_theme(theme_xml)
  let slide_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    #|  <p:cSld><p:spTree>
    #|    <p:sp>
    #|      <p:nvSpPr><p:cNvPr id="1" name="r"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    #|      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    #|      <p:style>
    #|        <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
    #|        <a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>
    #|        <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    #|        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
    #|      </p:style>
    #|      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    #|    </p:sp>
    #|  </p:spTree></p:cSld>
    #|</p:sld>
  let slide = parse_slide(slide_xml, SlideSize::default(), theme)
  let shape = slide.shapes[0]
  // fillRef idx=2 → gradFill with phClr=accent1. Two gradient stops expected.
  assert_false(shape.grad_fill.is_none())
  assert_eq(shape.grad_fill.stops.length(), 2)
  // accent1 = FF0000 in our test theme.
  assert_eq(shape.grad_fill.stops[0].color.r, 255)
  assert_eq(shape.grad_fill.stops[0].color.g, 0)
  assert_eq(shape.grad_fill.stops[0].color.b, 0)
  // solid fill must remain unset when a gradFill is applied.
  assert_true(shape.fill.is_none())
  // lnRef idx=1 → 19050 EMU stroke using accent1.
  assert_eq(shape.stroke_w, 19050)
  assert_eq(shape.stroke.r, 255)
  assert_eq(shape.stroke.g, 0)
  assert_eq(shape.stroke.b, 0)
}

///|
test "parse_slide captures lstStyle buAutoNum into level defaults" {
  // a:lstStyle/a:lvl1pPr/a:buAutoNum type="arabicPeriod" should populate
  // Shape.lst_style.levels[0].bullet_auto so that placeholder inheritance can
  // apply it to slide paragraphs.  The earlier bug: we parsed a:buChar but not
  // a:buAutoNum, so numbered-list layouts inherited nothing and fell back to
  // the default "•" bullet.
  let slide_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    #|  <p:cSld><p:spTree>
    #|    <p:sp>
    #|      <p:nvSpPr><p:cNvPr id="1" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="13"/></p:nvPr></p:nvSpPr>
    #|      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    #|      <p:txBody>
    #|        <a:bodyPr/>
    #|        <a:lstStyle>
    #|          <a:lvl1pPr marL="0" indent="302676"><a:buAutoNum type="arabicPeriod"/></a:lvl1pPr>
    #|          <a:lvl2pPr marL="100000" indent="0"><a:buAutoNum type="alphaLcPeriod"/></a:lvl2pPr>
    #|        </a:lstStyle>
    #|        <a:p><a:r><a:rPr lang="en-US"/><a:t>item</a:t></a:r></a:p>
    #|      </p:txBody>
    #|    </p:sp>
    #|  </p:spTree></p:cSld>
    #|</p:sld>
  let slide = parse_slide(slide_xml, SlideSize::default(), ThemeData::default())
  let lst = slide.shapes[0].lst_style
  assert_eq(lst.levels[0].bullet_auto, "arabicPeriod")
  assert_eq(lst.levels[1].bullet_auto, "alphaLcPeriod")
  // Other levels untouched
  assert_eq(lst.levels[2].bullet_auto, "")
}

///|
test "parse_slide explicit fill wins over p:style/a:fillRef" {
  // When spPr carries an explicit fill (solidFill or noFill), the p:style
  // fillRef must be ignored — explicit properties always take precedence.
  let theme_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
    #|  <a:themeElements>
    #|    <a:clrScheme name="Test">
    #|      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    #|      <a:dk2><a:srgbClr val="112233"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    #|      <a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2>
    #|      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
    #|      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
    #|      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F63"/></a:folHlink>
    #|    </a:clrScheme>
    #|    <a:fontScheme name="Test"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
    #|    <a:fmtScheme name="Test">
    #|      <a:fillStyleLst>
    #|        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    #|        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    #|        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    #|      </a:fillStyleLst>
    #|      <a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/>
    #|    </a:fmtScheme>
    #|  </a:themeElements>
    #|</a:theme>
  let theme = parse_theme(theme_xml)
  let slide_xml =
    #|<?xml version="1.0" encoding="UTF-8"?>
    #|<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    #|  <p:cSld><p:spTree>
    #|    <p:sp>
    #|      <p:nvSpPr><p:cNvPr id="1" name="r"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    #|      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></p:spPr>
    #|      <p:style>
    #|        <a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
    #|        <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
    #|        <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    #|        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
    #|      </p:style>
    #|      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    #|    </p:sp>
    #|  </p:spTree></p:cSld>
    #|</p:sld>
  let slide = parse_slide(slide_xml, SlideSize::default(), theme)
  let shape = slide.shapes[0]
  // Explicit solidFill 0000FF wins, not the theme accent1 (FF0000).
  assert_eq(shape.fill.r, 0)
  assert_eq(shape.fill.g, 0)
  assert_eq(shape.fill.b, 255)
}

///|
test "EffectList with blur is not none" {
  let e : EffectList = { ..EffectList::none(), blur: { rad: 76200 } }
  assert_true(!e.is_none())
  assert_eq(e.blur.rad, 76200)
}

///|
test "PresetShadow::none and is_none" {
  let ps = PresetShadow::none()
  assert_true(ps.is_none())
  assert_eq(ps.prst, "")
  let ps2 : PresetShadow = {
    prst: "shdw1",
    dist: 38100,
    dir: 2700000,
    color: parse_hex_color("000000"),
  }
  assert_true(!ps2.is_none())
  assert_eq(ps2.prst, "shdw1")
}

///|
test "PresetShadow::blur_rad by preset" {
  let ps1 : PresetShadow = {
    prst: "shdw1",
    dist: 0,
    dir: 0,
    color: Color::none(),
  }
  assert_eq(ps1.blur_rad(), 38100) // standard offset shadow
  let ps9 : PresetShadow = {
    prst: "shdw9",
    dist: 0,
    dir: 0,
    color: Color::none(),
  }
  assert_eq(ps9.blur_rad(), 76200) // center shadow — larger blur
  let ps17 : PresetShadow = {
    prst: "shdw17",
    dist: 0,
    dir: 0,
    color: Color::none(),
  }
  assert_eq(ps17.blur_rad(), 57150) // below perspective
}

///|
test "EffectList with prst_shadow is not none" {
  let e : EffectList = {
    ..EffectList::none(),
    prst_shadow: {
      prst: "shdw2",
      dist: 38100,
      dir: 5400000,
      color: parse_hex_color("333333"),
    },
  }
  assert_true(!e.is_none())
  assert_eq(e.prst_shadow.prst, "shdw2")
}

///|
test "FillOverlay::none is none" {
  let fo = FillOverlay::none()
  assert_true(fo.is_none())
  assert_eq(fo.blend, "")
}

///|
test "FillOverlay with blend and color is not none" {
  let fo : FillOverlay = { blend: "mult", color: parse_hex_color("FF0000") }
  assert_true(!fo.is_none())
  assert_eq(fo.blend, "mult")
  assert_eq(fo.color.r, 255)
}

///|
test "FillOverlay with over blend (default)" {
  let fo : FillOverlay = { blend: "over", color: parse_hex_color("0000FF") }
  assert_true(!fo.is_none())
  assert_eq(fo.blend, "over")
  assert_eq(fo.color.b, 255)
}

///|
test "EffectList with fill_overlay is not none" {
  let e : EffectList = {
    ..EffectList::none(),
    fill_overlay: { blend: "screen", color: parse_hex_color("00FF00") },
  }
  assert_true(!e.is_none())
  assert_eq(e.fill_overlay.blend, "screen")
}

// ── GradientFill / BlipFill / PatternFill ───────────────────────────────────

///|
test "GradientFill::none" {
  assert_true(GradientFill::none().is_none())
}

///|
test "BlipFill::none" {
  assert_true(BlipFill::none().is_none())
  assert_false(BlipFill::none().has_crop())
}

///|
test "PatternFill::none" {
  assert_true(PatternFill::none().is_none())
}

// ── Shape3d / Scene3d / Bevel ───────────────────────────────────────────────

///|
test "Shape3d::none" {
  assert_true(Shape3d::none().is_none())
}

///|
test "Scene3d::none" {
  assert_true(Scene3d::none().is_none())
}

///|
test "Bevel::none" {
  assert_true(Bevel::none().is_none())
}

// ── ChartEx parsing ─────────────────────────────────────────────────────────

///|
test "parse_chartex_xml waterfall" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    #|  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:strDim type="cat">
    #|        <cx:lvl ptCount="3">
    #|          <cx:pt idx="0">Q1</cx:pt>
    #|          <cx:pt idx="1">Q2</cx:pt>
    #|          <cx:pt idx="2">Total</cx:pt>
    #|        </cx:lvl>
    #|      </cx:strDim>
    #|      <cx:numDim type="val">
    #|        <cx:lvl ptCount="3" formatCode="General">
    #|          <cx:pt idx="0">100</cx:pt>
    #|          <cx:pt idx="1">50</cx:pt>
    #|          <cx:pt idx="2">150</cx:pt>
    #|        </cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="waterfall">
    #|          <cx:dataId val="0"/>
    #|          <cx:layoutPr>
    #|            <cx:subtotals>
    #|              <cx:idx val="2"/>
    #|            </cx:subtotals>
    #|          </cx:layoutPr>
    #|        </cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_eq(chart.groups.length(), 1)
  assert_true(
    match chart.groups[0].chart_type {
      WaterfallChart => true
      _ => false
    },
  )
  assert_eq(chart.groups[0].subtotals.length(), 1)
  assert_eq(chart.groups[0].subtotals[0], 2)
  assert_eq(chart.groups[0].series.length(), 1)
  // Check data was mapped correctly
  let cats = match chart.groups[0].series[0].cat {
    StrSource(_, data) => data.points
    _ => []
  }
  assert_eq(cats.length(), 3)
  assert_eq(cats[0].value, "Q1")
}

///|
test "parse_chartex_xml funnel" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:strDim type="cat">
    #|        <cx:lvl><cx:pt idx="0">Stage1</cx:pt><cx:pt idx="1">Stage2</cx:pt></cx:lvl>
    #|      </cx:strDim>
    #|      <cx:numDim type="val">
    #|        <cx:lvl><cx:pt idx="0">1000</cx:pt><cx:pt idx="1">500</cx:pt></cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="funnel"><cx:dataId val="0"/></cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|    <cx:legend pos="b"/>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_eq(chart.groups.length(), 1)
  assert_true(
    match chart.groups[0].chart_type {
      FunnelChart => true
      _ => false
    },
  )
  assert_eq(chart.legend.position, "b")
  assert_true(chart.legend.show)
}

///|
test "parse_chartex_xml treemap" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:strDim type="cat">
    #|        <cx:lvl><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt><cx:pt idx="2">C</cx:pt></cx:lvl>
    #|      </cx:strDim>
    #|      <cx:numDim type="val">
    #|        <cx:lvl><cx:pt idx="0">50</cx:pt><cx:pt idx="1">30</cx:pt><cx:pt idx="2">20</cx:pt></cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:title><cx:tx><cx:txData><cx:v>My Treemap</cx:v></cx:txData></cx:tx></cx:title>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="treemap"><cx:dataId val="0"/></cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_eq(chart.title, "My Treemap")
  assert_true(
    match chart.groups[0].chart_type {
      TreemapChart => true
      _ => false
    },
  )
}

///|
test "parse_chartex_xml all layout types" {
  // Test all 6 layout types map to correct ChartKind
  let layout_ids = [
    "waterfall", "treemap", "sunburst", "boxWhisker", "funnel", "clusteredColumn",
  ]
  let expected_types = [
    "WaterfallChart", "TreemapChart", "SunburstChart", "BoxWhiskerChart", "FunnelChart",
    "HistogramChart",
  ]
  for li in 0..<layout_ids.length() {
    let lid = layout_ids[li]
    let xml = "<?xml version=\"1.0\"?><cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chartData><cx:data id=\"0\"><cx:numDim type=\"val\"><cx:lvl><cx:pt idx=\"0\">10</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId=\"" +
      lid +
      "\"><cx:dataId val=\"0\"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"
    let chart = parse_chartex_xml(xml, ThemeData::default())
    assert_eq(chart.groups.length(), 1)
    let kind_str = match chart.groups[0].chart_type {
      WaterfallChart => "WaterfallChart"
      TreemapChart => "TreemapChart"
      SunburstChart => "SunburstChart"
      BoxWhiskerChart => "BoxWhiskerChart"
      FunnelChart => "FunnelChart"
      HistogramChart => "HistogramChart"
      _ => "other"
    }
    assert_eq(kind_str, expected_types[li])
  }
}

///|
test "parse_chartex_xml sunburst" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:strDim type="cat">
    #|        <cx:lvl><cx:pt idx="0">N</cx:pt><cx:pt idx="1">S</cx:pt><cx:pt idx="2">E</cx:pt></cx:lvl>
    #|      </cx:strDim>
    #|      <cx:numDim type="size">
    #|        <cx:lvl><cx:pt idx="0">40</cx:pt><cx:pt idx="1">30</cx:pt><cx:pt idx="2">20</cx:pt></cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="sunburst"><cx:dataId val="0"/></cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_true(
    match chart.groups[0].chart_type {
      SunburstChart => true
      _ => false
    },
  )
  // type="size" should be parsed as numDim
  let vals = match chart.groups[0].series[0].val {
    NumSource(_, data) => data.points
    _ => []
  }
  assert_eq(vals.length(), 3)
  assert_eq(vals[0].value, "40")
}

///|
test "parse_chartex_xml waterfall values and multiple subtotals" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:strDim type="cat">
    #|        <cx:lvl><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt><cx:pt idx="2">Sub</cx:pt><cx:pt idx="3">C</cx:pt><cx:pt idx="4">Total</cx:pt></cx:lvl>
    #|      </cx:strDim>
    #|      <cx:numDim type="val">
    #|        <cx:lvl><cx:pt idx="0">100</cx:pt><cx:pt idx="1">-30</cx:pt><cx:pt idx="2">70</cx:pt><cx:pt idx="3">50</cx:pt><cx:pt idx="4">120</cx:pt></cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="waterfall">
    #|          <cx:dataId val="0"/>
    #|          <cx:layoutPr>
    #|            <cx:subtotals><cx:idx val="2"/><cx:idx val="4"/></cx:subtotals>
    #|          </cx:layoutPr>
    #|        </cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  // Multiple subtotals
  assert_eq(chart.groups[0].subtotals.length(), 2)
  assert_eq(chart.groups[0].subtotals[0], 2)
  assert_eq(chart.groups[0].subtotals[1], 4)
  // Values parsed correctly
  let vals = match chart.groups[0].series[0].val {
    NumSource(_, data) => data.points
    _ => []
  }
  assert_eq(vals.length(), 5)
  assert_eq(vals[0].value, "100")
  assert_eq(vals[1].value, "-30")
}

///|
test "parse_chartex_xml empty chartSpace" {
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_eq(chart.groups.length(), 0)
  assert_true(!chart.chart_xml.is_empty())
}

///|
test "parse_chartex_xml histogram no strDim" {
  // Real histograms may have only numDim (no categories)
  let xml =
    #|<?xml version="1.0"?>
    #|<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
    #|  <cx:chartData>
    #|    <cx:data id="0">
    #|      <cx:numDim type="val">
    #|        <cx:lvl><cx:pt idx="0">5</cx:pt><cx:pt idx="1">15</cx:pt><cx:pt idx="2">25</cx:pt></cx:lvl>
    #|      </cx:numDim>
    #|    </cx:data>
    #|  </cx:chartData>
    #|  <cx:chart>
    #|    <cx:plotArea>
    #|      <cx:plotAreaRegion>
    #|        <cx:series layoutId="clusteredColumn"><cx:dataId val="0"/></cx:series>
    #|      </cx:plotAreaRegion>
    #|    </cx:plotArea>
    #|  </cx:chart>
    #|</cx:chartSpace>
  let chart = parse_chartex_xml(xml, ThemeData::default())
  assert_true(
    match chart.groups[0].chart_type {
      HistogramChart => true
      _ => false
    },
  )
  // No categories — cat should be NoData
  assert_true(
    match chart.groups[0].series[0].cat {
      NoData => true
      _ => false
    },
  )
}
