// ── Color constructors ──────────────────────────────────────────────────────

test "Color::white" {
  let c = Color::white()
  assert_eq(c.r, 255)
  assert_eq(c.g, 255)
  assert_eq(c.b, 255)
  assert_eq(c.alpha, 255)
  assert_false(c.is_none())
}

test "Color::black" {
  let c = Color::black()
  assert_eq(c.r, 0)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

test "Color::none" {
  let c = Color::none()
  assert_true(c.is_none())
}

// ── parse_hex_color ─────────────────────────────────────────────────────────

test "parse_hex_color 6-char" {
  let c = parse_hex_color("FF8800")
  assert_eq(c.r, 255)
  assert_eq(c.g, 136)
  assert_eq(c.b, 0)
  assert_eq(c.alpha, 255)
}

test "parse_hex_color 8-char with alpha" {
  let c = parse_hex_color("80FF0000")
  assert_eq(c.alpha, 128)
  assert_eq(c.r, 255)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

test "parse_hex_color short string" {
  // Falls back to black
  let c = parse_hex_color("FFF")
  assert_eq(c.r, 0)
  assert_eq(c.g, 0)
  assert_eq(c.b, 0)
}

// ── Color::to_hex / to_hex_rgb / to_css ─────────────────────────────────────

test "Color::to_hex roundtrip" {
  let c = parse_hex_color("4488CC")
  assert_eq(c.to_hex(), "4488CC")
  assert_eq(c.to_hex_rgb(), "4488CC")
}

test "Color::to_hex with alpha" {
  let c = parse_hex_color("80FF0000")
  assert_eq(c.to_hex(), "80FF0000")
  assert_eq(c.to_hex_rgb(), "FF0000")
}

test "Color::to_hex none" {
  assert_eq(Color::none().to_hex(), "none")
  assert_eq(Color::none().to_hex_rgb(), "none")
}

test "Color::to_css opaque" {
  let c = parse_hex_color("FF0000")
  assert_eq(c.to_css(), "rgb(255,0,0)")
}

test "Color::to_css with alpha" {
  let c = parse_hex_color("80FF0000")
  let css = c.to_css()
  assert_true(@xml.str_starts_with(css, "rgba(255,0,0,"))
}

test "Color::to_css none" {
  assert_eq(Color::none().to_css(), "none")
}

// ── alpha_to_css_str ────────────────────────────────────────────────────────

test "alpha_to_css_str" {
  assert_eq(alpha_to_css_str(255), "1")
  assert_eq(alpha_to_css_str(0), "0")
  assert_eq(alpha_to_css_str(128), "0.50")
}

// ── resolve_scheme_color ────────────────────────────────────────────────────

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

test "apply_color_modifiers empty" {
  let c = parse_hex_color("FF0000")
  let result = apply_color_modifiers(c, [])
  assert_eq(result.r, 255)
  assert_eq(result.g, 0)
  assert_eq(result.b, 0)
}

test "apply_color_modifiers alpha" {
  let c = parse_hex_color("FF0000")
  let nodes = @xml.parse_xml("<a:alpha val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // alpha 50000/100000 * 255 ≈ 127-128
  assert_true(result.alpha < 130)
  assert_true(result.alpha > 125)
}

test "apply_color_modifiers shade" {
  let c = Color::white()
  let nodes = @xml.parse_xml("<a:shade val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // 255 * 50000/100000 = 127
  assert_eq(result.r, 127)
  assert_eq(result.g, 127)
  assert_eq(result.b, 127)
}

test "apply_color_modifiers tint" {
  let c = Color::black()
  let nodes = @xml.parse_xml("<a:tint val=\"50000\"/>")
  let result = apply_color_modifiers(c, nodes)
  // 255 - (255-0)*50000/100000 = 128 (integer rounding)
  assert_eq(result.r, 128)
  assert_eq(result.g, 128)
  assert_eq(result.b, 128)
}

// ── parse_geom / to_prst ────────────────────────────────────────────────────

test "parse_geom basic shapes" {
  assert_eq(parse_geom("rect", []).to_prst(), "rect")
  assert_eq(parse_geom("ellipse", []).to_prst(), "ellipse")
  assert_eq(parse_geom("roundRect", []).to_prst(), "roundRect")
  assert_eq(parse_geom("line", []).to_prst(), "line")
  assert_eq(parse_geom("straightConnector1", []).to_prst(), "line")
}

test "parse_geom other" {
  assert_eq(parse_geom("star5", []).to_prst(), "star5")
}

test "parse_connector_geom" {
  assert_eq(parse_connector_geom("straightConnector1", []).to_prst(), "line")
  assert_eq(parse_connector_geom("bentConnector3", [50000]).to_prst(), "bentConnector3")
}

// ── parse_rels ──────────────────────────────────────────────────────────────

test "parse_rels basic" {
  let xml = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/></Relationships>"
  let rels = parse_rels(xml)
  assert_eq(rels.length(), 2)
  assert_eq(find_rel_target(rels, "rId1"), "../slideLayouts/slideLayout1.xml")
  assert_eq(find_rel_target(rels, "rId2"), "../media/image1.png")
  assert_eq(find_rel_target(rels, "rId99"), "")
}

// ── parse_slide_size ────────────────────────────────────────────────────────

test "parse_slide_size standard" {
  let xml = "<p:presentation><p:sldSz cx=\"9144000\" cy=\"6858000\"/></p:presentation>"
  let size = parse_slide_size(xml)
  assert_eq(size.cx, 9144000)
  assert_eq(size.cy, 6858000)
}

test "parse_slide_size widescreen" {
  let xml = "<p:presentation><p:sldSz cx=\"12192000\" cy=\"6858000\"/></p:presentation>"
  let size = parse_slide_size(xml)
  assert_eq(size.cx, 12192000)
}

// ── EffectList / Shadow / Glow ──────────────────────────────────────────────

test "EffectList::none" {
  let e = EffectList::none()
  assert_true(e.is_none())
  assert_true(e.outer_shadow.is_none())
  assert_true(e.inner_shadow.is_none())
  assert_true(e.glow.is_none())
  assert_true(e.soft_edge.is_none())
  assert_true(e.reflection.is_none())
}

// ── GradientFill / BlipFill / PatternFill ───────────────────────────────────

test "GradientFill::none" {
  assert_true(GradientFill::none().is_none())
}

test "BlipFill::none" {
  assert_true(BlipFill::none().is_none())
  assert_false(BlipFill::none().has_crop())
}

test "PatternFill::none" {
  assert_true(PatternFill::none().is_none())
}

// ── Shape3d / Scene3d / Bevel ───────────────────────────────────────────────

test "Shape3d::none" {
  assert_true(Shape3d::none().is_none())
}

test "Scene3d::none" {
  assert_true(Scene3d::none().is_none())
}

test "Bevel::none" {
  assert_true(Bevel::none().is_none())
}
