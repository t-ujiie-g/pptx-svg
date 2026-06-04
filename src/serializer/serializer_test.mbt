///|
using @test {assert_eq}

// ── serialize_slide minimal ──────────────────────────────────────────────────

///|
test "serialize_slide empty slide" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  // Should start with XML declaration
  assert_true(@xml.str_starts_with(xml, "<?xml version"))
  // Should contain p:sld root element
  assert_true(
    @xml.str_starts_with(
      xml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld",
    ),
  )
  // Should contain shape tree
  let nodes = @xml.parse_xml(xml)
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_tag(nodes[0]), "p:sld")
}

///|
test "serialize_slide hidden slide" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: true,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  assert_eq(@xml.get_attr(nodes[0], "show"), Some("0"))
}

///|
test "serialize_slide svg-only blip background omits r:embed" {
  // Regression: SVG-only blips (no r:embed, only asvg:svgBlip) were being
  // serialized as <a:blip r:embed=""> which PowerPoint rejects.
  let svg_only : @ooxml.BlipFill = {
    ..@ooxml.BlipFill::none(),
    svg_rid: "rId9",
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: svg_only,
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("asvg:svgBlip"))
  assert_true(xml.contains("r:embed=\"rId9\""))
  // Must NOT emit the empty r:embed attribute
  assert_false(xml.contains("r:embed=\"\""))
}

///|
test "serialize_slide raster blip background keeps r:embed" {
  let raster : @ooxml.BlipFill = { ..@ooxml.BlipFill::none(), rid: "rId7" }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: raster,
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("<a:blip r:embed=\"rId7\""))
}

///|
test "serialize_slide with colored background" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.parse_hex_color("FF0000"),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  // p:bg is inside p:cSld
  match @xml.find_path(nodes, ["p:sld", "p:cSld", "p:bg"]) {
    Some(_) => ()
    None => assert_true(false)
  }
}

///|
test "serialize_slide with single shape" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    {
      x: 100000,
      y: 200000,
      cx: 500000,
      cy: 300000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  let csld = @xml.get_children(nodes[0])
  match @xml.find_child(csld, "p:cSld") {
    Some(csld_node) => {
      let sp_tree = @xml.find_child(@xml.get_children(csld_node), "p:spTree")
      assert_false(sp_tree is None)
    }
    None => assert_true(false)
  }
}

///|
test "serialize_slide shape with blur effect" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    effects: { ..@ooxml.EffectList::none(), blur: { rad: 76200 } },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  // Should contain a:effectLst with a:blur
  assert_true(xml.contains("<a:effectLst>"))
  assert_true(xml.contains("<a:blur"))
  assert_true(xml.contains("rad=\"76200\""))
}

///|
test "serialize_slide shape with outer shadow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      outer_shadow: {
        blur_rad: 152400,
        dist: 114300,
        dir: 5400000,
        color: @ooxml.parse_hex_color("000000"),
        sx: 0,
        sy: 0,
        algn: "ctr",
        rot_with_shape: false,
      },
    },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("<a:outerShdw"))
  assert_true(xml.contains("blurRad=\"152400\""))
  assert_true(xml.contains("dir=\"5400000\""))
  assert_true(xml.contains("algn=\"ctr\""))
}

///|
test "serialize_slide shape with glow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      glow: { rad: 228600, color: @ooxml.parse_hex_color("FFC000") },
    },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("<a:glow"))
  assert_true(xml.contains("rad=\"228600\""))
}

///|
test "serialize_slide shape with preset shadow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      prst_shadow: {
        prst: "shdw1",
        dist: 38100,
        dir: 2700000,
        color: @ooxml.parse_hex_color("000000"),
      },
    },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("<a:prstShdw"))
  assert_true(xml.contains("prst=\"shdw1\""))
  assert_true(xml.contains("dist=\"38100\""))
  assert_true(xml.contains("dir=\"2700000\""))
}

///|
test "serialize_slide fill overlay" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
  )
  let shape = {
    ..shape,
    effects: {
      ..@ooxml.EffectList::none(),
      fill_overlay: { blend: "mult", color: @ooxml.parse_hex_color("FF0000") },
    },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  assert_true(xml.contains("<a:fillOverlay"))
  assert_true(xml.contains("blend=\"mult\""))
  // fillOverlay must wrap color in a:solidFill with RRGGBB val (not AARRGGBB)
  assert_true(
    xml.contains(
      "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:fillOverlay>",
    ),
  )
}

///|
test "serialize_slide fill overlay with alpha" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
  )
  let shape = {
    ..shape,
    effects: {
      ..@ooxml.EffectList::none(),
      fill_overlay: { blend: "over", color: { r: 255, g: 0, b: 0, alpha: 128 } },
    },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  // val must be RRGGBB (not AARRGGBB), alpha as child element
  assert_true(xml.contains("val=\"FF0000\""))
  assert_true(xml.contains("<a:alpha val="))
  assert_true(!xml.contains("val=\"80FF0000\""))
}

///|
test "serialize_slide no effectLst when none" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  // No effects → no a:effectLst in output
  assert_true(!xml.contains("<a:effectLst>"))
}

// ── restore round-trip idempotency (E6.1 Undo/Redo) ──────────────────────────
// The Undo/Redo history captures a slide's OOXML via get_slide_ooxml and restores
// it through restore_slide_ooxml (parse_slide → g_slides). This must be stable:
// serialize → parse_slide → serialize should be a fixed point, so no data is lost
// across an undo/redo cycle.

///|
test "restore round-trip: serialize is idempotent after parse_slide" {
  let shape1 : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 100000,
        y: 200000,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    fill: { r: 200, g: 100, b: 50, alpha: 255 },
    stroke: { r: 0, g: 0, b: 0, alpha: 255 },
    stroke_w: 12700,
  }
  let shape2 : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
      {
        x: 700000,
        y: 100000,
        cx: 400000,
        cy: 400000,
        rot: 5400000,
        flip_h: true,
        flip_v: false,
      },
    ),
    fill: { r: 30, g: 144, b: 255, alpha: 255 },
  }
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape1, shape2],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let s1 = serialize_slide(slide)
  let slide2 = @ooxml.parse_slide(
    s1,
    @ooxml.SlideSize::default(),
    @ooxml.ThemeData::default(),
  )
  let s2 = serialize_slide(slide2)
  assert_eq(s1, s2)
}

///|
test "restore round-trip: shape text survives parse_slide" {
  let xml =
    #|<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    #|<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello Undo</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>
  let slide = @ooxml.parse_slide(
    xml,
    @ooxml.SlideSize::default(),
    @ooxml.ThemeData::default(),
  )
  // Re-serialize and confirm the text is preserved.
  let out = serialize_slide(slide)
  assert_true(out.contains("Hello Undo"))
  // And the re-serialize is a fixed point.
  let slide2 = @ooxml.parse_slide(
    out,
    @ooxml.SlideSize::default(),
    @ooxml.ThemeData::default(),
  )
  assert_eq(out, serialize_slide(slide2))
}

// ── parse_shapes_xml round-trip (E6.5 copy/paste) ────────────────────────────

///|
test "parse_shapes_xml round-trips a serialized shape" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 100000,
        y: 200000,
        cx: 500000,
        cy: 300000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    fill: { r: 200, g: 100, b: 50, alpha: 255 },
  }
  let xml = serialize_shape(shape)
  let shapes = @ooxml.parse_shapes_xml(xml, @ooxml.ThemeData::default())
  assert_eq(shapes.length(), 1)
  let s = shapes[0]
  assert_eq(s.transform.x, 100000)
  assert_eq(s.transform.y, 200000)
  assert_eq(s.transform.cx, 500000)
  assert_eq(s.transform.cy, 300000)
  let is_rect = match s.kind {
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect) => true
    _ => false
  }
  assert_true(is_rect)
}
