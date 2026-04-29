// ── render_slide_svg basic ───────────────────────────────────────────────────

///|
fn make_empty_slide() -> @ooxml.SlideData {
  {
    slide_size: { cx: 9144000, cy: 6858000 },
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
}

///|
test "render_slide_svg empty slide" {
  let slide = make_empty_slide()
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // Should produce valid SVG with metadata
  assert_true(@xml.str_starts_with(svg, "<svg"))
  assert_true(@xml.str_ends_with(svg, "</svg>"))
  // Should contain data-ooxml-slide-cx
  let nodes = @xml.parse_xml(svg)
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_tag(nodes[0]), "svg")
  assert_eq(@xml.get_attr(nodes[0], "data-ooxml-slide-cx"), Some("9144000"))
  assert_eq(@xml.get_attr(nodes[0], "data-ooxml-slide-cy"), Some("6858000"))
}

///|
test "render_slide_svg hidden slide" {
  let slide : @ooxml.SlideData = { ..make_empty_slide(), hidden: true }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let nodes = @xml.parse_xml(svg)
  assert_eq(@xml.get_attr(nodes[0], "data-ooxml-hidden"), Some("1"))
}

///|
test "render_slide_svg with rect shape" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    {
      x: 914400,
      y: 914400,
      cx: 2743200,
      cy: 1828800,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // Should contain a group for the shape
  let nodes = @xml.parse_xml(svg)
  let svg_children = @xml.get_children(nodes[0])
  // Find a <g> with data-ooxml-shape-type
  let mut found_shape = false
  for i = 0; i < svg_children.length(); i = i + 1 {
    if @xml.get_tag(svg_children[i]) == "g" {
      let st = @xml.get_attr(svg_children[i], "data-ooxml-shape-type")
      if !(st is None) {
        found_shape = true
      }
    }
  }
  assert_true(found_shape)
}

///|
test "render_slide_svg with ellipse" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    {
      x: 0,
      y: 0,
      cx: 1000000,
      cy: 1000000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // SVG should contain an ellipse element
  let nodes = @xml.parse_xml(svg)
  let svg_children = @xml.get_children(nodes[0])
  let mut found_ellipse = false
  for i = 0; i < svg_children.length(); i = i + 1 {
    if @xml.get_tag(svg_children[i]) == "g" {
      let children = @xml.get_children(svg_children[i])
      for j = 0; j < children.length(); j = j + 1 {
        if @xml.get_tag(children[j]) == "ellipse" {
          found_ellipse = true
        }
      }
    }
  }
  assert_true(found_ellipse)
}

///|
test "render_slide_svg colored background" {
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    background: @ooxml.parse_hex_color("FF0000"),
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_eq(
    @xml.get_attr(@xml.parse_xml(svg)[0], "data-ooxml-bg"),
    Some("FF0000"),
  )
}

// ── Round-trip: render → parse → compare ────────────────────────────────────

///|
test "roundtrip empty slide" {
  let slide = make_empty_slide()
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.slide_size.cx, slide.slide_size.cx)
  assert_eq(parsed.slide_size.cy, slide.slide_size.cy)
  assert_eq(parsed.background.r, slide.background.r)
  assert_eq(parsed.background.g, slide.background.g)
  assert_eq(parsed.background.b, slide.background.b)
  assert_eq(parsed.hidden, false)
}

///|
test "roundtrip hidden slide" {
  let slide : @ooxml.SlideData = { ..make_empty_slide(), hidden: true }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_true(parsed.hidden)
}

///|
test "roundtrip colored background" {
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    background: @ooxml.parse_hex_color("3366CC"),
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.background.r, 0x33)
  assert_eq(parsed.background.g, 0x66)
  assert_eq(parsed.background.b, 0xCC)
}

///|
test "roundtrip shape transform" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    {
      x: 914400,
      y: 457200,
      cx: 2743200,
      cy: 1828800,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  assert_eq(parsed.shapes[0].transform.x, 914400)
  assert_eq(parsed.shapes[0].transform.y, 457200)
  assert_eq(parsed.shapes[0].transform.cx, 2743200)
  assert_eq(parsed.shapes[0].transform.cy, 1828800)
}

///|
test "roundtrip shape with blur effect" {
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
    effects: { ..@ooxml.EffectList::none(), blur: { rad: 76200 } },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  assert_eq(parsed.shapes[0].effects.blur.rad, 76200)
  assert_true(!parsed.shapes[0].effects.blur.is_none())
}

///|
test "roundtrip shape with outer shadow" {
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
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.outer_shadow.blur_rad, 152400)
  assert_eq(parsed.shapes[0].effects.outer_shadow.dist, 114300)
  assert_eq(parsed.shapes[0].effects.outer_shadow.dir, 5400000)
  assert_eq(parsed.shapes[0].effects.outer_shadow.algn, "ctr")
}

///|
test "roundtrip shape with glow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
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
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.glow.rad, 228600)
  assert_eq(parsed.shapes[0].effects.glow.color.r, 0xFF)
  assert_eq(parsed.shapes[0].effects.glow.color.g, 0xC0)
}

///|
test "roundtrip shape with soft edge" {
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
    effects: { ..@ooxml.EffectList::none(), soft_edge: { rad: 63500 } },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.soft_edge.rad, 63500)
}

///|
test "roundtrip shape with solid fill" {
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
    fill: @ooxml.parse_hex_color("FF6633"),
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].fill.r, 0xFF)
  assert_eq(parsed.shapes[0].fill.g, 0x66)
  assert_eq(parsed.shapes[0].fill.b, 0x33)
}

///|
test "roundtrip shape with stroke" {
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
    stroke: @ooxml.parse_hex_color("0000FF"),
    stroke_w: 25400,
    stroke_dash: "dash",
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].stroke.r, 0)
  assert_eq(parsed.shapes[0].stroke.g, 0)
  assert_eq(parsed.shapes[0].stroke.b, 0xFF)
  assert_eq(parsed.shapes[0].stroke_w, 25400)
  assert_eq(parsed.shapes[0].stroke_dash, "dash")
}

///|
test "roundtrip ellipse geometry" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    {
      x: 100000,
      y: 100000,
      cx: 400000,
      cy: 400000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  match parsed.shapes[0].kind {
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse) => ()
    _ => assert_true(false)
  }
}

///|
test "roundtrip shape with preset shadow" {
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
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.prst_shadow.prst, "shdw1")
  assert_eq(parsed.shapes[0].effects.prst_shadow.dist, 38100)
  assert_eq(parsed.shapes[0].effects.prst_shadow.dir, 2700000)
}

///|
test "roundtrip fill overlay effect" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
  )
  let shape = {
    ..shape,
    fill: @ooxml.parse_hex_color("0000FF"),
    effects: {
      ..@ooxml.EffectList::none(),
      fill_overlay: { blend: "mult", color: @ooxml.parse_hex_color("80FF0000") },
    },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.fill_overlay.blend, "mult")
  assert_eq(parsed.shapes[0].effects.fill_overlay.color.r, 255)
  assert_eq(parsed.shapes[0].effects.fill_overlay.color.g, 0)
}

///|
test "roundtrip justify alignment" {
  let run : @ooxml.TextRun = {
    text: "Hello world test",
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size: 1800,
    color: @ooxml.Color::none(),
    font_face: "",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "just",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    line_spacing: 0,
    mar_l: -1,
    indent: 0,
    bullet: "",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 5000000,
        cy: 2000000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // SVG should contain para-align="just"
  assert_true(svg.contains("para-align=\"just\""))
  // Round-trip: parse back
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].paragraphs[0].align, "just")
}

///|
/// Construct a text shape with a single bulleted paragraph for indent tests.
/// marL = 457200 EMU (0.5"), bullet = "•", font_size = 1800 (18pt).
fn make_bulleted_shape(indent : Int) -> @ooxml.Shape {
  let run : @ooxml.TextRun = {
    text: "Hello",
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size: 1800,
    color: @ooxml.Color::none(),
    font_face: "",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    line_spacing: 0,
    mar_l: 457200,
    indent,
    bullet: "•",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 5000000,
        cy: 2000000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
}

///|
test "bullet repositioning only applies to hanging indent" {
  // With hanging indent (indent < 0) the renderer must reposition the first
  // line's text anchor to marL so it lines up under continuation lines.
  // With positive first-line indent the text must flow naturally after the
  // bullet (regression fix: previous behavior made bullet + text overlap).
  let theme = @ooxml.ThemeData::default()
  let hanging_slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_bulleted_shape(-228600)],
  }
  let positive_slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_bulleted_shape(228600)],
  }
  let hanging_svg = render_slide_svg(hanging_slide, [], 1, theme, [], [], [], [])
  let positive_svg = render_slide_svg(
    positive_slide,
    [],
    1,
    theme,
    [],
    [],
    [],
    [],
  )
  // Both must contain the bullet glyph and text.
  assert_true(hanging_svg.contains("•"))
  assert_true(hanging_svg.contains("Hello"))
  assert_true(positive_svg.contains("•"))
  assert_true(positive_svg.contains("Hello"))
  // Regression guard: the two renderings must differ. The hanging-indent
  // variant emits an extra reposition tspan (to line up wrapped continuation
  // with the first line under marL), while the positive-indent variant must
  // skip that reposition so text flows naturally after the bullet.
  // If the renderer ever reintroduces the unconditional reposition, both
  // SVGs would become byte-identical and this assertion catches it.
  assert_true(hanging_svg != positive_svg)
}

///|
/// Construct a minimal text shape with the given shape-level rotation.
/// Uses noFill/noStroke so only the text itself is visible — matches the
/// slide 6 regression where rotation wasn't applied to the text.
fn make_rotated_text_shape(rot : Int) -> @ooxml.Shape {
  let run : @ooxml.TextRun = {
    text: "Tilted",
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size: 2400,
    color: @ooxml.Color::none(),
    font_face: "",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    line_spacing: 0,
    mar_l: -1,
    indent: 0,
    bullet: "",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 700000,
        y: 1500000,
        cx: 1284515,
        cy: 445828,
        rot,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
}

///|
test "shape xfrm rot rotates text, not only the shape rect" {
  // Regression: slide 6 of rails_study_updated.pptx used text boxes with
  // <a:xfrm rot="..."/> but noFill/noStroke. The renderer rotated the
  // (invisible) rect only and emitted the <text> upright. Fix: compose
  // t.rot + body_props.rot into the text's rotate() transform.
  let theme = @ooxml.ThemeData::default()
  let rotated : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_rotated_text_shape(300000)], // 5°
  }
  let straight : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_rotated_text_shape(0)],
  }
  let rotated_svg = render_slide_svg(rotated, [], 1, theme, [], [], [], [])
  let straight_svg = render_slide_svg(straight, [], 1, theme, [], [], [], [])
  // Rotated shape must emit a rotate(...) transform inside its <text>.
  assert_true(rotated_svg.contains("<text"))
  assert_true(rotated_svg.contains("transform=\"rotate(5.00"))
  // Straight shape must not emit any rotate() transform.
  assert_false(straight_svg.contains("transform=\"rotate"))
}

// ── ChartEx rendering tests ──────────────────────────────────────────────────

///|
fn make_chart_shape(chart : @ooxml.ChartData) -> @ooxml.Shape {
  @ooxml.Shape::default_with(@ooxml.ShapeKind::ChartShape(chart), {
    x: 0,
    y: 0,
    cx: 9144000,
    cy: 6858000,
    rot: 0,
    flip_h: false,
    flip_v: false,
  })
}

///|
fn make_cx_chart_xml(
  layout_id : String,
  cats : String,
  vals : String,
  extra? : String = "",
) -> String {
  let cat_dim = if cats.is_empty() {
    ""
  } else {
    "<cx:strDim type=\"cat\"><cx:lvl>" + cats + "</cx:lvl></cx:strDim>"
  }
  "<?xml version=\"1.0\"?><cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chartData><cx:data id=\"0\">" +
  cat_dim +
  "<cx:numDim type=\"val\"><cx:lvl>" +
  vals +
  "</cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId=\"" +
  layout_id +
  "\"><cx:dataId val=\"0\"/>" +
  extra +
  "</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"
}

///|
test "render waterfall chart produces rects and connectors" {
  let xml = make_cx_chart_xml(
    "waterfall",
    "<cx:pt idx=\"0\">Q1</cx:pt><cx:pt idx=\"1\">Q2</cx:pt><cx:pt idx=\"2\">Total</cx:pt>",
    "<cx:pt idx=\"0\">100</cx:pt><cx:pt idx=\"1\">-20</cx:pt><cx:pt idx=\"2\">80</cx:pt>",
    extra="<cx:layoutPr><cx:subtotals><cx:idx val=\"2\"/></cx:subtotals></cx:layoutPr>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("<rect"))
  assert_true(svg.contains("stroke-dasharray"))
  assert_true(svg.contains("Q1"))
  assert_true(svg.contains("Total"))
}

///|
test "render treemap chart produces labeled rects" {
  let xml = make_cx_chart_xml(
    "treemap", "<cx:pt idx=\"0\">A</cx:pt><cx:pt idx=\"1\">B</cx:pt><cx:pt idx=\"2\">C</cx:pt>",
    "<cx:pt idx=\"0\">50</cx:pt><cx:pt idx=\"1\">30</cx:pt><cx:pt idx=\"2\">20</cx:pt>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("<rect"))
  assert_true(svg.contains("stroke=\"white\""))
}

///|
test "render sunburst chart produces arc paths" {
  let xml = make_cx_chart_xml(
    "sunburst", "<cx:pt idx=\"0\">N</cx:pt><cx:pt idx=\"1\">S</cx:pt><cx:pt idx=\"2\">E</cx:pt>",
    "<cx:pt idx=\"0\">40</cx:pt><cx:pt idx=\"1\">30</cx:pt><cx:pt idx=\"2\">20</cx:pt>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("<path d=\"M"))
}

///|
test "render funnel chart produces centered rects" {
  let xml = make_cx_chart_xml(
    "funnel", "<cx:pt idx=\"0\">Prospects</cx:pt><cx:pt idx=\"1\">Qualified</cx:pt><cx:pt idx=\"2\">Closed</cx:pt>",
    "<cx:pt idx=\"0\">1000</cx:pt><cx:pt idx=\"1\">600</cx:pt><cx:pt idx=\"2\">100</cx:pt>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("<rect"))
  assert_true(svg.contains("Prospects"))
  assert_true(svg.contains("Closed"))
}

///|
test "render box-whisker chart produces lines and rects" {
  let xml = make_cx_chart_xml(
    "boxWhisker", "<cx:pt idx=\"0\">Data</cx:pt><cx:pt idx=\"1\">Data</cx:pt><cx:pt idx=\"2\">Data</cx:pt><cx:pt idx=\"3\">Data</cx:pt><cx:pt idx=\"4\">Data</cx:pt><cx:pt idx=\"5\">Data</cx:pt><cx:pt idx=\"6\">Data</cx:pt>",
    "<cx:pt idx=\"0\">10</cx:pt><cx:pt idx=\"1\">15</cx:pt><cx:pt idx=\"2\">20</cx:pt><cx:pt idx=\"3\">25</cx:pt><cx:pt idx=\"4\">30</cx:pt><cx:pt idx=\"5\">35</cx:pt><cx:pt idx=\"6\">40</cx:pt>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("fill-opacity=\"0.3\""))
  assert_true(svg.contains("<line"))
  assert_true(svg.contains("stroke-width=\"2\""))
}

///|
test "render histogram chart reuses bar renderer" {
  let xml = make_cx_chart_xml(
    "clusteredColumn", "<cx:pt idx=\"0\">0-20</cx:pt><cx:pt idx=\"1\">20-40</cx:pt><cx:pt idx=\"2\">40-60</cx:pt>",
    "<cx:pt idx=\"0\">5</cx:pt><cx:pt idx=\"1\">15</cx:pt><cx:pt idx=\"2\">25</cx:pt>",
  )
  let chart = @ooxml.parse_chartex_xml(xml, @ooxml.ThemeData::default())
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [make_chart_shape(chart)],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(svg.contains("<rect"))
}

///|
test "roundtrip multiple shapes" {
  let shape1 = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 200000, cy: 200000, rot: 0, flip_h: false, flip_v: false },
  )
  let shape2 = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    {
      x: 300000,
      y: 300000,
      cx: 200000,
      cy: 200000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    shapes: [shape1, shape2],
  }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 2)
  assert_eq(parsed.shapes[0].transform.x, 0)
  assert_eq(parsed.shapes[1].transform.x, 300000)
}

// ── sanitize_url (security: H1/H3) ──────────────────────────────────────────

///|
test "sanitize_url accepts safe schemes" {
  assert_eq(sanitize_url("http://example.com/x"), "http://example.com/x")
  assert_eq(sanitize_url("https://example.com"), "https://example.com")
  assert_eq(sanitize_url("mailto:a@b.com"), "mailto:a@b.com")
  assert_eq(sanitize_url("tel:+8190"), "tel:+8190")
  assert_eq(sanitize_url("ftp://host/file"), "ftp://host/file")
}

///|
test "sanitize_url accepts relative URLs and fragments" {
  assert_eq(sanitize_url("slide2.xml"), "slide2.xml")
  assert_eq(sanitize_url("../media/x.png"), "../media/x.png")
  assert_eq(sanitize_url("#anchor"), "#anchor")
  assert_eq(sanitize_url("/path/to"), "/path/to")
}

///|
test "sanitize_url accepts data:image/* only" {
  assert_eq(
    sanitize_url("data:image/png;base64,AAAA"),
    "data:image/png;base64,AAAA",
  )
  assert_eq(sanitize_url("data:text/html,<script>alert(1)</script>"), "")
  assert_eq(sanitize_url("data:application/javascript,alert(1)"), "")
}

///|
test "sanitize_url rejects script-execution schemes" {
  assert_eq(sanitize_url("javascript:alert(1)"), "")
  assert_eq(sanitize_url("vbscript:msgbox(1)"), "")
  assert_eq(sanitize_url("file:///etc/passwd"), "")
  assert_eq(sanitize_url("about:blank"), "")
}

///|
test "sanitize_url is case-insensitive on scheme" {
  // Browsers parse scheme case-insensitively: JaVaScRiPt: must be rejected.
  assert_eq(sanitize_url("JaVaScRiPt:alert(1)"), "")
  assert_eq(sanitize_url("HTTP://example.com"), "HTTP://example.com")
}

///|
test "sanitize_url strips leading whitespace and control bytes" {
  // Browsers tolerate leading whitespace/control chars in href scheme
  // parsing, so "\tjavascript:..." would still execute. We must strip them.
  assert_eq(sanitize_url("\tjavascript:alert(1)"), "")
  assert_eq(sanitize_url("\njavascript:alert(1)"), "")
  assert_eq(sanitize_url(" javascript:alert(1)"), "")
  assert_eq(sanitize_url(""), "")
}

// ── End-to-end attribute escape (security: H2) ──────────────────────────────

///|
/// Helper: find first occurrence of `pat` in `s`, or -1 if not found.
fn find_substring(s : String, pat : String) -> Int {
  let slen = s.length()
  let plen = pat.length()
  if plen == 0 {
    return 0
  }
  if slen < plen {
    return -1
  }
  let mut i = 0
  while i + plen <= slen {
    let mut k = 0
    let mut matched = true
    while k < plen {
      if s.get_char(i + k).unwrap() != pat.get_char(k).unwrap() {
        matched = false
        break
      }
      k = k + 1
    }
    if matched {
      return i
    }
    i = i + 1
  }
  -1
}

///|
test "rendered SVG escapes malicious ph_type in data attribute" {
  // ph_type flows through `da("ph-type", ...)`. After H2's fix, da() must
  // XML-escape, so a literal `"` in the value cannot break the attribute.
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 1000000,
        cy: 500000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    ph_type: "x\" onload=\"alert(1)",
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // The literal injection sequence must NOT appear unescaped.
  assert_eq(find_substring(svg, "\" onload=\""), -1)
  // The escaped form should appear (proves the value made it through).
  assert_true(find_substring(svg, "&quot; onload=&quot;") >= 0)
}

///|
test "rendered SVG drops javascript: hyperlink target" {
  // After H1's fix, an unsafe-scheme hyperlink target is dropped from the
  // <a href> wrapper entirely. The resulting SVG must contain neither
  // `javascript:` nor an <a href> wrapper around the run text.
  let run : @ooxml.TextRun = {
    text: "click",
    bold: false,
    italic: false,
    font_size: 1800,
    color: @ooxml.Color::black(),
    font_face: "Arial",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "rId1",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    bold_explicit: false,
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "l",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    line_spacing: 0,
    mar_l: -1,
    indent: 0,
    bullet: "",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 1000000,
        cy: 500000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let bad_rels : Array[@ooxml.Relationship] = [
    {
      id: "rId1",
      target: "javascript:alert(1)",
      type_: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      target_mode: "External",
    },
  ]
  let svg = render_slide_svg(
    slide,
    bad_rels,
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_eq(find_substring(svg, "javascript:"), -1)
  assert_eq(find_substring(svg, "<a href="), -1)
}

// ── Geometry-specific regression tests ──────────────────────────────────────

///|
/// roundRect must use OOXML default adj 16667 and apply it to min(w,h) — not w/20.
/// Regression: a wide-thin shape (cx >> cy) used to render as a full pill.
test "roundRect default adj uses min(w,h) * 16667 / 100000" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(
      @ooxml.ShapeGeom::RoundRect(@ooxml.round_rect_default_adj),
    ),
    {
      x: 0,
      y: 0,
      cx: 11520000, // very wide
      cy: 550000, //   very thin
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // The rendered <rect> should NOT have rx >= cy_p / 2 (pill).
  // With shape width 11520000 EMU and SVG width scaling to 960 px, cy_p ≈ 28
  // and the pill threshold rx >= 14. A correct radius should be ≈ 4-5 px.
  // Easier to verify: rx is far less than the rendered width.
  assert_true(find_substring(svg, "<rect") >= 0)
  // Pill shape would have rx equal to half the height (~14). Look for
  // suspicious large rx — e.g. rx="576" (the old buggy width/20 pixel result).
  assert_eq(find_substring(svg, "rx=\"576\""), -1)
  assert_eq(find_substring(svg, "rx=\"288\""), -1)
}

///|
/// Author-specified roundRect adj must round-trip through SVG → SlideData.
test "roundRect custom adj round-trips" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::RoundRect(50000)),
    {
      x: 0,
      y: 0,
      cx: 1000000,
      cy: 1000000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  match parsed.shapes[0].kind {
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::RoundRect(adj)) =>
      assert_eq(adj, 50000)
    _ => assert_true(false)
  }
}

///|
/// stripedRightArrow and leftRightArrowCallout used to fall back to a 4-corner
/// rectangle path. Verify the renderer now emits a complex path with multiple
/// segments per the OOXML preset definitions.
test "stripedRightArrow renders as multi-segment path" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(
      @ooxml.ShapeGeom::Other("stripedRightArrow", []),
    ),
    {
      x: 0,
      y: 0,
      cx: 2000000,
      cy: 1000000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // Three subpaths separated by `M` after first move.
  // Plain rect fallback would have only one subpath with 4 line segments.
  assert_true(find_substring(svg, "<path") >= 0)
  // Three subpaths (3 stripes + body+head) → at least 3 `Z` close commands.
  let mut count = 0
  let mut idx = 0
  while idx < svg.length() {
    if svg.get_char(idx).unwrap() == 'Z' {
      count = count + 1
    }
    idx = idx + 1
  }
  assert_true(count >= 3)
}

///|
/// leftRightArrowCallout has a 18-point closed path (callout box + two arrows).
test "leftRightArrowCallout renders as preset path" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(
      @ooxml.ShapeGeom::Other("leftRightArrowCallout", []),
    ),
    {
      x: 0,
      y: 0,
      cx: 2000000,
      cy: 1500000,
      rot: 0,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  assert_true(find_substring(svg, "<path") >= 0)
  // Plain rectangle fallback path has only 4 `L` line commands. The OOXML
  // preset for leftRightArrowCallout has 17 line segments. Count `L` chars
  // (followed by a digit or `-` to avoid matching `Lt;`/`<`).
  let mut count = 0
  let mut idx = 0
  while idx < svg.length() - 1 {
    if svg.get_char(idx).unwrap() == 'L' {
      let next = svg.get_char(idx + 1).unwrap()
      let nc = Char::to_int(next)
      let is_digit = nc >= 48 && nc <= 57
      if is_digit || next == '-' {
        count = count + 1
      }
    }
    idx = idx + 1
  }
  assert_true(count > 4)
}

// ── Bullet color inheritance (OOXML) ────────────────────────────────────────

///|
/// When a paragraph has no <a:buClr>, the bullet must inherit the first run's
/// color (per ECMA-376). Regression: bullets used to fall back to <text fill=
/// "black"> default and render as black on dark backgrounds.
test "bullet inherits first run color when buClr is absent" {
  let run : @ooxml.TextRun = {
    text: "hello",
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size: 0,
    color: ({ r: 255, g: 255, b: 255, alpha: 255 } : @ooxml.Color), // white
    font_face: "",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    mar_l: 342900,
    indent: -342900,
    line_spacing: 0,
    bullet: "•",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(), // ← intentionally unset
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 5000000,
        cy: 1000000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // The bullet tspan must explicitly set fill to the run's white color so it
  // does not inherit black from <text fill="black">.
  assert_true(find_substring(svg, "<tspan fill=\"rgb(255,255,255)\"") >= 0)
}

///|
/// Regression: Google Slides emits bentConnector3 with adj1 values like
/// -39687500 to push the bend point well beyond the bounding box, producing
/// a long visible trunk after rotation. The bend math
/// `(x2 - x1) * adj1 / 100000` overflows Int32 when (x2-x1) and adj1 are both
/// large in opposite signs, AND the legacy pixel-based math also collapsed to
/// zero when the bbox was thinner than one pixel. Both must be handled.
test "bentConnector3 with extreme adj1 keeps bend outside the bbox" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(
      @ooxml.ShapeGeom::Connector("bentConnector3", [-39687500]),
    ),
    {
      x: 3382695,
      y: 1829144,
      cx: 900, // sub-pixel — must not collapse the bend math
      cy: 2016900,
      rot: 5400000,
      flip_h: false,
      flip_v: false,
    },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // Look for the connector path. With Int32 overflow it would wrap to ~349;
  // with proper Int64 math the bend lands at ~317 (38 px outside the bbox).
  // We check that the path contains a coordinate clearly outside the (355±0)
  // bbox column.
  assert_true(find_substring(svg, "<path") >= 0)
  // x=355 is the bbox column; bend should be visibly to the left, e.g. 317.
  assert_true(
    find_substring(svg, "317") >= 0 || find_substring(svg, "318") >= 0,
  )
}

///|
/// Regression: an empty paragraph in an `<a:spAutoFit/>` text box used to
/// reserve the default 24 pt (~28 px) line height, ignoring the font size
/// declared on its `<a:endParaRPr>`. That made the auto-fit shape much
/// taller than its stored cy, pushing visible content downward and into
/// neighbouring footer shapes (slide 22 of HTL Japan deck).
test "empty paragraph uses endParaRPr font size for autofit height" {
  // Empty paragraph (no runs) with epr_font_size = 800 (= 8 pt).
  let empty_para : @ooxml.TextParagraph = {
    runs: [],
    align: "",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    mar_l: 0,
    indent: 0,
    line_spacing: 0,
    bullet: "",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 800,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  let body_props : @ooxml.BodyProps = {
    ..@ooxml.BodyProps::default(),
    auto_fit: true,
  }
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 4364608,
        cx: 4572000,
        cy: 707886, // ≈ 74 px at scale 9525
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [empty_para, empty_para, empty_para, empty_para],
    body_props,
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // The shape's bounding rect should keep its stored height (74 px), not
  // grow to the 4 × 28 = 112 px the bug produced.
  assert_true(find_substring(svg, "height=\"74\"") >= 0)
}

///|
/// Regression: bullet markers without an explicit `<a:buSzPct>` / `<a:buSzPts>`
/// used to inherit the parent `<text>` element's default 18 pt size, making
/// them visibly larger than the (smaller) run text. Per ECMA-376 they should
/// match the run's font size when no bullet size is specified.
test "bullet font-size matches run when bu_size is unspecified" {
  let run : @ooxml.TextRun = {
    text: "hello",
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size: 1200, // 12 pt → 16 px at scale 9525
    color: ({ r: 0, g: 0, b: 0, alpha: 255 } : @ooxml.Color),
    font_face: "",
    ea_font: "",
    cs_font: "",
    sym_font: "",
    underline: "",
    strike: "",
    baseline: 0,
    char_spacing: 0,
    kern: 0,
    cap: "",
    hlink_rid: "",
    hlink_mouse_over_rid: "",
    effects: @ooxml.EffectList::none(),
    outline_color: @ooxml.Color::none(),
    outline_w: 0,
    text_grad_fill: @ooxml.GradientFill::none(),
    text_patt_fill: @ooxml.PatternFill::none(),
    math_xml: "",
  }
  let para : @ooxml.TextParagraph = {
    runs: [run],
    align: "",
    level: 0,
    spc_before: 0,
    spc_after: 0,
    mar_l: 342900,
    indent: -342900,
    line_spacing: 0,
    bullet: "•",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0, // ← unspecified
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      {
        x: 0,
        y: 0,
        cx: 5000000,
        cy: 1000000,
        rot: 0,
        flip_h: false,
        flip_v: false,
      },
    ),
    paragraphs: [para],
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(
    slide,
    [],
    1,
    @ooxml.ThemeData::default(),
    [],
    [],
    [],
    [],
  )
  // The bullet tspan must explicitly set font-size to match the run
  // (run is 12 pt → 16 px at scale 9525).
  assert_true(find_substring(svg, "<tspan font-size=\"16\"") >= 0)
}
