// ── render_slide_svg basic ───────────────────────────────────────────────────

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

test "render_slide_svg empty slide" {
  let slide = make_empty_slide()
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
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

test "render_slide_svg hidden slide" {
  let slide : @ooxml.SlideData = { ..make_empty_slide(), hidden: true }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let nodes = @xml.parse_xml(svg)
  assert_eq(@xml.get_attr(nodes[0], "data-ooxml-hidden"), Some("1"))
}

test "render_slide_svg with rect shape" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 914400, y: 914400, cx: 2743200, cy: 1828800, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  // Should contain a group for the shape
  let nodes = @xml.parse_xml(svg)
  let svg_children = @xml.get_children(nodes[0])
  // Find a <g> with data-ooxml-shape-type
  let mut found_shape = false
  for i = 0; i < svg_children.length(); i = i + 1 {
    if @xml.get_tag(svg_children[i]) == "g" {
      let st = @xml.get_attr(svg_children[i], "data-ooxml-shape-type")
      if not(st is None) {
        found_shape = true
      }
    }
  }
  assert_true(found_shape)
}

test "render_slide_svg with ellipse" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    { x: 0, y: 0, cx: 1000000, cy: 1000000, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
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

test "render_slide_svg colored background" {
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    background: @ooxml.parse_hex_color("FF0000"),
  }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  assert_eq(
    @xml.get_attr(@xml.parse_xml(svg)[0], "data-ooxml-bg"),
    Some("FF0000"),
  )
}

// ── Round-trip: render → parse → compare ────────────────────────────────────

test "roundtrip empty slide" {
  let slide = make_empty_slide()
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.slide_size.cx, slide.slide_size.cx)
  assert_eq(parsed.slide_size.cy, slide.slide_size.cy)
  assert_eq(parsed.background.r, slide.background.r)
  assert_eq(parsed.background.g, slide.background.g)
  assert_eq(parsed.background.b, slide.background.b)
  assert_eq(parsed.hidden, false)
}

test "roundtrip hidden slide" {
  let slide : @ooxml.SlideData = { ..make_empty_slide(), hidden: true }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_true(parsed.hidden)
}

test "roundtrip colored background" {
  let slide : @ooxml.SlideData = {
    ..make_empty_slide(),
    background: @ooxml.parse_hex_color("3366CC"),
  }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.background.r, 0x33)
  assert_eq(parsed.background.g, 0x66)
  assert_eq(parsed.background.b, 0xCC)
}

test "roundtrip shape transform" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 914400, y: 457200, cx: 2743200, cy: 1828800, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  assert_eq(parsed.shapes[0].transform.x, 914400)
  assert_eq(parsed.shapes[0].transform.y, 457200)
  assert_eq(parsed.shapes[0].transform.cx, 2743200)
  assert_eq(parsed.shapes[0].transform.cy, 1828800)
}

test "roundtrip shape with blur effect" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 100000, y: 200000, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    effects: { ..@ooxml.EffectList::none(), blur: { rad: 76200 } },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  assert_eq(parsed.shapes[0].effects.blur.rad, 76200)
  assert_true(not(parsed.shapes[0].effects.blur.is_none()))
}

test "roundtrip shape with outer shadow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      outer_shadow: {
        blur_rad: 152400, dist: 114300, dir: 5400000,
        color: @ooxml.parse_hex_color("000000"),
        sx: 0, sy: 0, algn: "ctr", rot_with_shape: false,
      },
    },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.outer_shadow.blur_rad, 152400)
  assert_eq(parsed.shapes[0].effects.outer_shadow.dist, 114300)
  assert_eq(parsed.shapes[0].effects.outer_shadow.dir, 5400000)
  assert_eq(parsed.shapes[0].effects.outer_shadow.algn, "ctr")
}

test "roundtrip shape with glow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      glow: { rad: 228600, color: @ooxml.parse_hex_color("FFC000") },
    },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.glow.rad, 228600)
  assert_eq(parsed.shapes[0].effects.glow.color.r, 0xFF)
  assert_eq(parsed.shapes[0].effects.glow.color.g, 0xC0)
}

test "roundtrip shape with soft edge" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    effects: { ..@ooxml.EffectList::none(), soft_edge: { rad: 63500 } },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.soft_edge.rad, 63500)
}

test "roundtrip shape with solid fill" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    fill: @ooxml.parse_hex_color("FF6633"),
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].fill.r, 0xFF)
  assert_eq(parsed.shapes[0].fill.g, 0x66)
  assert_eq(parsed.shapes[0].fill.b, 0x33)
}

test "roundtrip shape with stroke" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    stroke: @ooxml.parse_hex_color("0000FF"),
    stroke_w: 25400,
    stroke_dash: "dash",
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].stroke.r, 0)
  assert_eq(parsed.shapes[0].stroke.g, 0)
  assert_eq(parsed.shapes[0].stroke.b, 0xFF)
  assert_eq(parsed.shapes[0].stroke_w, 25400)
  assert_eq(parsed.shapes[0].stroke_dash, "dash")
}

test "roundtrip ellipse geometry" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    { x: 100000, y: 100000, cx: 400000, cy: 400000, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 1)
  match parsed.shapes[0].kind {
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse) => ()
    _ => assert_true(false)
  }
}

test "roundtrip shape with preset shadow" {
  let shape : @ooxml.Shape = {
    ..@ooxml.Shape::default_with(
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
      { x: 0, y: 0, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
    ),
    effects: {
      ..@ooxml.EffectList::none(),
      prst_shadow: {
        prst: "shdw1", dist: 38100, dir: 2700000,
        color: @ooxml.parse_hex_color("000000"),
      },
    },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.prst_shadow.prst, "shdw1")
  assert_eq(parsed.shapes[0].effects.prst_shadow.dist, 38100)
  assert_eq(parsed.shapes[0].effects.prst_shadow.dir, 2700000)
}

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
      fill_overlay: {
        blend: "mult",
        color: @ooxml.parse_hex_color("80FF0000"),
      },
    },
  }
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes[0].effects.fill_overlay.blend, "mult")
  assert_eq(parsed.shapes[0].effects.fill_overlay.color.r, 255)
  assert_eq(parsed.shapes[0].effects.fill_overlay.color.g, 0)
}

test "roundtrip multiple shapes" {
  let shape1 = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 0, y: 0, cx: 200000, cy: 200000, rot: 0, flip_h: false, flip_v: false },
  )
  let shape2 = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Ellipse),
    { x: 300000, y: 300000, cx: 200000, cy: 200000, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = { ..make_empty_slide(), shapes: [shape1, shape2] }
  let svg = render_slide_svg(slide, [], 1, @ooxml.ThemeData::default(), [], [], [], [])
  let parsed = @svg_parser.parse_svg_to_slide(svg)
  assert_eq(parsed.shapes.length(), 2)
  assert_eq(parsed.shapes[0].transform.x, 0)
  assert_eq(parsed.shapes[1].transform.x, 300000)
}
