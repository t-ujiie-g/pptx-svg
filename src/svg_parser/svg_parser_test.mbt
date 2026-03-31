// ── parse_svg_to_slide ───────────────────────────────────────────────────────

///|
test "parse_svg_to_slide empty svg" {
  let slide = parse_svg_to_slide("<svg></svg>")
  assert_eq(slide.shapes.length(), 0)
  assert_false(slide.hidden)
}

///|
test "parse_svg_to_slide no svg element" {
  let slide = parse_svg_to_slide("<div>not svg</div>")
  assert_eq(slide.shapes.length(), 0)
}

///|
test "parse_svg_to_slide slide metadata" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\" data-ooxml-bg=\"FF0000\" data-ooxml-hidden=\"1\"></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_eq(slide.slide_size.cx, 9144000)
  assert_eq(slide.slide_size.cy, 6858000)
  assert_eq(slide.background.r, 255)
  assert_eq(slide.background.g, 0)
  assert_eq(slide.background.b, 0)
  assert_true(slide.hidden)
}

///|
test "parse_svg_to_slide default size" {
  let svg = "<svg></svg>"
  let slide = parse_svg_to_slide(svg)
  // Default slide size
  assert_eq(slide.slide_size.cx, 9144000)
  assert_eq(slide.slide_size.cy, 6858000)
}

///|
test "parse_svg_to_slide basic shape" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"100000\" data-ooxml-y=\"200000\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_eq(slide.shapes.length(), 1)
  assert_eq(slide.shapes[0].transform.x, 100000)
  assert_eq(slide.shapes[0].transform.y, 200000)
  assert_eq(slide.shapes[0].transform.cx, 500000)
  assert_eq(slide.shapes[0].transform.cy, 300000)
}

///|
test "parse_svg_to_slide shape with blur effect" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-eff-blur-rad=\"76200\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_eq(slide.shapes.length(), 1)
  assert_eq(slide.shapes[0].effects.blur.rad, 76200)
  assert_true(!(slide.shapes[0].effects.blur.is_none()))
}

///|
test "parse_svg_to_slide shape with outer shadow" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-eff-os-blur=\"152400\" data-ooxml-eff-os-dist=\"114300\" data-ooxml-eff-os-dir=\"5400000\" data-ooxml-eff-os-clr=\"000000\" data-ooxml-eff-os-algn=\"ctr\" data-ooxml-eff-os-rot=\"0\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  let os = slide.shapes[0].effects.outer_shadow
  assert_true(!(os.is_none()))
  assert_eq(os.blur_rad, 152400)
  assert_eq(os.dist, 114300)
  assert_eq(os.dir, 5400000)
  assert_eq(os.algn, "ctr")
}

///|
test "parse_svg_to_slide shape with glow" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-eff-gl-rad=\"228600\" data-ooxml-eff-gl-clr=\"FFC000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  let gl = slide.shapes[0].effects.glow
  assert_true(!(gl.is_none()))
  assert_eq(gl.rad, 228600)
  assert_eq(gl.color.r, 0xFF)
  assert_eq(gl.color.g, 0xC0)
  assert_eq(gl.color.b, 0x00)
}

///|
test "parse_svg_to_slide shape with fill and stroke" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-fill=\"FF6633\" data-ooxml-stroke=\"0000FF\" data-ooxml-stroke-w=\"25400\" data-ooxml-stroke-dash=\"dash\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_eq(slide.shapes[0].fill.r, 0xFF)
  assert_eq(slide.shapes[0].fill.g, 0x66)
  assert_eq(slide.shapes[0].fill.b, 0x33)
  assert_eq(slide.shapes[0].stroke.r, 0)
  assert_eq(slide.shapes[0].stroke.g, 0)
  assert_eq(slide.shapes[0].stroke.b, 0xFF)
  assert_eq(slide.shapes[0].stroke_w, 25400)
  assert_eq(slide.shapes[0].stroke_dash, "dash")
}

///|
test "parse_svg_to_slide shape with preset shadow" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-eff-ps-prst=\"shdw1\" data-ooxml-eff-ps-dist=\"38100\" data-ooxml-eff-ps-dir=\"2700000\" data-ooxml-eff-ps-clr=\"000000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  let ps = slide.shapes[0].effects.prst_shadow
  assert_true(!(ps.is_none()))
  assert_eq(ps.prst, "shdw1")
  assert_eq(ps.dist, 38100)
  assert_eq(ps.dir, 2700000)
}

///|
test "parse_svg_to_slide text run effects round-trip" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\"><rect/><text><tspan><tspan data-ooxml-bold=\"1\" data-ooxml-font-size=\"1800\" data-ooxml-reff-os-blur=\"50800\" data-ooxml-reff-os-dist=\"38100\" data-ooxml-reff-os-dir=\"2700000\" data-ooxml-reff-os-clr=\"000000\" data-ooxml-reff-blur-rad=\"76200\" data-ooxml-reff-ps-prst=\"shdw3\" data-ooxml-reff-ps-dist=\"25400\" data-ooxml-reff-ps-dir=\"5400000\" data-ooxml-reff-ps-clr=\"333333\">Hello</tspan></tspan></text></g></svg>"
  let slide = parse_svg_to_slide(svg)
  let run = slide.shapes[0].paragraphs[0].runs[0]
  // Outer shadow
  assert_true(!(run.effects.outer_shadow.is_none()))
  assert_eq(run.effects.outer_shadow.blur_rad, 50800)
  assert_eq(run.effects.outer_shadow.dist, 38100)
  // Blur
  assert_eq(run.effects.blur.rad, 76200)
  // Preset shadow
  assert_eq(run.effects.prst_shadow.prst, "shdw3")
  assert_eq(run.effects.prst_shadow.dist, 25400)
}

///|
test "parse_svg_to_slide fill overlay" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\" data-ooxml-eff-fo-blend=\"mult\" data-ooxml-eff-fo-clr=\"80FF0000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  let fo = slide.shapes[0].effects.fill_overlay
  assert_true(!(fo.is_none()))
  assert_eq(fo.blend, "mult")
  assert_eq(fo.color.r, 255)
  assert_eq(fo.color.g, 0)
  assert_eq(fo.color.b, 0)
}

///|
test "parse_svg_to_slide no effects when absent" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"0\" data-ooxml-y=\"0\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_true(slide.shapes[0].effects.is_none())
  assert_true(slide.shapes[0].effects.blur.is_none())
  assert_true(slide.shapes[0].effects.outer_shadow.is_none())
  assert_true(slide.shapes[0].effects.glow.is_none())
}
