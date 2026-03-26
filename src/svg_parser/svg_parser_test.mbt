// ── parse_svg_to_slide ───────────────────────────────────────────────────────

test "parse_svg_to_slide empty svg" {
  let slide = parse_svg_to_slide("<svg></svg>")
  assert_eq(slide.shapes.length(), 0)
  assert_false(slide.hidden)
}

test "parse_svg_to_slide no svg element" {
  let slide = parse_svg_to_slide("<div>not svg</div>")
  assert_eq(slide.shapes.length(), 0)
}

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

test "parse_svg_to_slide default size" {
  let svg = "<svg></svg>"
  let slide = parse_svg_to_slide(svg)
  // Default slide size
  assert_eq(slide.slide_size.cx, 9144000)
  assert_eq(slide.slide_size.cy, 6858000)
}

test "parse_svg_to_slide basic shape" {
  let svg = "<svg data-ooxml-slide-cx=\"9144000\" data-ooxml-slide-cy=\"6858000\"><g data-ooxml-shape-type=\"auto\" data-ooxml-prst=\"rect\" data-ooxml-x=\"100000\" data-ooxml-y=\"200000\" data-ooxml-cx=\"500000\" data-ooxml-cy=\"300000\"><rect/></g></svg>"
  let slide = parse_svg_to_slide(svg)
  assert_eq(slide.shapes.length(), 1)
  assert_eq(slide.shapes[0].transform.x, 100000)
  assert_eq(slide.shapes[0].transform.y, 200000)
  assert_eq(slide.shapes[0].transform.cx, 500000)
  assert_eq(slide.shapes[0].transform.cy, 300000)
}
