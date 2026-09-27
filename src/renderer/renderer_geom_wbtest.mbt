///|
using @test {assert_eq}

///|
fn geom_env(w : Double, h : Double) -> GuideEnv {
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  env
}

///|
/// Parse and evaluate one formula into guide `name`.
fn eval_formula(env : GuideEnv, name : String, fmla : String) -> Unit {
  eval_guides(env, parse_guides(name + "=" + fmla))
}

///|
fn near(a : Double, b : Double) -> Bool {
  let d = a - b
  d < 0.001 && d > -0.001
}

///|
test "guide builtins include hc/vc/wd32" {
  let env = geom_env(200.0, 100.0)
  assert_eq(env.resolve("hc"), 100.0)
  assert_eq(env.resolve("vc"), 50.0)
  assert_eq(env.resolve("wd32"), 6.25)
}

///|
test "undefined guide name resolves to 0, not its digits" {
  let env = geom_env(200.0, 100.0)
  assert_eq(env.resolve("x1"), 0.0)
  assert_eq(env.resolve("-42"), -42.0)
}

///|
test "trig guides take (value, angle) per ECMA-376" {
  let env = geom_env(200.0, 100.0)
  eval_formula(env, "c", "cos 100 5400000")
  eval_formula(env, "s", "sin 100 5400000")
  eval_formula(env, "t", "tan 100 2700000")
  assert_true(near(env.resolve("c"), 0.0))
  assert_true(near(env.resolve("s"), 100.0))
  assert_true(near(env.resolve("t"), 100.0))
}

///|
test "+/ guide is (a + b) / c" {
  let env = geom_env(200.0, 100.0)
  eval_formula(env, "v", "+/ 10 20 3")
  assert_eq(env.resolve("v"), 10.0)
}

///|
test "?: guide takes the then-branch only for a positive condition" {
  let env = geom_env(200.0, 100.0)
  eval_formula(env, "p", "?: 1 7 2")
  eval_formula(env, "z", "?: 0 7 2")
  eval_formula(env, "n", "?: -5 7 2")
  assert_eq(env.resolve("p"), 7.0)
  assert_eq(env.resolve("z"), 2.0)
  assert_eq(env.resolve("n"), 2.0)
}

///|
test "full-circle arcTo is split into two drawable half arcs" {
  let (d, ex, ey) = arcto_svg(
    0.0, 50.0, 100.0, 50.0, 10800000.0, 21600000.0, 0, 0,
  )
  assert_eq(d, "A100 50 0 0 1 200 50A100 50 0 0 1 0 50")
  assert_true(near(ex, 0.0))
  assert_true(near(ey, 50.0))
}

///|
test "arcTo angles are visual angles on an ellipse" {
  // At a visual 45°, the point lies on the diagonal ray: x == y.
  let (x, y) = ellipse_point(200.0, 100.0, 2700000.0)
  assert_true(near(x, y))
  assert_true(near(x * x / 40000.0 + y * y / 10000.0, 1.0))
}

///|
test "W token scales a path from its own coordinate space" {
  let env = geom_env(200.0, 100.0)
  let svg = render_styled_paths(
    env,
    parse_paths("W 21600 21600 M 0 0 L 21600 21600"),
    0,
    0,
    200.0,
    100.0,
    0.0,
    0.0,
    "",
    "",
    "",
  )
  assert_eq(svg, "<path d=\"M0 0L200 100\"/>")
}

///|
fn preset_path(prst : String, adj : Array[Int], w : Int, h : Int) -> String {
  render_preset_geom(prst, adj, w, h, 0, 0, "", "", "")
}

///|
test "donut renders a ring, not zero-length arcs" {
  let svg = preset_path("donut", [], 200, 100)
  assert_true(svg.contains("A100 50 0 0 1 200 50"))
  assert_true(!svg.contains("A100 50 0 1 1 0 50"))
}

///|
test "wedgeRectCallout attaches the tail to the edge nearest the tip" {
  // Tip above the box: the wedge leaves the top edge (y=0), so the tip vertex
  // follows a top-edge vertex.
  let above = preset_path("wedgeRectCallout", [30000, -80000], 180, 129)
  assert_true(above.contains("L105 0L144 -38.7L150 0"))
  // Tip left of the box: the wedge leaves the left edge (x=0).
  let left = preset_path("wedgeRectCallout", [-80000, 0], 180, 129)
  assert_true(left.contains("L0 53.8L-54 64.5L0 21.5"))
}

///|
test "formulas parse into typed operators with literal / guide operands" {
  guard parse_formula("*/ w adj 100000")
    is Some(MulDiv(Var("w"), Var("adj"), Lit(100000.0))) else {
    fail("*/ did not parse to MulDiv")
  }
  assert_true(parse_formula("") is None)
  assert_true(parse_formula("bogus 1 2") is Some(Unknown))
  // Too few operands is not a valid "+-"
  assert_true(parse_formula("+- a b") is Some(Unknown))
}

///|
test "path parsing skips a command with missing operands" {
  // "L 5" lacks its y operand: both tokens are skipped.
  let cmds = parse_path("M 0 0 L 5").cmds
  assert_eq(cmds.length(), 1)
  assert_true(cmds[0] is MoveTo(Lit(_), Lit(_)))
}

///|
test "preset definitions are parsed once and cached" {
  guard preset_geom("donut") is Some(first) else { fail("donut missing") }
  guard preset_geom("donut") is Some(second) else { fail("donut missing") }
  assert_true(physical_equal(first, second))
  assert_true(preset_geom("noSuchPreset") is None)
}

///|
fn styled_preset(prst : String) -> String {
  render_preset_geom(
    prst,
    [],
    200,
    100,
    0,
    0,
    " fill=\"#123456\"",
    " stroke=\"#000000\"",
    "",
  )
}

///|
fn count_paths(svg : String) -> Int {
  let mut n = 0
  let mut rest = svg
  while rest.find("<path ") is Some(i) {
    n = n + 1
    rest = @xml.str_substring(rest, i + 6, rest.length())
  }
  n
}

///|
test "plain preset stays a single filled-and-stroked path" {
  let svg = styled_preset("star5")
  assert_eq(count_paths(svg), 1)
  assert_true(svg.contains(" fill=\"#123456\" stroke=\"#000000\""))
}

///|
test "arc: filled wedge without stroke, then the stroked arc without fill" {
  let svg = styled_preset("arc")
  assert_eq(count_paths(svg), 2)
  assert_true(svg.contains("fill=\"#123456\" stroke=\"none\""))
  assert_true(svg.contains("fill=\"none\" stroke=\"#000000\""))
}

///|
test "connector presets are stroke-only" {
  let svg = styled_preset("bentConnector3")
  assert_eq(count_paths(svg), 1)
  assert_true(svg.contains("fill=\"none\" stroke=\"#000000\""))
}

///|
test "cube faces are shaded with overlays" {
  let svg = styled_preset("cube")
  assert_true(svg.contains("fill=\"#000000\" fill-opacity=\"0.2\""))
  assert_true(svg.contains("fill=\"#FFFFFF\" fill-opacity=\"0.2\""))
}

///|
test "shade overlays are skipped when the shape has no fill" {
  let svg = render_preset_geom(
    "cube",
    [],
    200,
    100,
    0,
    0,
    " fill=\"none\"",
    " stroke=\"#000000\"",
    "",
  )
  assert_true(!svg.contains("fill-opacity"))
}

///|
test "stroke-only sub-paths are drawn after the fills (chartPlus)" {
  let svg = styled_preset("chartPlus")
  let fill_at = svg.find("fill=\"#123456\"").unwrap()
  let line_at = svg.find("fill=\"none\"").unwrap()
  assert_true(line_at > fill_at)
}

///|
test "text rectangle comes from the spec a:rect" {
  // triangle: rect l=x1 t=vc r=x3 b=b → the lower half, inset horizontally
  assert_eq(
    compute_preset_text_rect("triangle", [], 200, 100),
    (50, 50, 100, 50),
  )
  // plaque: inset by x1·0.70711 on every side (x1 = ss·16.667%)
  assert_eq(compute_preset_text_rect("plaque", [], 200, 100), (11, 11, 177, 77))
  assert_eq(
    compute_preset_text_rect("noSuchPreset", [], 200, 100),
    (0, 0, 200, 100),
  )
}

///|
test "custGeom honours per-path fill modes and its own coordinate spaces" {
  let data : @ooxml.CustomGeomData = {
    gdlst: "",
    paths: "S false M 0 0 L 100 100 Z|W 50 50 F none M 0 50 L 50 0",
    path_w: 100,
    path_h: 100,
    rect_l: "",
    rect_t: "",
    rect_r: "",
    rect_b: "",
    cxn_lst: "",
  }
  let svg = render_custom_geom(
    data, 200, 100, 0, 0, " fill=\"#123456\"", " stroke=\"#000000\"", "",
  )
  assert_eq(
    svg,
    "<path d=\"M0 0L200 100Z\" fill=\"#123456\" stroke=\"none\"/>" +
    "<path d=\"M0 100L200 0\" fill=\"none\" stroke=\"#000000\"/>",
  )
}
