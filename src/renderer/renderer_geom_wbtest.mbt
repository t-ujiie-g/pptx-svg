///|
using @test {assert_eq}

///|
fn geom_env(w : Double, h : Double) -> GuideEnv {
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  env
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
  let d = render_geom_paths(
    env, "W 21600 21600 M 0 0 L 21600 21600", 0, 0, 200.0, 100.0, 0.0, 0.0,
  )
  assert_eq(d, "M0 0L200 100")
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
