// SVG renderer — Preset geometry evaluator
//
// Implements the DrawingML guide formula system (ECMA-376 §20.1.10)
// to convert preset shapes to SVG <path> elements.
//
// Same package as renderer.mbt — all constants and helpers are shared.

// ── Constants ────────────────────────────────────────────────────────────────

///|
/// OOXML angles are in 60000ths of a degree.
let ooxml_angle_to_rad : Double = 3.14159265358979323846 / 10800000.0

///|
/// 180° in OOXML angle units.
let half_circle_angle : Double = 10800000.0

// ── Guide variable storage ───────────────────────────────────────────────────

///|
/// Simple variable store using parallel arrays. Linear lookup is fine
///  since preset shapes have at most ~50 variables.
priv struct GuideEnv {
  names : Array[String]
  values : Array[Double]
}

///|
fn GuideEnv::new() -> GuideEnv {
  { names: [], values: [], }
}

///|
fn GuideEnv::set(self : GuideEnv, name : String, value : Double) -> Unit {
  let n = self.names.length()
  for i in 0..<n {
    if self.names[i] == name {
      self.values[i] = value
      return
    }
  }
  self.names.push(name)
  self.values.push(value)
}

///|
/// Value of an operand; an undefined guide name evaluates to 0.
fn GuideEnv::value(self : GuideEnv, op : Operand) -> Double {
  match op {
    Lit(v) => v
    Var(name) => {
      for i, n in self.names {
        if n == name {
          return self.values[i]
        }
      }
      0.0
    }
  }
}

///|
/// Value of a guide name or integer literal token (text-rect references).
fn GuideEnv::resolve(self : GuideEnv, token : String) -> Double {
  self.value(Operand::parse(token))
}

///|
/// True if `s` is an optionally negative decimal integer literal.
fn is_int_literal(s : String) -> Bool {
  let n = s.length()
  let start = if n > 0 && s.get_char(0).unwrap() == '-' { 1 } else { 0 }
  if start >= n {
    return false
  }
  for i in start..<n {
    let c = s.get_char(i).unwrap()
    if c < '0' || c > '9' {
      return false
    }
  }
  true
}

///|
/// Initialize built-in variables for a shape with given pixel dimensions.
fn GuideEnv::init_builtins(self : GuideEnv, w : Double, h : Double) -> Unit {
  // Dimensions
  self.set("w", w)
  self.set("h", h)
  // Edges
  self.set("l", 0.0)
  self.set("t", 0.0)
  self.set("r", w)
  self.set("b", h)
  self.set("hc", w / 2.0)
  self.set("vc", h / 2.0)
  // Half dimensions
  self.set("wd2", w / 2.0)
  self.set("hd2", h / 2.0)
  self.set("wd3", w / 3.0)
  self.set("hd3", h / 3.0)
  self.set("wd4", w / 4.0)
  self.set("hd4", h / 4.0)
  self.set("wd5", w / 5.0)
  self.set("hd5", h / 5.0)
  self.set("wd6", w / 6.0)
  self.set("hd6", h / 6.0)
  self.set("wd8", w / 8.0)
  self.set("hd8", h / 8.0)
  self.set("wd10", w / 10.0)
  self.set("hd10", h / 10.0)
  self.set("wd12", w / 12.0)
  self.set("hd12", h / 12.0)
  self.set("wd32", w / 32.0)
  // Short side
  let ss = if w < h { w } else { h }
  self.set("ss", ss)
  self.set("ssd2", ss / 2.0)
  self.set("ssd4", ss / 4.0)
  self.set("ssd6", ss / 6.0)
  self.set("ssd8", ss / 8.0)
  self.set("ssd16", ss / 16.0)
  self.set("ssd32", ss / 32.0)
  // Long side
  let ls = if w > h { w } else { h }
  self.set("ls", ls)
  // Angle constants (in 60000ths of degree)
  self.set("cd2", 10800000.0) // 180°
  self.set("cd4", 5400000.0) // 90°
  self.set("cd8", 2700000.0) // 45°
  // Full circle
  self.set("3cd4", 16200000.0) // 270°
  self.set("3cd8", 8100000.0) // 135°
  self.set("5cd8", 13500000.0) // 225°
  self.set("7cd8", 18900000.0) // 315°
}

// ── Formula evaluator ────────────────────────────────────────────────────────

///|
/// Parse a double from a string. Simple integer parsing since guide values are integers.
fn parse_double(s : String) -> Double {
  let n = s.length()
  if n == 0 {
    return 0.0
  }
  let mut neg = false
  let mut i = 0
  if s.get_char(0).unwrap() == '-' {
    neg = true
    i = 1
  }
  let mut result = 0.0
  while i < n {
    let c = s.get_char(i).unwrap()
    if c >= '0' && c <= '9' {
      result = result * 10.0 + (c.to_int() - 48).to_double()
    }
    i = i + 1
  }
  if neg {
    -result
  } else {
    result
  }
}

// ── Typed guide representation ───────────────────────────────────────────────

///|
/// A formula operand: an integer literal or a guide name.
priv enum Operand {
  Lit(Double)
  Var(String)
}

///|
fn Operand::parse(token : String) -> Operand {
  if is_int_literal(token) {
    Lit(parse_double(token))
  } else {
    Var(token)
  }
}

///|
/// A guide formula (ECMA-376 §20.1.9.11). Angles are in 60000ths of a degree.
priv enum Formula {
  Val(Double) // "val x"
  AddSub(Operand, Operand, Operand) // "+- a b c" = a + b - c
  AddDiv(Operand, Operand, Operand) // "+/ a b c" = (a + b) / c
  MulDiv(Operand, Operand, Operand) // "*/ a b c" = a · b / c
  IfElse(Operand, Operand, Operand) // "?: c a b" = c > 0 ? a : b
  Pin(Operand, Operand, Operand) // "pin lo x hi" = clamp
  Abs(Operand)
  Sqrt(Operand)
  Max(Operand, Operand)
  Min(Operand, Operand)
  Mod(Operand, Operand, Operand) // √(a² + b² + c²)
  Sin(Operand, Operand) // "sin x ang" = x · sin(ang)
  Cos(Operand, Operand) // "cos x ang" = x · cos(ang)
  Tan(Operand, Operand) // "tan x ang" = x · tan(ang)
  At2(Operand, Operand) // "at2 x y" = atan2(y, x), as an angle
  Cat2(Operand, Operand, Operand) // "cat2 x y z" = x · cos(atan2(z, y))
  Sat2(Operand, Operand, Operand) // "sat2 x y z" = x · sin(atan2(z, y))
  Unknown // unrecognised operator: evaluates to 0
}

///|
/// A named guide: `<a:gd name=… fmla=…/>`.
priv struct Guide {
  name : String
  fmla : Formula
}

///|
/// Parse a formula string; None for an empty one.
fn parse_formula(fmla : String) -> Formula? {
  let o = Operand::parse
  let f : Formula = match @xml.split_by_space(fmla) {
    [] => return None
    ["val", v, ..] => Val(parse_double(v))
    ["+-", a, b, c, ..] => AddSub(o(a), o(b), o(c))
    ["+/", a, b, c, ..] => AddDiv(o(a), o(b), o(c))
    ["*/", a, b, c, ..] => MulDiv(o(a), o(b), o(c))
    ["?:" | "if", c, a, b, ..] => IfElse(o(c), o(a), o(b))
    ["pin", lo, x, hi, ..] => Pin(o(lo), o(x), o(hi))
    ["abs", x, ..] => Abs(o(x))
    ["sqrt", x, ..] => Sqrt(o(x))
    ["max", a, b, ..] => Max(o(a), o(b))
    ["min", a, b, ..] => Min(o(a), o(b))
    ["mod", a, b, c, ..] => Mod(o(a), o(b), o(c))
    ["sin", x, ang, ..] => Sin(o(x), o(ang))
    ["cos", x, ang, ..] => Cos(o(x), o(ang))
    ["tan", x, ang, ..] => Tan(o(x), o(ang))
    ["at2", x, y, ..] => At2(o(x), o(y))
    ["cat2", x, y, z, ..] => Cat2(o(x), o(y), o(z))
    ["sat2", x, y, z, ..] => Sat2(o(x), o(y), o(z))
    _ => Unknown
  }
  Some(f)
}

///|
/// Parse a "name1=fmla1;name2=fmla2" guide list.
fn parse_guides(gdlst : String) -> Array[Guide] {
  let guides : Array[Guide] = []
  for entry in split_on(gdlst, ';') {
    if split_entry(entry) is Some((name, text)) &&
      parse_formula(text) is Some(fmla) {
      guides.push({ name, fmla, })
    }
  }
  guides
}

///|
/// `n / d`, or 0 when `d` is 0 (guide formulas never raise).
fn safe_div(n : Double, d : Double) -> Double {
  if d == 0.0 {
    0.0
  } else {
    n / d
  }
}

///|
fn Formula::eval(self : Formula, env : GuideEnv) -> Double {
  let v = fn(op : Operand) { env.value(op) }
  match self {
    Val(x) => x
    AddSub(a, b, c) => v(a) + v(b) - v(c)
    AddDiv(a, b, c) => safe_div(v(a) + v(b), v(c))
    MulDiv(a, b, c) => safe_div(v(a) * v(b), v(c))
    IfElse(c, a, b) => if v(c) > 0.0 { v(a) } else { v(b) }
    Pin(lo, x, hi) => {
      let (lo, x, hi) = (v(lo), v(x), v(hi))
      if x < lo {
        lo
      } else if x > hi {
        hi
      } else {
        x
      }
    }
    Abs(x) => {
      let x = v(x)
      if x < 0.0 {
        -x
      } else {
        x
      }
    }
    Sqrt(x) => {
      let x = v(x)
      if x < 0.0 {
        0.0
      } else {
        @ffi.ffi_sqrt(x)
      }
    }
    Max(a, b) => {
      let (a, b) = (v(a), v(b))
      if a > b {
        a
      } else {
        b
      }
    }
    Min(a, b) => {
      let (a, b) = (v(a), v(b))
      if a < b {
        a
      } else {
        b
      }
    }
    Mod(a, b, c) => {
      let (a, b, c) = (v(a), v(b), v(c))
      @ffi.ffi_sqrt(a * a + b * b + c * c)
    }
    Sin(x, ang) => v(x) * @ffi.ffi_sin(v(ang) * ooxml_angle_to_rad)
    Cos(x, ang) => v(x) * @ffi.ffi_cos(v(ang) * ooxml_angle_to_rad)
    Tan(x, ang) => {
      let rad = v(ang) * ooxml_angle_to_rad
      v(x) * safe_div(@ffi.ffi_sin(rad), @ffi.ffi_cos(rad))
    }
    At2(x, y) => @ffi.ffi_atan2(v(y), v(x)) / ooxml_angle_to_rad
    Cat2(x, y, z) => v(x) * @ffi.ffi_cos(@ffi.ffi_atan2(v(z), v(y)))
    Sat2(x, y, z) => v(x) * @ffi.ffi_sin(@ffi.ffi_atan2(v(z), v(y)))
    Unknown => 0.0
  }
}

///|
/// Evaluate guides in order, each seeing the ones before it.
fn eval_guides(env : GuideEnv, guides : Array[Guide]) -> Unit {
  for g in guides {
    env.set(g.name, g.fmla.eval(env))
  }
}

// ── Path rendering ───────────────────────────────────────────────────────────

///|
/// A path command (ECMA-376 §20.1.9.16). Coordinates are in the path's own
///  space until scaled to the shape.
priv enum PathCmd {
  Space(Double, Double) // this sub-path's coordinate space (a:path w/h)
  MoveTo(Operand, Operand)
  LnTo(Operand, Operand)
  CubicBezTo(Operand, Operand, Operand, Operand, Operand, Operand)
  QuadBezTo(Operand, Operand, Operand, Operand)
  ArcTo(Operand, Operand, Operand, Operand) // wR hR stAng swAng
  Close
}

///|
/// How a sub-path is filled (a:path fill, ECMA-376 §20.1.10.37). The shade
///  modes draw the shape fill darkened or lightened.
priv enum PathFill {
  Norm
  NoFill
  Darken
  DarkenLess
  Lighten
  LightenLess
} derive(Eq)

///|
fn PathFill::parse(mode : String) -> PathFill {
  match mode {
    "none" => NoFill
    "darken" => Darken
    "darkenLess" => DarkenLess
    "lighten" => Lighten
    "lightenLess" => LightenLess
    _ => Norm
  }
}

///|
/// One a:path: its commands plus how it is filled and whether it is stroked.
priv struct SubPath {
  fill : PathFill
  stroke : Bool
  cmds : Array[PathCmd]
}

///|
/// Parse one compact sub-path: "M x y", "L x y", "C x1 y1 x2 y2 x3 y3",
///  "Q x1 y1 x2 y2", "A wR hR stAng swAng", "Z", "W pw ph" for the sub-path's
///  own coordinate space, "F mode" for its fill and "S false" when it is not
///  stroked. Unrecognised tokens are skipped.
fn parse_path(cmds : String) -> SubPath {
  let o = Operand::parse
  let tokens = @xml.split_by_space(cmds)
  let out : Array[PathCmd] = []
  let mut fill = Norm
  let mut stroke = true
  let mut i = 0
  while i < tokens.length() {
    let (cmd, used) : (PathCmd?, Int) = match tokens[i:] {
      ["F", mode, ..] => {
        fill = PathFill::parse(mode)
        (None, 2)
      }
      ["S", "false", ..] => {
        stroke = false
        (None, 2)
      }
      ["W", pw, ph, ..] => (Some(Space(parse_double(pw), parse_double(ph))), 3)
      ["M", x, y, ..] => (Some(MoveTo(o(x), o(y))), 3)
      ["L", x, y, ..] => (Some(LnTo(o(x), o(y))), 3)
      ["C", x1, y1, x2, y2, x3, y3, ..] =>
        (Some(CubicBezTo(o(x1), o(y1), o(x2), o(y2), o(x3), o(y3))), 7)
      ["Q", x1, y1, x2, y2, ..] =>
        (Some(QuadBezTo(o(x1), o(y1), o(x2), o(y2))), 5)
      ["A", wr, hr, st, sw, ..] => (Some(ArcTo(o(wr), o(hr), o(st), o(sw))), 5)
      ["Z", ..] => (Some(Close), 1)
      _ => (None, 1)
    }
    if cmd is Some(c) {
      out.push(c)
    }
    i = i + used
  }
  { fill, stroke, cmds: out, }
}

///|
/// Parse "|"-separated sub-paths.
fn parse_paths(paths : String) -> Array[SubPath] {
  split_on(paths, '|').map(parse_path)
}

///|
/// Render one sub-path to SVG path data.
fn render_geom_path(
  env : GuideEnv,
  cmds : Array[PathCmd],
  x_off : Int,
  y_off : Int,
  path_w : Double,
  path_h : Double,
  shape_w : Double,
  shape_h : Double,
) -> String {
  let ox = x_off.to_double()
  let oy = y_off.to_double()
  // Scale factors if path has its own coordinate space
  let mut sx = if path_w > 0.0 { shape_w / path_w } else { 1.0 }
  let mut sy = if path_h > 0.0 { shape_h / path_h } else { 1.0 }
  // Point in shape-local pixels (before the x/y offset)
  let pt = fn(x : Operand, y : Operand) {
    (env.value(x) * sx, env.value(y) * sy)
  }
  // "x y" in slide pixels, for the SVG output
  let fmt = fn(p : (Double, Double)) { dts(p.0 + ox) + " " + dts(p.1 + oy) }
  // Current position for arcTo endpoint calculation
  let mut cur = (0.0, 0.0)
  let mut d = ""
  for cmd in cmds {
    match cmd {
      Space(pw, ph) => {
        sx = if pw > 0.0 { shape_w / pw } else { 1.0 }
        sy = if ph > 0.0 { shape_h / ph } else { 1.0 }
      }
      MoveTo(x, y) => {
        cur = pt(x, y)
        d = d + "M" + fmt(cur)
      }
      LnTo(x, y) => {
        cur = pt(x, y)
        d = d + "L" + fmt(cur)
      }
      CubicBezTo(x1, y1, x2, y2, x3, y3) => {
        let c1 = pt(x1, y1)
        let c2 = pt(x2, y2)
        cur = pt(x3, y3)
        d = d + "C" + fmt(c1) + " " + fmt(c2) + " " + fmt(cur)
      }
      QuadBezTo(x1, y1, x2, y2) => {
        let c1 = pt(x1, y1)
        cur = pt(x2, y2)
        d = d + "Q" + fmt(c1) + " " + fmt(cur)
      }
      ArcTo(wr, hr, st_ang, sw_ang) => {
        let (arc, end_x, end_y) = arcto_svg(
          cur.0,
          cur.1,
          env.value(wr) * sx,
          env.value(hr) * sy,
          env.value(st_ang),
          env.value(sw_ang),
          x_off,
          y_off,
        )
        cur = (end_x, end_y)
        d = d + arc
      }
      Close => d = d + "Z"
    }
  }
  d
}

///|
/// Convert double to string with 1 decimal place precision.
fn dts(v : Double) -> String {
  let neg = v < 0.0
  let abs_v = if neg { -v } else { v }
  let rounded = (abs_v * 10.0 + 0.5).to_int()
  let int_part = rounded / 10
  let frac = rounded % 10
  let prefix = if neg { "-" } else { "" }
  if frac == 0 {
    prefix + int_to_str(int_part)
  } else {
    prefix + int_to_str(int_part) + "." + int_to_str(frac)
  }
}

///|
/// Convert OOXML arcTo to SVG arc command.
///  Returns (svg_arc_string, end_x, end_y) in local coordinates.
fn arcto_svg(
  cur_x : Double,
  cur_y : Double,
  wr : Double,
  hr : Double,
  st_ang : Double,
  sw_ang : Double,
  x_off : Int,
  y_off : Int,
) -> (String, Double, Double) {
  if wr <= 0.0 || hr <= 0.0 {
    return ("", cur_x, cur_y)
  }
  // SVG draws nothing for an arc whose endpoints coincide (SVG 1.1 §F.6.2),
  // so a sweep over 180° (up to a full 360° ellipse) is emitted as two halves.
  let abs_sw = if sw_ang < 0.0 { -sw_ang } else { sw_ang }
  if abs_sw > half_circle_angle {
    let half = sw_ang / 2.0
    let (s1, x1, y1) = arcto_svg(
      cur_x, cur_y, wr, hr, st_ang, half, x_off, y_off,
    )
    let (s2, x2, y2) = arcto_svg(
      x1,
      y1,
      wr,
      hr,
      st_ang + half,
      half,
      x_off,
      y_off,
    )
    return (s1 + s2, x2, y2)
  }
  // The current point lies on the ellipse at stAng; the end point at stAng + swAng.
  let (sx, sy) = ellipse_point(wr, hr, st_ang)
  let (ex, ey) = ellipse_point(wr, hr, st_ang + sw_ang)
  let end_x = cur_x + ex - sx
  let end_y = cur_y + ey - sy
  // The split above keeps every emitted arc within 180°, so large-arc is 0.
  let sweep = if sw_ang > 0.0 { "1" } else { "0" }
  let ox = x_off.to_double()
  let oy = y_off.to_double()
  let s = "A" +
    dts(wr) +
    " " +
    dts(hr) +
    " 0 0 " +
    sweep +
    " " +
    dts(end_x + ox) +
    " " +
    dts(end_y + oy)
  (s, end_x, end_y)
}

///|
/// Point on an ellipse centred at the origin at OOXML angle `ang`.
///  arcTo angles are visual: the direction of the ray from the centre, not the
///  ellipse's parametric angle (ECMA-376 §20.1.9.3). The two agree only on a
///  circle or at multiples of 90°.
fn ellipse_point(wr : Double, hr : Double, ang : Double) -> (Double, Double) {
  let rad = ang * ooxml_angle_to_rad
  let c = @ffi.ffi_cos(rad)
  let s = @ffi.ffi_sin(rad)
  let hc = hr * c
  let ws = wr * s
  let d = @ffi.ffi_sqrt(hc * hc + ws * ws)
  if d == 0.0 {
    (0.0, 0.0)
  } else {
    (wr * hc / d, hr * ws / d)
  }
}

// ── Multiple path support ────────────────────────────────────────────────────

///|
/// Render all sub-paths into a single SVG path d string, ignoring their
///  per-path fill / stroke (used for a:custGeom).
fn render_geom_paths(
  env : GuideEnv,
  paths : Array[SubPath],
  x_off : Int,
  y_off : Int,
  w : Double,
  h : Double,
  path_w : Double,
  path_h : Double,
) -> String {
  let mut d = ""
  for sub in paths {
    d = d + render_geom_path(env, sub.cmds, x_off, y_off, path_w, path_h, w, h)
  }
  d
}

///|
/// Shade overlays for PathFill modes: (colour, opacity) drawn over the fill.
fn shade_overlay(mode : PathFill) -> (String, String)? {
  match mode {
    Darken => Some(("#000000", "0.4"))
    DarkenLess => Some(("#000000", "0.2"))
    Lighten => Some(("#FFFFFF", "0.4"))
    LightenLess => Some(("#FFFFFF", "0.2"))
    Norm | NoFill => None
  }
}

///|
/// SVG elements for path data `d` drawn with one sub-path style.
///  `fill` / `stroke` / `xform` are the shape's attribute strings.
fn styled_path_svg(
  d : String,
  mode : PathFill,
  stroked : Bool,
  fill : String,
  stroke : String,
  xform : String,
) -> String {
  let no_fill = a("fill", "none")
  let no_stroke = a("stroke", "none")
  let path = fn(f : String, s : String) {
    "<path d=\"" + d + "\"" + f + s + xform + "/>"
  }
  let outline = if stroked { path(no_fill, stroke) } else { "" }
  match (mode, shade_overlay(mode)) {
    (NoFill, _) => outline
    (_, None) => path(fill, if stroked { stroke } else { no_stroke })
    (_, Some((color, opacity))) => {
      let overlay = if fill == no_fill {
        ""
      } else {
        path(a("fill", color) + a("fill-opacity", opacity), no_stroke)
      }
      path(fill, no_stroke) + overlay + outline
    }
  }
}

///|
/// Render preset sub-paths as SVG elements, honouring each a:path's fill mode
///  and stroke flag. Adjacent sub-paths with the same style share one <path>,
///  so a plain filled-and-stroked preset is a single element. Stroke-only
///  (fill="none") sub-paths are drawn after the filled ones: chartPlus / chartX /
///  chartStar list their lines before the box that would otherwise cover them.
fn render_styled_paths(
  env : GuideEnv,
  paths : Array[SubPath],
  x_off : Int,
  y_off : Int,
  w : Double,
  h : Double,
  fill : String,
  stroke : String,
  xform : String,
) -> String {
  let mut filled = ""
  let mut lines = ""
  let emit = fn(d : String, mode : PathFill, stroked : Bool) {
    let svg = styled_path_svg(d, mode, stroked, fill, stroke, xform)
    if mode == NoFill {
      lines = lines + svg
    } else {
      filled = filled + svg
    }
  }
  let mut group_d = ""
  let mut group : (PathFill, Bool)? = None
  for sub in paths {
    let d = render_geom_path(env, sub.cmds, x_off, y_off, 0.0, 0.0, w, h)
    if d.is_empty() {
      continue
    }
    let style = (sub.fill, sub.stroke)
    if group == Some(style) {
      group_d = group_d + d
    } else {
      if group is Some((mode, stroked)) {
        emit(group_d, mode, stroked)
      }
      group_d = d
      group = Some(style)
    }
  }
  if group is Some((mode, stroked)) {
    emit(group_d, mode, stroked)
  }
  filled + lines
}

// ── Parsed preset cache ──────────────────────────────────────────────────────

///|
/// A preset definition with its guide list and paths parsed once.
priv struct ParsedGeom {
  adjs : Array[(String, Double)] // avLst defaults, in declaration order
  guides : Array[Guide]
  paths : Array[SubPath]
  rect : Array[Operand] // text rectangle l t r b; empty = the whole shape
}

///|
/// A preset definition as generated from presetShapeDefinitions.xml
///  (formats: scripts/gen_preset_geom.py).
priv struct PresetDef {
  avlst : String
  gdlst : String
  paths : String
  rect : String
}

///|
/// Parsed definitions by preset name; None records an unknown preset.
let preset_geom_cache : Map[String, ParsedGeom?] = Map([])

///|
/// Look up a preset and parse its definition on first use.
fn preset_geom(prst : String) -> ParsedGeom? {
  if preset_geom_cache.get(prst) is Some(cached) {
    return cached
  }
  let parsed = match get_preset_def(prst) {
    None => None
    Some(def) =>
      Some({
        adjs: parse_adjustments(def.avlst),
        guides: parse_guides(def.gdlst),
        paths: parse_paths(def.paths),
        rect: @xml.split_by_space(def.rect).map(Operand::parse),
      })
  }
  preset_geom_cache[prst] = parsed
  parsed
}

// ── Main entry point ─────────────────────────────────────────────────────────

///|
/// Evaluate a preset geometry and return the SVG element string.
///  Returns the full <path> element with fill, stroke, and transform attributes.
///  If the preset is not implemented, returns a fallback rectangle.
fn render_preset_geom(
  prst : String,
  adj_values : Array[Int],
  cx_p : Int,
  cy_p : Int,
  x_p : Int,
  y_p : Int,
  fill : String,
  stroke : String,
  rot : String,
) -> String {
  // An unknown preset, or one whose path evaluates to nothing, falls back to a rectangle.
  let svg = match preset_geom(prst) {
    None => ""
    Some(geom) => {
      let w = cx_p.to_double()
      let h = cy_p.to_double()
      let env = GuideEnv::new()
      env.init_builtins(w, h)
      apply_adjustments(env, geom.adjs, adj_values)
      eval_guides(env, geom.guides)
      render_styled_paths(env, geom.paths, x_p, y_p, w, h, fill, stroke, rot)
    }
  }
  if svg.is_empty() {
    "<rect" +
    ai("x", x_p) +
    ai("y", y_p) +
    ai("width", cx_p) +
    ai("height", cy_p) +
    fill +
    stroke +
    rot +
    "/>"
  } else {
    svg
  }
}

///|
/// Render custom geometry (a:custGeom) using the guide formula engine.
fn render_custom_geom(
  data : @ooxml.CustomGeomData,
  cx_p : Int,
  cy_p : Int,
  x_p : Int,
  y_p : Int,
  fill : String,
  stroke : String,
  rot : String,
) -> String {
  let w = cx_p.to_double()
  let h = cy_p.to_double()
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  // Evaluate guide list (includes both avLst and gdLst)
  eval_guides(env, parse_guides(data.gdlst))
  // Render paths with coordinate space scaling
  let pw = data.path_w.to_double()
  let ph = data.path_h.to_double()
  let d = render_geom_paths(
    env,
    parse_paths(data.paths),
    x_p,
    y_p,
    w,
    h,
    pw,
    ph,
  )
  if d.is_empty() {
    // Empty custGeom path — no visible shape to render.
    // Returning a filled rect here produces an unwanted colored block
    // (e.g. a solid blue square where an icon should be).
    // Better to hide the shape entirely so it doesn't obstruct the slide.
    ""
  } else {
    "<path d=\"" + d + "\"" + fill + stroke + rot + "/>"
  }
}

///|
/// Compute text rectangle for a preset geometry shape.
/// Returns (l, t, w, h) in pixels, where l/t are offsets from shape origin.
/// If the shape has no custom text rect, returns (0, 0, cx_p, cy_p).
fn compute_preset_text_rect(
  prst : String,
  adj_values : Array[Int],
  cx_p : Int,
  cy_p : Int,
) -> (Int, Int, Int, Int) {
  guard preset_geom(prst) is Some({ rect: [rl, rt, rr, rb], adjs, guides, .. }) else {
    return (0, 0, cx_p, cy_p)
  }
  let w = cx_p.to_double()
  let h = cy_p.to_double()
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  apply_adjustments(env, adjs, adj_values)
  eval_guides(env, guides)
  let rl = env.value(rl)
  let rt = env.value(rt)
  let rr = env.value(rr)
  let rb = env.value(rb)
  // Clamp to shape bounds
  let l = if rl < 0.0 { 0 } else { rl.to_int() }
  let t = if rt < 0.0 { 0 } else { rt.to_int() }
  let r_val = if rr > w { cx_p } else { rr.to_int() }
  let b_val = if rb > h { cy_p } else { rb.to_int() }
  (l, t, r_val - l, b_val - t)
}

///|
/// Compute text rectangle for custom geometry.
/// Uses guide formulas and the a:rect definition from the custom geom data.
fn compute_custom_text_rect(
  data : @ooxml.CustomGeomData,
  cx_p : Int,
  cy_p : Int,
) -> (Int, Int, Int, Int) {
  if data.rect_l.is_empty() &&
    data.rect_t.is_empty() &&
    data.rect_r.is_empty() &&
    data.rect_b.is_empty() {
    return (0, 0, cx_p, cy_p)
  }
  let w = cx_p.to_double()
  let h = cy_p.to_double()
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  env.set("3wd4", w * 3.0 / 4.0)
  env.set("3hd4", h * 3.0 / 4.0)
  eval_guides(env, parse_guides(data.gdlst))
  let rl = if data.rect_l.is_empty() { 0.0 } else { env.resolve(data.rect_l) }
  let rt = if data.rect_t.is_empty() { 0.0 } else { env.resolve(data.rect_t) }
  let rr = if data.rect_r.is_empty() { w } else { env.resolve(data.rect_r) }
  let rb = if data.rect_b.is_empty() { h } else { env.resolve(data.rect_b) }
  let l = if rl < 0.0 { 0 } else { rl.to_int() }
  let t = if rt < 0.0 { 0 } else { rt.to_int() }
  let r_val = if rr > w { cx_p } else { rr.to_int() }
  let b_val = if rb > h { cy_p } else { rb.to_int() }
  (l, t, r_val - l, b_val - t)
}

///|
/// Parse "name1=val1;name2=val2" adjustment defaults.
fn parse_adjustments(avlst : String) -> Array[(String, Double)] {
  let adjs : Array[(String, Double)] = []
  for entry in split_on(avlst, ';') {
    if split_entry(entry) is Some((name, v)) {
      adjs.push((name, parse_double(v)))
    }
  }
  adjs
}

///|
/// Set adjustment defaults, then override them by position with the shape's
///  own `a:avLst` values.
fn apply_adjustments(
  env : GuideEnv,
  adjs : Array[(String, Double)],
  adj_values : Array[Int],
) -> Unit {
  for adj in adjs {
    env.set(adj.0, adj.1)
  }
  for j, v in adj_values {
    if j < adjs.length() {
      env.set(adjs[j].0, v.to_double())
    }
  }
}

///|
/// Split `s` on `sep`, dropping empty segments.
fn split_on(s : String, sep : Char) -> Array[String] {
  let parts : Array[String] = []
  let len = s.length()
  let mut start = 0
  for i in 0..<=len {
    if i == len || s.get_char(i).unwrap() == sep {
      if i > start {
        parts.push(@xml.str_substring(s, start, i))
      }
      start = i + 1
    }
  }
  parts
}

///|
/// Split a "name=value" entry at its first '='; None if there is none.
fn split_entry(s : String) -> (String, String)? {
  let len = s.length()
  for i in 0..<len {
    if s.get_char(i).unwrap() == '=' {
      return Some(
        (@xml.str_substring(s, 0, i), @xml.str_substring(s, i + 1, len)),
      )
    }
  }
  None
}
