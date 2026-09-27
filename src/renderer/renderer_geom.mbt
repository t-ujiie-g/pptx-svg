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
fn GuideEnv::resolve(self : GuideEnv, token : String) -> Double {
  // First try variable lookup
  let n = self.names.length()
  for i in 0..<n {
    if self.names[i] == token {
      return self.values[i]
    }
  }
  // Then a literal integer; an undefined guide name evaluates to 0
  if is_int_literal(token) {
    parse_double(token)
  } else {
    0.0
  }
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

///|
/// Evaluate a single guide formula and store the result.
fn eval_formula(env : GuideEnv, name : String, fmla : String) -> Unit {
  let tokens = @xml.split_by_space(fmla)
  let nt = tokens.length()
  if nt == 0 {
    return
  }
  let op = tokens[0]
  let val = if op == "val" && nt >= 2 {
    parse_double(tokens[1])
  } else if op == "+-" && nt >= 4 {
    // a + b - c
    env.resolve(tokens[1]) + env.resolve(tokens[2]) - env.resolve(tokens[3])
  } else if op == "+/" && nt >= 4 {
    // (a + b) / c
    let c = env.resolve(tokens[3])
    if c == 0.0 {
      0.0
    } else {
      (env.resolve(tokens[1]) + env.resolve(tokens[2])) / c
    }
  } else if op == "*/" && nt >= 4 {
    // a * b / c
    let c = env.resolve(tokens[3])
    if c == 0.0 {
      0.0
    } else {
      env.resolve(tokens[1]) * env.resolve(tokens[2]) / c
    }
  } else if (op == "?:" || op == "if") && nt >= 4 {
    // if a > 0 then b else c
    if env.resolve(tokens[1]) > 0.0 {
      env.resolve(tokens[2])
    } else {
      env.resolve(tokens[3])
    }
  } else if op == "pin" && nt >= 4 {
    // clamp(val, min, max)
    let min_v = env.resolve(tokens[1])
    let val_v = env.resolve(tokens[2])
    let max_v = env.resolve(tokens[3])
    if val_v < min_v {
      min_v
    } else if val_v > max_v {
      max_v
    } else {
      val_v
    }
  } else if op == "abs" && nt >= 2 {
    let v = env.resolve(tokens[1])
    if v < 0.0 {
      -v
    } else {
      v
    }
  } else if op == "sqrt" && nt >= 2 {
    let v = env.resolve(tokens[1])
    if v < 0.0 {
      0.0
    } else {
      @ffi.ffi_sqrt(v)
    }
  } else if op == "max" && nt >= 3 {
    let a = env.resolve(tokens[1])
    let b = env.resolve(tokens[2])
    if a > b {
      a
    } else {
      b
    }
  } else if op == "min" && nt >= 3 {
    let a = env.resolve(tokens[1])
    let b = env.resolve(tokens[2])
    if a < b {
      a
    } else {
      b
    }
  } else if op == "mod" && nt >= 4 {
    // sqrt(a^2 + b^2 + c^2)
    let a = env.resolve(tokens[1])
    let b = env.resolve(tokens[2])
    let c = env.resolve(tokens[3])
    @ffi.ffi_sqrt(a * a + b * b + c * c)
  } else if op == "sin" && nt >= 3 {
    // "sin x y" = x * sin(y) — y in 60000ths of a degree (ECMA-376 §20.1.9.11)
    let dist = env.resolve(tokens[1])
    let ang = env.resolve(tokens[2]) * ooxml_angle_to_rad
    @ffi.ffi_sin(ang) * dist
  } else if op == "cos" && nt >= 3 {
    // "cos x y" = x * cos(y)
    let dist = env.resolve(tokens[1])
    let ang = env.resolve(tokens[2]) * ooxml_angle_to_rad
    @ffi.ffi_cos(ang) * dist
  } else if op == "tan" && nt >= 3 {
    // "tan x y" = x * tan(y)
    let dist = env.resolve(tokens[1])
    let ang = env.resolve(tokens[2]) * ooxml_angle_to_rad
    let cos_v = @ffi.ffi_cos(ang)
    if cos_v == 0.0 {
      0.0
    } else {
      @ffi.ffi_sin(ang) / cos_v * dist
    }
  } else if op == "at2" && nt >= 3 {
    // atan2(y, x) → angle in 60000ths of degree
    let x = env.resolve(tokens[1])
    let y = env.resolve(tokens[2])
    @ffi.ffi_atan2(y, x) / ooxml_angle_to_rad
  } else if op == "cat2" && nt >= 4 {
    // cos(atan2(c, d)) * a
    let a = env.resolve(tokens[1])
    let c = env.resolve(tokens[2])
    let d = env.resolve(tokens[3])
    let ang = @ffi.ffi_atan2(d, c)
    @ffi.ffi_cos(ang) * a
  } else if op == "sat2" && nt >= 4 {
    // sin(atan2(c, d)) * a
    let a = env.resolve(tokens[1])
    let c = env.resolve(tokens[2])
    let d = env.resolve(tokens[3])
    let ang = @ffi.ffi_atan2(d, c)
    @ffi.ffi_sin(ang) * a
  } else {
    0.0
  }
  env.set(name, val)
}

// ── Path rendering ───────────────────────────────────────────────────────────

///|
/// Convert guide-based coordinates to an SVG path d attribute string.
///  Path format: space-separated tokens where:
///  M x y = moveTo, L x y = lineTo, Z = close
///  A wR hR stAng swAng = arcTo (OOXML format, converted to SVG arc)
///  C x1 y1 x2 y2 x3 y3 = cubicBezTo
///  Q x1 y1 x2 y2 = quadBezTo
///  W pw ph = this path's own coordinate space (a:path w/h), overriding path_w/path_h
fn render_geom_path(
  env : GuideEnv,
  path_cmds : String,
  x_off : Int,
  y_off : Int,
  path_w : Double,
  path_h : Double,
  shape_w : Double,
  shape_h : Double,
) -> String {
  let tokens = @xml.split_by_space(path_cmds)
  let nt = tokens.length()
  let mut d = ""
  let mut i = 0
  // Scale factors if path has its own coordinate space
  let mut sx = if path_w > 0.0 { shape_w / path_w } else { 1.0 }
  let mut sy = if path_h > 0.0 { shape_h / path_h } else { 1.0 }
  // Current position for arcTo endpoint calculation
  let mut cur_x = 0.0
  let mut cur_y = 0.0
  while i < nt {
    let cmd = tokens[i]
    if cmd == "W" && i + 2 < nt {
      let pw = parse_double(tokens[i + 1])
      let ph = parse_double(tokens[i + 2])
      sx = if pw > 0.0 { shape_w / pw } else { 1.0 }
      sy = if ph > 0.0 { shape_h / ph } else { 1.0 }
      i = i + 3
    } else if cmd == "M" && i + 2 < nt {
      let x = env.resolve(tokens[i + 1]) * sx
      let y = env.resolve(tokens[i + 2]) * sy
      cur_x = x
      cur_y = y
      d = d +
        "M" +
        dts(x + x_off.to_double()) +
        " " +
        dts(y + y_off.to_double())
      i = i + 3
    } else if cmd == "L" && i + 2 < nt {
      let x = env.resolve(tokens[i + 1]) * sx
      let y = env.resolve(tokens[i + 2]) * sy
      cur_x = x
      cur_y = y
      d = d +
        "L" +
        dts(x + x_off.to_double()) +
        " " +
        dts(y + y_off.to_double())
      i = i + 3
    } else if cmd == "C" && i + 6 < nt {
      let x1 = env.resolve(tokens[i + 1]) * sx
      let y1 = env.resolve(tokens[i + 2]) * sy
      let x2 = env.resolve(tokens[i + 3]) * sx
      let y2 = env.resolve(tokens[i + 4]) * sy
      let x3 = env.resolve(tokens[i + 5]) * sx
      let y3 = env.resolve(tokens[i + 6]) * sy
      cur_x = x3
      cur_y = y3
      let ox = x_off.to_double()
      let oy = y_off.to_double()
      d = d +
        "C" +
        dts(x1 + ox) +
        " " +
        dts(y1 + oy) +
        " " +
        dts(x2 + ox) +
        " " +
        dts(y2 + oy) +
        " " +
        dts(x3 + ox) +
        " " +
        dts(y3 + oy)
      i = i + 7
    } else if cmd == "Q" && i + 4 < nt {
      let x1 = env.resolve(tokens[i + 1]) * sx
      let y1 = env.resolve(tokens[i + 2]) * sy
      let x2 = env.resolve(tokens[i + 3]) * sx
      let y2 = env.resolve(tokens[i + 4]) * sy
      cur_x = x2
      cur_y = y2
      let ox = x_off.to_double()
      let oy = y_off.to_double()
      d = d +
        "Q" +
        dts(x1 + ox) +
        " " +
        dts(y1 + oy) +
        " " +
        dts(x2 + ox) +
        " " +
        dts(y2 + oy)
      i = i + 5
    } else if cmd == "A" && i + 4 < nt {
      // arcTo: A wR hR stAng swAng (OOXML format)
      let wr = env.resolve(tokens[i + 1]) * sx
      let hr = env.resolve(tokens[i + 2]) * sy
      let st_ang = env.resolve(tokens[i + 3])
      let sw_ang = env.resolve(tokens[i + 4])
      let arc_svg = arcto_svg(
        cur_x, cur_y, wr, hr, st_ang, sw_ang, x_off, y_off,
      )
      cur_x = arc_svg.1
      cur_y = arc_svg.2
      d = d + arc_svg.0
      i = i + 5
    } else if cmd == "Z" {
      d = d + "Z"
      i = i + 1
    } else {
      i = i + 1
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
/// Render multiple paths separated by "|" into a single SVG path d string.
fn render_geom_paths(
  env : GuideEnv,
  paths : String,
  x_off : Int,
  y_off : Int,
  w : Double,
  h : Double,
  path_w : Double,
  path_h : Double,
) -> String {
  // Split by "|" for multiple paths
  let len = paths.length()
  let mut start = 0
  let mut d = ""
  for i in 0..<len {
    if paths.get_char(i).unwrap() == '|' {
      if i > start {
        let seg = @xml.str_substring(paths, start, i)
        d = d + render_geom_path(env, seg, x_off, y_off, path_w, path_h, w, h)
      }
      start = i + 1
    }
  }
  if start < len {
    let seg = @xml.str_substring(paths, start, len)
    d = d + render_geom_path(env, seg, x_off, y_off, path_w, path_h, w, h)
  }
  d
}

// ── Preset shape definitions ─────────────────────────────────────────────────

///|
/// Shape definition: adjustment defaults, guide formulas, and path commands.
///  avlst: "name1=val1;name2=val2" (empty if no adjustments)
///  gdlst: "name1=fmla1;name2=fmla2" (sequential evaluation)
///  paths: path commands (multiple paths separated by "|")
priv struct PresetDef {
  avlst : String
  gdlst : String
  paths : String
}

///|
/// Look up a preset shape definition by name.
///  Returns None if the shape is not implemented.
fn get_preset_def(prst : String) -> PresetDef? {
  // ── Basic shapes ──
  if prst == "triangle" || prst == "isosTriangle" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 100000;x1=*/ w a 100000;x2=*/ w a 200000",
      paths: "M l b L x1 t L r b Z",
    })
  } else if prst == "rtTriangle" {
    Some({ avlst: "", gdlst: "", paths: "M l b L l t L r b Z", })
  } else if prst == "diamond" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z", })
  } else if prst == "parallelogram" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 100000;x1=*/ w a 100000;dx1=+- r 0 x1",
      paths: "M x1 t L r t L dx1 b L l b Z",
    })
  } else if prst == "trapezoid" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 50000;x1=*/ w a 100000;dx1=+- r 0 x1",
      paths: "M l b L x1 t L dx1 t L r b Z",
    })
  } else if prst == "pentagon" {
    // Regular pentagon with top vertex pointing up
    // Vertices at OOXML angles: 270°, 342°, 54°, 126°, 198°
    // Use cos/sin with pre-computed angles in 60000ths
    Some({
      avlst: "",
      gdlst: "cx1=cos wd2 20520000;cy1=sin hd2 20520000;cx2=cos wd2 3240000;cy2=sin hd2 3240000;cx3=cos wd2 7560000;cy3=sin hd2 7560000;cx4=cos wd2 11880000;cy4=sin hd2 11880000;vx1=+- wd2 cx1 0;vy1=+- hd2 cy1 0;vx2=+- wd2 cx2 0;vy2=+- hd2 cy2 0;vx3=+- wd2 cx3 0;vy3=+- hd2 cy3 0;vx4=+- wd2 cx4 0;vy4=+- hd2 cy4 0",
      paths: "M wd2 t L vx1 vy1 L vx2 vy2 L vx3 vy3 L vx4 vy4 Z",
    })
  } else if prst == "hexagon" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 50000;x1=*/ w a 100000;dx1=+- r 0 x1",
      paths: "M x1 t L dx1 t L r hd2 L dx1 b L x1 b L l hd2 Z",
    })
  } else if prst == "octagon" {
    Some({
      avlst: "adj=29290",
      gdlst: "a=pin 0 adj 50000;x1=*/ w a 100000;dx1=+- r 0 x1;y1=*/ h a 100000;dy1=+- b 0 y1",
      paths: "M x1 t L dx1 t L r y1 L r dy1 L dx1 b L x1 b L l dy1 L l y1 Z",
    })
  } else if prst == "heptagon" {
    // Regular heptagon: simplified as 7-sided polygon
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 4;x2=*/ w 1 4;y1=*/ h 1 5;y2=*/ h 4 5",
      paths: "M wd2 t L x1 y1 L r hd2 L x1 y2 L x2 y2 L l hd2 L x2 y1 Z",
    })
  } else if prst == "plus" || prst == "cross" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 50000;x1=*/ w a 100000;dx1=+- r 0 x1;y1=*/ h a 100000;dy1=+- b 0 y1",
      paths: "M x1 t L dx1 t L dx1 y1 L r y1 L r dy1 L dx1 dy1 L dx1 b L x1 b L x1 dy1 L l dy1 L l y1 L x1 y1 Z",
    })
  } else if prst == "donut" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 50000;dr=*/ ss a 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr;idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z M dr vc A iwd2 ihd2 cd2 -5400000 A iwd2 ihd2 cd4 -5400000 A iwd2 ihd2 0 -5400000 A iwd2 ihd2 3cd4 -5400000 Z",
    })
  } else if prst == "noSmoking" {
    Some({
      avlst: "adj=18750",
      gdlst: "a=pin 0 adj 50000;dr=*/ ss a 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr;ang=at2 w h;ct=cos ihd2 ang;st=sin iwd2 ang;m=mod ct st 0;n=*/ iwd2 ihd2 m;drd2=*/ dr 1 2;dang=at2 n drd2;dang2=*/ dang 2 1;swAng=+- -10800000 dang2 0;t3=at2 w h;stAng1=+- t3 0 dang;stAng2=+- stAng1 0 cd2;ct1=cos ihd2 stAng1;st1=sin iwd2 stAng1;m1=mod ct1 st1 0;n1=*/ iwd2 ihd2 m1;dx1=cos n1 stAng1;dy1=sin n1 stAng1;x1=+- hc dx1 0;y1=+- vc dy1 0;x2=+- hc 0 dx1;y2=+- vc 0 dy1;idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z M x1 y1 A iwd2 ihd2 stAng1 swAng Z M x2 y2 A iwd2 ihd2 stAng2 swAng Z",
    })
  } else if prst == "blockArc" {
    Some({
      avlst: "adj1=10800000;adj2=0;adj3=25000",
      gdlst: "a3=pin 0 adj3 50000;dr=*/ ss a3 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr;sa=*/ adj1 1 1;ea=*/ adj2 1 1;csa=cos wd2 sa;ssa=sin hd2 sa;cea=cos wd2 ea;sea=sin hd2 ea;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0;icsa=cos iwd2 sa;issa=sin ihd2 sa;icea=cos iwd2 ea;isea=sin ihd2 ea;ix1=+- wd2 icsa 0;iy1=+- hd2 issa 0;ix2=+- wd2 icea 0;iy2=+- hd2 isea 0;sw=+- ea 0 sa",
      paths: "M x1 y1 A wd2 hd2 sa sw L ix2 iy2 A iwd2 ihd2 ea 0 Z",
    })
    // ── Block arrows ──
  } else if prst == "rightArrow" {
    // ECMA-376: arrowhead dimension uses ss (short side), not w
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 100000 w ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dx=*/ ss a2 100000;x1=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M l y1 L x1 y1 L x1 t L r hd2 L x1 b L x1 y2 L l y2 Z",
    })
  } else if prst == "leftArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 100000 w ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dx=*/ ss a2 100000;x1=+- l dx 0;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M r y1 L x1 y1 L x1 t L l hd2 L x1 b L x1 y2 L r y2 Z",
    })
  } else if prst == "upArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 100000 h ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dy=*/ ss a2 100000;y1=+- t dy 0;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M x1 b L x1 y1 L l y1 L wd2 t L r y1 L x2 y1 L x2 b Z",
    })
  } else if prst == "downArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 100000 h ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dy=*/ ss a2 100000;y1=+- b 0 dy;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M x1 t L x1 y1 L l y1 L wd2 b L r y1 L x2 y1 L x2 t Z",
    })
  } else if prst == "leftRightArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 50000 w ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dx=*/ ss a2 100000;x1=+- l dx 0;x2=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M l hd2 L x1 t L x1 y1 L x2 y1 L x2 t L r hd2 L x2 b L x2 y2 L x1 y2 L x1 b Z",
    })
  } else if prst == "upDownArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 50000 h ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dy=*/ ss a2 100000;y1=+- t dy 0;y2=+- b 0 dy;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M wd2 t L r y1 L x2 y1 L x2 y2 L r y2 L wd2 b L l y2 L x1 y2 L x1 y1 L l y1 Z",
    })
  } else if prst == "quadArrow" {
    Some({
      avlst: "adj1=22500;adj2=22500;adj3=22500",
      gdlst: "a2=pin 0 adj2 50000;maxAdj1=*/ a2 2 1;a1=pin 0 adj1 maxAdj1;q1=+- 100000 0 maxAdj1;maxAdj3=*/ q1 1 2;a3=pin 0 adj3 maxAdj3;x1=*/ ss a3 100000;dx2=*/ ss a2 100000;x2=+- hc 0 dx2;x5=+- hc dx2 0;dx3=*/ ss a1 200000;x3=+- hc 0 dx3;x4=+- hc dx3 0;x6=+- r 0 x1;y2=+- vc 0 dx2;y5=+- vc dx2 0;y3=+- vc 0 dx3;y4=+- vc dx3 0;y6=+- b 0 x1;il=*/ dx3 x1 dx2;ir=+- r 0 il",
      paths: "M l vc L x1 y2 L x1 y3 L x3 y3 L x3 x1 L x2 x1 L hc t L x5 x1 L x4 x1 L x4 y3 L x6 y3 L x6 y2 L r vc L x6 y5 L x6 y4 L x4 y4 L x4 y6 L x5 y6 L hc b L x2 y6 L x3 y6 L x3 y4 L x1 y4 L x1 y5 Z",
    })
    // ── Stars ──
  } else if prst == "star4" {
    Some({
      avlst: "adj=12500",
      gdlst: "a=pin 0 adj 50000;dx=*/ wd2 a 50000;dy=*/ hd2 a 50000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M wd2 t L x2 y1 L r hd2 L x2 y2 L wd2 b L x1 y2 L l hd2 L x1 y1 Z",
    })
  } else if prst == "star5" {
    // 5-pointed star: outer vertices at 270°,342°,54°,126°,198° (OOXML clockwise)
    // Inner vertices between them at 306°,18°,90°,162°,234°
    // adj controls inner radius (default 19098 = ~38.2% of half-size)
    Some({
      avlst: "adj=19098",
      gdlst: "a=pin 0 adj 50000;swd2=*/ wd2 a 50000;shd2=*/ hd2 a 50000;co1=cos wd2 20520000;so1=sin hd2 20520000;co2=cos wd2 3240000;so2=sin hd2 3240000;co3=cos wd2 7560000;so3=sin hd2 7560000;co4=cos wd2 11880000;so4=sin hd2 11880000;ox1=+- wd2 co1 0;oy1=+- hd2 so1 0;ox2=+- wd2 co2 0;oy2=+- hd2 so2 0;ox3=+- wd2 co3 0;oy3=+- hd2 so3 0;ox4=+- wd2 co4 0;oy4=+- hd2 so4 0;ci1=cos swd2 18360000;si1=sin shd2 18360000;ci2=cos swd2 1080000;si2=sin shd2 1080000;ci3=cos swd2 5400000;si3=sin shd2 5400000;ci4=cos swd2 9720000;si4=sin shd2 9720000;ci5=cos swd2 14040000;si5=sin shd2 14040000;ix1=+- wd2 ci1 0;iy1=+- hd2 si1 0;ix2=+- wd2 ci2 0;iy2=+- hd2 si2 0;ix3=+- wd2 ci3 0;iy3=+- hd2 si3 0;ix4=+- wd2 ci4 0;iy4=+- hd2 si4 0;ix5=+- wd2 ci5 0;iy5=+- hd2 si5 0",
      paths: "M wd2 t L ix1 iy1 L ox1 oy1 L ix2 iy2 L ox2 oy2 L ix3 iy3 L ox3 oy3 L ix4 iy4 L ox4 oy4 L ix5 iy5 Z",
    })
  } else if prst == "star6" {
    // 6-pointed star (Star of David): outer vertices at top/bottom and 4 side tips,
    // inner vertices form a hexagon
    Some({
      avlst: "adj=28868",
      gdlst: "a=pin 0 adj 50000;swd2=*/ wd2 a 50000;x1=+- wd2 0 swd2;x2=+- wd2 swd2 0;oy=*/ h 1 4;by=*/ h 3 4",
      paths: "M wd2 t L x2 oy L r oy L x2 hd2 L r by L x2 by L wd2 b L x1 by L l by L x1 hd2 L l oy L x1 oy Z",
    })
  } else if prst == "star8" {
    // 8-pointed star: outer at N/NE/E/SE/S/SW/W/NW, inner between them
    // cos(45°) = sin(45°) ≈ 0.7071, use */ with 7071/10000
    Some({
      avlst: "adj=37500",
      gdlst: "a=pin 0 adj 50000;dx=*/ wd2 a 50000;dy=*/ hd2 a 50000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dxa=*/ wd2 7071 10000;dya=*/ hd2 7071 10000;xa=+- wd2 0 dxa;xb=+- wd2 dxa 0;ya=+- hd2 0 dya;yb=+- hd2 dya 0;idxa=*/ dx 7071 10000;idya=*/ dy 7071 10000;ix1=+- wd2 0 idxa;ix2=+- wd2 idxa 0;iy1=+- hd2 0 idya;iy2=+- hd2 idya 0",
      paths: "M wd2 t L ix2 iy1 L xb ya L x2 y1 L r hd2 L x2 y2 L xb yb L ix2 iy2 L wd2 b L ix1 iy2 L xa yb L x1 y2 L l hd2 L x1 y1 L xa ya L ix1 iy1 Z",
    })
    // ── Flowchart shapes ──
  } else if prst == "flowChartProcess" {
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L r b L l b Z", })
  } else if prst == "flowChartDecision" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z", })
  } else if prst == "flowChartTerminator" {
    Some({
      avlst: "",
      gdlst: "ir=*/ w 18 21;il=*/ w 3 21",
      paths: "M il t A il hd2 10800000 10800000 L ir b A il hd2 0 10800000 Z",
    })
  } else if prst == "flowChartPredefinedProcess" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 8;x2=*/ w 7 8",
      paths: "M l t L r t L r b L l b Z|M x1 t L x1 b|M x2 t L x2 b",
    })
  } else if prst == "flowChartDocument" {
    Some({
      avlst: "",
      gdlst: "y1=*/ h 17322 21600;y2=*/ h 20172 21600",
      paths: "W 21600 21600 M 0 0 L 21600 0 L 21600 17322 C 10800 17322 10800 23922 0 20172 Z",
    })
  } else if prst == "flowChartInputOutput" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;dx1=+- r 0 x1",
      paths: "M x1 t L r t L dx1 b L l b Z",
    })
  } else if prst == "flowChartPreparation" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;dx1=+- r 0 x1",
      paths: "M x1 t L dx1 t L r hd2 L dx1 b L x1 b L l hd2 Z",
    })
  } else if prst == "flowChartManualInput" {
    Some({
      avlst: "",
      gdlst: "y1=*/ h 1 5",
      paths: "M l y1 L r t L r b L l b Z",
    })
  } else if prst == "flowChartManualOperation" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;dx1=+- r 0 x1",
      paths: "M l t L r t L dx1 b L x1 b Z",
    })
  } else if prst == "flowChartOffpageConnector" {
    Some({
      avlst: "",
      gdlst: "y1=*/ h 4 5",
      paths: "M l t L r t L r y1 L wd2 b L l y1 Z",
    })
  } else if prst == "flowChartConnector" {
    Some({
      avlst: "",
      gdlst: "idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z",
    })
  } else if prst == "flowChartSort" {
    Some({
      avlst: "",
      gdlst: "",
      paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z|M l hd2 L r hd2",
    })
  } else if prst == "flowChartMerge" {
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L wd2 b Z", })
  } else if prst == "flowChartExtract" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r b L l b Z", })
  } else if prst == "flowChartDelay" {
    Some({
      avlst: "",
      gdlst: "",
      paths: "M l t L wd2 t A wd2 hd2 16200000 10800000 L l b Z",
    })
    // ── Callouts ──
  } else if prst == "wedgeRectCallout" {
    Some({
      avlst: "adj1=-20833;adj2=62500",
      gdlst: "dxPos=*/ w adj1 100000;dyPos=*/ h adj2 100000;xPos=+- hc dxPos 0;yPos=+- vc dyPos 0;dx=+- xPos 0 hc;dy=+- yPos 0 vc;dq=*/ dxPos h w;ady=abs dyPos;adq=abs dq;dz=+- ady 0 adq;xg1=?: dxPos 7 2;xg2=?: dxPos 10 5;x1=*/ w xg1 12;x2=*/ w xg2 12;yg1=?: dyPos 7 2;yg2=?: dyPos 10 5;y1=*/ h yg1 12;y2=*/ h yg2 12;t1=?: dxPos l xPos;xl=?: dz l t1;t2=?: dyPos x1 xPos;xt=?: dz t2 x1;t3=?: dxPos xPos r;xr=?: dz r t3;t4=?: dyPos xPos x1;xb=?: dz t4 x1;t5=?: dxPos y1 yPos;yl=?: dz y1 t5;t6=?: dyPos t yPos;yt=?: dz t6 t;t7=?: dxPos yPos y1;yr=?: dz y1 t7;t8=?: dyPos yPos b;yb=?: dz t8 b",
      paths: "M l t L x1 t L xt yt L x2 t L r t L r y1 L xr yr L r y2 L r b L x2 b L xb yb L x1 b L l b L l y2 L xl yl L l y1 Z",
    })
  } else if prst == "wedgeRoundRectCallout" {
    Some({
      avlst: "adj1=-20833;adj2=62500",
      gdlst: "a1=*/ w adj1 100000;a2=*/ h adj2 100000;x1=+- wd2 a1 0;y1=+- hd2 a2 0;cr=*/ ss 1 12;dx1=+- wd2 0 wd8;dx2=+- wd2 wd8 0;dcr=+- r 0 cr;dbcr=+- b 0 cr",
      paths: "M cr t L dcr t A cr cr 16200000 5400000 L r dbcr A cr cr 0 5400000 L dx2 b L x1 y1 L dx1 b L cr b A cr cr 5400000 5400000 L l cr A cr cr 10800000 5400000 Z",
    })
  } else if prst == "wedgeEllipseCallout" {
    Some({
      avlst: "adj1=-20833;adj2=62500",
      gdlst: "dxPos=*/ w adj1 100000;dyPos=*/ h adj2 100000;xPos=+- hc dxPos 0;yPos=+- vc dyPos 0;sdx=*/ dxPos h 1;sdy=*/ dyPos w 1;pang=at2 sdx sdy;stAng=+- pang 660000 0;enAng=+- pang 0 660000;dx1=cos wd2 stAng;dy1=sin hd2 stAng;x1=+- hc dx1 0;y1=+- vc dy1 0;dx2=cos wd2 enAng;dy2=sin hd2 enAng;x2=+- hc dx2 0;y2=+- vc dy2 0;stAng1=at2 dx1 dy1;enAng1=at2 dx2 dy2;swAng1=+- enAng1 0 stAng1;swAng2=+- swAng1 21600000 0;swAng=?: swAng1 swAng1 swAng2;idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M xPos yPos L x1 y1 A wd2 hd2 stAng1 swAng Z",
    })
    // ── Decorative / other ──
  } else if prst == "heart" {
    // Heart shape using cubic bezier curves
    // Start at bottom point, curve up to top-center dip, then around each lobe
    Some({
      avlst: "",
      gdlst: "dx1=*/ w 49 48;dx2=*/ w 10 48;cx1=*/ w 1 6;cx2=*/ w 5 6;cy2=*/ h 1 6;cx3=*/ w 1 3;cx4=*/ w 2 3;cy4=*/ h 5 6",
      paths: "M wd2 b C cx3 cy4 l hd2 l hd4 C l 0 wd2 0 wd2 hd4 C wd2 0 r 0 r hd4 C r hd2 cx4 cy4 wd2 b Z",
    })
  } else if prst == "lightningBolt" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 5022 21600;x3=*/ w 8472 21600;x4=*/ w 8757 21600;x5=*/ w 10012 21600;x8=*/ w 12860 21600;x9=*/ w 13917 21600;x11=*/ w 16577 21600;y1=*/ h 3890 21600;y2=*/ h 6080 21600;y4=*/ h 7437 21600;y6=*/ h 9705 21600;y7=*/ h 12007 21600;y10=*/ h 14277 21600;y11=*/ h 14915 21600",
      paths: "W 21600 21600 M 8472 0 L 12860 6080 L 11050 6797 L 16577 12007 L 14767 12877 L 21600 21600 L 10012 14915 L 12222 13987 L 5022 9705 L 7602 8382 L 0 3890 Z",
    })
  } else if prst == "sun" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 12500 adj 46875;g0=+- 50000 0 a;g1=*/ g0 30274 32768;g2=*/ g0 12540 32768;g3=+- g1 50000 0;g4=+- g2 50000 0;g5=+- 50000 0 g1;g6=+- 50000 0 g2;g7=*/ g0 23170 32768;g8=+- 50000 g7 0;g9=+- 50000 0 g7;g10=*/ g5 3 4;g11=*/ g6 3 4;g12=+- g10 3662 0;g13=+- g11 3662 0;g14=+- g11 12500 0;g15=+- 100000 0 g10;g16=+- 100000 0 g12;g17=+- 100000 0 g13;g18=+- 100000 0 g14;ox1=*/ w 18436 21600;oy1=*/ h 3163 21600;ox2=*/ w 3163 21600;oy2=*/ h 18436 21600;x8=*/ w g8 100000;x9=*/ w g9 100000;x10=*/ w g10 100000;x12=*/ w g12 100000;x13=*/ w g13 100000;x14=*/ w g14 100000;x15=*/ w g15 100000;x16=*/ w g16 100000;x17=*/ w g17 100000;x18=*/ w g18 100000;x19=*/ w a 100000;wR=*/ w g0 100000;hR=*/ h g0 100000;y8=*/ h g8 100000;y9=*/ h g9 100000;y10=*/ h g10 100000;y12=*/ h g12 100000;y13=*/ h g13 100000;y14=*/ h g14 100000;y15=*/ h g15 100000;y16=*/ h g16 100000;y17=*/ h g17 100000;y18=*/ h g18 100000",
      paths: "M r vc L x15 y18 L x15 y14 Z M ox1 oy1 L x16 y13 L x17 y12 Z M hc t L x18 y10 L x14 y10 Z M ox2 oy1 L x13 y12 L x12 y13 Z M l vc L x10 y14 L x10 y18 Z M ox2 oy2 L x12 y17 L x13 y16 Z M hc b L x14 y15 L x18 y15 Z M ox1 oy2 L x17 y16 L x16 y17 Z M x19 vc A wR hR cd2 21600000 Z",
    })
  } else if prst == "moon" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 87500;g0=*/ ss a 100000;g0w=*/ g0 w ss;g1=+- ss 0 g0;g2=*/ g0 g0 g1;g3=*/ ss ss g1;g4=*/ g3 2 1;g5=+- g4 0 g2;g6=+- g5 0 g0;g6w=*/ g6 w ss;g7=*/ g5 1 2;g8=+- g7 0 g0;dy1=*/ g8 hd2 ss;g10h=+- vc 0 dy1;g11h=+- vc dy1 0;g12=*/ g0 9598 32768;g12w=*/ g12 w ss;g13=+- ss 0 g12;q1=*/ ss ss 1;q2=*/ g13 g13 1;q3=+- q1 0 q2;q4=sqrt q3;dy4=*/ q4 hd2 ss;g15h=+- vc 0 dy4;g16h=+- vc dy4 0;g17w=+- g6w 0 g0w;g18w=*/ g17w 1 2;dx2p=+- g0w g18w w;dx2=*/ dx2p -1 1;dy2=*/ hd2 -1 1;stAng1=at2 dx2 dy2;enAngp1=at2 dx2 hd2;enAng1=+- enAngp1 0 21600000;swAng1=+- enAng1 0 stAng1",
      paths: "M r b A w hd2 cd4 cd2 A g18w dy1 stAng1 swAng1 Z",
    })
  } else if prst == "smileyFace" {
    Some({
      avlst: "adj=4653",
      gdlst: "a=pin -4653 adj 4653;x1=*/ w 4969 21699;x2=*/ w 6215 21600;x3=*/ w 13135 21600;x4=*/ w 16640 21600;y1=*/ h 7570 21600;y3=*/ h 16515 21600;dy2=*/ h a 100000;y2=+- y3 0 dy2;y4=+- y3 dy2 0;dy3=*/ h a 50000;y5=+- y4 dy3 0;idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0;wR=*/ w 1125 21600;hR=*/ h 1125 21600",
      paths: "M l vc A wd2 hd2 cd2 21600000 Z|M x2 y1 A wR hR cd2 21600000 M x3 y1 A wR hR cd2 21600000|M x1 y2 Q hc y5 x4 y2|M l vc A wd2 hd2 cd2 21600000 Z",
    })
  } else if prst == "foldedCorner" {
    Some({
      avlst: "adj=16667",
      gdlst: "a=pin 0 adj 50000;dy=*/ h a 100000;y1=+- b 0 dy;dx=*/ w a 100000;x1=+- r 0 dx",
      paths: "M l t L r t L r y1 L x1 b L l b Z|M r y1 L x1 y1 L x1 b",
    })
  } else if prst == "frame" {
    Some({
      avlst: "adj1=12500",
      gdlst: "a=pin 0 adj1 50000;dx=*/ w a 100000;dy=*/ h a 100000;dx1=+- r 0 dx;dy1=+- b 0 dy",
      paths: "M l t L r t L r b L l b Z|M dx dy L dx1 dy L dx1 dy1 L dx dy1 Z",
    })
  } else if prst == "bevel" {
    Some({
      avlst: "adj=12500",
      gdlst: "a=pin 0 adj 50000;dx=*/ w a 100000;dy=*/ h a 100000;dx1=+- r 0 dx;dy1=+- b 0 dy",
      paths: "M l t L r t L dx1 dy L dx dy Z|M r t L r b L dx1 dy1 L dx1 dy Z|M r b L l b L dx dy1 L dx1 dy1 Z|M l b L l t L dx dy L dx dy1 Z|M dx dy L dx1 dy L dx1 dy1 L dx dy1 Z",
    })
  } else if prst == "can" || prst == "cylinder" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 50000;dy=*/ h a 100000;y1=+- t dy 0;y2=+- b 0 dy",
      paths: "M l y1 A wd2 dy 10800000 -10800000 L r y2 A wd2 dy 0 10800000 Z|M l y1 A wd2 dy 10800000 10800000",
    })
  } else if prst == "cube" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 0 adj 100000;y1=*/ ss a 100000;y4=+- b 0 y1;y2=*/ y4 1 2;y3=+/ y1 b 2;x4=+- r 0 y1;x2=*/ x4 1 2;x3=+/ y1 r 2",
      paths: "M l y1 L x4 y1 L x4 b L l b Z|M x4 y1 L r t L r y4 L x4 b Z|M l y1 L y1 t L r t L x4 y1 Z|M l y1 L y1 t L r t L r y4 L x4 b L l b Z M l y1 L x4 y1 L r t M x4 y1 L x4 b",
    })
  } else if prst == "ribbon" {
    Some({
      avlst: "adj1=16667;adj2=50000",
      gdlst: "a1=pin 0 adj1 33333;a2=pin 25000 adj2 75000;dy=*/ h a1 100000;y1=+- t dy 0;dx=*/ w a2 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;dy2=*/ dy 1 2;y2=+- b 0 dy;y3=+- b 0 dy2",
      paths: "M l y1 L x1 y1 L x1 t L x2 t L x2 y1 L r y1 L r y3 L x2 b L x2 y2 L x1 y2 L x1 b L l y3 Z",
    })
  } else if prst == "ribbon2" {
    Some({
      avlst: "adj1=16667;adj2=50000",
      gdlst: "a1=pin 0 adj1 33333;a2=pin 25000 adj2 75000;dy=*/ h a1 100000;y1=+- b 0 dy;dx=*/ w a2 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;dy2=*/ dy 1 2;y2=+- t dy 0;y3=+- t dy2 0",
      paths: "M l y3 L x1 t L x1 y2 L x2 y2 L x2 t L r y3 L r y1 L x2 y1 L x2 b L x1 b L x1 y1 L l y1 Z",
    })
  } else if prst == "chevron" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 100000;dx=*/ w a 100000;x1=+- r 0 dx",
      paths: "M l t L x1 t L r hd2 L x1 b L l b L dx hd2 Z",
    })
  } else if prst == "homePlate" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 100000;dx=*/ ss a 100000;x1=+- r 0 dx",
      paths: "M l t L x1 t L r hd2 L x1 b L l b Z",
    })
  } else if prst == "notchedRightArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 100000 w ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;dx=*/ ss a2 100000;x1=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx1=*/ h a2 200000;nx1=+- l dx1 0",
      paths: "M l y1 L x1 y1 L x1 t L r hd2 L x1 b L x1 y2 L l y2 L nx1 hd2 Z",
    })
  } else if prst == "stripedRightArrow" {
    // ECMA-376: three small stripes on the left, then arrow body+head on the right.
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "maxAdj2=*/ 84375 w ss;a1=pin 0 adj1 100000;a2=pin 0 adj2 maxAdj2;x4=*/ ss 5 32;dx5=*/ ss a2 100000;x5=+- r 0 dx5;dy1=*/ h a1 200000;y1=+- hd2 0 dy1;y2=+- hd2 dy1 0;dx2=*/ ss 1 32;x1=+- l dx2 0;dx3=*/ ss 1 16;x2=+- l dx3 0;x3=+- x2 dx2 0",
      paths: "M l y1 L x1 y1 L x1 y2 L l y2 Z|M x2 y1 L x3 y1 L x3 y2 L x2 y2 Z|M x4 y1 L x5 y1 L x5 t L r hd2 L x5 b L x5 y2 L x4 y2 Z",
    })
  } else if prst == "leftRightArrowCallout" {
    // ECMA-376: callout box with both ends pointing inward as arrowheads.
    // adj1=body height ratio, adj2=head height ratio, adj3=head width ratio,
    // adj4=callout box width ratio.
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000;adj4=48123",
      gdlst: "a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;q2=*/ a3 ss w;maxAdj1=+- 100000 0 q2;a1=pin 0 adj1 maxAdj1;dy2=*/ h a2 100000;dy1=*/ h a1 200000;y1=+- hd2 0 dy1;y2=+- hd2 0 dy2;y3=+- hd2 dy2 0;y4=+- hd2 dy1 0;a4=pin 0 adj4 100000;x1=*/ w a3 100000;x4=+- r 0 x1;dx2=*/ w a4 200000;x2=+- wd2 0 dx2;x3=+- wd2 dx2 0",
      paths: "M l hd2 L x1 y1 L x1 y2 L x2 y2 L x2 t L x3 t L x3 y2 L x4 y2 L x4 y1 L r hd2 L x4 y4 L x4 y3 L x3 y3 L x3 b L x2 b L x2 y3 L x1 y3 L x1 y4 Z",
    })
    // ── Math ──
  } else if prst == "mathPlus" {
    Some({
      avlst: "adj1=23520",
      gdlst: "a=pin 0 adj1 73490;dy=*/ h a 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx=*/ w a 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M x1 y1 L x2 y1 L x2 t L r y1 L r y2 L x2 y2 L x2 b L x1 b L x1 y2 L l y2 L l y1 L x1 y1 Z",
    })
  } else if prst == "mathMinus" {
    Some({
      avlst: "adj1=23520",
      gdlst: "a=pin 0 adj1 100000;dy=*/ h a 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M l y1 L r y1 L r y2 L l y2 Z",
    })
  } else if prst == "mathMultiply" {
    Some({
      avlst: "adj1=23520",
      gdlst: "a=pin 0 adj1 51965;dx=*/ wd2 a 100000;dy=*/ hd2 a 100000;ix1=+- wd2 0 dx;ix2=+- wd2 dx 0;iy1=+- hd2 0 dy;iy2=+- hd2 dy 0",
      paths: "M ix1 t L wd2 iy1 L ix2 t L r iy1 L ix2 hd2 L r iy2 L ix2 b L wd2 iy2 L ix1 b L l iy2 L ix1 hd2 L l iy1 Z",
    })
  } else if prst == "mathDivide" {
    Some({
      avlst: "adj1=23520;adj2=5880;adj3=11760",
      gdlst: "a1=pin 1000 adj1 36745;ma1=+- 0 0 a1;ma3h=+/ 73490 ma1 4;ma3w=*/ 36745 w h;maxAdj3=min ma3h ma3w;a3=pin 1000 adj3 maxAdj3;m4a3=*/ -4 a3 1;maxAdj2=+- 73490 m4a3 a1;a2=pin 0 adj2 maxAdj2;dy1=*/ h a1 200000;yg=*/ h a2 100000;rad=*/ h a3 100000;dx1=*/ w 73490 200000;y3=+- vc 0 dy1;y4=+- vc dy1 0;a=+- yg rad 0;y2=+- y3 0 a;y1=+- y2 0 rad;y5=+- b 0 y1;x1=+- hc 0 dx1;x3=+- hc dx1 0;x2=+- hc 0 rad",
      paths: "M hc y1 A rad rad 3cd4 21600000 Z M hc y5 A rad rad cd4 21600000 Z M x1 y3 L x3 y3 L x3 y4 L x1 y4 Z",
    })
  } else if prst == "mathEqual" {
    Some({
      avlst: "adj1=23520;adj2=11760",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dy1=*/ h a1 200000;dy2=*/ h a2 200000;y1=+- hd2 0 dy2;y2=+- y1 0 dy1;y3=+- hd2 dy2 0;y4=+- y3 dy1 0",
      paths: "M l y2 L r y2 L r y1 L l y1 Z|M l y3 L r y3 L r y4 L l y4 Z",
    })
    // ── Snip/Round rect variants ──
  } else if prst == "snip1Rect" {
    Some({
      avlst: "adj=16667",
      gdlst: "a=pin 0 adj 50000;dx=*/ ss a 100000;x1=+- r 0 dx",
      paths: "M l t L x1 t L r dx L r b L l b Z",
    })
  } else if prst == "snip2SameRect" {
    Some({
      avlst: "adj1=16667;adj2=0",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;dx1=*/ ss a1 100000;dx2=*/ ss a2 100000;x1=+- r 0 dx1;x2=+- r 0 dx2;y2=+- b 0 dx2",
      paths: "M dx1 t L x1 t L r dx1 L r y2 L x2 b L dx2 b L l y2 L l dx1 Z",
    })
  } else if prst == "round1Rect" {
    Some({
      avlst: "adj=16667",
      gdlst: "a=pin 0 adj 50000;dx=*/ ss a 100000;x1=+- r 0 dx",
      paths: "M l t L x1 t A dx dx 16200000 5400000 L r b L l b Z",
    })
  } else if prst == "round2SameRect" {
    Some({
      avlst: "adj1=16667;adj2=0",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;dx1=*/ ss a1 100000;dx2=*/ ss a2 100000;x1=+- r 0 dx1;x2=+- r 0 dx2;y2=+- b 0 dx2",
      paths: "M dx1 t L x1 t A dx1 dx1 16200000 5400000 L r y2 A dx2 dx2 0 5400000 L dx2 b A dx2 dx2 5400000 5400000 L l dx1 A dx1 dx1 10800000 5400000 Z",
    })
  } else if prst == "round2DiagRect" {
    Some({
      avlst: "adj1=16667;adj2=0",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;dx1=*/ ss a1 100000;dx2=*/ ss a2 100000;x1=+- r 0 dx1;x2=+- r 0 dx2;y1=+- b 0 dx1;y2=+- b 0 dx2",
      paths: "M dx2 t L x1 t A dx1 dx1 16200000 5400000 L r y2 A dx2 dx2 0 5400000 L dx1 b A dx1 dx1 5400000 5400000 L l dx2 A dx2 dx2 10800000 5400000 Z",
    })
  } else if prst == "snipRoundRect" {
    Some({
      avlst: "adj1=16667;adj2=16667",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;dx1=*/ ss a1 100000;dx2=*/ ss a2 100000;x1=+- r 0 dx1;x2=+- r 0 dx2",
      paths: "M dx1 t L x2 t L r dx2 L r b L l b L l dx1 A dx1 dx1 10800000 5400000 Z",
    })
    // ── Other common shapes ──
  } else if prst == "teardrop" {
    Some({
      avlst: "adj=100000",
      gdlst: "a=pin 0 adj 200000;r2=*/ ss a 200000;ox=+- wd2 r2 0;oy=+- hd2 0 r2",
      paths: "M l hd2 A wd2 hd2 10800000 5400000 L ox oy L wd2 t A wd2 hd2 16200000 -5400000 Z",
    })
  } else if prst == "pie" || prst == "arc" {
    Some({
      avlst: "adj1=0;adj2=16200000",
      gdlst: "sa=*/ adj1 1 1;ea=*/ adj2 1 1;sw=+- ea 0 sa;csa=cos wd2 sa;ssa=sin hd2 sa;cea=cos wd2 ea;sea=sin hd2 ea;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0",
      paths: "M wd2 hd2 L x1 y1 A wd2 hd2 sa sw L wd2 hd2 Z",
    })
  } else if prst == "chord" {
    Some({
      avlst: "adj1=2700000;adj2=16200000",
      gdlst: "sa=*/ adj1 1 1;ea=*/ adj2 1 1;sw=+- ea 0 sa;csa=cos wd2 sa;ssa=sin hd2 sa;cea=cos wd2 ea;sea=sin hd2 ea;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0",
      paths: "M x1 y1 A wd2 hd2 sa sw Z",
    })
  } else if prst == "decagon" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;x2=*/ w 4 5;y1=*/ h 3 20;y2=*/ h 17 20;x3=*/ w 1 10;x4=*/ w 9 10",
      paths: "M wd2 t L x4 y1 L r hd3 L r hd2 L x4 y2 L wd2 b L x3 y2 L l hd2 L l hd3 L x3 y1 Z",
    })
  } else if prst == "diagStripe" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 100000;dy=*/ h a 100000;dx=*/ w a 100000",
      paths: "M l dy L dx t L r t L l b Z",
    })
  } else if prst == "corner" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;x1=*/ w a2 100000;y1=*/ h a1 100000",
      paths: "M l t L x1 t L x1 y1 L r y1 L r b L l b Z",
    })
  } else if prst == "plaque" {
    Some({
      avlst: "adj=16667",
      gdlst: "a=pin 0 adj 50000;dx=*/ ss a 100000;x1=+- r 0 dx;y1=+- b 0 dx",
      paths: "M l dx A dx dx 5400000 -5400000 L x1 t A dx dx 10800000 -5400000 L r y1 A dx dx 16200000 -5400000 L dx b A dx dx 0 -5400000 Z",
    })
  } else if prst == "gear6" {
    // 6-tooth gear: alternating outer arcs (teeth) and inner arcs (valleys)
    // Each tooth sector = 60°, tooth top = 30° arc, valley = 30° arc
    // adj controls inner radius (tooth valley depth)
    Some({
      avlst: "adj=75000",
      gdlst: "a=pin 10000 adj 99000;ir=*/ wd2 a 100000;irh=*/ hd2 a 100000;c00=cos wd2 20700000;s00=sin hd2 20700000;mx=+- wd2 c00 0;my=+- hd2 s00 0;ci0=cos ir 900000;si0=sin irh 900000;px0=+- wd2 ci0 0;py0=+- hd2 si0 0;co1=cos wd2 2700000;so1=sin hd2 2700000;qx1=+- wd2 co1 0;qy1=+- hd2 so1 0;ci1=cos ir 4500000;si1=sin irh 4500000;px1=+- wd2 ci1 0;py1=+- hd2 si1 0;co2=cos wd2 6300000;so2=sin hd2 6300000;qx2=+- wd2 co2 0;qy2=+- hd2 so2 0;ci2=cos ir 8100000;si2=sin irh 8100000;px2=+- wd2 ci2 0;py2=+- hd2 si2 0;co3=cos wd2 9900000;so3=sin hd2 9900000;qx3=+- wd2 co3 0;qy3=+- hd2 so3 0;ci3=cos ir 11700000;si3=sin irh 11700000;px3=+- wd2 ci3 0;py3=+- hd2 si3 0;co4=cos wd2 13500000;so4=sin hd2 13500000;qx4=+- wd2 co4 0;qy4=+- hd2 so4 0;ci4=cos ir 15300000;si4=sin irh 15300000;px4=+- wd2 ci4 0;py4=+- hd2 si4 0;co5=cos wd2 17100000;so5=sin hd2 17100000;qx5=+- wd2 co5 0;qy5=+- hd2 so5 0;ci5=cos ir 18900000;si5=sin irh 18900000;px5=+- wd2 ci5 0;py5=+- hd2 si5 0;hr=*/ wd2 30000 100000;hrh=*/ hd2 30000 100000;hx=+- wd2 hr 0",
      paths: "M mx my A wd2 hd2 20700000 1800000 L px0 py0 A ir irh 900000 1800000 L qx1 qy1 A wd2 hd2 2700000 1800000 L px1 py1 A ir irh 4500000 1800000 L qx2 qy2 A wd2 hd2 6300000 1800000 L px2 py2 A ir irh 8100000 1800000 L qx3 qy3 A wd2 hd2 9900000 1800000 L px3 py3 A ir irh 11700000 1800000 L qx4 qy4 A wd2 hd2 13500000 1800000 L px4 py4 A ir irh 15300000 1800000 L qx5 qy5 A wd2 hd2 17100000 1800000 L px5 py5 A ir irh 18900000 1800000 Z|M hx hd2 A hr hrh 0 21600000 Z",
    })
  } else if prst == "gear9" {
    // 9-tooth gear: alternating outer arcs (teeth) and inner arcs (valleys)
    // Each tooth sector = 40°, tooth top = 20° arc, valley = 20° arc
    Some({
      avlst: "adj=75000",
      gdlst: "a=pin 10000 adj 99000;ir=*/ wd2 a 100000;irh=*/ hd2 a 100000;c00=cos wd2 21000000;s00=sin hd2 21000000;mx=+- wd2 c00 0;my=+- hd2 s00 0;ci0=cos ir 600000;si0=sin irh 600000;px0=+- wd2 ci0 0;py0=+- hd2 si0 0;co1=cos wd2 1800000;so1=sin hd2 1800000;qx1=+- wd2 co1 0;qy1=+- hd2 so1 0;ci1=cos ir 3000000;si1=sin irh 3000000;px1=+- wd2 ci1 0;py1=+- hd2 si1 0;co2=cos wd2 4200000;so2=sin hd2 4200000;qx2=+- wd2 co2 0;qy2=+- hd2 so2 0;ci2=cos ir 5400000;si2=sin irh 5400000;px2=+- wd2 ci2 0;py2=+- hd2 si2 0;co3=cos wd2 6600000;so3=sin hd2 6600000;qx3=+- wd2 co3 0;qy3=+- hd2 so3 0;ci3=cos ir 7800000;si3=sin irh 7800000;px3=+- wd2 ci3 0;py3=+- hd2 si3 0;co4=cos wd2 9000000;so4=sin hd2 9000000;qx4=+- wd2 co4 0;qy4=+- hd2 so4 0;ci4=cos ir 10200000;si4=sin irh 10200000;px4=+- wd2 ci4 0;py4=+- hd2 si4 0;co5=cos wd2 11400000;so5=sin hd2 11400000;qx5=+- wd2 co5 0;qy5=+- hd2 so5 0;ci5=cos ir 12600000;si5=sin irh 12600000;px5=+- wd2 ci5 0;py5=+- hd2 si5 0;co6=cos wd2 13800000;so6=sin hd2 13800000;qx6=+- wd2 co6 0;qy6=+- hd2 so6 0;ci6=cos ir 15000000;si6=sin irh 15000000;px6=+- wd2 ci6 0;py6=+- hd2 si6 0;co7=cos wd2 16200000;so7=sin hd2 16200000;qx7=+- wd2 co7 0;qy7=+- hd2 so7 0;ci7=cos ir 17400000;si7=sin irh 17400000;px7=+- wd2 ci7 0;py7=+- hd2 si7 0;co8=cos wd2 18600000;so8=sin hd2 18600000;qx8=+- wd2 co8 0;qy8=+- hd2 so8 0;ci8=cos ir 19800000;si8=sin irh 19800000;px8=+- wd2 ci8 0;py8=+- hd2 si8 0;hr=*/ wd2 25000 100000;hrh=*/ hd2 25000 100000;hx=+- wd2 hr 0",
      paths: "M mx my A wd2 hd2 21000000 1200000 L px0 py0 A ir irh 600000 1200000 L qx1 qy1 A wd2 hd2 1800000 1200000 L px1 py1 A ir irh 3000000 1200000 L qx2 qy2 A wd2 hd2 4200000 1200000 L px2 py2 A ir irh 5400000 1200000 L qx3 qy3 A wd2 hd2 6600000 1200000 L px3 py3 A ir irh 7800000 1200000 L qx4 qy4 A wd2 hd2 9000000 1200000 L px4 py4 A ir irh 10200000 1200000 L qx5 qy5 A wd2 hd2 11400000 1200000 L px5 py5 A ir irh 12600000 1200000 L qx6 qy6 A wd2 hd2 13800000 1200000 L px6 py6 A ir irh 15000000 1200000 L qx7 qy7 A wd2 hd2 16200000 1200000 L px7 py7 A ir irh 17400000 1200000 L qx8 qy8 A wd2 hd2 18600000 1200000 L px8 py8 A ir irh 19800000 1200000 Z|M hx hd2 A hr hrh 0 21600000 Z",
    })
  } else if prst == "funnel" {
    Some({
      avlst: "",
      gdlst: "y1=*/ h 1 4;x1=*/ w 1 4;x2=*/ w 3 4",
      paths: "M l t A wd2 y1 10800000 -10800000 L x2 y1 L x2 b L x1 b L x1 y1 Z",
    })
  } else if prst == "irregularSeal1" || prst == "irregularSeal2" {
    // Simplified irregular seal — render as rough polygon
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;x2=*/ w 2 5;x3=*/ w 3 5;x4=*/ w 4 5;y1=*/ h 1 5;y2=*/ h 2 5;y3=*/ h 3 5;y4=*/ h 4 5",
      paths: "M x2 t L x4 y1 L r y2 L x4 y3 L r y4 L x3 b L x1 y4 L l y3 L x1 y2 L l y1 Z",
    })
    // ── Wave shapes ──
  } else if prst == "wave" {
    Some({
      avlst: "adj1=12500;adj2=0",
      gdlst: "a1=pin 0 adj1 50000;dy=*/ h a1 100000;y1=+- t dy 0;y2=+- b 0 dy",
      paths: "M l y1 C wd4 t wd2 y1 r t L r y2 C wd2 b wd4 y2 l b Z",
    })
  } else if prst == "doubleWave" {
    Some({
      avlst: "adj1=6250;adj2=0",
      gdlst: "a1=pin 0 adj1 50000;dy=*/ h a1 100000;y1=+- t dy 0;y2=+- b 0 dy",
      paths: "M l y1 C wd6 t wd3 y1 wd2 t C wd3 y1 wd6 t r y1 L r y2 C wd3 b wd6 y2 wd2 b C wd6 y2 wd3 b l y2 Z",
    })
    // ── Tab / bracket shapes ──
  } else if prst == "bracketPair" {
    Some({
      avlst: "adj=16667",
      gdlst: "a=pin 0 adj 50000;dx=*/ ss a 100000;x1=+- r 0 dx;y1=+- b 0 dx",
      paths: "M l dx A dx dx 10800000 5400000 L l y1 A dx dx 5400000 5400000|M x1 t A dx dx 16200000 5400000 L r y1 A dx dx 0 5400000",
    })
  } else if prst == "bracePair" {
    Some({
      avlst: "adj=8333",
      gdlst: "a=pin 0 adj 25000;dx=*/ ss a 100000;dx2=*/ dx 1 2;x1=+- r 0 dx;y1=+- hd2 0 dx2;y2=+- hd2 dx2 0;y3=+- b 0 dx",
      paths: "M dx t A dx dx 10800000 5400000 L l y1 A dx2 dx2 16200000 -5400000 L l y2 A dx2 dx2 5400000 -5400000 L l y3 A dx dx 5400000 5400000|M x1 t A dx dx 16200000 5400000 L r y1 A dx2 dx2 10800000 -5400000 L r y2 A dx2 dx2 0 -5400000 L r y3 A dx dx 0 5400000",
    })
    // ── Stars (remaining) ──
  } else if prst == "star10" {
    // 10-pointed star: outer at 5 points, inner at 5 between
    Some({
      avlst: "adj=38268",
      gdlst: "a=pin 0 adj 50000;swd2=*/ wd2 a 50000;shd2=*/ hd2 a 50000;co0=cos wd2 16200000;so0=sin hd2 16200000;co1=cos wd2 19800000;so1=sin hd2 19800000;co2=cos wd2 1800000;so2=sin hd2 1800000;co3=cos wd2 5400000;so3=sin hd2 5400000;co4=cos wd2 9000000;so4=sin hd2 9000000;ox0=+- wd2 co0 0;oy0=+- hd2 so0 0;ox1=+- wd2 co1 0;oy1=+- hd2 so1 0;ox2=+- wd2 co2 0;oy2=+- hd2 so2 0;ox3=+- wd2 co3 0;oy3=+- hd2 so3 0;ox4=+- wd2 co4 0;oy4=+- hd2 so4 0;ci0=cos swd2 18000000;si0=sin shd2 18000000;ci1=cos swd2 21600000;si1=sin shd2 21600000;ci2=cos swd2 3600000;si2=sin shd2 3600000;ci3=cos swd2 7200000;si3=sin shd2 7200000;ci4=cos swd2 10800000;si4=sin shd2 10800000;ix0=+- wd2 ci0 0;iy0=+- hd2 si0 0;ix1=+- wd2 ci1 0;iy1=+- hd2 si1 0;ix2=+- wd2 ci2 0;iy2=+- hd2 si2 0;ix3=+- wd2 ci3 0;iy3=+- hd2 si3 0;ix4=+- wd2 ci4 0;iy4=+- hd2 si4 0",
      paths: "M ox0 oy0 L ix0 iy0 L ox1 oy1 L ix1 iy1 L ox2 oy2 L ix2 iy2 L ox3 oy3 L ix3 iy3 L ox4 oy4 L ix4 iy4 Z",
    })
  } else if prst == "star12" {
    Some({
      avlst: "adj=37500",
      gdlst: "a=pin 0 adj 50000;dx1=cos wd2 1800000;dy1=sin hd2 3600000;x1=+- hc 0 dx1;x3=*/ w 3 4;x4=+- hc dx1 0;y1=+- vc 0 dy1;y3=*/ h 3 4;y4=+- vc dy1 0;iwd2=*/ wd2 a 50000;ihd2=*/ hd2 a 50000;sdx1=cos iwd2 900000;sdx2=cos iwd2 2700000;sdx3=cos iwd2 4500000;sdy1=sin ihd2 4500000;sdy2=sin ihd2 2700000;sdy3=sin ihd2 900000;sx1=+- hc 0 sdx1;sx2=+- hc 0 sdx2;sx3=+- hc 0 sdx3;sx4=+- hc sdx3 0;sx5=+- hc sdx2 0;sx6=+- hc sdx1 0;sy1=+- vc 0 sdy1;sy2=+- vc 0 sdy2;sy3=+- vc 0 sdy3;sy4=+- vc sdy3 0;sy5=+- vc sdy2 0;sy6=+- vc sdy1 0;yAdj=+- vc 0 ihd2",
      paths: "M l vc L sx1 sy3 L x1 hd4 L sx2 sy2 L wd4 y1 L sx3 sy1 L hc t L sx4 sy1 L x3 y1 L sx5 sy2 L x4 hd4 L sx6 sy3 L r vc L sx6 sy4 L x4 y3 L sx5 sy5 L x3 y4 L sx4 sy6 L hc b L sx3 sy6 L wd4 y4 L sx2 sy5 L x1 y3 L sx1 sy4 Z",
    })
  } else if prst == "star16" {
    Some({
      avlst: "adj=37500",
      gdlst: "a=pin 0 adj 50000;dx1=*/ wd2 92388 100000;dx2=*/ wd2 70711 100000;dx3=*/ wd2 38268 100000;dy1=*/ hd2 92388 100000;dy2=*/ hd2 70711 100000;dy3=*/ hd2 38268 100000;x1=+- hc 0 dx1;x2=+- hc 0 dx2;x3=+- hc 0 dx3;x4=+- hc dx3 0;x5=+- hc dx2 0;x6=+- hc dx1 0;y1=+- vc 0 dy1;y2=+- vc 0 dy2;y3=+- vc 0 dy3;y4=+- vc dy3 0;y5=+- vc dy2 0;y6=+- vc dy1 0;iwd2=*/ wd2 a 50000;ihd2=*/ hd2 a 50000;sdx1=*/ iwd2 98079 100000;sdx2=*/ iwd2 83147 100000;sdx3=*/ iwd2 55557 100000;sdx4=*/ iwd2 19509 100000;sdy1=*/ ihd2 98079 100000;sdy2=*/ ihd2 83147 100000;sdy3=*/ ihd2 55557 100000;sdy4=*/ ihd2 19509 100000;sx1=+- hc 0 sdx1;sx2=+- hc 0 sdx2;sx3=+- hc 0 sdx3;sx4=+- hc 0 sdx4;sx5=+- hc sdx4 0;sx6=+- hc sdx3 0;sx7=+- hc sdx2 0;sx8=+- hc sdx1 0;sy1=+- vc 0 sdy1;sy2=+- vc 0 sdy2;sy3=+- vc 0 sdy3;sy4=+- vc 0 sdy4;sy5=+- vc sdy4 0;sy6=+- vc sdy3 0;sy7=+- vc sdy2 0;sy8=+- vc sdy1 0;idx=cos iwd2 2700000;idy=sin ihd2 2700000;il=+- hc 0 idx;it=+- vc 0 idy;ir=+- hc idx 0;ib=+- vc idy 0;yAdj=+- vc 0 ihd2",
      paths: "M l vc L sx1 sy4 L x1 y3 L sx2 sy3 L x2 y2 L sx3 sy2 L x3 y1 L sx4 sy1 L hc t L sx5 sy1 L x4 y1 L sx6 sy2 L x5 y2 L sx7 sy3 L x6 y3 L sx8 sy4 L r vc L sx8 sy5 L x6 y4 L sx7 sy6 L x5 y5 L sx6 sy7 L x4 y6 L sx5 sy8 L hc b L sx4 sy8 L x3 y6 L sx3 sy7 L x2 y5 L sx2 sy6 L x1 y4 L sx1 sy5 Z",
    })
  } else if prst == "star24" {
    Some({
      avlst: "adj=37500",
      gdlst: "a=pin 0 adj 50000;dx1=cos wd2 900000;dx2=cos wd2 1800000;dx3=cos wd2 2700000;dx4=val wd4;dx5=cos wd2 4500000;dy1=sin hd2 4500000;dy2=sin hd2 3600000;dy3=sin hd2 2700000;dy4=val hd4;dy5=sin hd2 900000;x1=+- hc 0 dx1;x2=+- hc 0 dx2;x3=+- hc 0 dx3;x4=+- hc 0 dx4;x5=+- hc 0 dx5;x6=+- hc dx5 0;x7=+- hc dx4 0;x8=+- hc dx3 0;x9=+- hc dx2 0;x10=+- hc dx1 0;y1=+- vc 0 dy1;y2=+- vc 0 dy2;y3=+- vc 0 dy3;y4=+- vc 0 dy4;y5=+- vc 0 dy5;y6=+- vc dy5 0;y7=+- vc dy4 0;y8=+- vc dy3 0;y9=+- vc dy2 0;y10=+- vc dy1 0;iwd2=*/ wd2 a 50000;ihd2=*/ hd2 a 50000;sdx1=*/ iwd2 99144 100000;sdx2=*/ iwd2 92388 100000;sdx3=*/ iwd2 79335 100000;sdx4=*/ iwd2 60876 100000;sdx5=*/ iwd2 38268 100000;sdx6=*/ iwd2 13053 100000;sdy1=*/ ihd2 99144 100000;sdy2=*/ ihd2 92388 100000;sdy3=*/ ihd2 79335 100000;sdy4=*/ ihd2 60876 100000;sdy5=*/ ihd2 38268 100000;sdy6=*/ ihd2 13053 100000;sx1=+- hc 0 sdx1;sx2=+- hc 0 sdx2;sx3=+- hc 0 sdx3;sx4=+- hc 0 sdx4;sx5=+- hc 0 sdx5;sx6=+- hc 0 sdx6;sx7=+- hc sdx6 0;sx8=+- hc sdx5 0;sx9=+- hc sdx4 0;sx10=+- hc sdx3 0;sx11=+- hc sdx2 0;sx12=+- hc sdx1 0;sy1=+- vc 0 sdy1;sy2=+- vc 0 sdy2;sy3=+- vc 0 sdy3;sy4=+- vc 0 sdy4;sy5=+- vc 0 sdy5;sy6=+- vc 0 sdy6;sy7=+- vc sdy6 0;sy8=+- vc sdy5 0;sy9=+- vc sdy4 0;sy10=+- vc sdy3 0;sy11=+- vc sdy2 0;sy12=+- vc sdy1 0;idx=cos iwd2 2700000;idy=sin ihd2 2700000;il=+- hc 0 idx;it=+- vc 0 idy;ir=+- hc idx 0;ib=+- vc idy 0;yAdj=+- vc 0 ihd2",
      paths: "M l vc L sx1 sy6 L x1 y5 L sx2 sy5 L x2 y4 L sx3 sy4 L x3 y3 L sx4 sy3 L x4 y2 L sx5 sy2 L x5 y1 L sx6 sy1 L hc t L sx7 sy1 L x6 y1 L sx8 sy2 L x7 y2 L sx9 sy3 L x8 y3 L sx10 sy4 L x9 y4 L sx11 sy5 L x10 y5 L sx12 sy6 L r vc L sx12 sy7 L x10 y6 L sx11 sy8 L x9 y7 L sx10 sy9 L x8 y8 L sx9 sy10 L x7 y9 L sx8 sy11 L x6 y10 L sx7 sy12 L hc b L sx6 sy12 L x5 y10 L sx5 sy11 L x4 y9 L sx4 sy10 L x3 y8 L sx3 sy9 L x2 y7 L sx2 sy8 L x1 y6 L sx1 sy7 Z",
    })
  } else if prst == "star32" {
    Some({
      avlst: "adj=37500",
      gdlst: "a=pin 0 adj 50000;dx1=*/ wd2 98079 100000;dx2=*/ wd2 92388 100000;dx3=*/ wd2 83147 100000;dx4=cos wd2 2700000;dx5=*/ wd2 55557 100000;dx6=*/ wd2 38268 100000;dx7=*/ wd2 19509 100000;dy1=*/ hd2 98079 100000;dy2=*/ hd2 92388 100000;dy3=*/ hd2 83147 100000;dy4=sin hd2 2700000;dy5=*/ hd2 55557 100000;dy6=*/ hd2 38268 100000;dy7=*/ hd2 19509 100000;x1=+- hc 0 dx1;x2=+- hc 0 dx2;x3=+- hc 0 dx3;x4=+- hc 0 dx4;x5=+- hc 0 dx5;x6=+- hc 0 dx6;x7=+- hc 0 dx7;x8=+- hc dx7 0;x9=+- hc dx6 0;x10=+- hc dx5 0;x11=+- hc dx4 0;x12=+- hc dx3 0;x13=+- hc dx2 0;x14=+- hc dx1 0;y1=+- vc 0 dy1;y2=+- vc 0 dy2;y3=+- vc 0 dy3;y4=+- vc 0 dy4;y5=+- vc 0 dy5;y6=+- vc 0 dy6;y7=+- vc 0 dy7;y8=+- vc dy7 0;y9=+- vc dy6 0;y10=+- vc dy5 0;y11=+- vc dy4 0;y12=+- vc dy3 0;y13=+- vc dy2 0;y14=+- vc dy1 0;iwd2=*/ wd2 a 50000;ihd2=*/ hd2 a 50000;sdx1=*/ iwd2 99518 100000;sdx2=*/ iwd2 95694 100000;sdx3=*/ iwd2 88192 100000;sdx4=*/ iwd2 77301 100000;sdx5=*/ iwd2 63439 100000;sdx6=*/ iwd2 47140 100000;sdx7=*/ iwd2 29028 100000;sdx8=*/ iwd2 9802 100000;sdy1=*/ ihd2 99518 100000;sdy2=*/ ihd2 95694 100000;sdy3=*/ ihd2 88192 100000;sdy4=*/ ihd2 77301 100000;sdy5=*/ ihd2 63439 100000;sdy6=*/ ihd2 47140 100000;sdy7=*/ ihd2 29028 100000;sdy8=*/ ihd2 9802 100000;sx1=+- hc 0 sdx1;sx2=+- hc 0 sdx2;sx3=+- hc 0 sdx3;sx4=+- hc 0 sdx4;sx5=+- hc 0 sdx5;sx6=+- hc 0 sdx6;sx7=+- hc 0 sdx7;sx8=+- hc 0 sdx8;sx9=+- hc sdx8 0;sx10=+- hc sdx7 0;sx11=+- hc sdx6 0;sx12=+- hc sdx5 0;sx13=+- hc sdx4 0;sx14=+- hc sdx3 0;sx15=+- hc sdx2 0;sx16=+- hc sdx1 0;sy1=+- vc 0 sdy1;sy2=+- vc 0 sdy2;sy3=+- vc 0 sdy3;sy4=+- vc 0 sdy4;sy5=+- vc 0 sdy5;sy6=+- vc 0 sdy6;sy7=+- vc 0 sdy7;sy8=+- vc 0 sdy8;sy9=+- vc sdy8 0;sy10=+- vc sdy7 0;sy11=+- vc sdy6 0;sy12=+- vc sdy5 0;sy13=+- vc sdy4 0;sy14=+- vc sdy3 0;sy15=+- vc sdy2 0;sy16=+- vc sdy1 0;idx=cos iwd2 2700000;idy=sin ihd2 2700000;il=+- hc 0 idx;it=+- vc 0 idy;ir=+- hc idx 0;ib=+- vc idy 0;yAdj=+- vc 0 ihd2",
      paths: "M l vc L sx1 sy8 L x1 y7 L sx2 sy7 L x2 y6 L sx3 sy6 L x3 y5 L sx4 sy5 L x4 y4 L sx5 sy4 L x5 y3 L sx6 sy3 L x6 y2 L sx7 sy2 L x7 y1 L sx8 sy1 L hc t L sx9 sy1 L x8 y1 L sx10 sy2 L x9 y2 L sx11 sy3 L x10 y3 L sx12 sy4 L x11 y4 L sx13 sy5 L x12 y5 L sx14 sy6 L x13 y6 L sx15 sy7 L x14 y7 L sx16 sy8 L r vc L sx16 sy9 L x14 y8 L sx15 sy10 L x13 y9 L sx14 sy11 L x12 y10 L sx13 sy12 L x11 y11 L sx12 sy13 L x10 y12 L sx11 sy14 L x9 y13 L sx10 sy15 L x8 y14 L sx9 sy16 L hc b L sx8 sy16 L x7 y14 L sx7 sy15 L x6 y13 L sx6 sy14 L x5 y12 L sx5 sy13 L x4 y11 L sx4 sy12 L x3 y10 L sx3 sy11 L x2 y9 L sx2 sy10 L x1 y8 L sx1 sy9 Z",
    })
    // ── Remaining flowchart shapes ──
  } else if prst == "flowChartAlternateProcess" {
    Some({
      avlst: "",
      gdlst: "cr=*/ ss 1 10;dcr=+- r 0 cr;dbcr=+- b 0 cr",
      paths: "M cr t L dcr t A cr cr 16200000 5400000 L r dbcr A cr cr 0 5400000 L cr b A cr cr 5400000 5400000 L l cr A cr cr 10800000 5400000 Z",
    })
  } else if prst == "flowChartMultidocument" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 10;x2=*/ w 9 10;y1=*/ h 1 10;y2=*/ h 2 10;y3=*/ h 8 10;cy=*/ h 9 10",
      paths: "M l y2 L x2 y2 L x2 y1 L r y1 L r y3 C r cy wd2 b wd2 y3 C wd2 cy l b l y3 Z|M x2 y2 L x2 y1|M x1 y1 L x1 t L r t L r y1",
    })
  } else if prst == "flowChartInternalStorage" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 8;y1=*/ h 1 8",
      paths: "M l t L r t L r b L l b Z|M x1 t L x1 b|M l y1 L r y1",
    })
  } else if prst == "flowChartMagneticDisk" {
    Some({
      avlst: "",
      gdlst: "dy=*/ h 1 6;y1=+- t dy 0;y2=+- b 0 dy",
      paths: "M l y1 A wd2 dy 10800000 -10800000 L r y2 A wd2 dy 0 10800000 Z|M l y1 A wd2 dy 10800000 10800000",
    })
  } else if prst == "flowChartMagneticDrum" {
    Some({
      avlst: "",
      gdlst: "dx=*/ w 1 6;x1=+- l dx 0;x2=+- r 0 dx",
      paths: "M x1 t L x2 t A dx hd2 16200000 -10800000 L x1 b A dx hd2 5400000 10800000 Z|M x2 t A dx hd2 16200000 10800000",
    })
  } else if prst == "flowChartMagneticTape" {
    Some({
      avlst: "",
      gdlst: "idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0;ang1=at2 w h",
      paths: "M hc b A wd2 hd2 cd4 cd4 A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 ang1 L r ib L r b Z",
    })
  } else if prst == "flowChartOnlineStorage" {
    Some({
      avlst: "",
      gdlst: "dx=*/ w 1 6;x1=+- r 0 dx",
      paths: "M dx t L r t A dx hd2 16200000 -10800000 L dx b A dx hd2 5400000 -10800000 Z",
    })
  } else if prst == "flowChartDisplay" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 6;x2=*/ w 5 6",
      paths: "M x1 t L x2 t A x1 hd2 16200000 10800000 L x1 b L l hd2 Z",
    })
  } else if prst == "flowChartPunchedCard" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 1 5;y1=*/ h 1 5",
      paths: "M l y1 L x1 t L r t L r b L l b Z",
    })
  } else if prst == "flowChartPunchedTape" {
    Some({
      avlst: "",
      gdlst: "y1=*/ h 1 10;y2=*/ h 9 10;cy1=*/ h 1 5;cy2=*/ h 4 5",
      paths: "M l y1 C wd4 0 wd2 cy1 r y1 L r y2 C wd2 b wd4 cy2 l y2 Z",
    })
  } else if prst == "flowChartSummingJunction" {
    Some({
      avlst: "",
      gdlst: "idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z|M il it L ir ib M ir it L il ib|M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z",
    })
  } else if prst == "flowChartOr" {
    Some({
      avlst: "",
      gdlst: "idx=cos wd2 2700000;idy=sin hd2 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z|M hc t L hc b M l vc L r vc|M l vc A wd2 hd2 cd2 cd4 A wd2 hd2 3cd4 cd4 A wd2 hd2 0 cd4 A wd2 hd2 cd4 cd4 Z",
    })
  } else if prst == "flowChartCollate" {
    Some({
      avlst: "",
      gdlst: "",
      paths: "M l t L r t L l b L r b Z|M l hd2 L r hd2",
    })
    // ── Remaining callouts ──
  } else if prst == "cloudCallout" {
    Some({
      avlst: "adj1=-20833;adj2=62500",
      gdlst: "dxPos=*/ w adj1 100000;dyPos=*/ h adj2 100000;xPos=+- hc dxPos 0;yPos=+- vc dyPos 0;ht=cat2 hd2 dxPos dyPos;wt=sat2 wd2 dxPos dyPos;g2=cat2 wd2 ht wt;g3=sat2 hd2 ht wt;g4=+- hc g2 0;g5=+- vc g3 0;g6=+- g4 0 xPos;g7=+- g5 0 yPos;g8=mod g6 g7 0;g9=*/ ss 6600 21600;g10=+- g8 0 g9;g11=*/ g10 1 3;g12=*/ ss 1800 21600;g13=+- g11 g12 0;g14=*/ g13 g6 g8;g15=*/ g13 g7 g8;g16=+- g14 xPos 0;g17=+- g15 yPos 0;g18=*/ ss 4800 21600;g19=*/ g11 2 1;g20=+- g18 g19 0;g21=*/ g20 g6 g8;g22=*/ g20 g7 g8;g23=+- g21 xPos 0;g24=+- g22 yPos 0;g25=*/ ss 1200 21600;g26=*/ ss 600 21600;x23=+- xPos g26 0;x24=+- g16 g25 0;x25=+- g23 g12 0;il=*/ w 2977 21600;it=*/ h 3262 21600;ir=*/ w 17087 21600;ib=*/ h 17337 21600;g27=*/ w 67 21600;g28=*/ h 21577 21600;g29=*/ w 21582 21600;g30=*/ h 1235 21600;pang=at2 dxPos dyPos",
      paths: "W 43200 43200 M 3900 14370 A 6753 9190 -11429249 7426832 A 5333 7267 -8646143 5396714 A 4365 5945 -8748475 5983381 A 4857 6595 -7859164 7034504 A 5333 7273 -4722533 6541615 A 6775 9220 -2776035 7816140 A 5785 7867 37501 6842000 A 6752 9215 1347096 6910353 A 7720 10543 3974558 4542661 A 4360 5918 -16496525 8804134 A 4345 5945 -14809710 9151131 Z|M x23 yPos A g26 g26 0 21600000 Z|M x24 g17 A g25 g25 0 21600000 Z|M x25 g24 A g12 g12 0 21600000 Z|W 43200 43200 M 4693 26177 A 4345 5945 5204520 1585770 M 6928 34899 A 4360 5918 4416628 686848 M 16478 39090 A 6752 9215 8257449 844866 M 28827 34751 A 6752 9215 387196 959901 M 34129 22954 A 5785 7867 -4217541 4255042 M 41798 15354 A 5333 7273 1819082 1665090 M 38324 5426 A 4857 6595 -824660 891534 M 29078 3952 A 4857 6595 -8950887 1091722 M 22141 4720 A 4365 5945 -9809656 1061181 M 14000 5192 A 6753 9190 -4002417 739161 M 4127 15789 A 6753 9190 9459261 711490",
    })
  } else if prst == "borderCallout1" || prst == "callout1" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=112500;adj4=-38333",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0",
      paths: "M l t L r t L r b L l b Z|M x1 y1 L x2 y2",
    })
  } else if prst == "borderCallout2" || prst == "callout2" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=18750;adj4=-16667;adj5=112500;adj6=-46667",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;a5=*/ h adj5 100000;a6=*/ w adj6 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0;y3=+- hd2 a5 0;x3=+- wd2 a6 0",
      paths: "M l t L r t L r b L l b Z|M x1 y1 L x2 y2 L x3 y3",
    })
  } else if prst == "borderCallout3" || prst == "callout3" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=18750;adj4=-16667;adj5=100000;adj6=-16667;adj7=112500;adj8=-46667",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;a5=*/ h adj5 100000;a6=*/ w adj6 100000;a7=*/ h adj7 100000;a8=*/ w adj8 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0;y3=+- hd2 a5 0;x3=+- wd2 a6 0;y4=+- hd2 a7 0;x4=+- wd2 a8 0",
      paths: "M l t L r t L r b L l b Z|M x1 y1 L x2 y2 L x3 y3 L x4 y4",
    })
  } else if prst == "accentCallout1" || prst == "accentBorderCallout1" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=112500;adj4=-38333",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0",
      paths: "M l t L r t L r b L l b Z|M x1 t L x1 b|M x1 y1 L x2 y2",
    })
  } else if prst == "accentCallout2" || prst == "accentBorderCallout2" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=18750;adj4=-16667;adj5=112500;adj6=-46667",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;a5=*/ h adj5 100000;a6=*/ w adj6 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0;y3=+- hd2 a5 0;x3=+- wd2 a6 0",
      paths: "M l t L r t L r b L l b Z|M x1 t L x1 b|M x1 y1 L x2 y2 L x3 y3",
    })
  } else if prst == "accentCallout3" || prst == "accentBorderCallout3" {
    Some({
      avlst: "adj1=18750;adj2=-8333;adj3=18750;adj4=-16667;adj5=100000;adj6=-16667;adj7=112500;adj8=-46667",
      gdlst: "a1=*/ h adj1 100000;a2=*/ w adj2 100000;a3=*/ h adj3 100000;a4=*/ w adj4 100000;a5=*/ h adj5 100000;a6=*/ w adj6 100000;a7=*/ h adj7 100000;a8=*/ w adj8 100000;y1=+- hd2 a1 0;x1=+- wd2 a2 0;y2=+- hd2 a3 0;x2=+- wd2 a4 0;y3=+- hd2 a5 0;x3=+- wd2 a6 0;y4=+- hd2 a7 0;x4=+- wd2 a8 0",
      paths: "M l t L r t L r b L l b Z|M x1 t L x1 b|M x1 y1 L x2 y2 L x3 y3 L x4 y4",
    })
    // ── Remaining block arrows ──
  } else if prst == "bentArrow" {
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000;adj4=43750",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;a4=pin 0 adj4 100000;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx=*/ w a2 100000;x1=+- r 0 dx;x2=*/ w a4 100000;dy2=*/ h a3 100000;y3=+- t dy2 0",
      paths: "M l y2 L x2 y2 L x2 y1 L x1 y1 L x1 t L r y3 L x1 dy2 L x1 y2 L l y2 Z",
    })
  } else if prst == "bentUpArrow" {
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;dy=*/ h a2 100000;y1=+- b 0 dy;dx2=*/ w a3 100000;x3=+- r 0 dx2",
      paths: "M l y1 L x3 y1 L x3 x1 L x1 x1 L x1 t L r x1 L x1 dy L x1 y1 L l y1 Z",
    })
  } else if prst == "curvedRightArrow" || prst == "curvedLeftArrow" {
    Some({
      avlst: "adj1=25000;adj2=50000;adj3=25000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 100000;a3=pin 0 adj3 50000;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx=*/ w a3 100000;x1=+- r 0 dx",
      paths: "M l y1 C wd2 t r t r hd2 L x1 hd2 L x1 t L r hd2 L x1 b L x1 hd2 C r hd2 r b wd2 b L l y2 C l y2 l hd2 l y1 Z",
    })
  } else if prst == "curvedUpArrow" || prst == "curvedDownArrow" {
    Some({
      avlst: "adj1=25000;adj2=50000;adj3=25000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 100000;a3=pin 0 adj3 50000;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;dy=*/ h a3 100000;y1=+- b 0 dy",
      paths: "M x1 b C t l t wd2 wd2 t L wd2 y1 L l y1 L wd2 b L r y1 L wd2 y1 C wd2 t r t r wd2 L x2 b Z",
    })
  } else if prst == "circularArrow" {
    Some({
      avlst: "adj1=12500;adj2=1142319;adj3=20457681;adj4=10800000;adj5=12500",
      gdlst: "a5=pin 0 adj5 25000;maxAdj1=*/ a5 2 1;a1=pin 0 adj1 maxAdj1;enAng=pin 1 adj3 21599999;stAng=pin 0 adj4 21599999;th=*/ ss a1 100000;thh=*/ ss a5 100000;th2=*/ th 1 2;rw1=+- wd2 th2 thh;rh1=+- hd2 th2 thh;rw2=+- rw1 0 th;rh2=+- rh1 0 th;rw3=+- rw2 th2 0;rh3=+- rh2 th2 0;wtH=sin rw3 enAng;htH=cos rh3 enAng;dxH=cat2 rw3 htH wtH;dyH=sat2 rh3 htH wtH;xH=+- hc dxH 0;yH=+- vc dyH 0;rI=min rw2 rh2;u1=*/ dxH dxH 1;u2=*/ dyH dyH 1;u3=*/ rI rI 1;u4=+- u1 0 u3;u5=+- u2 0 u3;u6=*/ u4 u5 u1;u7=*/ u6 1 u2;u8=+- 1 0 u7;u9=sqrt u8;u10=*/ u4 1 dxH;u11=*/ u10 1 dyH;u12=+/ 1 u9 u11;u13=at2 1 u12;u14=+- u13 21600000 0;u15=?: u13 u13 u14;u16=+- u15 0 enAng;u17=+- u16 21600000 0;u18=?: u16 u16 u17;u19=+- u18 0 cd2;u20=+- u18 0 21600000;u21=?: u19 u20 u18;maxAng=abs u21;aAng=pin 0 adj2 maxAng;ptAng=+- enAng aAng 0;wtA=sin rw3 ptAng;htA=cos rh3 ptAng;dxA=cat2 rw3 htA wtA;dyA=sat2 rh3 htA wtA;xA=+- hc dxA 0;yA=+- vc dyA 0;wtE=sin rw1 stAng;htE=cos rh1 stAng;dxE=cat2 rw1 htE wtE;dyE=sat2 rh1 htE wtE;xE=+- hc dxE 0;yE=+- vc dyE 0;dxG=cos thh ptAng;dyG=sin thh ptAng;xG=+- xH dxG 0;yG=+- yH dyG 0;dxB=cos thh ptAng;dyB=sin thh ptAng;xB=+- xH 0 dxB 0;yB=+- yH 0 dyB 0;sx1=+- xB 0 hc;sy1=+- yB 0 vc;sx2=+- xG 0 hc;sy2=+- yG 0 vc;rO=min rw1 rh1;x1O=*/ sx1 rO rw1;y1O=*/ sy1 rO rh1;x2O=*/ sx2 rO rw1;y2O=*/ sy2 rO rh1;dxO=+- x2O 0 x1O;dyO=+- y2O 0 y1O;dO=mod dxO dyO 0;q1=*/ x1O y2O 1;q2=*/ x2O y1O 1;DO=+- q1 0 q2;q3=*/ rO rO 1;q4=*/ dO dO 1;q5=*/ q3 q4 1;q6=*/ DO DO 1;q7=+- q5 0 q6;q8=max q7 0;sdelO=sqrt q8;ndyO=*/ dyO -1 1;sdyO=?: ndyO -1 1;q9=*/ sdyO dxO 1;q10=*/ q9 sdelO 1;q11=*/ DO dyO 1;dxF1=+/ q11 q10 q4;q12=+- q11 0 q10;dxF2=*/ q12 1 q4;adyO=abs dyO;q13=*/ adyO sdelO 1;q14=*/ DO dxO -1;dyF1=+/ q14 q13 q4;q15=+- q14 0 q13;dyF2=*/ q15 1 q4;q16=+- x2O 0 dxF1;q17=+- x2O 0 dxF2;q18=+- y2O 0 dyF1;q19=+- y2O 0 dyF2;q20=mod q16 q18 0;q21=mod q17 q19 0;q22=+- q21 0 q20;dxF=?: q22 dxF1 dxF2;dyF=?: q22 dyF1 dyF2;sdxF=*/ dxF rw1 rO;sdyF=*/ dyF rh1 rO;xF=+- hc sdxF 0;yF=+- vc sdyF 0;x1I=*/ sx1 rI rw2;y1I=*/ sy1 rI rh2;x2I=*/ sx2 rI rw2;y2I=*/ sy2 rI rh2;dxI=+- x2I 0 x1I;dyI=+- y2I 0 y1I;dI=mod dxI dyI 0;v1=*/ x1I y2I 1;v2=*/ x2I y1I 1;DI=+- v1 0 v2;v3=*/ rI rI 1;v4=*/ dI dI 1;v5=*/ v3 v4 1;v6=*/ DI DI 1;v7=+- v5 0 v6;v8=max v7 0;sdelI=sqrt v8;v9=*/ sdyO dxI 1;v10=*/ v9 sdelI 1;v11=*/ DI dyI 1;dxC1=+/ v11 v10 v4;v12=+- v11 0 v10;dxC2=*/ v12 1 v4;adyI=abs dyI;v13=*/ adyI sdelI 1;v14=*/ DI dxI -1;dyC1=+/ v14 v13 v4;v15=+- v14 0 v13;dyC2=*/ v15 1 v4;v16=+- x1I 0 dxC1;v17=+- x1I 0 dxC2;v18=+- y1I 0 dyC1;v19=+- y1I 0 dyC2;v20=mod v16 v18 0;v21=mod v17 v19 0;v22=+- v21 0 v20;dxC=?: v22 dxC1 dxC2;dyC=?: v22 dyC1 dyC2;sdxC=*/ dxC rw2 rI;sdyC=*/ dyC rh2 rI;xC=+- hc sdxC 0;yC=+- vc sdyC 0;ist0=at2 sdxC sdyC;ist1=+- ist0 21600000 0;istAng=?: ist0 ist0 ist1;isw1=+- stAng 0 istAng;isw2=+- isw1 0 21600000;iswAng=?: isw1 isw2 isw1;p1=+- xF 0 xC;p2=+- yF 0 yC;p3=mod p1 p2 0;p4=*/ p3 1 2;p5=+- p4 0 thh;xGp=?: p5 xF xG;yGp=?: p5 yF yG;xBp=?: p5 xC xB;yBp=?: p5 yC yB;en0=at2 sdxF sdyF;en1=+- en0 21600000 0;en2=?: en0 en0 en1;sw0=+- en2 0 stAng;sw1=+- sw0 21600000 0;swAng=?: sw0 sw0 sw1;wtI=sin rw3 stAng;htI=cos rh3 stAng;dxI=cat2 rw3 htI wtI;dyI=sat2 rh3 htI wtI;xI=+- hc dxI 0;yI=+- vc dyI 0;aI=+- stAng 0 cd4;aA=+- ptAng cd4 0;aB=+- ptAng cd2 0;idx=cos rw1 2700000;idy=sin rh1 2700000;il=+- hc 0 idx;ir=+- hc idx 0;it=+- vc 0 idy;ib=+- vc idy 0",
      paths: "M xE yE A rw1 rh1 stAng swAng L xGp yGp L xA yA L xBp yBp L xC yC A rw2 rh2 istAng iswAng Z",
    })
  } else if prst == "uturnArrow" {
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000;adj4=43750;adj5=75000",
      gdlst: "a2=pin 0 adj2 25000;maxAdj1=*/ a2 2 1;a1=pin 0 adj1 maxAdj1;q2=*/ a1 ss h;q3=+- 100000 0 q2;maxAdj3=*/ q3 h ss;a3=pin 0 adj3 maxAdj3;q1=+- a3 a1 0;minAdj5=*/ q1 ss h;a5=pin minAdj5 adj5 100000;th=*/ ss a1 100000;aw2=*/ ss a2 100000;th2=*/ th 1 2;dh2=+- aw2 0 th2;y5=*/ h a5 100000;ah=*/ ss a3 100000;y4=+- y5 0 ah;x9=+- r 0 dh2;bw=*/ x9 1 2;bs=min bw y4;maxAdj4=*/ bs 100000 ss;a4=pin 0 adj4 maxAdj4;bd=*/ ss a4 100000;bd3=+- bd 0 th;bd2=max bd3 0;x3=+- th bd2 0;x8=+- r 0 aw2;x6=+- x8 0 aw2;x7=+- x6 dh2 0;x4=+- x9 0 bd;x5=+- x7 0 bd2;cx=+/ th x7 2",
      paths: "M l b L l bd A bd bd cd2 cd4 L x4 t A bd bd 3cd4 cd4 L x9 y4 L r y4 L x8 y5 L x6 y4 L x7 y4 L x7 x3 A bd2 bd2 0 -5400000 L x3 th A bd2 bd2 3cd4 -5400000 L th b Z",
    })
  } else if prst == "leftUpArrow" {
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx=*/ w a2 100000;x1=+- l dx 0;dx2=*/ w a3 100000;x2=+- l dx2 0",
      paths: "M l y1 L x2 t L x2 y1 L x1 y1 L x1 y2 L r y2 L r b L x1 b L x1 y2 L l y2 Z",
    })
  } else if prst == "upDownArrowCallout" {
    Some({
      avlst: "adj1=25000;adj2=25000;adj3=25000;adj4=48123",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;a4=pin 0 adj4 50000;dy1=*/ h a2 100000;y1=+- t dy1 0;y2=+- b 0 dy1;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0;dy2=*/ h a4 100000;y3=+- t dy2 0;y4=+- b 0 dy2;dx2=*/ w a3 200000;x3=+- wd2 0 dx2;x4=+- wd2 dx2 0",
      paths: "M wd2 t L r y1 L x2 y1 L x2 y3 L x4 y3 L x4 y4 L x2 y4 L x2 y2 L r y2 L wd2 b L l y2 L x1 y2 L x1 y4 L x3 y4 L x3 y3 L x1 y3 L x1 y1 L l y1 Z",
    })
  } else if prst == "quadArrowCallout" {
    Some({
      avlst: "adj1=18515;adj2=18515;adj3=18515;adj4=48123",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;a4=pin 0 adj4 50000;dx=*/ w a2 100000;dy=*/ h a2 100000;dh=*/ w a1 200000;dv=*/ h a1 200000;ah=*/ w a3 100000;av=*/ h a3 100000;x1=+- wd2 0 dh;x2=+- wd2 dh 0;y1=+- hd2 0 dv;y2=+- hd2 dv 0;xe=+- r 0 dx;ye=+- b 0 dy;ddx=*/ w a4 200000;ddy=*/ h a4 200000;x3=+- wd2 0 ddx;x4=+- wd2 ddx 0;y3=+- hd2 0 ddy;y4=+- hd2 ddy 0",
      paths: "M wd2 t L r dy L x2 dy L x2 y3 L x4 y3 L x4 y1 L xe y1 L xe hd2 L r hd2 L xe y2 L x4 y2 L x4 y4 L x2 y4 L x2 ye L wd2 b L l ye L x1 ye L x1 y4 L x3 y4 L x3 y2 L dx y2 L dx hd2 L l hd2 L dx y1 L x3 y1 L x3 y3 L x1 y3 L x1 dy L wd2 t Z",
    })
    // ── Remaining other shapes ──
  } else if prst == "cloud" {
    // Cloud shape using bezier curves
    Some({
      avlst: "",
      gdlst: "",
      paths: "M wd4 hd2 C wd4 hd4 wd3 hd5 wd2 hd4 C wd2 hd5 wd3 hd4 wd4 hd4 C wd6 hd6 wd8 hd4 l hd2 C l hd3 wd8 hd2 wd4 hd2 Z|M wd2 hd4 C wd2 hd6 wd4 0 wd2 0 C wd2 hd6 r hd6 r hd4 C r hd3 r hd2 wd2 hd2 L wd4 hd2 C wd4 hd3 wd2 hd3 wd2 hd4 Z|M wd4 hd2 C wd4 hd2 wd6 b wd4 b C wd3 b wd2 b wd2 hd2 L r hd2 C r hd2 r b wd2 b C wd3 b wd4 hd2 wd4 hd2 Z",
    })
  } else if prst == "leftBracket" {
    Some({
      avlst: "adj=8333",
      gdlst: "a=pin 0 adj 50000;dy=*/ h a 100000;y1=+- b 0 dy",
      paths: "M r t A wd2 dy 16200000 -5400000 L l y1 A wd2 dy 5400000 -5400000",
    })
  } else if prst == "rightBracket" {
    Some({
      avlst: "adj=8333",
      gdlst: "a=pin 0 adj 50000;dy=*/ h a 100000;y1=+- b 0 dy",
      paths: "M l t A wd2 dy 16200000 5400000 L r y1 A wd2 dy 5400000 5400000",
    })
  } else if prst == "leftBrace" {
    Some({
      avlst: "adj1=8333;adj2=50000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 100000;dy=*/ h a1 100000;dy2=*/ dy 1 2;ym=*/ h a2 100000;y1=+- ym 0 dy2;y2=+- ym dy2 0;y3=+- b 0 dy",
      paths: "M r t A wd2 dy 16200000 -5400000 L l y1 A wd4 dy2 0 -5400000 L l y2 A wd4 dy2 10800000 -5400000 L l y3 A wd2 dy 5400000 -5400000",
    })
  } else if prst == "rightBrace" {
    Some({
      avlst: "adj1=8333;adj2=50000",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 100000;dy=*/ h a1 100000;dy2=*/ dy 1 2;ym=*/ h a2 100000;y1=+- ym 0 dy2;y2=+- ym dy2 0;y3=+- b 0 dy",
      paths: "M l t A wd2 dy 16200000 5400000 L r y1 A wd4 dy2 10800000 5400000 L r y2 A wd4 dy2 0 5400000 L r y3 A wd2 dy 5400000 5400000",
    })
  } else if prst == "horzScroll" {
    Some({
      avlst: "adj=12500",
      gdlst: "a=pin 0 adj 25000;dx=*/ w a 100000;x1=+- r 0 dx;dx2=*/ dx 1 2;x2=+- l dx2 0;x3=+- r 0 dx2;y1=+- t dx 0;y2=+- b 0 dx;y3=+- t dx2 0;y4=+- b 0 dx2",
      paths: "M x2 y1 L x1 y1 A dx2 dx2 16200000 5400000 L r y4 A dx2 dx2 0 -5400000 L dx y2 A dx2 dx2 5400000 5400000 L l y3 A dx2 dx2 10800000 -5400000 Z|M l y3 A dx2 dx2 10800000 10800000|M x1 y1 A dx2 dx2 16200000 -10800000",
    })
  } else if prst == "vertScroll" {
    Some({
      avlst: "adj=12500",
      gdlst: "a=pin 0 adj 25000;dy=*/ h a 100000;y1=+- b 0 dy;dy2=*/ dy 1 2;x1=+- l dy 0;x2=+- r 0 dy;x3=+- l dy2 0;x4=+- r 0 dy2;y2=+- t dy2 0;y3=+- b 0 dy2",
      paths: "M x1 t L x1 y1 A dy2 dy2 5400000 5400000 L x3 b A dy2 dy2 10800000 -5400000 L x2 b L x2 dy A dy2 dy2 16200000 5400000 L x4 t A dy2 dy2 0 -5400000 Z|M x4 t A dy2 dy2 0 10800000|M x1 y1 A dy2 dy2 5400000 -10800000",
    })
  } else if prst == "snip2DiagRect" {
    Some({
      avlst: "adj1=0;adj2=16667",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;dx1=*/ ss a1 100000;dx2=*/ ss a2 100000;x1=+- r 0 dx1;x2=+- r 0 dx2;y2=+- b 0 dx2;y1=+- b 0 dx1",
      paths: "M dx1 t L x2 t L r dx2 L r y1 L x1 b L dx2 b L l y2 L l dx1 Z",
    })
  } else if prst == "plaqueTabs" {
    Some({
      avlst: "",
      gdlst: "dx=*/ w 1 5;dy=*/ h 1 5;x1=+- r 0 dx;y1=+- b 0 dy",
      paths: "M l t A dx dy 10800000 -5400000|M x1 t A dx dy 16200000 5400000|M l y1 A dx dy 5400000 -5400000|M x1 y1 A dx dy 0 5400000",
    })
  } else if prst == "mathNotEqual" {
    Some({
      avlst: "adj1=23520;adj2=6600;adj3=11760",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 50000;a3=pin 0 adj3 100000;dy1=*/ h a1 200000;dy2=*/ h a2 100000;dy3=*/ h a3 200000;y1=+- hd2 0 dy3;y2=+- y1 0 dy1;y3=+- hd2 dy3 0;y4=+- y3 dy1 0;dx=*/ w 1 8",
      paths: "M l y2 L r y2 L r y1 L l y1 Z|M l y3 L r y3 L r y4 L l y4 Z|M dx t L r y3 L r y4 L dx hd2 L l y2 L l y1 Z",
    })
    // ── Action buttons ──
  } else if prst == "actionButtonBlank" {
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L r b L l b Z", })
  } else if prst == "actionButtonHome" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 10;x2=*/ w 7 10;y1=*/ h 3 10;y2=*/ h 7 10;x3=*/ w 4 10;x4=*/ w 6 10;y3=*/ h 45 100;y4=*/ h 55 100",
      paths: "M l t L r t L r b L l b Z|M wd2 y1 L x2 hd2 L x2 y2 L x1 y2 L x1 hd2 Z|M x3 y3 L x4 y3 L x4 y2 L x3 y2 Z",
    })
  } else if prst == "actionButtonBackPrevious" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 10;x2=*/ w 7 10;y1=*/ h 3 10;y2=*/ h 7 10",
      paths: "M l t L r t L r b L l b Z|M x2 hd2 L x1 y1 L x1 y2 Z",
    })
  } else if prst == "actionButtonForwardNext" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 10;x2=*/ w 7 10;y1=*/ h 3 10;y2=*/ h 7 10",
      paths: "M l t L r t L r b L l b Z|M x1 y1 L x2 hd2 L x1 y2 Z",
    })
  } else if prst == "actionButtonBeginning" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 10;x2=*/ w 35 100;x3=*/ w 7 10;y1=*/ h 3 10;y2=*/ h 7 10",
      paths: "M l t L r t L r b L l b Z|M x3 hd2 L x2 y1 L x2 y2 Z|M x2 y1 L x1 y1 L x1 y2 L x2 y2",
    })
  } else if prst == "actionButtonEnd" {
    Some({
      avlst: "",
      gdlst: "x1=*/ w 3 10;x2=*/ w 65 100;x3=*/ w 7 10;y1=*/ h 3 10;y2=*/ h 7 10",
      paths: "M l t L r t L r b L l b Z|M x1 y1 L x2 hd2 L x1 y2 Z|M x2 y1 L x3 y1 L x3 y2 L x2 y2",
    })
  } else if prst == "actionButtonReturn" ||
    prst == "actionButtonMovie" ||
    prst == "actionButtonDocument" ||
    prst == "actionButtonSound" ||
    prst == "actionButtonHelp" ||
    prst == "actionButtonInformation" {
    // Simplified — render as rect with bevel effect
    Some({
      avlst: "",
      gdlst: "dx=*/ w 1 20;dy=*/ h 1 20;x1=+- r 0 dx;y1=+- b 0 dy",
      paths: "M l t L r t L r b L l b Z|M dx dy L x1 dy L x1 y1 L dx y1 Z",
    })
  } else {
    None
  }
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
  match get_preset_def(prst) {
    None =>
      // Fallback: render as rectangle
      "<rect" +
      ai("x", x_p) +
      ai("y", y_p) +
      ai("width", cx_p) +
      ai("height", cy_p) +
      fill +
      stroke +
      rot +
      "/>"
    Some(def) => {
      let w = cx_p.to_double()
      let h = cy_p.to_double()
      let env = GuideEnv::new()
      env.init_builtins(w, h)
      // Parse and apply adjustment defaults, then override with provided values
      apply_adjustments(env, def.avlst, adj_values)
      // Evaluate guide list
      eval_guide_list(env, def.gdlst)
      // Render path
      let d = render_geom_paths(env, def.paths, x_p, y_p, w, h, 0.0, 0.0)
      if d.is_empty() {
        // Fallback if path generation failed
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
        "<path d=\"" + d + "\"" + fill + stroke + rot + "/>"
      }
    }
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
  eval_guide_list(env, data.gdlst)
  // Render paths with coordinate space scaling
  let pw = data.path_w.to_double()
  let ph = data.path_h.to_double()
  let d = render_geom_paths(env, data.paths, x_p, y_p, w, h, pw, ph)
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
/// Text rectangle definition: "l t r b" using guide variable names.
/// Returns (rect_l, rect_t, rect_r, rect_b) as guide var names, or empty for default.
fn get_preset_text_rect(prst : String) -> String {
  // Text rects for shapes where text area differs from bounding box.
  // Values are guide variable names resolved after evaluating the preset's gdlst.
  if prst == "triangle" || prst == "isosTriangle" {
    "trl t trr b"
  } else if prst == "rtTriangle" {
    "l trt trr b"
  } else if prst == "diamond" || prst == "flowChartDecision" {
    "wd4 hd4 3wd4 3hd4"
  } else if prst == "pentagon" || prst == "homePlate" {
    "wd6 t 5wd6 b"
  } else if prst == "hexagon" {
    "hxl t hxr b"
  } else if prst == "parallelogram" {
    "prl t prr b"
  } else if prst == "trapezoid" {
    "tzl t tzr b"
  } else if prst == "rightArrow" || prst == "notchedRightArrow" {
    "l ady1 arx b"
  } else if prst == "leftArrow" {
    "alx ady1 r b"
  } else if prst == "upArrow" {
    "adx1 aly t adx2 b"
  } else if prst == "downArrow" {
    "adx1 t ary adx2 b"
  } else if prst == "leftRightArrow" {
    "alx ady1 arx ady2"
  } else if prst == "upDownArrow" {
    "adx1 aly adx2 ary"
  } else if prst == "chevron" {
    "l t chx b"
  } else if prst == "star4" {
    "x1 y1 x2 y2"
  } else if prst == "star5" ||
    prst == "star6" ||
    prst == "star8" ||
    prst == "star10" {
    "x1 y1 x2 y2"
  } else if prst == "pie" || prst == "arc" {
    "wd4 hd4 3wd4 3hd4"
  } else if prst == "donut" ||
    prst == "cloudCallout" ||
    prst == "circularArrow" ||
    prst == "star16" ||
    prst == "star24" ||
    prst == "star32" {
    "il it ir ib"
  } else if prst == "star12" {
    "sx2 sy2 sx5 sy5"
  } else if prst == "quadArrow" {
    "il y3 ir y4"
  } else if prst == "lightningBolt" {
    "x4 y4 x9 y10"
  } else if prst == "sun" {
    "x9 y9 x8 y8"
  } else if prst == "moon" {
    "g12w g15h g0w g16h"
  } else if prst == "cube" {
    "l y1 x4 b"
  } else if prst == "flowChartDocument" {
    "l t r y1"
  } else if prst == "gear6" || prst == "gear9" {
    "wd4 hd4 3wd4 3hd4"
  } else if prst == "octagon" {
    "ol t orr b"
  } else if prst == "flowChartInputOutput" || prst == "flowChartManualInput" {
    "wd6 hd4 5wd6 b"
  } else if prst == "flowChartExtract" || prst == "flowChartMerge" {
    "wd4 hd4 3wd4 3hd4"
  } else if prst == "flowChartOffpageConnector" {
    "l t r ofy"
  } else {
    ""
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
  let rect_def = get_preset_text_rect(prst)
  if rect_def.is_empty() {
    return (0, 0, cx_p, cy_p)
  }
  // Get the preset definition to evaluate guides
  let def = match get_preset_def(prst) {
    None => return (0, 0, cx_p, cy_p)
    Some(d) => d
  }
  let w = cx_p.to_double()
  let h = cy_p.to_double()
  let env = GuideEnv::new()
  env.init_builtins(w, h)
  // Add convenience variables for text rect formulas
  env.set("3wd4", w * 3.0 / 4.0)
  env.set("3hd4", h * 3.0 / 4.0)
  env.set("5wd6", w * 5.0 / 6.0)
  apply_adjustments(env, def.avlst, adj_values)
  eval_guide_list(env, def.gdlst)
  // Add text-rect-specific guide variables for common shapes
  add_text_rect_guides(env, prst, w, h)
  // Parse rect definition: "l t r b"
  let parts = @xml.split_by_space(rect_def)
  if parts.length() < 4 {
    return (0, 0, cx_p, cy_p)
  }
  let rl = env.resolve(parts[0])
  let rt = env.resolve(parts[1])
  let rr = env.resolve(parts[2])
  let rb = env.resolve(parts[3])
  // Clamp to shape bounds
  let l = if rl < 0.0 { 0 } else { rl.to_int() }
  let t = if rt < 0.0 { 0 } else { rt.to_int() }
  let r_val = if rr > w { cx_p } else { rr.to_int() }
  let b_val = if rb > h { cy_p } else { rb.to_int() }
  (l, t, r_val - l, b_val - t)
}

///|
/// Add text-rect-specific guide variables that aren't part of the shape's gdlst.
fn add_text_rect_guides(
  env : GuideEnv,
  prst : String,
  w : Double,
  h : Double,
) -> Unit {
  if prst == "triangle" || prst == "isosTriangle" {
    // Inscribed rectangle: narrower at top, wider at bottom
    let adj = env.resolve("adj")
    let dx = adj * w / 200000.0
    env.set("trl", w / 2.0 - dx)
    env.set("trr", w / 2.0 + dx)
  } else if prst == "rtTriangle" {
    // Right triangle: text in lower-left area
    env.set("trt", h / 2.0)
    env.set("trr", w / 2.0)
  } else if prst == "hexagon" {
    let adj = env.resolve("adj")
    let dx = adj * w / 100000.0
    env.set("hxl", dx)
    env.set("hxr", w - dx)
  } else if prst == "parallelogram" {
    let adj = env.resolve("adj")
    let dx = adj * w / 200000.0
    env.set("prl", dx)
    env.set("prr", w - dx)
  } else if prst == "trapezoid" {
    let adj = env.resolve("adj")
    let dx = adj * w / 200000.0
    env.set("tzl", dx)
    env.set("tzr", w - dx)
  } else if prst == "rightArrow" || prst == "notchedRightArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dy = adj1 * h / 200000.0
    env.set("ady1", h / 2.0 - dy)
    env.set("arx", w - adj2 * ss / 100000.0)
  } else if prst == "leftArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dy = adj1 * h / 200000.0
    env.set("ady1", h / 2.0 - dy)
    env.set("alx", adj2 * ss / 100000.0)
  } else if prst == "upArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dx = adj1 * w / 200000.0
    env.set("adx1", w / 2.0 - dx)
    env.set("adx2", w / 2.0 + dx)
    env.set("aly", adj2 * ss / 100000.0)
  } else if prst == "downArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dx = adj1 * w / 200000.0
    env.set("adx1", w / 2.0 - dx)
    env.set("adx2", w / 2.0 + dx)
    env.set("ary", h - adj2 * ss / 100000.0)
  } else if prst == "leftRightArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dy = adj1 * h / 200000.0
    env.set("ady1", h / 2.0 - dy)
    env.set("ady2", h / 2.0 + dy)
    env.set("alx", adj2 * ss / 100000.0)
    env.set("arx", w - adj2 * ss / 100000.0)
  } else if prst == "upDownArrow" {
    let adj1 = env.resolve("adj1")
    let adj2 = env.resolve("adj2")
    let ss = if w < h { w } else { h }
    let dx = adj1 * w / 200000.0
    env.set("adx1", w / 2.0 - dx)
    env.set("adx2", w / 2.0 + dx)
    env.set("aly", adj2 * ss / 100000.0)
    env.set("ary", h - adj2 * ss / 100000.0)
  } else if prst == "chevron" {
    let adj = env.resolve("adj")
    env.set("chx", w - adj * w / 100000.0)
  } else if prst == "octagon" {
    let adj = env.resolve("adj")
    let dx = adj * w / 100000.0
    env.set("ol", dx / 2.0)
    env.set("orr", w - dx / 2.0)
  } else if prst == "flowChartOffpageConnector" {
    env.set("ofy", h * 4.0 / 5.0)
  } else {
    ignore(w)
    ignore(h)
  }
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
  eval_guide_list(env, data.gdlst)
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
/// Parse adjustment defaults from the definition and apply user overrides.
fn apply_adjustments(
  env : GuideEnv,
  avlst_str : String,
  adj_values : Array[Int],
) -> Unit {
  if avlst_str.is_empty() {
    return
  }
  // Parse "name1=val1;name2=val2;..."
  let adj_names : Array[String] = []
  let len = avlst_str.length()
  let mut start = 0
  for i in 0..<len {
    let c = avlst_str.get_char(i).unwrap()
    if c == ';' {
      parse_one_adj(env, avlst_str, start, i, adj_names)
      start = i + 1
    }
  }
  if start < len {
    parse_one_adj(env, avlst_str, start, len, adj_names)
  }
  // Override with user-provided values (positional)
  let n_overrides = adj_values.length()
  let n_names = adj_names.length()
  for j in 0..<n_overrides {
    if j < n_names {
      env.set(adj_names[j], adj_values[j].to_double())
    }
  }
}

///|
/// Parse a single "name=value" adjustment entry.
fn parse_one_adj(
  env : GuideEnv,
  s : String,
  start : Int,
  end_ : Int,
  adj_names : Array[String],
) -> Unit {
  // Find '='
  let mut eq = start
  while eq < end_ {
    if s.get_char(eq).unwrap() == '=' {
      break
    }
    eq = eq + 1
  }
  if eq >= end_ {
    return
  }
  let name = @xml.str_substring(s, start, eq)
  let val_str = @xml.str_substring(s, eq + 1, end_)
  adj_names.push(name)
  env.set(name, parse_double(val_str))
}

///|
/// Evaluate a semicolon-separated guide list.
fn eval_guide_list(env : GuideEnv, gdlst_str : String) -> Unit {
  if gdlst_str.is_empty() {
    return
  }
  let len = gdlst_str.length()
  let mut start = 0
  for i in 0..<len {
    if gdlst_str.get_char(i).unwrap() == ';' {
      eval_one_guide(env, gdlst_str, start, i)
      start = i + 1
    }
  }
  if start < len {
    eval_one_guide(env, gdlst_str, start, len)
  }
}

///|
/// Parse and evaluate a single "name=formula" guide entry.
fn eval_one_guide(env : GuideEnv, s : String, start : Int, end_ : Int) -> Unit {
  let mut eq = start
  while eq < end_ {
    if s.get_char(eq).unwrap() == '=' {
      break
    }
    eq = eq + 1
  }
  if eq >= end_ {
    return
  }
  let name = @xml.str_substring(s, start, eq)
  let fmla = @xml.str_substring(s, eq + 1, end_)
  eval_formula(env, name, fmla)
}
