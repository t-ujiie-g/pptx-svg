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
  { names: [], values: [] }
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
  // Then try parsing as literal integer
  parse_double(token)
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
/// Split a formula string by spaces. Returns array of tokens.
fn split_formula(fmla : String) -> Array[String] {
  let tokens : Array[String] = []
  let len = fmla.length()
  let mut start = 0
  let mut i = 0
  while i < len {
    if fmla.get_char(i).unwrap() == ' ' {
      if i > start {
        tokens.push(@xml.str_substring(fmla, start, i))
      }
      start = i + 1
    }
    i = i + 1
  }
  if start < len {
    tokens.push(@xml.str_substring(fmla, start, len))
  }
  tokens
}

///|
/// Evaluate a single guide formula and store the result.
fn eval_formula(env : GuideEnv, name : String, fmla : String) -> Unit {
  let tokens = split_formula(fmla)
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
    // sin(angle) * distance — angle in 60000ths of degree
    let ang = env.resolve(tokens[1]) * ooxml_angle_to_rad
    let dist = env.resolve(tokens[2])
    @ffi.ffi_sin(ang) * dist
  } else if op == "cos" && nt >= 3 {
    // cos(angle) * distance
    let ang = env.resolve(tokens[1]) * ooxml_angle_to_rad
    let dist = env.resolve(tokens[2])
    @ffi.ffi_cos(ang) * dist
  } else if op == "tan" && nt >= 3 {
    // tan(angle) * distance
    let ang = env.resolve(tokens[1]) * ooxml_angle_to_rad
    let dist = env.resolve(tokens[2])
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
  let tokens = split_formula(path_cmds)
  let nt = tokens.length()
  let mut d = ""
  let mut i = 0
  // Scale factors if path has its own coordinate space
  let sx = if path_w > 0.0 { shape_w / path_w } else { 1.0 }
  let sy = if path_h > 0.0 { shape_h / path_h } else { 1.0 }
  // Current position for arcTo endpoint calculation
  let mut cur_x = 0.0
  let mut cur_y = 0.0
  while i < nt {
    let cmd = tokens[i]
    if cmd == "M" && i + 2 < nt {
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
  let st_rad = st_ang * ooxml_angle_to_rad
  let sw_rad = sw_ang * ooxml_angle_to_rad
  let end_rad = st_rad + sw_rad
  // Compute end point relative to current position
  // Current pos is on the ellipse at angle stAng
  // End pos is on the ellipse at angle stAng + swAng
  let dx = wr * (@ffi.ffi_cos(end_rad) - @ffi.ffi_cos(st_rad))
  let dy = hr * (@ffi.ffi_sin(end_rad) - @ffi.ffi_sin(st_rad))
  let end_x = cur_x + dx
  let end_y = cur_y + dy
  // SVG arc flags
  let abs_sw = if sw_ang < 0.0 { -sw_ang } else { sw_ang }
  let large_arc = if abs_sw > 10800000.0 { "1" } else { "0" }
  let sweep = if sw_ang > 0.0 { "1" } else { "0" }
  let ox = x_off.to_double()
  let oy = y_off.to_double()
  let s = "A" +
    dts(wr) +
    " " +
    dts(hr) +
    " 0 " +
    large_arc +
    " " +
    sweep +
    " " +
    dts(end_x + ox) +
    " " +
    dts(end_y + oy)
  (s, end_x, end_y)
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
) -> String {
  // Split by "|" for multiple paths
  let len = paths.length()
  let mut start = 0
  let mut d = ""
  let mut i = 0
  while i < len {
    if paths.get_char(i).unwrap() == '|' {
      if i > start {
        let seg = @xml.str_substring(paths, start, i)
        d = d + render_geom_path(env, seg, x_off, y_off, 0.0, 0.0, w, h)
      }
      start = i + 1
    }
    i = i + 1
  }
  if start < len {
    let seg = @xml.str_substring(paths, start, len)
    d = d + render_geom_path(env, seg, x_off, y_off, 0.0, 0.0, w, h)
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
    Some({ avlst: "", gdlst: "", paths: "M l b L l t L r b Z" })
  } else if prst == "diamond" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z" })
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
      gdlst: "cx1=cos 20520000 wd2;cy1=sin 20520000 hd2;cx2=cos 3240000 wd2;cy2=sin 3240000 hd2;cx3=cos 7560000 wd2;cy3=sin 7560000 hd2;cx4=cos 11880000 wd2;cy4=sin 11880000 hd2;vx1=+- wd2 cx1 0;vy1=+- hd2 cy1 0;vx2=+- wd2 cx2 0;vy2=+- hd2 cy2 0;vx3=+- wd2 cx3 0;vy3=+- hd2 cy3 0;vx4=+- wd2 cx4 0;vy4=+- hd2 cy4 0",
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
      gdlst: "a=pin 0 adj 50000;dr=*/ ss a 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z|M iwd2 hd2 A iwd2 ihd2 0 -21600000 Z",
    })
  } else if prst == "noSmoking" {
    Some({
      avlst: "adj=18750",
      gdlst: "a=pin 0 adj 50000;dr=*/ ss a 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z",
    })
  } else if prst == "blockArc" {
    Some({
      avlst: "adj1=10800000;adj2=0;adj3=25000",
      gdlst: "a3=pin 0 adj3 50000;dr=*/ ss a3 100000;iwd2=+- wd2 0 dr;ihd2=+- hd2 0 dr;sa=*/ adj1 1 1;ea=*/ adj2 1 1;csa=cos sa wd2;ssa=sin sa hd2;cea=cos ea wd2;sea=sin ea hd2;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0;icsa=cos sa iwd2;issa=sin sa ihd2;icea=cos ea iwd2;isea=sin ea ihd2;ix1=+- wd2 icsa 0;iy1=+- hd2 issa 0;ix2=+- wd2 icea 0;iy2=+- hd2 isea 0;sw=+- ea 0 sa",
      paths: "M x1 y1 A wd2 hd2 sa sw L ix2 iy2 A iwd2 ihd2 ea 0 Z",
    })
    // ── Block arrows ──
  } else if prst == "rightArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dx=*/ w a2 100000;x1=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M l y1 L x1 y1 L x1 t L r hd2 L x1 b L x1 y2 L l y2 Z",
    })
  } else if prst == "leftArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dx=*/ w a2 100000;x1=+- l dx 0;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M r y1 L x1 y1 L x1 t L l hd2 L x1 b L x1 y2 L r y2 Z",
    })
  } else if prst == "upArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dy=*/ h a2 100000;y1=+- t dy 0;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M x1 b L x1 y1 L l y1 L wd2 t L r y1 L x2 y1 L x2 b Z",
    })
  } else if prst == "downArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dy=*/ h a2 100000;y1=+- b 0 dy;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M x1 t L x1 y1 L l y1 L wd2 b L r y1 L x2 y1 L x2 t Z",
    })
  } else if prst == "leftRightArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 50000;dx=*/ w a2 100000;x1=+- l dx 0;x2=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0",
      paths: "M l hd2 L x1 t L x1 y1 L x2 y1 L x2 t L r hd2 L x2 b L x2 y2 L x1 y2 L x1 b Z",
    })
  } else if prst == "upDownArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 50000;dy=*/ h a2 100000;y1=+- t dy 0;y2=+- b 0 dy;dx=*/ w a1 200000;x1=+- wd2 0 dx;x2=+- wd2 dx 0",
      paths: "M wd2 t L r y1 L x2 y1 L x2 y2 L r y2 L wd2 b L l y2 L x1 y2 L x1 y1 L l y1 Z",
    })
  } else if prst == "quadArrow" {
    Some({
      avlst: "adj1=22500;adj2=22500;adj3=22500",
      gdlst: "a1=pin 0 adj1 50000;a2=pin 0 adj2 50000;a3=pin 0 adj3 50000;dx=*/ w a2 100000;dy=*/ h a2 100000;dh=*/ w a1 200000;dv=*/ h a1 200000;ah=*/ w a3 100000;av=*/ h a3 100000;x1=+- wd2 0 dh;x2=+- wd2 dh 0;y1=+- hd2 0 dv;y2=+- hd2 dv 0;xe=+- r 0 dx;ye=+- b 0 dy",
      paths: "M wd2 t L r dy L x2 dy L x2 y1 L xe y1 L xe hd2 L r hd2|M xe hd2 L xe y2 L x2 y2 L x2 ye L wd2 b L l ye L x1 ye L x1 y2 L dx y2 L dx hd2|M l hd2 L dx y1 L x1 y1 L x1 dy L wd2 t Z",
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
      gdlst: "a=pin 0 adj 50000;swd2=*/ wd2 a 50000;shd2=*/ hd2 a 50000;co1=cos 20520000 wd2;so1=sin 20520000 hd2;co2=cos 3240000 wd2;so2=sin 3240000 hd2;co3=cos 7560000 wd2;so3=sin 7560000 hd2;co4=cos 11880000 wd2;so4=sin 11880000 hd2;ox1=+- wd2 co1 0;oy1=+- hd2 so1 0;ox2=+- wd2 co2 0;oy2=+- hd2 so2 0;ox3=+- wd2 co3 0;oy3=+- hd2 so3 0;ox4=+- wd2 co4 0;oy4=+- hd2 so4 0;ci1=cos 18360000 swd2;si1=sin 18360000 shd2;ci2=cos 1080000 swd2;si2=sin 1080000 shd2;ci3=cos 5400000 swd2;si3=sin 5400000 shd2;ci4=cos 9720000 swd2;si4=sin 9720000 shd2;ci5=cos 14040000 swd2;si5=sin 14040000 shd2;ix1=+- wd2 ci1 0;iy1=+- hd2 si1 0;ix2=+- wd2 ci2 0;iy2=+- hd2 si2 0;ix3=+- wd2 ci3 0;iy3=+- hd2 si3 0;ix4=+- wd2 ci4 0;iy4=+- hd2 si4 0;ix5=+- wd2 ci5 0;iy5=+- hd2 si5 0",
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
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L r b L l b Z" })
  } else if prst == "flowChartDecision" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z" })
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
      gdlst: "y1=*/ h 17 20;cy1=*/ h 19 20",
      paths: "M l t L r t L r y1 C r cy1 wd2 b wd2 y1 C wd2 cy1 l b l y1 Z",
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
      gdlst: "",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z",
    })
  } else if prst == "flowChartSort" {
    Some({
      avlst: "",
      gdlst: "",
      paths: "M wd2 t L r hd2 L wd2 b L l hd2 Z|M l hd2 L r hd2",
    })
  } else if prst == "flowChartMerge" {
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L wd2 b Z" })
  } else if prst == "flowChartExtract" {
    Some({ avlst: "", gdlst: "", paths: "M wd2 t L r b L l b Z" })
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
      gdlst: "a1=*/ w adj1 100000;a2=*/ h adj2 100000;x1=+- wd2 a1 0;y1=+- hd2 a2 0;dx1=+- wd2 0 wd8;dx2=+- wd2 wd8 0",
      paths: "M l t L r t L r b L dx2 b L x1 y1 L dx1 b L l b Z",
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
      gdlst: "a1=*/ w adj1 100000;a2=*/ h adj2 100000;x1=+- wd2 a1 0;y1=+- hd2 a2 0",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z|M wd2 b L x1 y1 L wd2 b",
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
      gdlst: "x1=*/ w 4269 21600;y1=*/ h 3 21;x2=*/ w 8640 21600;y2=*/ h 8 21;x3=*/ w 5765 21600;y3=*/ h 9 21;x4=*/ w 12960 21600;y4=*/ h 13 21;x5=*/ w 8323 21600;y5=*/ h 14 21;x6=*/ w 14400 21600;y6=*/ h 17 21;x7=*/ w 4320 21600",
      paths: "M x1 t L x2 y2 L x3 y3 L x4 y4 L x5 y5 L r y6 L x7 b L x5 y5 L x1 y3 L x3 y2 Z",
    })
  } else if prst == "sun" {
    Some({
      avlst: "adj=25000",
      gdlst: "a=pin 12500 adj 46875;cir=*/ ss a 100000;ir=+- wd2 0 cir;ib=+- hd2 0 cir;or2=+- wd2 cir 0;ob2=+- hd2 cir 0",
      paths: "M wd2 t L or2 ib L r hd2 L or2 ob2 L wd2 b L ir ob2 L l hd2 L ir ib Z|M ir hd2 A cir cir 10800000 21600000 Z",
    })
  } else if prst == "moon" {
    Some({
      avlst: "adj=50000",
      gdlst: "a=pin 0 adj 87500;x1=*/ w a 100000",
      paths: "M r t C x1 hd4 x1 hd2 C x1 hd2 x1 b r b A wd2 hd2 0 -10800000 Z",
    })
  } else if prst == "smileyFace" {
    Some({
      avlst: "adj=4653",
      gdlst: "a=pin -4653 adj 4653;ey=*/ h 3 4;cx1=*/ w 3 10;cx2=*/ w 7 10;ey2=+- ey 0 hd10;er=*/ w 3 20;sy=+- ey 0 hd10;sdy=*/ h a 100000;sy1=+- ey sdy 0;sy2=+- ey sdy 0",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z|M cx1 ey A er er 0 21600000 Z|M cx2 ey A er er 0 21600000 Z|M cx1 sy1 C cx1 sy2 cx2 sy2 cx2 sy1",
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
      gdlst: "a=pin 0 adj 100000;dx=*/ w a 100000;y1=*/ h a 100000;dx1=+- r 0 dx",
      paths: "M l y1 L dx t L r t L r b L dx1 b L l b Z|M l y1 L dx y1 L dx t|M dx y1 L r b",
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
      gdlst: "a=pin 0 adj 100000;dx=*/ w a 100000;x1=+- r 0 dx",
      paths: "M l t L x1 t L r hd2 L x1 b L l b Z",
    })
  } else if prst == "notchedRightArrow" {
    Some({
      avlst: "adj1=50000;adj2=50000",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;dx=*/ w a2 100000;x1=+- r 0 dx;dy=*/ h a1 200000;y1=+- hd2 0 dy;y2=+- hd2 dy 0;dx1=*/ h a2 200000;nx1=+- l dx1 0",
      paths: "M l y1 L x1 y1 L x1 t L r hd2 L x1 b L x1 y2 L l y2 L nx1 hd2 Z",
    })
  } else if prst == "leftRightArrowCallout" || prst == "stripedRightArrow" {
    // Simplified — render as rectangle for less common shapes
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L r b L l b Z" })
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
      avlst: "adj1=23520;adj2=5765;adj3=11760",
      gdlst: "a1=pin 0 adj1 100000;a2=pin 0 adj2 100000;a3=pin 0 adj3 100000;dy1=*/ h a1 200000;y1=+- hd2 0 dy1;y2=+- hd2 dy1 0;dr=*/ ss a2 200000;dy2=*/ h a3 100000;y3=+- y1 0 dy2;y4=+- y2 dy2 0",
      paths: "M l y1 L r y1 L r y2 L l y2 Z|M wd2 y3 A dr dr 0 21600000 Z|M wd2 y4 A dr dr 0 21600000 Z",
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
      gdlst: "sa=*/ adj1 1 1;ea=*/ adj2 1 1;sw=+- ea 0 sa;csa=cos sa wd2;ssa=sin sa hd2;cea=cos ea wd2;sea=sin ea hd2;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0",
      paths: "M wd2 hd2 L x1 y1 A wd2 hd2 sa sw L wd2 hd2 Z",
    })
  } else if prst == "chord" {
    Some({
      avlst: "adj1=2700000;adj2=16200000",
      gdlst: "sa=*/ adj1 1 1;ea=*/ adj2 1 1;sw=+- ea 0 sa;csa=cos sa wd2;ssa=sin sa hd2;cea=cos ea wd2;sea=sin ea hd2;x1=+- wd2 csa 0;y1=+- hd2 ssa 0;x2=+- wd2 cea 0;y2=+- hd2 sea 0",
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
  } else if prst == "gear6" || prst == "gear9" {
    // Simplified — render as circle for gear shapes
    Some({
      avlst: "",
      gdlst: "",
      paths: "M l hd2 A wd2 hd2 10800000 21600000 Z",
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
    // ── Action buttons ──
  } else if prst == "actionButtonBlank" {
    Some({ avlst: "", gdlst: "", paths: "M l t L r t L r b L l b Z" })
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
      let d = render_geom_paths(env, def.paths, x_p, y_p, w, h)
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
  let mut i = 0
  while i < len {
    let c = avlst_str.get_char(i).unwrap()
    if c == ';' {
      parse_one_adj(env, avlst_str, start, i, adj_names)
      start = i + 1
    }
    i = i + 1
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
  let mut i = 0
  while i < len {
    if gdlst_str.get_char(i).unwrap() == ';' {
      eval_one_guide(env, gdlst_str, start, i)
      start = i + 1
    }
    i = i + 1
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
