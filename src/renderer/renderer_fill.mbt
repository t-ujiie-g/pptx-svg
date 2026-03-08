// SVG renderer — Fill rendering (gradient, pattern, blip tile)
//
// Same package as renderer.mbt — all constants and helpers are shared.

///|
let pattern_tile_sz : Int = 8 // default SVG <pattern> tile size in px

// ── Gradient helpers ─────────────────────────────────────────────────────────

///|
/// Render a single gradient <defs> definition.
///  Returns a <linearGradient> or <radialGradient> element string.
fn render_gradient_def(gf : @ooxml.GradientFill, grad_id : String) -> String {
  // Build stop elements
  let mut stops_svg = ""
  let ns = gf.stops.length()
  let mut si = 0
  while si < ns {
    let stop = gf.stops[si]
    // pos is 0..100000 → percentage 0..100
    let pct = stop.pos / ooxml_stop_pct
    let stop_opacity = if stop.color.alpha < 255 {
      " stop-opacity=\"" + @ooxml.alpha_to_css_str(stop.color.alpha) + "\""
    } else {
      ""
    }
    // Use rgb() for stop-color (ignoring alpha) and separate stop-opacity
    let stop_css = "rgb(" +
      int_to_str(stop.color.r) +
      "," +
      int_to_str(stop.color.g) +
      "," +
      int_to_str(stop.color.b) +
      ")"
    stops_svg = stops_svg +
      "<stop offset=\"" +
      int_to_str(pct) +
      "%\" stop-color=\"" +
      stop_css +
      "\"" +
      stop_opacity +
      "/>"
    si = si + 1
  }
  // tileFlip → SVG spreadMethod="reflect" (x, y, or xy all map to reflect)
  let spread_attr = if gf.tile_flip == "x" ||
    gf.tile_flip == "y" ||
    gf.tile_flip == "xy" {
    " spreadMethod=\"reflect\""
  } else {
    ""
  }
  if gf.angle >= 0 && gf.path_type.is_empty() {
    // Linear gradient
    // OOXML angle is in 60000th of a degree. SVG gradientTransform rotate()
    // uses degrees centered at (0.5, 0.5). Direct division gives the SVG angle.
    let deg = gf.angle / ooxml_angle_unit
    "<linearGradient id=\"" +
    grad_id +
    "\" gradientUnits=\"objectBoundingBox\" x1=\"0\" y1=\"0\" x2=\"1\" y2=\"0\" gradientTransform=\"rotate(" +
    int_to_str(deg) +
    ", 0.5, 0.5)\"" +
    spread_attr +
    ">" +
    stops_svg +
    "</linearGradient>"
  } else if not(gf.path_type.is_empty()) {
    // Radial / path gradient
    let cx_s = if gf.fill_to_l >= 0 && gf.fill_to_r >= 0 {
      int_to_str((gf.fill_to_l + (100000 - gf.fill_to_r)) / ooxml_fill_to_pct)
    } else {
      "50"
    }
    let cy_s = if gf.fill_to_t >= 0 && gf.fill_to_b >= 0 {
      int_to_str((gf.fill_to_t + (100000 - gf.fill_to_b)) / ooxml_fill_to_pct)
    } else {
      "50"
    }
    "<radialGradient id=\"" +
    grad_id +
    "\" gradientUnits=\"objectBoundingBox\" cx=\"" +
    cx_s +
    "%\" cy=\"" +
    cy_s +
    "%\" r=\"50%\"" +
    spread_attr +
    ">" +
    stops_svg +
    "</radialGradient>"
  } else {
    // Fallback linear
    "<linearGradient id=\"" +
    grad_id +
    "\" gradientUnits=\"objectBoundingBox\"" +
    spread_attr +
    ">" +
    stops_svg +
    "</linearGradient>"
  }
}

///|
/// Serialize gradient stop data for data-ooxml-grad-stops attribute.
///  Format: "pos,RRGGBB;pos,RRGGBB;..."
fn grad_stops_attr(gf : @ooxml.GradientFill) -> String {
  let mut out = ""
  let ns = gf.stops.length()
  let mut si = 0
  while si < ns {
    if si > 0 {
      out = out + ";"
    }
    out = out + int_to_str(gf.stops[si].pos) + "," + gf.stops[si].color.to_hex()
    si = si + 1
  }
  out
}

// ── Pattern fill SVG helpers ─────────────────────────────────────────────────

///|
/// SVG circle element for pattern dot.
fn patt_dot(cx : String, cy : String, r : String, fg : String) -> String {
  "<circle cx=\"" +
  cx +
  "\" cy=\"" +
  cy +
  "\" r=\"" +
  r +
  "\" fill=\"" +
  fg +
  "\"/>"
}

///|
/// SVG line element for pattern line.
fn patt_line(
  x1 : String,
  y1 : String,
  x2 : String,
  y2 : String,
  fg : String,
  sw : String,
) -> String {
  "<line x1=\"" +
  x1 +
  "\" y1=\"" +
  y1 +
  "\" x2=\"" +
  x2 +
  "\" y2=\"" +
  y2 +
  "\" stroke=\"" +
  fg +
  "\" stroke-width=\"" +
  sw +
  "\"/>"
}

///|
/// SVG rect element for pattern fill.
fn patt_rect(
  x : String,
  y : String,
  w : String,
  h : String,
  fg : String,
) -> String {
  "<rect x=\"" +
  x +
  "\" y=\"" +
  y +
  "\" width=\"" +
  w +
  "\" height=\"" +
  h +
  "\" fill=\"" +
  fg +
  "\"/>"
}

// ── Pattern fill rendering ────────────────────────────────────────────────────

///|
/// Generate SVG content for a pattern preset. Returns SVG shapes to place inside <pattern>.
fn patt_svg_content(
  prst : String,
  fg : String,
  bg : String,
  sz : Int,
) -> String {
  let s = int_to_str(sz)
  let h = int_to_str(sz / 2)
  let q = int_to_str(sz / 4)
  let q3 = int_to_str(sz * 3 / 4)
  // Background rect
  let br = "<rect width=\"" +
    s +
    "\" height=\"" +
    s +
    "\" fill=\"" +
    bg +
    "\"/>"
  // Foreground-filled rect (for inverted patterns)
  let fr = "<rect width=\"" +
    s +
    "\" height=\"" +
    s +
    "\" fill=\"" +
    fg +
    "\"/>"

  // ── Percentage fills ──
  if prst == "pct5" {
    br + patt_dot(h, h, "1", fg)
  } else if prst == "pct10" {
    br + patt_dot(h, h, "1", fg)
  } else if prst == "pct20" || prst == "pct25" {
    br + patt_dot(h, h, "2", fg)
  } else if prst == "pct30" {
    br + patt_dot(q, q, "1", fg) + patt_dot(q3, q3, "1", fg)
  } else if prst == "pct40" {
    br +
    patt_dot(q, q, "1", fg) +
    patt_dot(q3, q, "1", fg) +
    patt_dot(h, q3, "1", fg)
  } else if prst == "pct50" {
    br + patt_rect("0", "0", h, h, fg)
  } else if prst == "pct60" {
    br + patt_rect("0", "0", h, h, fg) + patt_dot(q3, q3, "1", fg)
  } else if prst == "pct70" || prst == "pct75" {
    fr + patt_dot(h, h, "2", bg)
  } else if prst == "pct80" {
    fr + patt_dot(h, h, "1", bg)
  } else if prst == "pct90" {
    fr + patt_dot(h, q, "1", bg)

    // ── Diagonal lines ──
  } else if prst == "ltDnDiag" {
    br + patt_line("0", "0", s, s, fg, "1")
  } else if prst == "dkDnDiag" {
    br + patt_line("0", "0", s, s, fg, "2")
  } else if prst == "ltUpDiag" {
    br + patt_line(s, "0", "0", s, fg, "1")
  } else if prst == "dkUpDiag" {
    br + patt_line(s, "0", "0", s, fg, "2")
  } else if prst == "dnDiag" {
    br +
    patt_line("0", "0", s, s, fg, "1") +
    patt_line(int_to_str(-sz), "0", "0", s, fg, "1") +
    patt_line(s, "0", int_to_str(sz * 2), s, fg, "1")
  } else if prst == "upDiag" {
    br +
    patt_line(s, "0", "0", s, fg, "1") +
    patt_line(int_to_str(sz * 2), "0", s, s, fg, "1") +
    patt_line("0", "0", int_to_str(-sz), s, fg, "1")
  } else if prst == "wdDnDiag" {
    br + patt_line("0", "0", s, s, fg, "3")
  } else if prst == "wdUpDiag" {
    br + patt_line(s, "0", "0", s, fg, "3")
  } else if prst == "dashDnDiag" {
    br + patt_line("0", "0", h, h, fg, "1")
  } else if prst == "dashUpDiag" {
    br + patt_line(s, "0", h, h, fg, "1")

    // ── Horizontal / Vertical lines ──
  } else if prst == "ltHorz" {
    br + patt_line("0", h, s, h, fg, "1")
  } else if prst == "dkHorz" {
    br + patt_line("0", h, s, h, fg, "2")
  } else if prst == "ltVert" {
    br + patt_line(h, "0", h, s, fg, "1")
  } else if prst == "dkVert" {
    br + patt_line(h, "0", h, s, fg, "2")
  } else if prst == "horz" {
    br + patt_line("0", h, s, h, fg, "1") + patt_line("0", "0", s, "0", fg, "1")
  } else if prst == "vert" {
    br + patt_line(h, "0", h, s, fg, "1") + patt_line("0", "0", "0", s, fg, "1")
  } else if prst == "narHorz" {
    br + patt_line("0", q, s, q, fg, "1") + patt_line("0", q3, s, q3, fg, "1")
  } else if prst == "narVert" {
    br + patt_line(q, "0", q, s, fg, "1") + patt_line(q3, "0", q3, s, fg, "1")
  } else if prst == "dashHorz" {
    br + patt_line("0", h, h, h, fg, "1")
  } else if prst == "dashVert" {
    br + patt_line(h, "0", h, h, fg, "1")

    // ── Cross patterns ──
  } else if prst == "cross" {
    br + patt_line(h, "0", h, s, fg, "1") + patt_line("0", h, s, h, fg, "1")
  } else if prst == "dnCross" {
    br + patt_line("0", "0", s, s, fg, "1") + patt_line(s, "0", "0", s, fg, "1")

    // ── Checkerboard / Diamond ──
  } else if prst == "smCheck" || prst == "lgCheck" {
    br + patt_rect("0", "0", h, h, fg) + patt_rect(h, h, h, h, fg)
  } else if prst == "openDmnd" {
    br +
    "<polygon points=\"" +
    h +
    ",0 " +
    s +
    "," +
    h +
    " " +
    h +
    "," +
    s +
    " 0," +
    h +
    "\" fill=\"none\" stroke=\"" +
    fg +
    "\" stroke-width=\"1\"/>"
  } else if prst == "solidDmnd" {
    br +
    "<polygon points=\"" +
    h +
    ",0 " +
    s +
    "," +
    h +
    " " +
    h +
    "," +
    s +
    " 0," +
    h +
    "\" fill=\"" +
    fg +
    "\"/>"

    // ── Grid patterns ──
  } else if prst == "smGrid" {
    br + patt_line(h, "0", h, s, fg, "1") + patt_line("0", h, s, h, fg, "1")
  } else if prst == "lgGrid" {
    br +
    patt_line("0", "0", s, "0", fg, "1") +
    patt_line("0", "0", "0", s, fg, "1")
  } else if prst == "dotGrid" {
    br +
    patt_dot("0", "0", "1", fg) +
    patt_dot(h, "0", "1", fg) +
    patt_dot("0", h, "1", fg) +
    patt_dot(h, h, "1", fg)
  } else if prst == "dotDmnd" {
    br +
    patt_dot(h, "0", "1", fg) +
    patt_dot("0", h, "1", fg) +
    patt_dot(s, h, "1", fg) +
    patt_dot(h, s, "1", fg)

    // ── Diagonal cross ──
  } else if prst == "ltDnDiagCross" {
    br + patt_line("0", "0", s, s, fg, "1") + patt_line(s, "0", "0", s, fg, "1")
  } else if prst == "dkDnDiagCross" {
    br + patt_line("0", "0", s, s, fg, "2") + patt_line(s, "0", "0", s, fg, "2")

    // ── Complex patterns ──
  } else if prst == "trellis" {
    br + patt_line("0", "0", s, s, fg, "2") + patt_line(s, "0", "0", s, fg, "2")
  } else if prst == "sphere" {
    br +
    "<circle cx=\"" +
    h +
    "\" cy=\"" +
    h +
    "\" r=\"" +
    q +
    "\" fill=\"" +
    fg +
    "\" opacity=\"0.7\"/>" +
    patt_dot(q, q, "1", bg)
  } else if prst == "weave" {
    br +
    patt_line("0", "0", h, h, fg, "2") +
    patt_line(s, "0", h, h, fg, "2") +
    patt_line("0", s, h, h, fg, "2") +
    patt_line(s, s, h, h, fg, "2")
  } else if prst == "plaid" {
    br +
    patt_rect("0", "0", h, h, fg) +
    patt_line("0", q3, s, q3, fg, "1") +
    patt_line(q3, "0", q3, s, fg, "1")
  } else if prst == "divot" {
    br + patt_line(q, q, h, q, fg, "1") + patt_line(h, q3, q3, q3, fg, "1")
  } else if prst == "zigZag" {
    br +
    "<polyline points=\"0," +
    h +
    " " +
    q +
    ",0 " +
    h +
    "," +
    h +
    " " +
    q3 +
    ",0 " +
    s +
    "," +
    h +
    "\" fill=\"none\" stroke=\"" +
    fg +
    "\" stroke-width=\"1\"/>"
  } else if prst == "wave" {
    br +
    "<path d=\"M0," +
    h +
    " Q" +
    q +
    ",0 " +
    h +
    "," +
    h +
    " Q" +
    q3 +
    "," +
    s +
    " " +
    s +
    "," +
    h +
    "\" fill=\"none\" stroke=\"" +
    fg +
    "\" stroke-width=\"1\"/>"
  } else if prst == "diagBrick" {
    br + patt_line("0", "0", s, s, fg, "1") + patt_line("0", h, h, "0", fg, "1")
  } else if prst == "horzBrick" {
    br +
    patt_line("0", h, s, h, fg, "1") +
    patt_line(h, "0", h, h, fg, "1") +
    patt_line("0", s, "0", h, fg, "1")
  } else if prst == "shingle" {
    br + patt_line("0", s, s, "0", fg, "1") + patt_line("0", h, h, "0", fg, "1")

    // ── Fallback ──
  } else {
    patt_rect("0", "0", s, s, fg)
  }
}

///|
/// Render a SVG <pattern> definition for a PatternFill.
fn render_pattern_def(pf : @ooxml.PatternFill, id : String) -> String {
  let sz = pattern_tile_sz
  let fg = pf.fg_color.to_css()
  let bg = pf.bg_color.to_css()
  let content = patt_svg_content(pf.prst, fg, bg, sz)
  "<pattern id=\"" +
  id +
  "\" patternUnits=\"userSpaceOnUse\" width=\"" +
  int_to_str(sz) +
  "\" height=\"" +
  int_to_str(sz) +
  "\">" +
  content +
  "</pattern>"
}

// ── Effect filter rendering ──────────────────────────────────────────────────

///|
/// Return CSS opacity string for an effect color, with a fallback default.
fn eff_alpha_str(clr : @ooxml.Color, fallback : String) -> String {
  if clr.alpha < 255 && not(clr.is_none()) {
    @ooxml.alpha_to_css_str(clr.alpha)
  } else {
    fallback
  }
}

///|
/// Convert OOXML EMU to SVG pixel value for filter attributes.
fn emu_to_filter_px(emu : Int, scale : Int) -> Int {
  if scale > 0 { emu / scale } else { emu / default_emu_per_px }
}

///|
/// Render SVG filter definition for shape effects (a:effectLst).
/// Returns (filter_defs, filter_attr) — defs SVG string and filter="url(#id)" attribute.
fn render_effect_filter(
  eff : @ooxml.EffectList,
  filter_id : String,
  scale : Int,
) -> (String, String) {
  if eff.is_none() {
    return ("", "")
  }
  let mut filters = ""
  // Outer shadow: SVG feDropShadow or feGaussianBlur+feOffset+feFlood+feComposite
  if not(eff.outer_shadow.is_none()) {
    let os = eff.outer_shadow
    let blur = emu_to_filter_px(os.blur_rad, scale)
    let dist = emu_to_filter_px(os.dist, scale)
    // dir is in 60000ths of a degree — convert to radians for dx/dy
    // 0 = right, 5400000 = bottom, etc.
    let angle_deg = os.dir / 60000
    let dx = dist * cos_deg(angle_deg) / 100
    let dy = dist * sin_deg(angle_deg) / 100
    let std_dev = blur / 2
    let clr = os.color
    let alpha_str = eff_alpha_str(clr, "0.5")
    filters = filters +
      "<feDropShadow dx=\"" +
      int_to_str(dx) +
      "\" dy=\"" +
      int_to_str(dy) +
      "\" stdDeviation=\"" +
      int_to_str(if std_dev > 0 { std_dev } else { 1 }) +
      "\" flood-color=\"" +
      clr.to_css() +
      "\" flood-opacity=\"" +
      alpha_str +
      "\" result=\"shadow\"/>"
    ()
  }
  // Glow: feGaussianBlur + feFlood + feComposite + feMerge
  if not(eff.glow.is_none()) {
    let rad = emu_to_filter_px(eff.glow.rad, scale)
    let std_dev = rad / 2
    let clr = eff.glow.color
    let alpha_str = eff_alpha_str(clr, "0.7")
    filters = filters +
      "<feGaussianBlur in=\"SourceAlpha\" stdDeviation=\"" +
      int_to_str(if std_dev > 0 { std_dev } else { 1 }) +
      "\" result=\"glowBlur\"/>" +
      "<feFlood flood-color=\"" +
      clr.to_css() +
      "\" flood-opacity=\"" +
      alpha_str +
      "\" result=\"glowColor\"/>" +
      "<feComposite in=\"glowColor\" in2=\"glowBlur\" operator=\"in\" result=\"glowFill\"/>" +
      "<feMerge><feMergeNode in=\"glowFill\"/><feMergeNode in=\"SourceGraphic\"/></feMerge>"
    ()
  }
  // Soft edge: feGaussianBlur on alpha, composite with source
  if not(eff.soft_edge.is_none()) {
    let rad = emu_to_filter_px(eff.soft_edge.rad, scale)
    let std_dev = rad / 2
    filters = filters +
      "<feGaussianBlur in=\"SourceAlpha\" stdDeviation=\"" +
      int_to_str(if std_dev > 0 { std_dev } else { 1 }) +
      "\" result=\"softAlpha\"/>" +
      "<feComposite in=\"SourceGraphic\" in2=\"softAlpha\" operator=\"in\"/>"
    ()
  }
  // Inner shadow: feGaussianBlur+feOffset on inverted alpha, flood+composite, merge
  if not(eff.inner_shadow.is_none()) {
    let ishdw = eff.inner_shadow
    let blur = emu_to_filter_px(ishdw.blur_rad, scale)
    let dist = emu_to_filter_px(ishdw.dist, scale)
    let angle_deg = ishdw.dir / 60000
    let dx = dist * cos_deg(angle_deg) / 100
    let dy = dist * sin_deg(angle_deg) / 100
    let std_dev = blur / 2
    let clr = ishdw.color
    let alpha_str = eff_alpha_str(clr, "0.5")
    // Inner shadow: blur source alpha, offset, invert, flood+composite, merge with source
    filters = filters +
      "<feComponentTransfer in=\"SourceAlpha\" result=\"invAlpha\">" +
      "<feFuncA type=\"table\" tableValues=\"1 0\"/>" +
      "</feComponentTransfer>" +
      "<feGaussianBlur in=\"invAlpha\" stdDeviation=\"" +
      int_to_str(if std_dev > 0 { std_dev } else { 1 }) +
      "\" result=\"isBlur\"/>" +
      "<feOffset in=\"isBlur\" dx=\"" +
      int_to_str(dx) +
      "\" dy=\"" +
      int_to_str(dy) +
      "\" result=\"isOff\"/>" +
      "<feFlood flood-color=\"" +
      clr.to_css() +
      "\" flood-opacity=\"" +
      alpha_str +
      "\"/>" +
      "<feComposite in2=\"isOff\" operator=\"in\" result=\"isFill\"/>" +
      "<feComposite in=\"isFill\" in2=\"SourceAlpha\" operator=\"in\" result=\"isClip\"/>" +
      "<feMerge><feMergeNode in=\"SourceGraphic\"/><feMergeNode in=\"isClip\"/></feMerge>"
    ()
  }
  // Note: Reflection requires structural approach (shape duplication), not SVG filter.
  let def = "<filter id=\"" +
    filter_id +
    "\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\">" +
    filters +
    "</filter>"
  (def, " filter=\"url(#" + filter_id + ")\"")
}

///|
/// Simple cos approximation for shadow direction (degrees → *100 result for int math).
fn cos_deg(deg : Int) -> Int {
  // cos for 0,45,90,135,180,225,270,315 quadrants
  let d = ((deg % 360) + 360) % 360
  if d == 0 { 100 } else if d == 90 { 0 } else if d == 180 {
    -100
  } else if d == 270 {
    0
  } else if d < 90 {
    // Linear interpolation: cos(45)≈71
    100 - d * 100 / 90
  } else if d < 180 {
    -(d - 90) * 100 / 90
  } else if d < 270 {
    -100 + (d - 180) * 100 / 90
  } else {
    (d - 270) * 100 / 90
  }
}

///|
/// Simple sin approximation for shadow direction (degrees → *100 result for int math).
fn sin_deg(deg : Int) -> Int {
  cos_deg(deg - 90)
}
