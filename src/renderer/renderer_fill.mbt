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
