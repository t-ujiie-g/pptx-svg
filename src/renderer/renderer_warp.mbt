// SVG renderer — Text warp rendering
//
// Implements visual rendering of prstTxWarp presets using SVG <textPath> for
// curved warps and SVG transforms for linear warps (slant, fade, cascade).
//
// Same package as renderer.mbt — all constants and helpers are shared.

// ── Default adjustment values ────────────────────────────────────────────────

///|
/// Return default adjustment value for a warp preset (OOXML 60000th-degree for
/// angles, 0-100000 for percentages).  idx=1 for av1, idx=2 for av2.
fn warp_default_av(prst : String, idx : Int) -> Int {
  if idx == 1 {
    // av1 defaults
    if prst == "textArchUp" || prst == "textArchUpPour" {
      return 10800000
    }
    if prst == "textArchDown" || prst == "textArchDownPour" {
      return 0
    }
    if prst == "textCircle" || prst == "textCirclePour" {
      return 10800000
    }
    if prst == "textButton" || prst == "textButtonPour" {
      return 10800000
    }
    if prst == "textCurveUp" || prst == "textCurveDown" {
      return 45977
    }
    if prst == "textCanUp" || prst == "textCanDown" {
      return 18000
    }
    if prst == "textWave1" ||
      prst == "textWave2" ||
      prst == "textWave4" ||
      prst == "textDoubleWave1" {
      return 12500
    }
    if prst == "textInflate" ||
      prst == "textDeflate" ||
      prst == "textInflateBottom" ||
      prst == "textInflateTop" ||
      prst == "textDeflateBottom" ||
      prst == "textDeflateTop" {
      return 18750
    }
    if prst == "textDeflateInflate" {
      return 25000
    }
    if prst == "textDeflateInflateDeflate" {
      return 16667
    }
    if prst == "textChevron" || prst == "textChevronInverted" {
      return 25000
    }
    if prst == "textTriangle" || prst == "textTriangleInverted" {
      return 50000
    }
    if prst == "textRingInside" || prst == "textRingOutside" {
      return 10800000
    }
    if prst == "textSlantUp" || prst == "textSlantDown" {
      return 55556
    }
    if prst == "textFadeRight" ||
      prst == "textFadeLeft" ||
      prst == "textFadeUp" ||
      prst == "textFadeDown" {
      return 33333
    }
    if prst == "textCascadeUp" || prst == "textCascadeDown" {
      return 44444
    }
    if prst == "textStop" {
      return 25000
    }
    return 50000 // generic default
  }
  if idx == 2 {
    // av2 defaults — most presets only use av1
    if prst == "textWave1" ||
      prst == "textWave2" ||
      prst == "textWave4" ||
      prst == "textDoubleWave1" {
      return 0
    }
    if prst == "textArchUpPour" ||
      prst == "textArchDownPour" ||
      prst == "textCirclePour" ||
      prst == "textButtonPour" {
      return 50000
    }
    return 0
  }
  0
}

// ── Text flattening ──────────────────────────────────────────────────────────

///|
/// Flatten all paragraph/run text into a single string for warp rendering.
/// Returns (text, font_size_px, font_family, fill_color, bold, italic).
fn flatten_warp_text(
  paragraphs : Array[@ooxml.TextParagraph],
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
) -> (String, Int, String, @ooxml.Color, Bool, Bool) {
  let mut text = ""
  let mut first_font_size = -1
  let mut first_ff = ""
  let mut first_color = @ooxml.Color::{ r: -1, g: -1, b: -1, alpha: 255 }
  let mut first_bold = false
  let mut first_italic = false
  let mut found_style = false
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    if pi > 0 && !(text.is_empty()) {
      text = text + " "
    }
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      if run.text != "\n" {
        let display = if run.cap == "all" {
          to_upper_ascii(run.text)
        } else {
          run.text
        }
        text = text + display
        if !(found_style) && !(run.text.is_empty()) && run.text != "\n" {
          first_font_size = run.font_size
          first_ff = if !(run.font_face.is_empty()) {
            run.font_face
          } else if !(run.ea_font.is_empty()) {
            run.ea_font
          } else {
            theme.minor_font
          }
          first_color = run.color
          first_bold = run.bold
          first_italic = run.italic
          found_style = true
        }
      }
      ri = ri + 1
    }
    pi = pi + 1
  }
  // Resolve font size
  let def_sz = px(default_font_size_emu(), scale)
  let fs_raw = if first_font_size > 0 {
    px(first_font_size * @ooxml.emu_per_hundredth_pt, scale)
  } else {
    def_sz
  }
  // Apply font scale from normAutofit
  let font_sc = body_props.font_scale
  let fs = apply_font_scale(fs_raw, font_sc)
  let ff = if first_ff.is_empty() { theme.minor_font } else { first_ff }
  (text, fs, ff, first_color, first_bold, first_italic)
}

// ── Warp category ────────────────────────────────────────────────────────────

///|
/// Returns true if the warp preset needs SVG <textPath> (curved path).
/// Returns false for transform-only warps (slant, fade, cascade).
fn is_path_warp(prst : String) -> Bool {
  // Transform-only warps
  if prst == "textSlantUp" ||
    prst == "textSlantDown" ||
    prst == "textFadeRight" ||
    prst == "textFadeLeft" ||
    prst == "textFadeUp" ||
    prst == "textFadeDown" ||
    prst == "textCascadeUp" ||
    prst == "textCascadeDown" {
    return false
  }
  // Everything else is path-based
  true
}

// ── SVG path generation ──────────────────────────────────────────────────────

///|
/// Generate SVG path d-string for a warp preset.
/// w, h are shape dimensions in pixels (as Double).
/// av1, av2 are resolved adjustment values (OOXML units).
fn generate_warp_path(
  prst : String,
  w : Double,
  h : Double,
  av1 : Double,
  av2 : Double,
) -> String {
  let cx = w / 2.0
  let cy = h / 2.0
  // ── Arch warps ──
  if prst == "textArchUp" {
    // Semicircular arc, text follows top half
    let rx = cx
    let ry = cy
    return "M " +
      dts(cx - rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx + rx) +
      "," +
      dts(cy)
  }
  if prst == "textArchDown" {
    let rx = cx
    let ry = cy
    return "M " +
      dts(cx + rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx - rx) +
      "," +
      dts(cy)
  }
  if prst == "textArchUpPour" {
    // Same as archUp but with adjustable radius
    let ratio = av2 / 100000.0
    let rx = cx * ratio
    let ry = cy * ratio
    if rx < 1.0 || ry < 1.0 {
      return "M 0," + dts(cy) + " L " + dts(w) + "," + dts(cy)
    }
    return "M " +
      dts(cx - rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx + rx) +
      "," +
      dts(cy)
  }
  if prst == "textArchDownPour" {
    let ratio = av2 / 100000.0
    let rx = cx * ratio
    let ry = cy * ratio
    if rx < 1.0 || ry < 1.0 {
      return "M 0," + dts(cy) + " L " + dts(w) + "," + dts(cy)
    }
    return "M " +
      dts(cx + rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx - rx) +
      "," +
      dts(cy)
  }
  // ── Circle ──
  if prst == "textCircle" || prst == "textCirclePour" {
    let rx = cx
    let ry = cy
    // Full circle = two semicircular arcs
    return "M " +
      dts(cx - rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx + rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx - rx) +
      "," +
      dts(cy)
  }
  // ── Ring warps ──
  if prst == "textRingInside" {
    let rx = cx
    let ry = cy
    return "M " +
      dts(cx - rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx + rx) +
      "," +
      dts(cy)
  }
  if prst == "textRingOutside" {
    let rx = cx
    let ry = cy
    return "M " +
      dts(cx + rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 1,1 " +
      dts(cx - rx) +
      "," +
      dts(cy)
  }
  // ── Button warps ──
  if prst == "textButton" || prst == "textButtonPour" {
    // Button = arch across top
    let rx = cx
    let ry = h * 0.25
    return "M " +
      dts(cx - rx) +
      "," +
      dts(cy) +
      " A " +
      dts(rx) +
      "," +
      dts(ry) +
      " 0 0,1 " +
      dts(cx + rx) +
      "," +
      dts(cy)
  }
  // ── Wave warps ──
  if prst == "textWave1" {
    let amp = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.25) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w * 0.75) +
      "," +
      dts(cy + amp) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textWave2" {
    let amp = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.25) +
      "," +
      dts(cy + amp) +
      " " +
      dts(w * 0.75) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textWave4" {
    // More pronounced wave
    let amp = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.125) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w * 0.375) +
      "," +
      dts(cy + amp) +
      " " +
      dts(cx) +
      "," +
      dts(cy) +
      " C " +
      dts(w * 0.625) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w * 0.875) +
      "," +
      dts(cy + amp) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textDoubleWave1" {
    // Double wave = two periods
    let amp = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.125) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w * 0.375) +
      "," +
      dts(cy + amp) +
      " " +
      dts(cx) +
      "," +
      dts(cy) +
      " C " +
      dts(w * 0.625) +
      "," +
      dts(cy - amp) +
      " " +
      dts(w * 0.875) +
      "," +
      dts(cy + amp) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  // ── Curve warps ──
  if prst == "textCurveUp" {
    let adj = h * av1 / 100000.0
    return "M 0," +
      dts(cy + adj) +
      " Q " +
      dts(cx) +
      "," +
      dts(cy - adj) +
      " " +
      dts(w) +
      "," +
      dts(cy + adj)
  }
  if prst == "textCurveDown" {
    let adj = h * av1 / 100000.0
    return "M 0," +
      dts(cy - adj) +
      " Q " +
      dts(cx) +
      "," +
      dts(cy + adj) +
      " " +
      dts(w) +
      "," +
      dts(cy - adj)
  }
  // ── Can warps ──
  if prst == "textCanUp" {
    let adj = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " A " +
      dts(cx) +
      "," +
      dts(adj) +
      " 0 0,1 " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textCanDown" {
    let adj = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " A " +
      dts(cx) +
      "," +
      dts(adj) +
      " 0 0,0 " +
      dts(w) +
      "," +
      dts(cy)
  }
  // ── Chevron warps ──
  if prst == "textChevron" {
    let adj_y = h * av1 / 100000.0
    return "M 0," +
      dts(adj_y) +
      " L " +
      dts(cx) +
      ",0 L " +
      dts(w) +
      "," +
      dts(adj_y)
  }
  if prst == "textChevronInverted" {
    let adj_y = h * av1 / 100000.0
    return "M 0," +
      dts(h - adj_y) +
      " L " +
      dts(cx) +
      "," +
      dts(h) +
      " L " +
      dts(w) +
      "," +
      dts(h - adj_y)
  }
  // ── Triangle warps ──
  if prst == "textTriangle" {
    let adj_y = h * av1 / 100000.0
    return "M 0," +
      dts(adj_y) +
      " L " +
      dts(cx) +
      ",0 L " +
      dts(w) +
      "," +
      dts(adj_y)
  }
  if prst == "textTriangleInverted" {
    let adj_y = h * av1 / 100000.0
    return "M 0," +
      dts(h - adj_y) +
      " L " +
      dts(cx) +
      "," +
      dts(h) +
      " L " +
      dts(w) +
      "," +
      dts(h - adj_y)
  }
  // ── Inflate/Deflate warps ──
  if prst == "textDeflate" {
    let dip = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textInflate" {
    let bow = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(cy - bow) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(cy - bow) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textDeflateBottom" {
    let dip = h * av1 / 100000.0
    return "M 0," +
      dts(h * 0.75) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(h * 0.75 + dip) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(h * 0.75 + dip) +
      " " +
      dts(w) +
      "," +
      dts(h * 0.75)
  }
  if prst == "textDeflateTop" {
    let dip = h * av1 / 100000.0
    return "M 0," +
      dts(h * 0.25) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(h * 0.25 - dip) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(h * 0.25 - dip) +
      " " +
      dts(w) +
      "," +
      dts(h * 0.25)
  }
  if prst == "textInflateBottom" {
    let bow = h * av1 / 100000.0
    return "M 0," +
      dts(h * 0.75) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(h * 0.75 - bow) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(h * 0.75 - bow) +
      " " +
      dts(w) +
      "," +
      dts(h * 0.75)
  }
  if prst == "textInflateTop" {
    let bow = h * av1 / 100000.0
    return "M 0," +
      dts(h * 0.25) +
      " C " +
      dts(w / 3.0) +
      "," +
      dts(h * 0.25 + bow) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(h * 0.25 + bow) +
      " " +
      dts(w) +
      "," +
      dts(h * 0.25)
  }
  if prst == "textDeflateInflate" {
    // Two-segment: deflate then inflate
    let dip = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.167) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w * 0.333) +
      "," +
      dts(cy + dip) +
      " " +
      dts(cx) +
      "," +
      dts(cy) +
      " C " +
      dts(w * 0.667) +
      "," +
      dts(cy - dip) +
      " " +
      dts(w * 0.833) +
      "," +
      dts(cy - dip) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  if prst == "textDeflateInflateDeflate" {
    let dip = h * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " C " +
      dts(w * 0.111) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w * 0.222) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w / 3.0) +
      "," +
      dts(cy) +
      " C " +
      dts(w * 0.444) +
      "," +
      dts(cy - dip) +
      " " +
      dts(w * 0.556) +
      "," +
      dts(cy - dip) +
      " " +
      dts(w * 2.0 / 3.0) +
      "," +
      dts(cy) +
      " C " +
      dts(w * 0.778) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w * 0.889) +
      "," +
      dts(cy + dip) +
      " " +
      dts(w) +
      "," +
      dts(cy)
  }
  // ── Stop ──
  if prst == "textStop" {
    // Octagonal path approximated as straight segments
    let adj = w * av1 / 100000.0
    return "M 0," +
      dts(cy) +
      " L " +
      dts(adj) +
      "," +
      dts(cy * 0.3) +
      " L " +
      dts(w - adj) +
      "," +
      dts(cy * 0.3) +
      " L " +
      dts(w) +
      "," +
      dts(cy)
  }
  // ── Fallback: horizontal line through center ──
  "M 0," + dts(cy) + " L " + dts(w) + "," + dts(cy)
}

// ── Transform generation for linear warps ────────────────────────────────────

///|
/// Generate SVG transform string for linear (non-path) warps.
/// Returns the transform attribute content (without the "transform=" wrapper).
fn generate_warp_transform(
  prst : String,
  w : Double,
  h : Double,
  av1 : Double,
) -> String {
  if prst == "textSlantUp" {
    // Skew upward: map av1 (0-100000) to skew angle
    let ratio = av1 / 100000.0
    let angle = -ratio * 30.0 // max 30 degrees
    return "skewY(" + dts(angle) + ")"
  }
  if prst == "textSlantDown" {
    let ratio = av1 / 100000.0
    let angle = ratio * 30.0
    return "skewY(" + dts(angle) + ")"
  }
  if prst == "textCascadeUp" {
    let ratio = av1 / 100000.0
    let angle = -ratio * 20.0
    return "skewY(" + dts(angle) + ")"
  }
  if prst == "textCascadeDown" {
    let ratio = av1 / 100000.0
    let angle = ratio * 20.0
    return "skewY(" + dts(angle) + ")"
  }
  if prst == "textFadeRight" {
    // Perspective fade to the right — approximate with scale matrix
    let ratio = 1.0 - av1 / 100000.0
    let sy = if ratio < 0.1 { 0.1 } else { ratio }
    let ty = h * (1.0 - sy) / 2.0
    return "translate(" +
      dts(w) +
      ",0) scale(-1,1) translate(" +
      dts(-w) +
      ",0) matrix(1,0,0," +
      dts(sy) +
      ",0," +
      dts(ty) +
      ") translate(" +
      dts(w) +
      ",0) scale(-1,1) translate(" +
      dts(-w) +
      ",0)"
  }
  if prst == "textFadeLeft" {
    let ratio = 1.0 - av1 / 100000.0
    let sy = if ratio < 0.1 { 0.1 } else { ratio }
    let ty = h * (1.0 - sy) / 2.0
    return "matrix(1,0,0," + dts(sy) + ",0," + dts(ty) + ")"
  }
  if prst == "textFadeUp" {
    let ratio = 1.0 - av1 / 100000.0
    let sx = if ratio < 0.1 { 0.1 } else { ratio }
    let tx = w * (1.0 - sx) / 2.0
    return "matrix(" + dts(sx) + ",0,0,1," + dts(tx) + ",0)"
  }
  if prst == "textFadeDown" {
    let ratio = 1.0 - av1 / 100000.0
    let sx = if ratio < 0.1 { 0.1 } else { ratio }
    let tx = w * (1.0 - sx) / 2.0
    return "translate(0," +
      dts(h) +
      ") scale(1,-1) matrix(" +
      dts(sx) +
      ",0,0,1," +
      dts(tx) +
      ",0) scale(1,-1) translate(0," +
      dts(-h) +
      ")"
  }
  // Fallback: no transform
  ""
}

// ── Main warp text renderer ──────────────────────────────────────────────────

///|
/// Render warped text as SVG. Called when body_props.warp_prst is a non-trivial
/// preset (not empty, not textNoShape, not textPlain).
fn render_warp_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
  rels : Array[@ooxml.Relationship],
  slide_num : Int,
  shape_idx : Int,
) -> String {
  let (text, font_sz, ff, color, bold, italic) = flatten_warp_text(
    paragraphs, scale, theme, body_props,
  )
  if text.is_empty() {
    return ""
  }
  let prst = body_props.warp_prst
  // Resolve adjustment values
  let av1 = if body_props.warp_av1 >= 0 {
    body_props.warp_av1
  } else {
    warp_default_av(prst, 1)
  }
  let av2 = if body_props.warp_av2 >= 0 {
    body_props.warp_av2
  } else {
    warp_default_av(prst, 2)
  }
  let x_p = px(t.x, scale)
  let y_p = px(t.y, scale)
  let w_p = px(t.cx, scale)
  let h_p = px(t.cy, scale)
  let w = w_p.to_double()
  let h = h_p.to_double()
  let av1_d = av1.to_double()
  let av2_d = av2.to_double()
  // Build style attributes
  let fill_attr = if !(color.is_none()) {
    " fill=\"" + color.to_css() + "\""
  } else {
    " fill=\"" + default_text_fill + "\""
  }
  let bold_attr = if bold { " font-weight=\"bold\"" } else { "" }
  let italic_attr = if italic { " font-style=\"italic\"" } else { "" }
  let font_family = ff + ", sans-serif"
  if is_path_warp(prst) {
    // ── Path-based warp ──
    let path_id = "warp-s" + int_to_str(slide_num) + "-" + int_to_str(shape_idx)
    let d = generate_warp_path(prst, w, h, av1_d, av2_d)
    // Determine text-anchor and startOffset based on warp type
    let (start_offset, text_anchor) = if prst == "textChevron" ||
      prst == "textChevronInverted" ||
      prst == "textTriangle" ||
      prst == "textTriangleInverted" ||
      prst == "textStop" {
      ("50%", "middle")
    } else if prst == "textArchUp" ||
      prst == "textArchDown" ||
      prst == "textArchUpPour" ||
      prst == "textArchDownPour" ||
      prst == "textCircle" ||
      prst == "textCirclePour" ||
      prst == "textRingInside" ||
      prst == "textRingOutside" ||
      prst == "textButton" ||
      prst == "textButtonPour" {
      ("25%", "middle")
    } else {
      ("0%", "start")
    }
    // dominant-baseline for vertical positioning
    let baseline = " dominant-baseline=\"central\""
    "<defs><path id=\"" +
    path_id +
    "\" d=\"" +
    d +
    "\" fill=\"none\" transform=\"translate(" +
    int_to_str(x_p) +
    "," +
    int_to_str(y_p) +
    ")\"/></defs>" +
    "<text" +
    " font-family=\"" +
    svg_escape(font_family) +
    "\"" +
    ai("font-size", font_sz) +
    fill_attr +
    bold_attr +
    italic_attr +
    baseline +
    " xml:space=\"preserve\"" +
    "><textPath href=\"#" +
    path_id +
    "\" startOffset=\"" +
    start_offset +
    "\" text-anchor=\"" +
    text_anchor +
    "\">" +
    svg_escape(text) +
    "</textPath></text>"
  } else {
    // ── Transform-based warp ──
    let xform = generate_warp_transform(prst, w, h, av1_d)
    if xform.is_empty() {
      // Unknown transform warp — fallback to normal rendering
      return render_text(
        paragraphs, t, scale, theme, body_props, rels, slide_num, shape_idx,
      )
    }
    // Wrap normal text rendering in a transform group
    // Use a modified body_props with cleared warp to avoid infinite recursion
    let plain_bp : @ooxml.BodyProps = {
      anchor: body_props.anchor,
      l_ins: body_props.l_ins,
      t_ins: body_props.t_ins,
      r_ins: body_props.r_ins,
      b_ins: body_props.b_ins,
      auto_fit: body_props.auto_fit,
      font_scale: body_props.font_scale,
      ln_spc_reduction: body_props.ln_spc_reduction,
      wrap: body_props.wrap,
      rot: body_props.rot,
      vert: body_props.vert,
      num_cols: body_props.num_cols,
      col_spacing: body_props.col_spacing,
      warp_prst: "", // cleared to avoid recursion
      warp_av1: -1,
      warp_av2: -1,
    }
    let inner = render_text(
      paragraphs, t, scale, theme, plain_bp, rels, slide_num, shape_idx,
    )
    // Apply transform centered on the shape
    let tcx = x_p + w_p / 2
    let tcy = y_p + h_p / 2
    "<g transform=\"translate(" +
    int_to_str(tcx) +
    "," +
    int_to_str(tcy) +
    ") " +
    xform +
    " translate(" +
    int_to_str(-tcx) +
    "," +
    int_to_str(-tcy) +
    ")\">" +
    inner +
    "</g>"
  }
}
