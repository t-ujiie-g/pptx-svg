// SVG renderer for PPTX slides — Round-trip capable
//
// All OOXML coordinates are in EMU (English Metric Units).
// This renderer converts EMU → pixels at a fixed target width of 960 px,
// producing an SVG with explicit pixel dimensions and NO viewBox transform.
//
// Round-trip support: each SVG element carries data-ooxml-* attributes
// preserving original OOXML values (EMU, hundredths-of-a-point, etc.)
// so the SVG can be parsed back to SlideData losslessly.
//
// NOTE: No StringBuilder.  All strings built via `+` (wasm:js-string "concat"),
// which is polyfillable in Tier-2/3 browsers.

// ── EMU constants ────────────────────────────────────────────────────────────
// 1 pt = 12700 EMU, 1 inch = 914400 EMU

let default_font_size_emu_ : Int = 228600  // 18 pt
let default_l_inset : Int = 91440          // 0.1 in (left / right)
let default_t_inset : Int = 45720          // 0.05 in (top / bottom)
let level_indent_emu : Int = 342900        // indent per bullet level
let svg_target_width : Int = 960           // output SVG width in px
let default_emu_per_px : Int = 9525        // EMU per pixel at 96 DPI
let line_height_numer : Int = 12           // line-height = font-size * line_height_numer / line_height_denom
let line_height_denom : Int = 10

// ── Shared helpers (from @xml) ───────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }

// ── SVG attribute helpers ─────────────────────────────────────────────────────

fn a(name : String, value : String) -> String {
  " " + name + "=\"" + value + "\""
}

fn ai(name : String, value : Int) -> String {
  a(name, int_to_str(value))
}

fn bool_str(b : Bool) -> String {
  if b { "1" } else { "0" }
}

// Escape text for SVG content.
fn svg_escape(s : String) -> String {
  let len = s.length()
  let mut needs = false
  let mut k = 0
  while k < len {
    let c = s.get_char(k).unwrap()
    if c == '<' || c == '>' || c == '&' || c == '"' {
      needs = true
      break
    }
    k = k + 1
  }
  if not(needs) { return s }
  let mut result = ""
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '<' {
      result = result + "&lt;"
    } else if c == '>' {
      result = result + "&gt;"
    } else if c == '&' {
      result = result + "&amp;"
    } else if c == '"' {
      result = result + "&quot;"
    } else {
      result = result + @ffi.ffi_char_code_to_str(Char::to_int(c))
    }
    i = i + 1
  }
  result
}

// ── EMU → pixel conversion ────────────────────────────────────────────────────

fn px(emu : Int, scale : Int) -> Int {
  if scale <= 0 { 0 } else { emu / scale }
}

// ── Font size ─────────────────────────────────────────────────────────────────

fn hundredth_pt_to_emu(sz : Int) -> Int {
  sz * 127
}

fn default_font_size_emu() -> Int {
  default_font_size_emu_
}

// ── Rotation transform ────────────────────────────────────────────────────────

fn rotation_attr(t : @ooxml.ShapeTransform, scale : Int) -> String {
  if t.rot == 0 { return "" }
  let deg_x100 = t.rot / 600
  let deg_int  = deg_x100 / 100
  let deg_frac = deg_x100 % 100
  let frac_str = if deg_frac < 10 {
    "0" + int_to_str(deg_frac)
  } else {
    int_to_str(deg_frac)
  }
  let cx_p = px(t.x + t.cx / 2, scale)
  let cy_p = px(t.y + t.cy / 2, scale)
  " transform=\"rotate(" +
  int_to_str(deg_int) + "." + frac_str + "," +
  int_to_str(cx_p) + "," + int_to_str(cy_p) + ")\""
}

// ── Image helpers ─────────────────────────────────────────────────────────────

fn resolve_image_path(target : String) -> String {
  if target.length() > 3 &&
    target.get_char(0).unwrap() == '.' &&
    target.get_char(1).unwrap() == '.' &&
    target.get_char(2).unwrap() == '/' {
    "ppt/" + suffix(target, 3)
  } else {
    target
  }
}

fn suffix(s : String, from : Int) -> String { @xml.str_suffix(s, from) }

fn mime_from_ext(path : String) -> String {
  let len = path.length()
  let mut dot = len - 1
  while dot >= 0 {
    if path.get_char(dot).unwrap() == '.' { break }
    dot = dot - 1
  }
  if dot < 0 { return "image/png" }
  let ext = suffix(path, dot + 1)
  if ext == "png" { "image/png" }
  else if ext == "jpg" || ext == "jpeg" { "image/jpeg" }
  else if ext == "gif" { "image/gif" }
  else if ext == "svg" || ext == "svgz" { "image/svg+xml" }
  else { "image/png" }
}

// ── data-ooxml-* attribute helpers ───────────────────────────────────────────

fn shape_type_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(_) => "autoshape"
    @ooxml.ShapeKind::Picture(_) => "picture"
    @ooxml.ShapeKind::TableShape(_) => "table"
    @ooxml.ShapeKind::GroupShape => "group"
    @ooxml.ShapeKind::Other => "other"
  }
}

fn geom_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(g) => g.to_prst()
    _ => "rect"
  }
}

// ── Bullet auto-numbering ─────────────────────────────────────────────────────

fn auto_number_text(type_ : String, idx : Int) -> String {
  let n = idx + 1
  if type_ == "arabicPeriod" {
    int_to_str(n) + "."
  } else if type_ == "arabicParenR" {
    int_to_str(n) + ")"
  } else if type_ == "alphaLcPeriod" {
    alpha_lower(n) + "."
  } else if type_ == "alphaLcParenR" {
    alpha_lower(n) + ")"
  } else if type_ == "alphaUcPeriod" {
    alpha_upper(n) + "."
  } else if type_ == "alphaUcParenR" {
    alpha_upper(n) + ")"
  } else if type_ == "romanLcPeriod" {
    roman_lower(n) + "."
  } else if type_ == "romanUcPeriod" {
    roman_upper(n) + "."
  } else {
    int_to_str(n) + "."
  }
}

fn alpha_lower(n : Int) -> String {
  if n >= 1 && n <= 26 {
    @ffi.ffi_char_code_to_str(96 + n)
  } else {
    int_to_str(n)
  }
}

fn alpha_upper(n : Int) -> String {
  if n >= 1 && n <= 26 {
    @ffi.ffi_char_code_to_str(64 + n)
  } else {
    int_to_str(n)
  }
}

fn roman_lower(n : Int) -> String {
  roman_convert(n, true)
}

fn roman_upper(n : Int) -> String {
  roman_convert(n, false)
}

fn roman_convert(n : Int, lower : Bool) -> String {
  if n <= 0 || n > 3999 { return int_to_str(n) }
  let vals  = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
  let syms  = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
  let syml  = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
  let mut result = ""
  let mut rem = n
  let mut i = 0
  while i < 13 {
    while rem >= vals[i] {
      result = result + (if lower { syml[i] } else { syms[i] })
      rem = rem - vals[i]
    }
    i = i + 1
  }
  result
}

// ── Text rendering ────────────────────────────────────────────────────────────

fn para_max_font_emu(para : @ooxml.TextParagraph) -> Int {
  let nr = para.runs.length()
  let mut max_fs = 0
  let mut i = 0
  while i < nr {
    let fs = para.runs[i].font_size
    if fs > max_fs { max_fs = fs }
    i = i + 1
  }
  if max_fs > 0 { hundredth_pt_to_emu(max_fs) } else { 0 }
}

fn para_font_size(para : @ooxml.TextParagraph, scale : Int, def_sz : Int) -> Int {
  let emu = para_max_font_emu(para)
  if emu > 0 { px(emu, scale) } else { def_sz }
}

///| Apply normAutofit font_scale to a pixel value.
///| font_scale: -1=inactive, 0-100000=percentage (100000=100%).
fn apply_font_scale(val : Int, font_scale : Int) -> Int {
  if font_scale >= 0 && font_scale < @ooxml.pct_100 {
    val * font_scale / @ooxml.pct_100
  } else {
    val
  }
}

///| Apply normAutofit lnSpcReduction to a line height.
///| ln_spc_reduction: -1=inactive, 0-100000=reduction percentage.
fn apply_ln_spc_reduction(val : Int, ln_spc_reduction : Int) -> Int {
  if ln_spc_reduction > 0 {
    val * (@ooxml.pct_100 - ln_spc_reduction) / @ooxml.pct_100
  } else {
    val
  }
}

///| Apply both normAutofit adjustments: font_scale then lnSpcReduction.
fn apply_autofit_scale(val : Int, font_scale : Int, ln_spc_reduction : Int) -> Int {
  apply_ln_spc_reduction(apply_font_scale(val, font_scale), ln_spc_reduction)
}

fn para_line_height(para : @ooxml.TextParagraph, scale : Int, def_line_h : Int) -> Int {
  if para.line_spacing > 0 {
    // Positive = EMU absolute value
    px(para.line_spacing, scale)
  } else if para.line_spacing < 0 {
    // Negative = percentage (-100000 = 100% = single spacing)
    // In OOXML, 100% means the default single-line spacing (font_size * 1.2)
    let pct = -(para.line_spacing) // e.g. 150000 for 150%
    let emu = para_max_font_emu(para)
    let base_line_h = if emu > 0 { px(emu, scale) * line_height_numer / line_height_denom } else { def_line_h }
    base_line_h * pct / 100000
  } else {
    // Default: use font size * line_height ratio
    let emu = para_max_font_emu(para)
    if emu > 0 { px(emu, scale) * line_height_numer / line_height_denom } else { def_line_h }
  }
}

fn render_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps
) -> String {
  if paragraphs.length() == 0 { return "" }
  // Insets: use body_props if specified, otherwise default (91440 / 45720 EMU)
  let l_ins = if body_props.l_ins >= 0 { body_props.l_ins } else { default_l_inset }
  let t_ins = if body_props.t_ins >= 0 { body_props.t_ins } else { default_t_inset }
  let r_ins = if body_props.r_ins >= 0 { body_props.r_ins } else { default_l_inset }

  let pad_l   = px(l_ins, scale)
  let pad_t   = px(t_ins, scale)
  let pad_r   = px(r_ins, scale)
  let text_x  = px(t.x, scale) + pad_l
  let text_w  = px(t.cx, scale) - pad_l - pad_r
  let def_sz  = px(default_font_size_emu(), scale) // 24 px for standard slides (unscaled)
  let def_line_h  = def_sz * line_height_numer / line_height_denom // unscaled default line height
  let fs = body_props.font_scale
  let lr = body_props.ln_spc_reduction

  // Calculate total text height for vertical alignment
  let np = paragraphs.length()
  let mut total_text_h = 0
  let mut ti = 0
  while ti < np {
    let para = paragraphs[ti]
    let nr = para.runs.length()
    let para_line_h = apply_autofit_scale(para_line_height(para, scale, def_line_h), fs, lr)
    if nr == 0 {
      total_text_h = total_text_h + para_line_h
    } else {
      // Add spacing before
      if para.spc_before > 0 && ti > 0 {
        total_text_h = total_text_h + px(para.spc_before, scale)
      }
      total_text_h = total_text_h + para_line_h
      // Add spacing after
      if para.spc_after > 0 {
        total_text_h = total_text_h + px(para.spc_after, scale)
      }
    }
    ti = ti + 1
  }

  // Vertical anchor offset
  let shape_h = px(t.cy, scale)
  let b_ins = if body_props.b_ins >= 0 { body_props.b_ins } else { default_t_inset }
  let pad_b = px(b_ins, scale)
  let avail_h = shape_h - pad_t - pad_b
  let anchor_offset = if body_props.anchor == "ctr" {
    (avail_h - total_text_h) / 2
  } else if body_props.anchor == "b" {
    avail_h - total_text_h
  } else {
    0
  }
  let anchor_offset_clamped = if anchor_offset > 0 { anchor_offset } else { 0 }

  let font_face = theme.minor_font + ", sans-serif"
  let scaled_def_sz = apply_font_scale(def_sz, fs)
  let mut out = "<text" +
    ai("x", text_x) +
    a("xml:space", "preserve") +
    a("font-family", font_face) +
    ai("font-size", scaled_def_sz) +
    a("fill", "black") +
    a("data-ooxml-font-face", theme.minor_font) +
    ">"

  let mut cur_y = px(t.y, scale) + pad_t + anchor_offset_clamped
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    let p_line_h = apply_autofit_scale(para_line_height(para, scale, def_line_h), fs, lr)
    let p_font_sz = apply_font_scale(para_font_size(para, scale, def_sz), fs)
    // Apply spacing before paragraph
    if para.spc_before > 0 && pi > 0 {
      cur_y = cur_y + px(para.spc_before, scale)
    }
    // Advance to baseline (ascent ≈ font_size)
    cur_y = cur_y + p_font_sz
    if nr == 0 {
      cur_y = cur_y + (p_line_h - p_font_sz)
      pi = pi + 1
    } else {
      // Indent: use marL from paragraph if set, otherwise use level-based
      let indent_emu = if para.mar_l > 0 { para.mar_l } else { para.level * level_indent_emu }
      let indent = px(indent_emu, scale)
      let (anchor_x, anchor_attr) = if para.align == "ctr" || para.align == "c" {
        (text_x + indent + (text_w - indent) / 2, " text-anchor=\"middle\"")
      } else if para.align == "r" {
        (text_x + text_w, " text-anchor=\"end\"")
      } else {
        (text_x + indent, " text-anchor=\"start\"")
      }
      // Determine bullet text
      let bullet_text = if para.bullet_none {
        ""
      } else if not(para.bullet.is_empty()) {
        para.bullet
      } else if not(para.bullet_auto.is_empty()) {
        auto_number_text(para.bullet_auto, pi)
      } else {
        ""
      }
      // Paragraph-level tspan with data-ooxml-* attributes
      let para_align = if para.align.is_empty() { "l" } else { para.align }
      out = out + "<tspan" +
        ai("x", anchor_x) +
        ai("y", cur_y) +
        anchor_attr +
        a("data-ooxml-para-idx", int_to_str(pi)) +
        a("data-ooxml-para-align", para_align) +
        ai("data-ooxml-para-lvl", para.level) +
        ">"
      // Render bullet if present
      if not(bullet_text.is_empty()) {
        out = out + "<tspan>" + svg_escape(bullet_text) + " </tspan>"
      }
      let mut ri = 0
      while ri < nr {
        let run = para.runs[ri]
        let mut span = "<tspan"
        // Visual attributes
        if run.bold   { span = span + " font-weight=\"bold\""  }
        if run.italic { span = span + " font-style=\"italic\"" }
        if run.font_size > 0 {
          span = span + ai("font-size", apply_font_scale(px(hundredth_pt_to_emu(run.font_size), scale), fs))
        }
        // font-family: ea_font, latin_font, sans-serif
        if not(run.font_face.is_empty()) || not(run.ea_font.is_empty()) {
          let mut ff = ""
          if not(run.ea_font.is_empty()) { ff = run.ea_font }
          if not(run.font_face.is_empty()) {
            if not(ff.is_empty()) { ff = ff + ", " }
            ff = ff + run.font_face
          }
          span = span + a("font-family", ff + ", sans-serif")
        }
        if not(run.color.is_none()) {
          span = span + a("fill", run.color.to_css())
        }
        // Character spacing
        if run.char_spacing != 0 {
          // hundredths-of-pt → px: val * 96 / 7200
          let ls_px = run.char_spacing * 96 / 7200
          if ls_px != 0 {
            span = span + ai("letter-spacing", ls_px)
          }
        }
        // Text decorations
        let mut deco = ""
        if not(run.underline.is_empty()) && run.underline != "none" {
          deco = "underline"
        }
        if run.strike == "sngStrike" || run.strike == "dblStrike" {
          if deco.is_empty() { deco = "line-through" } else { deco = deco + " line-through" }
        }
        if not(deco.is_empty()) {
          span = span + a("text-decoration", deco)
        }
        // Superscript / subscript via baseline-shift
        if run.baseline > 0 {
          span = span + a("baseline-shift", "super")
          span = span + a("font-size", "70%")
        } else if run.baseline < 0 {
          span = span + a("baseline-shift", "sub")
          span = span + a("font-size", "70%")
        }
        // data-ooxml-* attributes for round-trip
        span = span +
          a("data-ooxml-run-idx", int_to_str(ri)) +
          a("data-ooxml-bold", bool_str(run.bold)) +
          a("data-ooxml-font-size", int_to_str(run.font_size)) +
          a("data-ooxml-color", run.color.to_hex()) +
          a("data-ooxml-run-font", run.font_face) +
          a("data-ooxml-ea-font", run.ea_font) +
          a("data-ooxml-underline", run.underline) +
          a("data-ooxml-strike", run.strike) +
          ai("data-ooxml-baseline", run.baseline) +
          ai("data-ooxml-char-spacing", run.char_spacing)
        out = out + span + ">" + svg_escape(run.text) + "</tspan>"
        ri = ri + 1
      }
      out = out + "</tspan>"
      cur_y = cur_y + (p_line_h - p_font_sz)
      // Apply spacing after paragraph
      if para.spc_after > 0 {
        cur_y = cur_y + px(para.spc_after, scale)
      }
      pi = pi + 1
    }
  }
  out + "</text>"
}

// ── Text height estimation ──────────────────────────────────────────────────

fn estimate_text_height(
  paragraphs : Array[@ooxml.TextParagraph],
  scale : Int,
  body_props : @ooxml.BodyProps
) -> Int {
  let def_sz = px(default_font_size_emu(), scale)
  let def_line_h = def_sz * line_height_numer / line_height_denom
  let t_ins = if body_props.t_ins >= 0 { body_props.t_ins } else { default_t_inset }
  let b_ins = if body_props.b_ins >= 0 { body_props.b_ins } else { default_t_inset }
  let np = paragraphs.length()
  let mut total = 0
  let mut i = 0
  while i < np {
    let para = paragraphs[i]
    let plh = para_line_height(para, scale, def_line_h)
    if para.spc_before > 0 && i > 0 {
      total = total + px(para.spc_before, scale)
    }
    total = total + plh
    if para.spc_after > 0 {
      total = total + px(para.spc_after, scale)
    }
    i = i + 1
  }
  total + px(t_ins, scale) + px(b_ins, scale)
}

// ── Shape rendering ───────────────────────────────────────────────────────────

fn render_shape(
  shape : @ooxml.Shape,
  rels : Array[@ooxml.Relationship],
  _slide_num : Int,
  scale : Int,
  shape_idx : Int,
  theme : @ooxml.ThemeData
) -> String {
  let t    = shape.transform
  let x_p  = px(t.x,  scale)
  let y_p  = px(t.y,  scale)
  let cx_p = px(t.cx, scale)
  let mut cy_p = px(t.cy, scale)
  // Apply autoFit: expand shape height if text overflows (never shrink)
  if shape.body_props.auto_fit && shape.paragraphs.length() > 0 {
    let text_h = estimate_text_height(shape.paragraphs, scale, shape.body_props)
    if text_h > cy_p {
      cy_p = text_h
    }
  }
  if cx_p <= 0 || cy_p <= 0 { return "" }
  let rot = rotation_attr(t, scale)

  // Build data-ooxml-* attributes for the <g> wrapper
  let g_attrs =
    a("data-ooxml-shape-type", shape_type_str(shape.kind)) +
    a("data-ooxml-geom", geom_str(shape.kind)) +
    ai("data-ooxml-x", t.x) +
    ai("data-ooxml-y", t.y) +
    ai("data-ooxml-cx", t.cx) +
    ai("data-ooxml-cy", t.cy) +
    ai("data-ooxml-rot", t.rot) +
    a("data-ooxml-fill", shape.fill.to_hex()) +
    a("data-ooxml-stroke", shape.stroke.to_hex()) +
    ai("data-ooxml-stroke-w", shape.stroke_w) +
    ai("data-ooxml-shape-idx", shape_idx) +
    a("data-ooxml-ph-type", shape.ph_type) +
    ai("data-ooxml-ph-idx", shape.ph_idx) +
    a("data-ooxml-anchor", shape.body_props.anchor) +
    ai("data-ooxml-l-ins", shape.body_props.l_ins) +
    ai("data-ooxml-t-ins", shape.body_props.t_ins) +
    ai("data-ooxml-r-ins", shape.body_props.r_ins) +
    ai("data-ooxml-b-ins", shape.body_props.b_ins) +
    a("data-ooxml-auto-fit", bool_str(shape.body_props.auto_fit)) +
    ai("data-ooxml-font-scale", shape.body_props.font_scale) +
    ai("data-ooxml-ln-spc-reduction", shape.body_props.ln_spc_reduction)

  let inner = match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => {
      let fill   = a("fill", shape.fill.to_css())
      let stroke = if shape.stroke.is_none() {
        " stroke=\"none\""
      } else {
        a("stroke", shape.stroke.to_css()) +
        ai("stroke-width",
          px(if shape.stroke_w > 0 { shape.stroke_w } else { @ooxml.default_stroke_w }, scale))
      }
      let rect_svg = match geom {
        @ooxml.ShapeGeom::Ellipse => {
          "<ellipse" +
          ai("cx", x_p + cx_p / 2) + ai("cy", y_p + cy_p / 2) +
          ai("rx", cx_p / 2) + ai("ry", cy_p / 2) +
          fill + stroke + rot + "/>"
        }
        _ => {
          let rx = match geom {
            @ooxml.ShapeGeom::RoundRect => int_to_str(cx_p / 20)
            _ => "0"
          }
          let base = "<rect" +
            ai("x", x_p) + ai("y", y_p) +
            ai("width", cx_p) + ai("height", cy_p) +
            fill + stroke + rot
          if rx == "0" {
            base + "/>"
          } else {
            base + a("rx", rx) + a("ry", rx) + "/>"
          }
        }
      }
      // For autoFit, use adjusted cy for text rendering if shape was expanded
      let text_t : @ooxml.ShapeTransform = if shape.body_props.auto_fit && cy_p > px(t.cy, scale) {
        { x: t.x, y: t.y, cx: t.cx, cy: cy_p * scale, rot: t.rot, flip_h: t.flip_h, flip_v: t.flip_v }
      } else {
        t
      }
      rect_svg + render_text(shape.paragraphs, text_t, scale, theme, shape.body_props)
    }
    @ooxml.ShapeKind::Picture(rid) => {
      let target = @ooxml.find_rel_target(rels, rid)
      let placeholder = "<rect" +
        ai("x", x_p) + ai("y", y_p) +
        ai("width", cx_p) + ai("height", cy_p) +
        " fill=\"#dddddd\"" + rot + "/>"
      if target.is_empty() {
        placeholder
      } else {
        let img_path = resolve_image_path(target)
        let b64 = @ffi.ffi_get_file_base64(img_path)
        if b64.is_empty() {
          placeholder
        } else {
          let mime = mime_from_ext(img_path)
          "<image" +
          ai("x", x_p) + ai("y", y_p) +
          ai("width", cx_p) + ai("height", cy_p) +
          " preserveAspectRatio=\"xMidYMid meet\"" +
          a("href", "data:" + mime + ";base64," + b64) +
          rot + "/>"
        }
      }
    }
    @ooxml.ShapeKind::TableShape(table) => {
      render_table(table, t, scale, theme)
    }
    _ => ""
  }

  if inner.is_empty() {
    ""
  } else {
    "<g" + g_attrs + ">" + inner + "</g>"
  }
}

// ── Table rendering ─────────────────────────────────────────────────────────

fn render_table_cell_text(
  paragraphs : Array[@ooxml.TextParagraph],
  cell_x : Int,
  cell_y : Int,
  _cell_w : Int,
  cell_h : Int,
  scale : Int,
  theme : @ooxml.ThemeData
) -> String {
  if paragraphs.length() == 0 { return "" }
  let pad = 4
  let tx = cell_x + pad
  let def_sz = px(default_font_size_emu(), scale)
  let font_sz = if def_sz > cell_h - pad * 2 { cell_h - pad * 2 } else { def_sz }
  let line_h = font_sz * line_height_numer / line_height_denom
  let font_face = theme.minor_font + ", sans-serif"
  let mut out = "<text" +
    ai("x", tx) +
    a("xml:space", "preserve") +
    a("font-family", font_face) +
    ai("font-size", font_sz) +
    a("fill", "black") +
    ">"
  let mut cur_y = cell_y + pad + font_sz
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    if nr == 0 {
      cur_y = cur_y + line_h
    } else {
      let mut ri = 0
      while ri < nr {
        let run = para.runs[ri]
        let mut span = "<tspan"
        span = span + ai("x", tx) + ai("y", cur_y)
        if run.bold { span = span + " font-weight=\"bold\"" }
        if run.italic { span = span + " font-style=\"italic\"" }
        if run.font_size > 0 {
          span = span + ai("font-size", px(hundredth_pt_to_emu(run.font_size), scale))
        }
        if not(run.font_face.is_empty()) || not(run.ea_font.is_empty()) {
          let mut ff = ""
          if not(run.ea_font.is_empty()) { ff = run.ea_font }
          if not(run.font_face.is_empty()) {
            if not(ff.is_empty()) { ff = ff + ", " }
            ff = ff + run.font_face
          }
          span = span + a("font-family", ff + ", sans-serif")
        }
        if not(run.color.is_none()) {
          span = span + a("fill", run.color.to_css())
        }
        out = out + span + ">" + svg_escape(run.text) + "</tspan>"
        ri = ri + 1
      }
      cur_y = cur_y + line_h
    }
    pi = pi + 1
    // Stop if we overflow the cell
    if cur_y > cell_y + cell_h - pad { break }
  }
  out + "</text>"
}

fn render_table(
  table : @ooxml.TableData,
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData
) -> String {
  let base_x = px(t.x, scale)
  let base_y = px(t.y, scale)
  let mut out = ""
  // Precompute column x-offsets in pixels
  let ncols = table.col_widths.length()
  let col_x : Array[Int] = []
  let mut cx_accum = 0
  let mut ci = 0
  while ci < ncols {
    col_x.push(px(cx_accum, scale))
    cx_accum = cx_accum + table.col_widths[ci]
    ci = ci + 1
  }
  // Render rows
  let nrows = table.rows.length()
  let mut row_y_emu = 0
  let mut ri = 0
  while ri < nrows {
    let row = table.rows[ri]
    let ry = base_y + px(row_y_emu, scale)
    let rh = px(row.height, scale)
    let ncells = row.cells.length()
    let mut cj = 0
    while cj < ncells && cj < ncols {
      let cell = row.cells[cj]
      let cx_px = base_x + col_x[cj]
      let cw = px(table.col_widths[cj], scale)
      // Cell background
      let cell_fill = if cell.fill.is_none() { "none" } else { cell.fill.to_css() }
      out = out + "<rect" +
        ai("x", cx_px) + ai("y", ry) +
        ai("width", cw) + ai("height", rh) +
        a("fill", cell_fill) +
        " stroke=\"#999999\" stroke-width=\"1\"/>"
      // Cell text
      out = out + render_table_cell_text(cell.paragraphs, cx_px, ry, cw, rh, scale, theme)
      cj = cj + 1
    }
    row_y_emu = row_y_emu + row.height
    ri = ri + 1
  }
  out
}

// ── Public API ────────────────────────────────────────────────────────────────

///| Render a SlideData to an SVG string with data-ooxml-* attributes.
pub fn render_slide_svg(
  slide : @ooxml.SlideData,
  rels : Array[@ooxml.Relationship],
  slide_num : Int,
  theme : @ooxml.ThemeData,
  bg_shapes : Array[@ooxml.Shape]
) -> String {
  let sw = slide.slide_size.cx
  let sh = slide.slide_size.cy
  let w = svg_target_width
  let scale = if sw > 0 { sw / w } else { default_emu_per_px }
  let h = if scale > 0 { sh / scale } else { svg_target_width * 9 / 16 }
  let header = "<svg" +
    " xmlns=\"http://www.w3.org/2000/svg\"" +
    " xmlns:xlink=\"http://www.w3.org/1999/xlink\"" +
    ai("width", w) + ai("height", h) +
    " style=\"background:" + slide.background.to_css() + ";display:block\"" +
    // data-ooxml-* metadata on SVG root for round-trip
    ai("data-ooxml-slide-cx", sw) +
    ai("data-ooxml-slide-cy", sh) +
    a("data-ooxml-bg", slide.background.to_hex()) +
    ai("data-ooxml-scale", scale) +
    ">"
  let bg_rect = "<rect" +
    ai("width", w) + ai("height", h) +
    a("fill", slide.background.to_css()) +
    "/>"
  // Render master/layout background shapes (non-placeholder)
  let mut bg_shapes_svg = ""
  let nbg = bg_shapes.length()
  let mut bi = 0
  while bi < nbg {
    bg_shapes_svg = bg_shapes_svg + render_shape(bg_shapes[bi], rels, slide_num, scale, -(bi + 1), theme)
    bi = bi + 1
  }
  // Render slide shapes
  let mut shapes_svg = ""
  let n = slide.shapes.length()
  let mut i = 0
  while i < n {
    shapes_svg = shapes_svg + render_shape(slide.shapes[i], rels, slide_num, scale, i, theme)
    i = i + 1
  }
  header + bg_rect + bg_shapes_svg + shapes_svg + "</svg>"
}
