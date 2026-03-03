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
let table_cell_pad : Int = 4              // padding inside table cells (px)
let default_text_fill : String = "black"
let image_placeholder_fill : String = "#dddddd"
let table_border_color : String = "#999999"

// ── Shared helpers (from @xml) ───────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }

// ── SVG attribute helpers ─────────────────────────────────────────────────────

fn a(name : String, value : String) -> String {
  " " + name + "=\"" + value + "\""
}

fn ai(name : String, value : Int) -> String {
  a(name, int_to_str(value))
}

///| Shorthand for data-ooxml-* attributes.
fn da(name : String, value : String) -> String {
  " data-ooxml-" + name + "=\"" + value + "\""
}

fn dai(name : String, value : Int) -> String {
  da(name, int_to_str(value))
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
  sz * @ooxml.emu_per_hundredth_pt
}

fn default_font_size_emu() -> Int {
  default_font_size_emu_
}

// ── Rotation transform ────────────────────────────────────────────────────────

/// Format rotation in 60000ths of a degree as "D.DD" string (e.g. "45.00").
/// Handles negative rotations correctly.
fn format_rotation_deg(rot_60k : Int) -> String {
  let deg_x100 = rot_60k / @ooxml.rot_to_centideg
  let deg_int = deg_x100 / @ooxml.centideg_per_deg
  let deg_frac_raw = deg_x100 % @ooxml.centideg_per_deg
  let deg_frac = if deg_frac_raw < 0 { -deg_frac_raw } else { deg_frac_raw }
  let frac_str = if deg_frac < 10 { "0" + int_to_str(deg_frac) } else { int_to_str(deg_frac) }
  int_to_str(deg_int) + "." + frac_str
}

/// Build SVG rotate() transform attribute string centered on (cx, cy).
fn rotation_transform(rot_60k : Int, cx : Int, cy : Int) -> String {
  if rot_60k == 0 { return "" }
  " transform=\"rotate(" + format_rotation_deg(rot_60k) + "," +
  int_to_str(cx) + "," + int_to_str(cy) + ")\""
}

fn rotation_attr(t : @ooxml.ShapeTransform, scale : Int) -> String {
  let cx_p = px(t.x + t.cx / 2, scale)
  let cy_p = px(t.y + t.cy / 2, scale)
  rotation_transform(t.rot, cx_p, cy_p)
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

let alpha_count : Int = 26        // number of letters in alphabet
let ascii_lower_a_minus1 : Int = 96  // 'a' - 1
let ascii_upper_a_minus1 : Int = 64  // 'A' - 1
let ascii_lower_a : Int = 97
let ascii_lower_z : Int = 122
let ascii_case_shift : Int = 32   // 'a' - 'A'

fn alpha_lower(n : Int) -> String {
  if n >= 1 && n <= alpha_count {
    @ffi.ffi_char_code_to_str(ascii_lower_a_minus1 + n)
  } else {
    int_to_str(n)
  }
}

fn alpha_upper(n : Int) -> String {
  if n >= 1 && n <= alpha_count {
    @ffi.ffi_char_code_to_str(ascii_upper_a_minus1 + n)
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
  let n_pairs = vals.length()
  let mut i = 0
  while i < n_pairs {
    while rem >= vals[i] {
      result = result + (if lower { syml[i] } else { syms[i] })
      rem = rem - vals[i]
    }
    i = i + 1
  }
  result
}

// ── Text helpers ─────────────────────────────────────────────────────────────

///| Convert ASCII lowercase letters to uppercase.
fn to_upper_ascii(s : String) -> String {
  let len = s.length()
  let mut result = ""
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    let code = Char::to_int(c)
    if code >= ascii_lower_a && code <= ascii_lower_z {
      // a-z → A-Z
      result = result + @ffi.ffi_char_code_to_str(code - ascii_case_shift)
    } else {
      result = result + @ffi.ffi_char_code_to_str(code)
    }
    i = i + 1
  }
  result
}

///| Check if a character is CJK (Chinese/Japanese/Korean).
fn is_cjk(c : Char) -> Bool {
  let code = Char::to_int(c)
  // CJK Unified Ideographs, CJK Compat Ideographs, Fullwidth forms, CJK Ext-B+
  (code >= 0x3000 && code <= 0x9FFF) ||
  (code >= 0xF900 && code <= 0xFAFF) ||
  (code >= 0xFF00 && code <= 0xFFEF) ||
  (code >= 0x20000 && code <= 0x2FA1F)
}

// ── Text wrapping ────────────────────────────────────────────────────────────

///| A visual line fragment: (run_index, text_fragment)
///  Used to split paragraph runs into wrapped lines.

///| Split paragraph runs into visual lines that fit within max_width pixels.
///| Returns Array of lines, each line being Array of (run_idx, text).
fn wrap_paragraph(
  para : @ooxml.TextParagraph,
  max_width : Int,
  scale : Int,
  def_sz : Int,
  font_scale : Int,
  theme : @ooxml.ThemeData
) -> Array[Array[(Int, String)]] {
  let lines : Array[Array[(Int, String)]] = []
  let mut cur_line : Array[(Int, String)] = []
  let mut cur_width : Double = 0.0
  let max_w : Double = max_width.to_double()
  let nr = para.runs.length()
  let mut ri = 0
  while ri < nr {
    let run = para.runs[ri]
    let fs_raw = if run.font_size > 0 { run.font_size } else { if def_sz > 0 { def_sz * scale / @ooxml.emu_per_hundredth_pt } else { 1800 } }
    let fs_pt = apply_font_scale(fs_raw, font_scale).to_double() / 100.0
    let ff = if not(run.ea_font.is_empty()) {
      run.ea_font
    } else if not(run.font_face.is_empty()) {
      run.font_face
    } else {
      theme.minor_font
    }
    let text = if run.cap == "all" { to_upper_ascii(run.text) } else { run.text }
    let text_len = text.length()
    // Split text into segments (words for Latin, chars for CJK)
    let mut seg_start = 0
    while seg_start < text_len {
      let c = text.get_char(seg_start).unwrap()
      let mut seg_end = seg_start
      if is_cjk(c) {
        // Single CJK character as one segment
        seg_end = seg_start + 1
      } else {
        // Collect word (non-CJK chars until space/CJK)
        seg_end = seg_start
        while seg_end < text_len {
          let sc = text.get_char(seg_end).unwrap()
          if is_cjk(sc) { break }
          seg_end = seg_end + 1
          // Break after whitespace
          if sc == ' ' || sc == '\t' { break }
        }
      }
      // Extract segment text
      let seg_text = str_sub(text, seg_start, seg_end)
      let seg_w = @ffi.ffi_measure_text(seg_text, ff, fs_pt)
      // Check if adding this segment would overflow
      if cur_width > 0.0 && cur_width + seg_w > max_w && max_w > 0.0 {
        // Start new line
        if cur_line.length() > 0 {
          lines.push(cur_line)
        }
        cur_line = []
        cur_width = 0.0
      }
      cur_line.push((ri, seg_text))
      cur_width = cur_width + seg_w
      seg_start = seg_end
    }
    ri = ri + 1
  }
  if cur_line.length() > 0 {
    lines.push(cur_line)
  }
  if lines.length() == 0 {
    // Empty paragraph - one empty line
    lines.push([])
  }
  lines
}

///| Extract substring from start (inclusive) to end (exclusive).
fn str_sub(s : String, start : Int, end_ : Int) -> String {
  let mut result = ""
  let mut i = start
  while i < end_ {
    let c = s.get_char(i).unwrap()
    result = result + @ffi.ffi_char_code_to_str(Char::to_int(c))
    i = i + 1
  }
  result
}

// ── Tab helpers ──────────────────────────────────────────────────────────

///| Check if a string contains a tab character.
fn contains_tab(s : String) -> Bool {
  let len = s.length()
  let mut i = 0
  while i < len {
    if s.get_char(i).unwrap() == '\t' { return true }
    i = i + 1
  }
  false
}

///| Split a string by tab characters. Returns array of parts.
fn split_by_tab(s : String) -> Array[String] {
  let parts : Array[String] = []
  let len = s.length()
  let mut start = 0
  let mut i = 0
  while i < len {
    if s.get_char(i).unwrap() == '\t' {
      parts.push(str_sub(s, start, i))
      start = i + 1
    }
    i = i + 1
  }
  parts.push(str_sub(s, start, len))
  parts
}

///| Render a run's text, splitting by tab characters and positioning at tab stops.
///  Returns (svg_output, updated_tab_idx).
fn render_run_with_tabs(
  run : @ooxml.TextRun,
  run_idx : Int,
  display_text : String,
  tab_stops : Array[@ooxml.TabStop],
  tab_idx : Int,
  text_x : Int,
  cur_y : Int,
  scale : Int,
  fs : Int,
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData
) -> (String, Int) {
  if tab_stops.length() > 0 && contains_tab(display_text) {
    let parts = split_by_tab(display_text)
    let nparts = parts.length()
    let mut out = ""
    let mut ti = tab_idx
    let mut parti = 0
    while parti < nparts {
      let part = parts[parti]
      if parti > 0 && ti < tab_stops.length() {
        let tab_x = text_x + px(tab_stops[ti].pos, scale)
        out = out + "</tspan><tspan" + ai("x", tab_x) + ai("y", cur_y) + ">"
        ti = ti + 1
      }
      out = out + render_run_tspan(run, run_idx, part, scale, fs, rels, theme)
      parti = parti + 1
    }
    (out, ti)
  } else {
    (render_run_tspan(run, run_idx, display_text, scale, fs, rels, theme), tab_idx)
  }
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
    base_line_h * pct / @ooxml.pct_100
  } else {
    // Default: use font size * line_height ratio
    let emu = para_max_font_emu(para)
    if emu > 0 { px(emu, scale) * line_height_numer / line_height_denom } else { def_line_h }
  }
}

///| Render a single text run as a <tspan> element.
fn render_run_tspan(
  run : @ooxml.TextRun,
  ri : Int,
  display_text : String,
  scale : Int,
  font_scale : Int,
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData
) -> String {
  let mut span = "<tspan"
  // Visual attributes
  if run.bold   { span = span + " font-weight=\"bold\""  }
  if run.italic { span = span + " font-style=\"italic\"" }
  if run.font_size > 0 {
    span = span + ai("font-size", apply_font_scale(px(hundredth_pt_to_emu(run.font_size), scale), font_scale))
  }
  // font-family: sym_font, cs_font, ea_font, latin_font, sans-serif
  let has_ff = not(run.font_face.is_empty()) || not(run.ea_font.is_empty()) || not(run.cs_font.is_empty()) || not(run.sym_font.is_empty())
  if has_ff {
    let mut ff = ""
    if not(run.sym_font.is_empty()) { ff = run.sym_font }
    if not(run.cs_font.is_empty()) {
      if not(ff.is_empty()) { ff = ff + ", " }
      ff = ff + run.cs_font
    }
    if not(run.ea_font.is_empty()) {
      if not(ff.is_empty()) { ff = ff + ", " }
      ff = ff + run.ea_font
    }
    if not(run.font_face.is_empty()) {
      if not(ff.is_empty()) { ff = ff + ", " }
      ff = ff + run.font_face
    }
    span = span + a("font-family", ff + ", sans-serif")
  }
  // Color: explicit color, or theme hlink/folHlink color for hyperlinks
  if not(run.color.is_none()) {
    span = span + a("fill", run.color.to_css())
  } else if not(run.hlink_rid.is_empty()) && not(theme.hlink.is_none()) {
    span = span + a("fill", theme.hlink.to_css())
  }
  // Character spacing
  if run.char_spacing != 0 {
    let ls_px = run.char_spacing * @ooxml.char_spc_dpi / @ooxml.hundredth_pt_per_inch
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
  // Capitalization
  if run.cap == "small" {
    span = span + a("font-variant", "small-caps")
  }
  // data-ooxml-* attributes for round-trip
  span = span +
    da("run-idx", int_to_str(ri)) +
    da("bold", bool_str(run.bold)) +
    da("font-size", int_to_str(run.font_size)) +
    da("color", run.color.to_hex()) +
    da("run-font", run.font_face) +
    da("ea-font", run.ea_font) +
    da("cs-font", run.cs_font) +
    da("sym-font", run.sym_font) +
    da("underline", run.underline) +
    da("strike", run.strike) +
    dai("baseline", run.baseline) +
    dai("char-spacing", run.char_spacing) +
    dai("kern", run.kern) +
    da("cap", run.cap) +
    da("hlink-rid", run.hlink_rid) +
    da("hlink-hover-rid", run.hlink_mouse_over_rid)
  // Wrap in hyperlink if present
  let inner = svg_escape(display_text)
  if not(run.hlink_rid.is_empty()) {
    let url = @ooxml.find_rel_target(rels, run.hlink_rid)
    if not(url.is_empty()) {
      span + "><a href=\"" + svg_escape(url) + "\">" + inner + "</a></tspan>"
    } else {
      span + ">" + inner + "</tspan>"
    }
  } else {
    span + ">" + inner + "</tspan>"
  }
}

fn render_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
  rels : Array[@ooxml.Relationship]
) -> String {
  if paragraphs.length() == 0 { return "" }
  // Insets: use body_props if specified, otherwise default (91440 / 45720 EMU)
  let l_ins = if body_props.l_ins >= 0 { body_props.l_ins } else { default_l_inset }
  let t_ins = if body_props.t_ins >= 0 { body_props.t_ins } else { default_t_inset }
  let r_ins = if body_props.r_ins >= 0 { body_props.r_ins } else { default_l_inset }

  let pad_l   = px(l_ins, scale)
  let pad_t   = px(t_ins, scale)
  let pad_r   = px(r_ins, scale)
  let base_text_x  = px(t.x, scale) + pad_l
  let full_text_w  = px(t.cx, scale) - pad_l - pad_r
  let mut text_x = base_text_x
  let mut text_w = full_text_w
  let def_sz  = px(default_font_size_emu(), scale) // 24 px for standard slides (unscaled)
  let def_line_h  = def_sz * line_height_numer / line_height_denom // unscaled default line height
  let fs = body_props.font_scale
  let lr = body_props.ln_spc_reduction

  let do_wrap = body_props.wrap != "none"

  // Multi-column support
  let num_cols = if body_props.num_cols > 1 { body_props.num_cols } else { 1 }
  let col_spacing_px = if num_cols > 1 { px(body_props.col_spacing, scale) } else { 0 }
  let col_w = if num_cols > 1 { (full_text_w - col_spacing_px * (num_cols - 1)) / num_cols } else { full_text_w }
  if num_cols > 1 {
    text_w = col_w
  }

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
      // Count visual lines for wrapped text (uses col_w for multi-col)
      let num_visual_lines = if do_wrap {
        let indent_emu = if para.mar_l > 0 { para.mar_l } else { para.level * level_indent_emu }
        let indent_px = px(indent_emu, scale)
        let wrap_w = col_w - indent_px
        let vlines = wrap_paragraph(para, wrap_w, scale, def_sz, fs, theme)
        vlines.length()
      } else {
        1
      }
      total_text_h = total_text_h + para_line_h * num_visual_lines
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

  // Vertical text mode
  let is_vert = body_props.vert == "vert" || body_props.vert == "eaVert" || body_props.vert == "vert270" || body_props.vert == "wordArtVert"

  // Text rotation from bodyPr rot (60000ths of a degree)
  let bp_rot = body_props.rot

  // Image bullets — render before <text> as separate <image> elements, collect them
  let mut img_bullets_svg = ""

  let max_col_h = if num_cols > 1 { total_text_h / num_cols + 1 } else { 0 }

  let font_face = theme.minor_font + ", sans-serif"
  let scaled_def_sz = apply_font_scale(def_sz, fs)

  fn make_text_header(tx : Int) -> String {
    let mut h = "<text" +
      ai("x", tx) +
      a("xml:space", "preserve") +
      a("font-family", font_face) +
      ai("font-size", scaled_def_sz) +
      a("fill", default_text_fill) +
      da("font-face", theme.minor_font)
    if is_vert {
      h = h + a("writing-mode", "tb")
      if body_props.vert == "vert" || body_props.vert == "eaVert" {
        h = h + a("glyph-orientation-vertical", "auto")
      }
    }
    if bp_rot != 0 {
      let txt_cx = px(t.x + t.cx / 2, scale)
      let txt_cy = px(t.y + t.cy / 2, scale)
      h = h + rotation_transform(bp_rot, txt_cx, txt_cy)
    }
    h + ">"
  }

  let mut out = make_text_header(text_x)

  let base_y = px(t.y, scale) + pad_t + anchor_offset_clamped
  let mut cur_y = base_y
  let mut cur_col = 0

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
    // Multi-column: check if we need to move to the next column (paragraph boundary)
    if num_cols > 1 && cur_col < num_cols - 1 && (cur_y - base_y) >= max_col_h {
      out = out + "</text>"
      cur_col = cur_col + 1
      text_x = base_text_x + (col_w + col_spacing_px) * cur_col
      cur_y = base_y

      out = out + make_text_header(text_x)
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
      // RTL: default alignment is right
      // In SVG with direction="rtl", text-anchor semantics are swapped:
      //   "start" = right edge (text flows leftward from x)
      //   "end"   = left edge (text flows rightward from x)
      let effective_align = if para.rtl && para.align.is_empty() { "r" } else { para.align }
      let (anchor_x, anchor_attr) = if effective_align == "ctr" || effective_align == "c" {
        (text_x + indent + (text_w - indent) / 2, " text-anchor=\"middle\"")
      } else if effective_align == "r" || (para.rtl && effective_align != "l") {
        if para.rtl {
          // RTL right-aligned: use text-anchor="start" so text flows leftward from right edge
          (text_x + text_w, " text-anchor=\"start\"")
        } else {
          (text_x + text_w, " text-anchor=\"end\"")
        }
      } else {
        if para.rtl {
          // RTL left-aligned: use text-anchor="end" so text ends at left edge
          (text_x + indent, " text-anchor=\"end\"")
        } else {
          (text_x + indent, " text-anchor=\"start\"")
        }
      }
      // RTL direction attribute
      let rtl_attr = if para.rtl { " direction=\"rtl\"" } else { "" }
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
      // Image bullet: render <image> element before text
      if not(para.bullet_img_rid.is_empty()) {
        let bu_target = @ooxml.find_rel_target(rels, para.bullet_img_rid)
        if not(bu_target.is_empty()) {
          let bu_img_path = resolve_image_path(bu_target)
          let bu_b64 = @ffi.ffi_get_file_base64(bu_img_path)
          if not(bu_b64.is_empty()) {
            let bu_mime = mime_from_ext(bu_img_path)
            let bu_img_sz = p_font_sz // size bullet image to font size
            let bu_img_x = if para.rtl { anchor_x + 4 } else { anchor_x - bu_img_sz - 4 }
            let bu_img_y = cur_y - bu_img_sz + 2
            img_bullets_svg = img_bullets_svg + "<image" +
              ai("x", bu_img_x) + ai("y", bu_img_y) +
              ai("width", bu_img_sz) + ai("height", bu_img_sz) +
              a("href", "data:" + bu_mime + ";base64," + bu_b64) +
              "/>"
          }
        }
      }
      // Paragraph-level tspan with data-ooxml-* attributes
      let para_align = if para.align.is_empty() { "l" } else { para.align }
      out = out + "<tspan" +
        ai("x", anchor_x) +
        ai("y", cur_y) +
        anchor_attr +
        rtl_attr +
        da("para-idx", int_to_str(pi)) +
        da("para-align", para_align) +
        dai("para-lvl", para.level) +
        da("bu-font", para.bullet_font) +
        dai("bu-size", para.bullet_size) +
        da("bu-color", para.bullet_color.to_hex()) +
        da("rtl", bool_str(para.rtl)) +
        ">"
      // Render bullet if present
      if not(bullet_text.is_empty()) {
        let mut bu_span = "<tspan"
        if not(para.bullet_font.is_empty()) {
          bu_span = bu_span + a("font-family", para.bullet_font + ", sans-serif")
        }
        if para.bullet_size > 0 {
          // Percentage of run font size
          let run_fs = if p_font_sz > 0 { p_font_sz } else { scaled_def_sz }
          let bu_fs = run_fs * para.bullet_size / @ooxml.pct_100
          bu_span = bu_span + ai("font-size", bu_fs)
        } else if para.bullet_size < 0 {
          // Absolute size in hundredths-of-a-point (negated)
          let bu_fs_emu = hundredth_pt_to_emu(-(para.bullet_size))
          bu_span = bu_span + ai("font-size", apply_font_scale(px(bu_fs_emu, scale), fs))
        }
        if not(para.bullet_color.is_none()) {
          bu_span = bu_span + a("fill", para.bullet_color.to_css())
        }
        out = out + bu_span + ">" + svg_escape(bullet_text) + " </tspan>"
      }
      // Text wrapping: split runs into visual lines
      if do_wrap {
        let wrap_w = text_w - indent
        let vlines = wrap_paragraph(para, wrap_w, scale, def_sz, fs, theme)
        let nvl = vlines.length()
        let mut vli = 0
        let mut tab_idx = 0
        let mut anchor_x = anchor_x
        while vli < nvl {
          let vline = vlines[vli]
          // For wrapped lines after the first, emit a new positioning tspan
          if vli > 0 {
            cur_y = cur_y + p_line_h
            // Multi-column: check if this line overflows to next column
            if num_cols > 1 && cur_col < num_cols - 1 && (cur_y - base_y) >= max_col_h {
              // Close current paragraph tspan and <text>
              out = out + "</tspan></text>"
              cur_col = cur_col + 1
              text_x = base_text_x + (col_w + col_spacing_px) * cur_col
              cur_y = base_y + p_font_sz
              // Recompute anchor_x for new column
              anchor_x = if effective_align == "ctr" || effective_align == "c" {
                text_x + indent + (text_w - indent) / 2
              } else if effective_align == "r" {
                text_x + text_w
              } else {
                text_x + indent
              }
              // Open new <text> and paragraph tspan
              out = out + make_text_header(text_x)
              out = out + "<tspan" +
                ai("x", anchor_x) +
                ai("y", cur_y) +
                anchor_attr +
                rtl_attr +
                ">"
            } else {
              out = out + "</tspan><tspan" +
                ai("x", anchor_x) +
                ai("y", cur_y) +
                anchor_attr +
                rtl_attr +
                ">"
            }
          }
          // Render fragments in this visual line
          let nf = vline.length()
          let mut fi = 0
          while fi < nf {
            let (run_idx, frag_text) = vline[fi]
            let run = para.runs[run_idx]
            let (tab_svg, new_ti) = render_run_with_tabs(run, run_idx, frag_text, para.tab_stops, tab_idx, text_x, cur_y, scale, fs, rels, theme)
            out = out + tab_svg
            tab_idx = new_ti
            fi = fi + 1
          }
          vli = vli + 1
        }
      } else {
        // No wrapping: render runs directly, handling tab stops
        let mut tab_idx = 0
        let mut ri = 0
        while ri < nr {
          let run = para.runs[ri]
          let display_text = if run.cap == "all" { to_upper_ascii(run.text) } else { run.text }
          let (tab_svg, new_ti) = render_run_with_tabs(run, ri, display_text, para.tab_stops, tab_idx, text_x, cur_y, scale, fs, rels, theme)
          out = out + tab_svg
          tab_idx = new_ti
          ri = ri + 1
        }
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
  img_bullets_svg + out + "</text>"
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
    da("shape-type", shape_type_str(shape.kind)) +
    da("geom", geom_str(shape.kind)) +
    dai("x", t.x) +
    dai("y", t.y) +
    dai("cx", t.cx) +
    dai("cy", t.cy) +
    dai("rot", t.rot) +
    da("fill", shape.fill.to_hex()) +
    da("stroke", shape.stroke.to_hex()) +
    dai("stroke-w", shape.stroke_w) +
    dai("shape-idx", shape_idx) +
    da("ph-type", shape.ph_type) +
    dai("ph-idx", shape.ph_idx) +
    da("anchor", shape.body_props.anchor) +
    dai("l-ins", shape.body_props.l_ins) +
    dai("t-ins", shape.body_props.t_ins) +
    dai("r-ins", shape.body_props.r_ins) +
    dai("b-ins", shape.body_props.b_ins) +
    da("auto-fit", bool_str(shape.body_props.auto_fit)) +
    dai("font-scale", shape.body_props.font_scale) +
    dai("ln-spc-reduction", shape.body_props.ln_spc_reduction) +
    da("wrap", shape.body_props.wrap) +
    dai("bp-rot", shape.body_props.rot) +
    da("bp-vert", shape.body_props.vert) +
    dai("bp-num-cols", shape.body_props.num_cols) +
    dai("bp-col-spacing", shape.body_props.col_spacing)

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
      rect_svg + render_text(shape.paragraphs, text_t, scale, theme, shape.body_props, rels)
    }
    @ooxml.ShapeKind::Picture(rid) => {
      let target = @ooxml.find_rel_target(rels, rid)
      let placeholder = "<rect" +
        ai("x", x_p) + ai("y", y_p) +
        ai("width", cx_p) + ai("height", cy_p) +
        " fill=\"" + image_placeholder_fill + "\"" + rot + "/>"
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
      render_table(table, t, scale, theme, rels)
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
  theme : @ooxml.ThemeData,
  rels : Array[@ooxml.Relationship]
) -> String {
  if paragraphs.length() == 0 { return "" }
  let tx = cell_x + table_cell_pad
  let def_sz = px(default_font_size_emu(), scale)
  let font_sz = if def_sz > cell_h - table_cell_pad * 2 { cell_h - table_cell_pad * 2 } else { def_sz }
  let line_h = font_sz * line_height_numer / line_height_denom
  let font_face = theme.minor_font + ", sans-serif"
  let mut out = "<text" +
    ai("x", tx) +
    a("xml:space", "preserve") +
    a("font-family", font_face) +
    ai("font-size", font_sz) +
    a("fill", default_text_fill) +
    ">"
  let mut cur_y = cell_y + table_cell_pad + font_sz
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    if nr == 0 {
      cur_y = cur_y + line_h
    } else {
      // Paragraph-level tspan for positioning
      out = out + "<tspan" + ai("x", tx) + ai("y", cur_y) + ">"
      let mut ri = 0
      while ri < nr {
        let run = para.runs[ri]
        let display_text = if run.cap == "all" { to_upper_ascii(run.text) } else { run.text }
        out = out + render_run_tspan(run, ri, display_text, scale, -1, rels, theme)
        ri = ri + 1
      }
      out = out + "</tspan>"
      cur_y = cur_y + line_h
    }
    pi = pi + 1
    // Stop if we overflow the cell
    if cur_y > cell_y + cell_h - table_cell_pad { break }
  }
  out + "</text>"
}

fn render_table(
  table : @ooxml.TableData,
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  rels : Array[@ooxml.Relationship]
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
        " stroke=\"" + table_border_color + "\" stroke-width=\"1\"/>"
      // Cell text
      out = out + render_table_cell_text(cell.paragraphs, cx_px, ry, cw, rh, scale, theme, rels)
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
    dai("slide-cx", sw) +
    dai("slide-cy", sh) +
    da("bg", slide.background.to_hex()) +
    dai("scale", scale) +
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
