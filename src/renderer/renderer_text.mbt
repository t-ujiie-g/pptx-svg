// SVG renderer — Text rendering
//
// Handles paragraph layout, bullet rendering, text wrapping (Latin word-break
// and CJK character-break), line spacing, tab stops, super/subscript,
// capitalization, hyperlinks, and text-level effects (shadow/glow via CSS filter).
//
// Same package as renderer.mbt — all constants and helpers are shared.

// ── Text helpers ─────────────────────────────────────────────────────────────

///|
/// Build a comma-separated font-family string from sym/cs/ea/latin fonts.
/// Returns empty string if no fonts are specified.
fn build_font_family(
  sym : String,
  cs : String,
  ea : String,
  latin : String,
) -> String {
  let mut ff = ""
  if not(sym.is_empty()) { ff = sym }
  if not(cs.is_empty()) { ff = if ff.is_empty() { cs } else { ff + ", " + cs } }
  if not(ea.is_empty()) {
    ff = if ff.is_empty() { ea } else { ff + ", " + ea }
    // Append platform fallbacks for EA font
    let fb = @ffi.ffi_get_font_fallback(ea)
    if not(fb.is_empty()) {
      ff = ff + ", " + fb
    }
  }
  if not(latin.is_empty()) {
    ff = if ff.is_empty() { latin } else { ff + ", " + latin }
    // Append platform fallbacks for Latin font (if different from EA)
    if latin != ea {
      let fb = @ffi.ffi_get_font_fallback(latin)
      if not(fb.is_empty()) {
        ff = ff + ", " + fb
      }
    }
  }
  ff
}

///|
/// Convert ASCII lowercase letters to uppercase.
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
    i = i + @xml.char_step(c)
  }
  result
}

///|
/// Check if a character is CJK (Chinese/Japanese/Korean).
fn is_cjk(c : Char) -> Bool {
  let code = Char::to_int(c)
  // CJK Unified Ideographs, CJK Compat Ideographs, Fullwidth forms, CJK Ext-B+
  (code >= 0x3000 && code <= 0x9FFF) ||
  (code >= 0xF900 && code <= 0xFAFF) ||
  (code >= 0xFF00 && code <= 0xFFEF) ||
  (code >= 0x20000 && code <= 0x2FA1F)
}

///|
/// Check if a character is a kinsoku line-start-prohibited character
/// (行頭禁止文字). These must not appear at the beginning of a line.
fn is_kinsoku_start_prohibited(c : Char) -> Bool {
  // Japanese punctuation, closing brackets, etc.
  c == '。' || c == '、' || c == '，' || c == '．' || c == '.' || c == ',' ||
  c == ';' || c == ':' || c == '!' || c == '?' || c == '！' || c == '？' ||
  c == '）' || c == '」' || c == '』' || c == '】' || c == '》' || c == '〉' ||
  c == '｝' || c == ')' || c == '>' || c == ']' || c == '}' ||
  c == 'ー' || c == '～' || c == '…' || c == '‥' ||
  c == '・' || c == '：' || c == '；' ||
  c == '゛' || c == '゜' || c == 'ゝ' || c == 'ゞ' || c == 'ヽ' || c == 'ヾ' ||
  c == '々' || c == '〃' ||
  // Small kana
  c == 'ぁ' || c == 'ぃ' || c == 'ぅ' || c == 'ぇ' || c == 'ぉ' ||
  c == 'っ' || c == 'ゃ' || c == 'ゅ' || c == 'ょ' || c == 'ゎ' ||
  c == 'ァ' || c == 'ィ' || c == 'ゥ' || c == 'ェ' || c == 'ォ' ||
  c == 'ッ' || c == 'ャ' || c == 'ュ' || c == 'ョ' || c == 'ヮ'
}

// ── Text wrapping ────────────────────────────────────────────────────────────

///| A visual line fragment: (run_index, text_fragment)
///  Used to split paragraph runs into wrapped lines.

///| Split paragraph runs into visual lines that fit within max_width pixels.

///|
/// Returns Array of lines, each line being Array of (run_idx, text).
fn wrap_paragraph(
  para : @ooxml.TextParagraph,
  max_width : Int,
  scale : Int,
  def_sz : Int,
  font_scale : Int,
  theme : @ooxml.ThemeData,
) -> Array[Array[(Int, String)]] {
  let lines : Array[Array[(Int, String)]] = []
  let mut cur_line : Array[(Int, String)] = []
  let mut cur_width : Double = 0.0
  let max_w : Double = max_width.to_double()
  // Track accumulated same-font text on the current line for accurate measurement.
  // Measuring CJK characters individually over-counts width due to per-glyph
  // side-bearings; measuring the concatenated string gives the true rendered width.
  let mut run_accum : String = "" // accumulated text for current font
  let mut run_ff : String = "" // font face for accumulated text
  let mut run_fs : Double = 0.0 // font size for accumulated text
  let nr = para.runs.length()
  let mut ri = 0
  while ri < nr {
    let run = para.runs[ri]
    let fs_raw = if run.font_size > 0 {
      run.font_size
    } else if def_sz > 0 {
      def_sz * scale / @ooxml.emu_per_hundredth_pt
    } else {
      1800
    }
    let fs_px = px(
      apply_font_scale(fs_raw, font_scale) * @ooxml.emu_per_hundredth_pt,
      scale,
    ).to_double()
    let ff = if not(run.ea_font.is_empty()) {
      run.ea_font
    } else if not(run.font_face.is_empty()) {
      run.font_face
    } else {
      theme.minor_font
    }
    let text = if run.cap == "all" {
      to_upper_ascii(run.text)
    } else {
      run.text
    }
    // Handle explicit line break runs (from <a:br>)
    if text == "\n" {
      if cur_line.length() > 0 {
        lines.push(cur_line)
      } else {
        lines.push([])
      }
      cur_line = []
      cur_width = 0.0
      run_accum = ""
      ri = ri + 1
      continue
    }
    // When font changes, lock in the accumulated width and reset
    if ff != run_ff || fs_px != run_fs {
      if not(run_accum.is_empty()) {
        cur_width = cur_width + @ffi.ffi_measure_text(run_accum, run_ff, run_fs)
        run_accum = ""
      }
      run_ff = ff
      run_fs = fs_px
    }
    let text_len = text.length()
    let mut seg_start = 0
    while seg_start < text_len {
      let c = text.get_char(seg_start).unwrap()
      let mut seg_end = seg_start
      if is_cjk(c) || @xml.is_astral(c) {
        // CJK or astral-plane character (emoji): treat as single segment
        seg_end = seg_start + @xml.char_step(c)
      } else {
        seg_end = seg_start
        while seg_end < text_len {
          let sc = text.get_char(seg_end).unwrap()
          if is_cjk(sc) || @xml.is_astral(sc) {
            break
          }
          seg_end = seg_end + @xml.char_step(sc)
          if sc == ' ' || sc == '\t' {
            break
          }
        }
      }
      let seg_text = @xml.str_substring(text, seg_start, seg_end)
      // Measure candidate: accumulated same-font text + new segment
      let candidate = run_accum + seg_text
      let candidate_w = @ffi.ffi_measure_text(candidate, ff, fs_px)
      let total_w = cur_width + candidate_w
      if cur_line.length() > 0 && total_w > max_w && max_w > 0.0 {
        // Kinsoku: line-start-prohibited chars stay on current line
        let first_char = seg_text.get_char(0).unwrap()
        if is_kinsoku_start_prohibited(first_char) {
          cur_line.push((ri, seg_text))
          run_accum = candidate
          seg_start = seg_end
          continue
        }
        // Start new line — lock in accumulated width first
        if cur_line.length() > 0 {
          lines.push(cur_line)
        }
        cur_line = []
        cur_width = 0.0
        run_accum = ""
      }
      cur_line.push((ri, seg_text))
      run_accum = run_accum + seg_text
      seg_start = seg_end
    }
    ri = ri + 1
  }
  // Lock in any remaining accumulated text
  if cur_line.length() > 0 {
    lines.push(cur_line)
  }
  if lines.length() == 0 {
    lines.push([])
  }
  lines
}

// ── Tab helpers ──────────────────────────────────────────────────────────

///|
/// Check if a string contains a tab character.
fn contains_tab(s : String) -> Bool {
  let len = s.length()
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '\t' {
      return true
    }
    i = i + @xml.char_step(c)
  }
  false
}

///|
/// Split a string by tab characters. Returns array of parts.
fn split_by_tab(s : String) -> Array[String] {
  let parts : Array[String] = []
  let len = s.length()
  let mut start = 0
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    let step = @xml.char_step(c)
    if c == '\t' {
      parts.push(@xml.str_substring(s, start, i))
      start = i + step
    }
    i = i + step
  }
  parts.push(@xml.str_substring(s, start, len))
  parts
}

///|
/// Render a run's text, splitting by tab characters and positioning at tab stops.
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
  theme : @ooxml.ThemeData,
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
    (
      render_run_tspan(run, run_idx, display_text, scale, fs, rels, theme),
      tab_idx,
    )
  }
}

// ── Text rendering ────────────────────────────────────────────────────────────

///|
fn para_max_font_emu(para : @ooxml.TextParagraph) -> Int {
  let nr = para.runs.length()
  let mut max_fs = 0
  let mut i = 0
  while i < nr {
    let fs = para.runs[i].font_size
    if fs > max_fs {
      max_fs = fs
    }
    i = i + 1
  }
  if max_fs > 0 {
    hundredth_pt_to_emu(max_fs)
  } else {
    0
  }
}

///|
fn para_font_size(
  para : @ooxml.TextParagraph,
  scale : Int,
  def_sz : Int,
) -> Int {
  let emu = para_max_font_emu(para)
  if emu > 0 {
    px(emu, scale)
  } else {
    def_sz
  }
}

///| Apply normAutofit font_scale to a pixel value.

///|
/// font_scale: -1=inactive, 0-100000=percentage (100000=100%).
fn apply_font_scale(val : Int, font_scale : Int) -> Int {
  if font_scale >= 0 && font_scale < @ooxml.pct_100 {
    val * font_scale / @ooxml.pct_100
  } else {
    val
  }
}

///| Apply normAutofit lnSpcReduction to a line height.

///|
/// ln_spc_reduction: -1=inactive, 0-100000=reduction percentage.
fn apply_ln_spc_reduction(val : Int, ln_spc_reduction : Int) -> Int {
  if ln_spc_reduction > 0 {
    val * (@ooxml.pct_100 - ln_spc_reduction) / @ooxml.pct_100
  } else {
    val
  }
}

///|
/// Apply both normAutofit adjustments: font_scale then lnSpcReduction.
fn apply_autofit_scale(
  val : Int,
  font_scale : Int,
  ln_spc_reduction : Int,
) -> Int {
  apply_ln_spc_reduction(apply_font_scale(val, font_scale), ln_spc_reduction)
}

///| Compute extra spacing needed above/below a paragraph due to super/subscript.

///|
/// Returns (extra_above, extra_below) in pixels.
fn para_baseline_shift_extra(
  para : @ooxml.TextParagraph,
  scale : Int,
  def_sz : Int,
) -> (Int, Int) {
  let nr = para.runs.length()
  let mut has_super = false
  let mut has_sub = false
  let mut i = 0
  while i < nr {
    if para.runs[i].baseline > 0 {
      has_super = true
    }
    if para.runs[i].baseline < 0 {
      has_sub = true
    }
    i = i + 1
  }
  if not(has_super) && not(has_sub) {
    return (0, 0)
  }
  let font_emu = para_max_font_emu(para)
  let font_px = if font_emu > 0 { px(font_emu, scale) } else { def_sz }
  // SVG baseline-shift:super moves up ~40% of font size, sub moves down ~20%
  // The shifted text is at 70% size, so it extends beyond the normal line bounds
  let extra = font_px * baseline_shift_extra_numer / baseline_shift_extra_denom
  let above = if has_super { extra } else { 0 }
  let below = if has_sub { extra } else { 0 }
  (above, below)
}

///|
fn para_line_height(
  para : @ooxml.TextParagraph,
  scale : Int,
  def_line_h : Int,
) -> Int {
  if para.line_spacing > 0 {
    // Positive = EMU absolute value
    px(para.line_spacing, scale)
  } else if para.line_spacing < 0 {
    // Negative = percentage (-100000 = 100% = single spacing)
    // In OOXML, percentage is relative to "normal" line spacing (ascent + descent + gap).
    // Use pct_line_height_numer/denom (1.2) as the base "single spacing" factor.
    // OOXML single spacing ≈ font (ascent+descent)/em ≈ 1.2× for common CJK fonts.
    let pct = -para.line_spacing // e.g. 130000 for 130%
    let emu = para_max_font_emu(para)
    let base_line_h = if emu > 0 {
      px(emu, scale) * pct_line_height_numer / pct_line_height_denom
    } else {
      def_line_h
    }
    base_line_h * pct / @ooxml.pct_100
  } else {
    // Default: use font size * line_height ratio
    let emu = para_max_font_emu(para)
    if emu > 0 {
      px(emu, scale) * line_height_numer / line_height_denom
    } else {
      def_line_h
    }
  }
}

///|
/// Build CSS rgba() color string from Color + alpha fallback.
fn css_rgba(clr : @ooxml.Color, alpha_fallback : String) -> String {
  "rgba(" + int_to_str(clr.r) + "," + int_to_str(clr.g) + "," +
  int_to_str(clr.b) + "," + eff_alpha_str(clr, alpha_fallback) + ")"
}

///|
/// Generate CSS style attribute with filter for text-level effects (shadow/glow).
/// Returns empty string if no visual effects, or ` style="filter: ..."` attribute.
fn render_text_effect_style(eff : @ooxml.EffectList, scale : Int) -> String {
  if eff.is_none() {
    return ""
  }
  let mut filters = ""
  // Outer shadow → CSS drop-shadow(dx dy blur color)
  if not(eff.outer_shadow.is_none()) {
    let os = eff.outer_shadow
    let blur = emu_to_filter_px(os.blur_rad, scale)
    let dist = emu_to_filter_px(os.dist, scale)
    let angle_deg = os.dir / ooxml_angle_unit
    let dx = dist * cos_deg(angle_deg) / 100
    let dy = dist * sin_deg(angle_deg) / 100
    let std_dev = if blur / 2 > 0 { blur / 2 } else { 1 }
    filters = filters + "drop-shadow(" +
      int_to_str(dx) + "px " + int_to_str(dy) + "px " +
      int_to_str(std_dev) + "px " + css_rgba(os.color, "0.5") + ")"
  }
  // Glow → approximation as drop-shadow with no offset
  if not(eff.glow.is_none()) {
    let rad = emu_to_filter_px(eff.glow.rad, scale)
    let std_dev = if rad / 2 > 0 { rad / 2 } else { 1 }
    if not(filters.is_empty()) {
      filters = filters + " "
    }
    filters = filters + "drop-shadow(0px 0px " +
      int_to_str(std_dev) + "px " + css_rgba(eff.glow.color, "0.7") + ")"
  }
  if filters.is_empty() {
    ""
  } else {
    " style=\"filter: " + filters + "\""
  }
}

///|
/// Render a single text run as a <tspan> element.
fn render_run_tspan(
  run : @ooxml.TextRun,
  ri : Int,
  display_text : String,
  scale : Int,
  font_scale : Int,
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData,
) -> String {
  let mut span = "<tspan"
  // Visual attributes
  if run.bold {
    span = span + " font-weight=\"bold\""
  }
  if run.italic {
    span = span + " font-style=\"italic\""
  }
  if run.font_size > 0 {
    span = span +
      ai(
        "font-size",
        apply_font_scale(
          px(hundredth_pt_to_emu(run.font_size), scale),
          font_scale,
        ),
      )
  }
  // font-family: sym_font, cs_font, ea_font, latin_font, sans-serif
  let ff = build_font_family(run.sym_font, run.cs_font, run.ea_font, run.font_face)
  if not(ff.is_empty()) {
    span = span + a("font-family", ff + ", sans-serif")
  }
  // Color: explicit color, or theme hlink/folHlink color for hyperlinks
  // For text gradient fill, use the first stop color as fallback display
  if not(run.text_grad_fill.is_none()) && run.text_grad_fill.stops.length() > 0 {
    span = span + a("fill", run.text_grad_fill.stops[0].color.to_css())
  } else if not(run.color.is_none()) {
    span = span + a("fill", run.color.to_css())
  } else if not(run.hlink_rid.is_empty()) && not(theme.hlink.is_none()) {
    span = span + a("fill", theme.hlink.to_css())
  }
  // Text outline: render as SVG stroke on text
  if not(run.outline_color.is_none()) && run.outline_w > 0 {
    span = span + a("stroke", run.outline_color.to_css())
    let ow_px = px(run.outline_w, scale)
    if ow_px > 0 {
      span = span + ai("stroke-width", ow_px)
    } else {
      span = span + a("stroke-width", "0.5")
    }
    span = span + a("paint-order", "stroke")
  }
  // Character spacing
  if run.char_spacing != 0 {
    let ls_px = run.char_spacing *
      @ooxml.char_spc_dpi /
      @ooxml.hundredth_pt_per_inch
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
    if deco.is_empty() {
      deco = "line-through"
    } else {
      deco = deco + " line-through"
    }
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
    da("hlink-hover-rid", run.hlink_mouse_over_rid) +
    // Text outline round-trip
    da("outline-color", if run.outline_color.is_none() { "" } else { run.outline_color.to_hex() }) +
    dai("outline-w", run.outline_w) +
    // Text gradient fill round-trip
    (if not(run.text_grad_fill.is_none()) {
      da("tgrad-stops", grad_stops_attr(run.text_grad_fill)) +
      dai("tgrad-angle", run.text_grad_fill.angle) +
      da("tgrad-path", run.text_grad_fill.path_type) +
      da("tgrad-rot", if run.text_grad_fill.rot_with_shape { "1" } else { "0" })
    } else { "" }) +
    // Text pattern fill round-trip
    (if not(run.text_patt_fill.is_none()) {
      da("tpatt-prst", run.text_patt_fill.prst) +
      da("tpatt-fg", run.text_patt_fill.fg_color.to_hex()) +
      da("tpatt-bg", run.text_patt_fill.bg_color.to_hex())
    } else { "" }) +
    // Text-level effect data attributes (round-trip only)
    (if not(run.effects.is_none()) {
      let eff = run.effects
      let mut ea = ""
      if not(eff.outer_shadow.is_none()) {
        let os = eff.outer_shadow
        ea = ea + dai("reff-os-blur", os.blur_rad) +
          dai("reff-os-dist", os.dist) +
          dai("reff-os-dir", os.dir) +
          da("reff-os-clr", os.color.to_hex())
      }
      if not(eff.glow.is_none()) {
        ea = ea + dai("reff-gl-rad", eff.glow.rad) +
          da("reff-gl-clr", eff.glow.color.to_hex())
      }
      ea
    } else { "" }) +
    // OMML math equation raw XML for round-trip
    (if run.math_xml.length() > 0 {
      da("math-xml", @xml.xml_escape(run.math_xml))
    } else { "" }) +
    // CSS filter for text-level effects (shadow/glow visual rendering)
    render_text_effect_style(run.effects, scale)
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

///|
fn render_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
  rels : Array[@ooxml.Relationship],
) -> String {
  if paragraphs.length() == 0 {
    return ""
  }
  // Insets: use body_props if specified, otherwise default (91440 / 45720 EMU)
  let l_ins = if body_props.l_ins >= 0 {
    body_props.l_ins
  } else {
    default_l_inset
  }
  let t_ins = if body_props.t_ins >= 0 {
    body_props.t_ins
  } else {
    default_t_inset
  }
  let r_ins = if body_props.r_ins >= 0 {
    body_props.r_ins
  } else {
    default_l_inset
  }

  let pad_l = px(l_ins, scale)
  let pad_t = px(t_ins, scale)
  let pad_r = px(r_ins, scale)
  let base_text_x = px(t.x, scale) + pad_l
  let full_text_w = px(t.cx, scale) - pad_l - pad_r
  let mut text_x = base_text_x
  let mut text_w = full_text_w
  let def_sz = px(default_font_size_emu(), scale) // 24 px for standard slides (unscaled)
  let def_line_h = def_sz * line_height_numer / line_height_denom // unscaled default line height
  let mut fs = body_props.font_scale
  let mut lr = body_props.ln_spc_reduction

  let do_wrap = body_props.wrap != "none"

  // Multi-column support
  let num_cols = if body_props.num_cols > 1 { body_props.num_cols } else { 1 }
  let col_spacing_px = if num_cols > 1 {
    px(body_props.col_spacing, scale)
  } else {
    0
  }
  let col_w = if num_cols > 1 {
    (full_text_w - col_spacing_px * (num_cols - 1)) / num_cols
  } else {
    full_text_w
  }
  if num_cols > 1 {
    text_w = col_w
  }

  // Compute available height (needed for normAutofit and anchor offset)
  let shape_h = px(t.cy, scale)
  let b_ins = if body_props.b_ins >= 0 {
    body_props.b_ins
  } else {
    default_t_inset
  }
  let pad_b = px(b_ins, scale)
  let avail_h = shape_h - pad_t - pad_b

  // normAutofit: dynamically adjust font_scale / ln_spc_reduction to fit text.
  // PowerPoint stores hints in the PPTX, but our font metrics may differ, so
  // we recompute if needed.  font_scale >= 0 means <a:normAutofit> was present.
  let is_norm_autofit = body_props.font_scale >= 0

  // Calculate total text height for vertical alignment
  let np = paragraphs.length()
  let mut total_text_h = 0
  // Track last paragraph's trailing space for vertical centering correction
  let mut last_para_line_h = 0
  let mut last_para_font_sz = 0
  let mut last_para_spc_after = 0
  let mut autofit_pass = 0
  while autofit_pass < 6 {
    total_text_h = 0
    last_para_line_h = 0
    last_para_font_sz = 0
    last_para_spc_after = 0
    let mut ti = 0
    while ti < np {
      let para = paragraphs[ti]
      let nr = para.runs.length()
      let para_line_h = apply_autofit_scale(
        para_line_height(para, scale, def_line_h),
        fs,
        lr,
      )
      let p_fsz = apply_font_scale(para_font_size(para, scale, def_sz), fs)
      if nr == 0 {
        total_text_h = total_text_h + para_line_h
        last_para_line_h = para_line_h
        last_para_font_sz = p_fsz
        last_para_spc_after = 0
      } else {
        // Add spacing before
        if para.spc_before > 0 && ti > 0 {
          total_text_h = total_text_h + px(para.spc_before, scale)
        }
        // Extra space for super/subscript (must match rendering loop)
        let (bl_extra_above, bl_extra_below) = para_baseline_shift_extra(
          para, scale, def_sz,
        )
        total_text_h = total_text_h + bl_extra_above
        // Count visual lines for wrapped text (uses col_w for multi-col)
        let num_visual_lines = if do_wrap {
          let indent_emu = if para.mar_l >= 0 {
            para.mar_l
          } else {
            para.level * level_indent_emu
          }
          let indent_px = px(indent_emu, scale)
          let wrap_w = col_w - indent_px
          let vlines = wrap_paragraph(para, wrap_w, scale, def_sz, fs, theme)
          vlines.length()
        } else {
          // Count explicit line breaks (from <a:br>) to get visual line count
          let mut br_count = 0
          let mut bri = 0
          while bri < nr {
            if para.runs[bri].text == "\n" {
              br_count = br_count + 1
            }
            bri = bri + 1
          }
          1 + br_count
        }
        total_text_h = total_text_h + para_line_h * num_visual_lines
        total_text_h = total_text_h + bl_extra_below
        last_para_line_h = para_line_h
        last_para_font_sz = p_fsz
        // Add spacing after
        last_para_spc_after = if para.spc_after > 0 {
          let sa = px(para.spc_after, scale)
          total_text_h = total_text_h + sa
          sa
        } else {
          0
        }
      }
      ti = ti + 1
    }
    // normAutofit: if text overflows, adjust parameters and retry
    if is_norm_autofit && total_text_h > avail_h && avail_h > 0 {
      // First try increasing line spacing reduction (up to 20%)
      if lr < 20000 {
        lr = lr + 5000
        if lr > 20000 { lr = 20000 }
        autofit_pass = autofit_pass + 1
        continue
      }
      // Then reduce font scale proportionally
      let needed_fs = fs * avail_h / total_text_h
      if needed_fs < fs && needed_fs > 25000 {
        fs = needed_fs
        autofit_pass = autofit_pass + 1
        continue
      }
    }
    break
  }
  // For vertical centering/bottom alignment, the trailing space of the last line
  // (line_height - font_size) and spc_after should not count toward the text extent.
  // The rendering places text at the baseline (font_size from line top), so the visual
  // bottom of text is at the last baseline, not at the last baseline + trailing space.
  let trailing_space = if last_para_line_h > last_para_font_sz {
    last_para_line_h - last_para_font_sz + last_para_spc_after
  } else {
    last_para_spc_after
  }
  let text_h_for_anchor = total_text_h - trailing_space

  // Vertical anchor offset
  let anchor_offset = if body_props.anchor == "ctr" {
    (avail_h - text_h_for_anchor) / 2
  } else if body_props.anchor == "b" {
    avail_h - text_h_for_anchor
  } else {
    0
  }
  // For anchor="ctr", allow negative offset so text centers even when it
  // overflows the padding area (PowerPoint behavior: center in full shape).
  // But clamp to -pad_t so text never starts above the shape's top edge.
  // For "t"/"b", clamp to 0 so text stays within the padding boundary.
  let anchor_offset_clamped = if body_props.anchor == "ctr" {
    if anchor_offset >= 0 - pad_t { anchor_offset } else { 0 - pad_t }
  } else if anchor_offset > 0 {
    anchor_offset
  } else {
    0
  }

  // Vertical text mode
  let is_vert = body_props.vert == "vert" ||
    body_props.vert == "eaVert" ||
    body_props.vert == "vert270" ||
    body_props.vert == "wordArtVert"

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

  // Auto-number counters per level (reset when a lower level is encountered)
  let auto_num_counters : Array[Int] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
  let mut prev_para_level = -1

  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    let p_line_h = apply_autofit_scale(
      para_line_height(para, scale, def_line_h),
      fs,
      lr,
    )
    let p_font_sz = apply_font_scale(para_font_size(para, scale, def_sz), fs)
    // Apply spacing before paragraph
    if para.spc_before > 0 && pi > 0 {
      cur_y = cur_y + px(para.spc_before, scale)
    }
    // Multi-column: check if we need to move to the next column (paragraph boundary)
    if num_cols > 1 && cur_col < num_cols - 1 && cur_y - base_y >= max_col_h {
      out = out + "</text>"
      cur_col = cur_col + 1
      text_x = base_text_x + (col_w + col_spacing_px) * cur_col
      cur_y = base_y

      out = out + make_text_header(text_x)
    }
    // Extra space for super/subscript to avoid overlap with adjacent lines
    let (bl_extra_above, bl_extra_below) = para_baseline_shift_extra(
      para, scale, def_sz,
    )
    cur_y = cur_y + bl_extra_above
    // Advance to baseline (ascent ≈ font_size)
    cur_y = cur_y + p_font_sz
    if nr == 0 {
      cur_y = cur_y + (p_line_h - p_font_sz)
      pi = pi + 1
    } else {
      // Indent: marL = left margin for continuation lines
      //         para.indent = first-line offset from marL (negative = hanging indent)
      let mar_l_emu = if para.mar_l >= 0 {
        para.mar_l
      } else {
        para.level * level_indent_emu
      }
      let indent = px(mar_l_emu, scale)
      // First-line indent: marL + indent (e.g., marL=171450, indent=-171450 → 0)
      let first_line_indent_emu = mar_l_emu + para.indent
      let first_line_indent = if first_line_indent_emu > 0 {
        px(first_line_indent_emu, scale)
      } else {
        0
      }
      // RTL: default alignment is right
      // In SVG with direction="rtl", text-anchor semantics are swapped:
      //   "start" = right edge (text flows leftward from x)
      //   "end"   = left edge (text flows rightward from x)
      let effective_align = if para.rtl && para.align.is_empty() {
        "r"
      } else {
        para.align
      }
      // anchor_x: first line uses first_line_indent, continuation uses indent (marL)
      let (anchor_x, cont_anchor_x_init, anchor_attr) = if effective_align == "ctr" ||
        effective_align == "c" {
        (
          text_x + first_line_indent + (text_w - first_line_indent) / 2,
          text_x + indent + (text_w - indent) / 2,
          " text-anchor=\"middle\"",
        )
      } else if effective_align == "r" || (para.rtl && effective_align != "l") {
        if para.rtl {
          (text_x + text_w, text_x + text_w, " text-anchor=\"start\"")
        } else {
          (text_x + text_w, text_x + text_w, " text-anchor=\"end\"")
        }
      } else if para.rtl {
        (text_x + first_line_indent, text_x + indent, " text-anchor=\"end\"")
      } else {
        (text_x + first_line_indent, text_x + indent, " text-anchor=\"start\"")
      }
      let mut cont_anchor_x = cont_anchor_x_init
      // RTL direction attribute
      let rtl_attr = if para.rtl { " direction=\"rtl\"" } else { "" }
      // Reset auto-number counters for levels deeper than current
      let cur_level = para.level
      if cur_level < prev_para_level {
        let mut rl = cur_level + 1
        while rl < 10 {
          auto_num_counters[rl] = 0
          rl = rl + 1
        }
      }
      prev_para_level = cur_level
      // Determine bullet text
      let bullet_text = if para.bullet_none {
        ""
      } else if not(para.bullet.is_empty()) {
        para.bullet
      } else if not(para.bullet_auto.is_empty()) {
        let auto_idx = auto_num_counters[cur_level]
        auto_num_counters[cur_level] = auto_idx + 1
        auto_number_text(para.bullet_auto, auto_idx)
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
            let bu_img_x = if para.rtl {
              anchor_x + 4
            } else {
              anchor_x - bu_img_sz - 4
            }
            let bu_img_y = cur_y - bu_img_sz + 2
            img_bullets_svg = img_bullets_svg +
              "<image" +
              ai("x", bu_img_x) +
              ai("y", bu_img_y) +
              ai("width", bu_img_sz) +
              ai("height", bu_img_sz) +
              a("href", "data:" + bu_mime + ";base64," + bu_b64) +
              "/>"
          }
        }
      }
      // Paragraph-level tspan with data-ooxml-* attributes
      let para_align = if para.align.is_empty() { "l" } else { para.align }
      out = out +
        "<tspan" +
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
          bu_span = bu_span +
            a("font-family", para.bullet_font + ", sans-serif")
        }
        if para.bullet_size > 0 {
          // Percentage of run font size
          let run_fs = if p_font_sz > 0 { p_font_sz } else { scaled_def_sz }
          let bu_fs = run_fs * para.bullet_size / @ooxml.pct_100
          bu_span = bu_span + ai("font-size", bu_fs)
        } else if para.bullet_size < 0 {
          // Absolute size in hundredths-of-a-point (negated)
          let bu_fs_emu = hundredth_pt_to_emu(-para.bullet_size)
          bu_span = bu_span +
            ai("font-size", apply_font_scale(px(bu_fs_emu, scale), fs))
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
        // Use first-line anchor for first line, continuation anchor for rest
        let mut cur_anchor_x = anchor_x
        while vli < nvl {
          let vline = vlines[vli]
          // For wrapped lines after the first, emit a new positioning tspan
          if vli > 0 {
            cur_y = cur_y + p_line_h
            // Multi-column: check if this line overflows to next column
            if num_cols > 1 &&
              cur_col < num_cols - 1 &&
              cur_y - base_y >= max_col_h {
              // Close current paragraph tspan and <text>
              out = out + "</tspan></text>"
              cur_col = cur_col + 1
              text_x = base_text_x + (col_w + col_spacing_px) * cur_col
              cur_y = base_y + p_font_sz
              // Recompute anchor for new column (use continuation indent)
              cur_anchor_x = if effective_align == "ctr" || effective_align == "c" {
                text_x + indent + (text_w - indent) / 2
              } else if effective_align == "r" {
                text_x + text_w
              } else {
                text_x + indent
              }
              // Update cont_anchor_x for subsequent lines in new column
              cont_anchor_x = cur_anchor_x
              // Open new <text> and paragraph tspan
              out = out + make_text_header(text_x)
              out = out +
                "<tspan" +
                ai("x", cur_anchor_x) +
                ai("y", cur_y) +
                anchor_attr +
                rtl_attr +
                ">"
            } else {
              // Switch to continuation indent (marL) for wrapped lines
              cur_anchor_x = cont_anchor_x
              out = out +
                "</tspan><tspan" +
                ai("x", cur_anchor_x) +
                ai("y", cur_y) +
                anchor_attr +
                rtl_attr +
                ">"
            }
          }
          // On first line with bullet, reposition to cont_anchor_x (= marL)
          // so text aligns with continuation lines (OOXML text starts at marL)
          if vli == 0 && not(bullet_text.is_empty()) {
            out = out +
              "</tspan><tspan" +
              ai("x", cont_anchor_x) +
              ai("y", cur_y) +
              anchor_attr +
              rtl_attr +
              ">"
          }
          // Render fragments in this visual line
          let nf = vline.length()
          let mut fi = 0
          while fi < nf {
            let (run_idx, frag_text) = vline[fi]
            let run = para.runs[run_idx]
            let (tab_svg, new_ti) = render_run_with_tabs(
              run,
              run_idx,
              frag_text,
              para.tab_stops,
              tab_idx,
              text_x,
              cur_y,
              scale,
              fs,
              rels,
              theme,
            )
            out = out + tab_svg
            tab_idx = new_ti
            fi = fi + 1
          }
          vli = vli + 1
        }
      } else {
        // No wrapping: render runs directly, handling tab stops and line breaks
        let mut tab_idx = 0
        let mut ri = 0
        while ri < nr {
          let run = para.runs[ri]
          let display_text = if run.cap == "all" {
            to_upper_ascii(run.text)
          } else {
            run.text
          }
          // Handle line break runs (from <a:br>)
          if display_text == "\n" {
            cur_y = cur_y + p_line_h
            out = out +
              "</tspan><tspan" +
              ai("x", anchor_x) +
              ai("y", cur_y) +
              anchor_attr +
              rtl_attr +
              ">"
            ri = ri + 1
            continue
          }
          let (tab_svg, new_ti) = render_run_with_tabs(
            run,
            ri,
            display_text,
            para.tab_stops,
            tab_idx,
            text_x,
            cur_y,
            scale,
            fs,
            rels,
            theme,
          )
          out = out + tab_svg
          tab_idx = new_ti
          ri = ri + 1
        }
      }
      out = out + "</tspan>"
      cur_y = cur_y + (p_line_h - p_font_sz) + bl_extra_below
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

///|
fn estimate_text_height(
  paragraphs : Array[@ooxml.TextParagraph],
  scale : Int,
  body_props : @ooxml.BodyProps,
) -> Int {
  let def_sz = px(default_font_size_emu(), scale)
  let def_line_h = def_sz * line_height_numer / line_height_denom
  let t_ins = if body_props.t_ins >= 0 {
    body_props.t_ins
  } else {
    default_t_inset
  }
  let b_ins = if body_props.b_ins >= 0 {
    body_props.b_ins
  } else {
    default_t_inset
  }
  let np = paragraphs.length()
  let mut total = 0
  let mut i = 0
  while i < np {
    let para = paragraphs[i]
    let plh = para_line_height(para, scale, def_line_h)
    let (bl_above, bl_below) = para_baseline_shift_extra(para, scale, def_sz)
    if para.spc_before > 0 && i > 0 {
      total = total + px(para.spc_before, scale)
    }
    total = total + plh + bl_above + bl_below
    if para.spc_after > 0 {
      total = total + px(para.spc_after, scale)
    }
    i = i + 1
  }
  total + px(t_ins, scale) + px(b_ins, scale)
}
