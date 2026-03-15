// Main — Placeholder inheritance + text style defaults + utilities
//
// Same package as main.mbt — all globals and helpers are shared.

// ── Placeholder inheritance ──────────────────────────────────────────────────

///|
/// Find a matching placeholder shape in a shapes array by ph_type and ph_idx.
fn find_placeholder(
  shapes : Array[@ooxml.Shape],
  ph_type : String,
  ph_idx : Int,
) -> @ooxml.Shape? {
  let n = shapes.length()
  let mut i = 0
  // First try matching by ph_idx (if set)
  if ph_idx >= 0 {
    while i < n {
      if shapes[i].ph_idx == ph_idx {
        return Some(shapes[i])
      }
      i = i + 1
    }
  }
  // Fallback: match by ph_type
  i = 0
  while i < n {
    if shapes[i].ph_type == ph_type {
      return Some(shapes[i])
    }
    i = i + 1
  }
  None
}

///|
/// Resolve placeholder transforms and fills from layout/master for a slide's shapes.
///  Modifies shapes in-place by replacing zero-size transforms with inherited ones.
///|
/// Try to inherit transform and body_props from a source placeholder.
/// Applies what it can and returns (got_transform, got_body_props).
fn try_inherit_from_source(
  shapes : Array[@ooxml.Shape],
  idx : Int,
  shape : @ooxml.Shape,
  source_shapes : Array[@ooxml.Shape],
  needs_transform : Bool,
  needs_body_props : Bool,
) -> (Bool, Bool) {
  match find_placeholder(source_shapes, shape.ph_type, shape.ph_idx) {
    None => (false, false)
    Some(ph) => {
      let mut new_shape = shape
      let got_t = needs_transform && ph.transform.cx > 0 && ph.transform.cy > 0
      if got_t {
        new_shape = { ..new_shape, transform: ph.transform }
      }
      if needs_body_props {
        new_shape = { ..new_shape, body_props: ph.body_props }
      }
      shapes[idx] = new_shape
      (got_t, needs_body_props)
    }
  }
}

fn resolve_placeholder_inheritance(
  shapes : Array[@ooxml.Shape],
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    if shape.ph_type.is_empty() {
      i = i + 1
      continue
    }
    let mut needs_transform = shape.transform.cx == 0 && shape.transform.cy == 0
    let mut needs_body_props = shape.body_props.l_ins < 0 &&
      shape.body_props.t_ins < 0 &&
      shape.body_props.r_ins < 0 &&
      shape.body_props.b_ins < 0 &&
      shape.body_props.anchor.is_empty()
    if not(needs_transform) && not(needs_body_props) {
      i = i + 1
      continue
    }
    // Try layout first
    match layout_opt {
      Some(layout) => {
        let (got_t, got_bp) = try_inherit_from_source(
          shapes, i, shapes[i], layout.shapes, needs_transform, needs_body_props,
        )
        if got_t {
          needs_transform = false
        }
        if got_bp {
          needs_body_props = false
        }
      }
      None => ()
    }
    // Try master for anything still needed
    if needs_transform || needs_body_props {
      match master_opt {
        Some(master) =>
          ignore(
            try_inherit_from_source(
              shapes, i, shapes[i], master.shapes, needs_transform,
              needs_body_props,
            ),
          )
        None => ()
      }
    }
    i = i + 1
  }
}

// ── Inline lstStyle application ──────────────────────────────────────────────

///|
/// Apply the shape's own inline lstStyle (from p:txBody/a:lstStyle) to its
/// paragraphs. This handles non-placeholder text boxes that define their own
/// per-level text defaults (font size, line spacing, bullets, etc.).
fn apply_inline_lst_style(shape : @ooxml.Shape) -> Unit {
  let style = shape.lst_style
  // Check if the style has any non-empty levels
  let nl = style.levels.length()
  if nl == 0 {
    return
  }
  apply_text_style_to_runs(shape.paragraphs, style)
  apply_para_alignment_from_style(shape.paragraphs, style)
  apply_para_spacing_from_style(shape.paragraphs, style)
}

// ── Text style defaults application ─────────────────────────────────────────

///|
/// Apply text style defaults from a source (layout placeholder or master style)
/// to runs that are missing font_size, color, font_face, or ea_font.
fn apply_text_style_to_runs(
  paragraphs : Array[@ooxml.TextParagraph],
  style : @ooxml.TextStyleGroup,
) -> Unit {
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let lvl = para.level
    let lvl_idx = if lvl >= 0 && lvl < @ooxml.max_outline_levels {
      lvl
    } else {
      0
    }
    let defaults = style.levels[lvl_idx]
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      let need_font_size = run.font_size == 0 && defaults.font_size > 0
      let need_color = run.color.is_none() && not(defaults.color.is_none())
      let need_font_face = run.font_face.is_empty() &&
        not(defaults.font_face.is_empty())
      let need_ea_font = run.ea_font.is_empty() &&
        not(defaults.ea_font.is_empty())
      let need_bold = not(run.bold_explicit) && not(run.bold) && defaults.bold
      if need_font_size || need_color || need_font_face || need_ea_font || need_bold {
        let new_run : @ooxml.TextRun = {
          ..run,
          font_size: if need_font_size {
            defaults.font_size
          } else {
            run.font_size
          },
          color: if need_color {
            defaults.color
          } else {
            run.color
          },
          font_face: if need_font_face {
            defaults.font_face
          } else {
            run.font_face
          },
          ea_font: if need_ea_font {
            defaults.ea_font
          } else {
            run.ea_font
          },
          bold: if need_bold {
            true
          } else {
            run.bold
          },
        }
        para.runs[ri] = new_run
      }
      ri = ri + 1
    }
    pi = pi + 1
  }
}

///|
/// Apply a placeholder's lst_style to a shape's paragraphs as fallback.
/// Searches source_shapes for a matching placeholder, then applies its
/// lst_style for run defaults (font_size, color, bold) and alignment.
fn apply_ph_lst_style_defaults(
  shape : @ooxml.Shape,
  source_shapes : Array[@ooxml.Shape],
) -> Unit {
  match find_placeholder(source_shapes, shape.ph_type, shape.ph_idx) {
    None => ()
    Some(ph) => {
      let ph_style = ph.lst_style
      apply_text_style_to_runs(shape.paragraphs, ph_style)
      apply_para_alignment_from_style(shape.paragraphs, ph_style)
      apply_para_spacing_from_style(shape.paragraphs, ph_style)
    }
  }
}

///|
/// Apply layout placeholder text styles to a shape's paragraphs.
fn apply_layout_text_defaults(
  shape : @ooxml.Shape,
  layout_opt : @ooxml.SlideLayoutData?,
) -> Unit {
  if shape.ph_type.is_empty() {
    return
  }
  match layout_opt {
    None => ()
    Some(layout) => apply_ph_lst_style_defaults(shape, layout.shapes)
  }
}

///|
/// Apply paragraph alignment from a TextStyleGroup when the paragraph
/// has no explicit alignment.
fn apply_para_alignment_from_style(
  paragraphs : Array[@ooxml.TextParagraph],
  style : @ooxml.TextStyleGroup,
) -> Unit {
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    if para.align.is_empty() {
      let lvl_idx = if para.level >= 0 && para.level < @ooxml.max_outline_levels {
        para.level
      } else {
        0
      }
      let align = style.levels[lvl_idx].align
      if not(align.is_empty()) {
        paragraphs[pi] = { ..para, align, }
      }
    }
    pi = pi + 1
  }
}

///|
/// Apply paragraph-level spacing and bullet defaults from a TextStyleGroup.
/// Inherits line_spacing, spc_before, spc_after, mar_l, indent, and bullet
/// properties when the paragraph has no explicit value.
fn apply_para_spacing_from_style(
  paragraphs : Array[@ooxml.TextParagraph],
  style : @ooxml.TextStyleGroup,
) -> Unit {
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let lvl_idx = if para.level >= 0 && para.level < @ooxml.max_outline_levels {
      para.level
    } else {
      0
    }
    let defaults = style.levels[lvl_idx]
    let mut changed = false
    let mut new_para = para
    // line_spacing: 0 = unset
    if para.line_spacing == 0 && defaults.line_spacing != 0 {
      new_para = { ..new_para, line_spacing: defaults.line_spacing }
      changed = true
    }
    // spc_before: 0 = unset
    if para.spc_before == 0 && defaults.spc_before != 0 {
      new_para = { ..new_para, spc_before: defaults.spc_before }
      changed = true
    }
    // spc_after: 0 = unset
    if para.spc_after == 0 && defaults.spc_after != 0 {
      new_para = { ..new_para, spc_after: defaults.spc_after }
      changed = true
    }
    // mar_l: 0 = unset (note: -1 means explicitly unset in LevelTextDefaults)
    if para.mar_l == 0 && defaults.mar_l > 0 {
      new_para = { ..new_para, mar_l: defaults.mar_l }
      changed = true
    }
    // indent: 0 = unset
    if para.indent == 0 && defaults.indent != 0 {
      new_para = { ..new_para, indent: defaults.indent }
      changed = true
    }
    // Note: bullet properties (bullet, bullet_font, bullet_size, bullet_color)
    // are NOT inherited here. They are already set explicitly in the slide XML
    // for paragraphs that need them, and inheriting from style defaults can
    // incorrectly add bullets to non-bullet paragraphs (e.g. title lines).
    if changed {
      paragraphs[pi] = new_para
    }
    pi = pi + 1
  }
}

///|
/// Apply master placeholder text styles to a shape's paragraphs.
fn apply_master_ph_text_defaults(
  shape : @ooxml.Shape,
  master_opt : @ooxml.SlideMasterData?,
) -> Unit {
  if shape.ph_type.is_empty() {
    return
  }
  match master_opt {
    None => ()
    Some(master) => apply_ph_lst_style_defaults(shape, master.shapes)
  }
}

///|
/// Apply master text style defaults to a shape's paragraphs based on ph_type.
fn apply_text_defaults(
  shape : @ooxml.Shape,
  text_styles : @ooxml.MasterTextStyles,
) -> Unit {
  let ph = shape.ph_type
  if ph.is_empty() {
    return
  }
  let style = if ph == "title" || ph == "ctrTitle" {
    text_styles.title_style
  } else if ph == "body" || ph == "subTitle" || ph == "obj" {
    text_styles.body_style
  } else {
    text_styles.other_style
  }
  apply_text_style_to_runs(shape.paragraphs, style)
  apply_para_spacing_from_style(shape.paragraphs, style)
}

// ── Placeholder auto-content ─────────────────────────────────────────────

///|
/// Check if all paragraphs in a shape have no text.
fn has_no_text(paragraphs : Array[@ooxml.TextParagraph]) -> Bool {
  let np = paragraphs.length()
  if np == 0 {
    return true
  }
  let mut i = 0
  while i < np {
    let nr = paragraphs[i].runs.length()
    let mut j = 0
    while j < nr {
      if not(paragraphs[i].runs[j].text.is_empty()) {
        return false
      }
      j = j + 1
    }
    i = i + 1
  }
  true
}

///|
/// Inject auto-content text for empty dt/sldNum/ftr placeholders.
/// slide_num is 1-indexed.
fn inject_placeholder_auto_content(
  shapes : Array[@ooxml.Shape],
  slide_num : Int,
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    let ph = shape.ph_type
    if ph.is_empty() || not(has_no_text(shape.paragraphs)) {
      i = i + 1
      continue
    }
    // Only process dt/sldNum/ftr
    let auto_text = if ph == "sldNum" {
      int_to_str(slide_num)
    } else if ph == "dt" {
      // Try to get date text from layout/master placeholder
      let ph_text = find_ph_text(ph, shape.ph_idx, layout_opt, master_opt)
      ph_text
    } else if ph == "ftr" {
      // Try to get footer text from layout/master placeholder
      find_ph_text(ph, shape.ph_idx, layout_opt, master_opt)
    } else {
      ""
    }
    if not(auto_text.is_empty()) {
      // Create a single paragraph with the auto text
      let run : @ooxml.TextRun = {
        text: auto_text,
        bold: false,
        italic: false,
        font_size: 0,
        color: @ooxml.Color::none(),
        font_face: "",
        ea_font: "",
        cs_font: "",
        sym_font: "",
        underline: "",
        strike: "",
        baseline: 0,
        char_spacing: 0,
        kern: 0,
        cap: "",
        hlink_rid: "",
        hlink_mouse_over_rid: "",
        effects: @ooxml.EffectList::none(),
        outline_color: @ooxml.Color::none(),
        outline_w: 0,
        text_grad_fill: @ooxml.GradientFill::none(),
        text_patt_fill: @ooxml.PatternFill::none(),
        bold_explicit: false,
      }
      let para : @ooxml.TextParagraph = {
        runs: [run],
        align: "",
        level: 0,
        spc_before: 0,
        spc_after: 0,
        line_spacing: 0,
        mar_l: 0,
        indent: 0,
        bullet: "",
        bullet_auto: "",
        bullet_none: false,
        bullet_font: "",
        bullet_size: 0,
        bullet_color: @ooxml.Color::none(),
        bullet_img_rid: "",
        tab_stops: [],
        rtl: false,
      }
      shapes[i] = { ..shape, paragraphs: [para] }
    }
    i = i + 1
  }
}

///|
/// Find text content from a placeholder in layout or master.
fn find_ph_text(
  ph_type : String,
  ph_idx : Int,
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> String {
  // Try layout first
  match layout_opt {
    Some(layout) =>
      match find_placeholder(layout.shapes, ph_type, ph_idx) {
        Some(ph) => {
          let text = extract_text(ph.paragraphs)
          if not(text.is_empty()) {
            return text
          }
        }
        None => ()
      }
    None => ()
  }
  // Try master
  match master_opt {
    Some(master) =>
      match find_placeholder(master.shapes, ph_type, ph_idx) {
        Some(ph) => extract_text(ph.paragraphs)
        None => ""
      }
    None => ""
  }
}

///|
/// Extract plain text from paragraphs.
fn extract_text(paragraphs : Array[@ooxml.TextParagraph]) -> String {
  let np = paragraphs.length()
  let mut result = ""
  let mut pi = 0
  while pi < np {
    let nr = paragraphs[pi].runs.length()
    let mut ri = 0
    while ri < nr {
      result = result + paragraphs[pi].runs[ri].text
      ri = ri + 1
    }
    pi = pi + 1
  }
  result
}

// ── Theme path resolution ────────────────────────────────────────────────────

///|
/// Find the theme file path from presentation.xml.rels.
fn find_theme_path(rels : Array[@ooxml.Relationship]) -> String {
  let n = rels.length()
  let mut i = 0
  while i < n {
    let rel = rels[i]
    // Theme relationship type ends with "/theme"
    if ends_with(rel.type_, "/theme") {
      // Target is relative to ppt/ directory
      if starts_with(rel.target, "/") {
        return suffix(rel.target, 1)
      } else {
        return "ppt/" + rel.target
      }
    }
    i = i + 1
  }
  // Fallback: try the standard location
  ""
}

// ── Package-local string/int helpers (shared by main.mbt) ────────────────────

///|
fn ends_with(s : String, pat : String) -> Bool {
  @xml.str_ends_with(s, pat)
}

///|
fn starts_with(s : String, pat : String) -> Bool {
  @xml.str_starts_with(s, pat)
}

///|
fn suffix(s : String, from : Int) -> String {
  @xml.str_suffix(s, from)
}

///|
fn int_to_str(n : Int) -> String {
  @xml.int_to_str(n)
}

// ── Slide counting ────────────────────────────────────────────────────────────

///|
fn count_slide_ids(xml : String) -> Int {
  count_occurrences(xml, "<p:sldId ") +
  count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") +
  count_occurrences(xml, "<p:sldId/>")
}

///|
fn count_occurrences(text : String, pattern : String) -> Int {
  let plen = pattern.length()
  let tlen = text.length()
  if plen == 0 || tlen < plen {
    return 0
  }
  let mut count = 0
  let mut i = 0
  while i <= tlen - plen {
    if match_at(text, pattern, i) {
      count = count + 1
      i = i + plen
    } else {
      i = i + 1
    }
  }
  count
}

///|
fn match_at(text : String, pattern : String, pos : Int) -> Bool {
  let plen = pattern.length()
  for j = 0; j < plen; j = j + 1 {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}
