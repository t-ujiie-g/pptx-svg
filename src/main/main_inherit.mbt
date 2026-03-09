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
/// Create a new shape with inherited transform and optionally body_props from a placeholder source.
fn shape_with_inherited_transform(
  shape : @ooxml.Shape,
  source_ph : @ooxml.Shape,
) -> @ooxml.Shape {
  {
    ..shape,
    transform: source_ph.transform,
    body_props: if shape.body_props.anchor.is_empty() &&
      not(source_ph.body_props.anchor.is_empty()) {
      source_ph.body_props
    } else {
      shape.body_props
    },
  }
}

///|
/// Resolve placeholder transforms and fills from layout/master for a slide's shapes.
///  Modifies shapes in-place by replacing zero-size transforms with inherited ones.
fn resolve_placeholder_inheritance(
  shapes : Array[@ooxml.Shape],
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    if shape.ph_type.is_empty() ||
      shape.transform.cx != 0 ||
      shape.transform.cy != 0 {
      i = i + 1
      continue
    }
    // Try layout first, then master
    let mut resolved = false
    match layout_opt {
      Some(layout) =>
        match find_placeholder(layout.shapes, shape.ph_type, shape.ph_idx) {
          Some(ph) =>
            if ph.transform.cx > 0 && ph.transform.cy > 0 {
              shapes[i] = shape_with_inherited_transform(shape, ph)
              resolved = true
            }
          None => ()
        }
      None => ()
    }
    if not(resolved) {
      match master_opt {
        Some(master) =>
          match find_placeholder(master.shapes, shape.ph_type, shape.ph_idx) {
            Some(ph) =>
              if ph.transform.cx > 0 && ph.transform.cy > 0 {
                shapes[i] = shape_with_inherited_transform(shape, ph)
              }
            None => ()
          }
        None => ()
      }
    }
    i = i + 1
  }
}

// ── Text style defaults application ─────────────────────────────────────────

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
  let np = shape.paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = shape.paragraphs[pi]
    let lvl = para.level
    let lvl_idx = if lvl >= 0 && lvl < @ooxml.max_outline_levels {
      lvl
    } else {
      0
    }
    let defaults = style.levels[lvl_idx]
    // Apply defaults to each run
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      // Apply defaults if any field is unset
      let need_font_size = run.font_size == 0 && defaults.font_size > 0
      let need_color = run.color.is_none() && not(defaults.color.is_none())
      let need_font_face = run.font_face.is_empty() &&
        not(defaults.font_face.is_empty())
      let need_ea_font = run.ea_font.is_empty() &&
        not(defaults.ea_font.is_empty())
      if need_font_size || need_color || need_font_face || need_ea_font {
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
        }
        para.runs[ri] = new_run
      }
      ri = ri + 1
    }
    pi = pi + 1
  }
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

// ── Integer → String conversion ───────────────────────────────────────────────

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
