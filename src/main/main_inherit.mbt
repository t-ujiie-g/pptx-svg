// Main — Placeholder inheritance + text style defaults + utilities
//
// Same package as main.mbt — all globals and helpers are shared.

// ── Placeholder inheritance ──────────────────────────────────────────────────

///| Find a matching placeholder shape in a shapes array by ph_type and ph_idx.
fn find_placeholder(shapes : Array[@ooxml.Shape], ph_type : String, ph_idx : Int) -> @ooxml.Shape? {
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

///| Resolve placeholder transforms and fills from layout/master for a slide's shapes.
///  Modifies shapes in-place by replacing zero-size transforms with inherited ones.
fn resolve_placeholder_inheritance(
  shapes : Array[@ooxml.Shape],
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    if shape.ph_type.is_empty() {
      i = i + 1
      continue
    }
    let t = shape.transform
    let needs_transform = t.cx == 0 && t.cy == 0
    if not(needs_transform) {
      i = i + 1
      continue
    }
    // Try layout first, then master
    let mut resolved = false
    match layout_opt {
      Some(layout) => {
        match find_placeholder(layout.shapes, shape.ph_type, shape.ph_idx) {
          Some(layout_ph) => {
            if layout_ph.transform.cx > 0 && layout_ph.transform.cy > 0 {
              shapes[i] = {
                kind: shape.kind,
                transform: layout_ph.transform,
                fill: shape.fill,
                grad_fill: shape.grad_fill,
                blip_fill: shape.blip_fill,
                patt_fill: shape.patt_fill,
                stroke: shape.stroke,
                stroke_w: shape.stroke_w,
                paragraphs: shape.paragraphs,
                body_props: if shape.body_props.anchor.is_empty() && not(layout_ph.body_props.anchor.is_empty()) {
                  layout_ph.body_props
                } else {
                  shape.body_props
                },
                ph_type: shape.ph_type,
                ph_idx: shape.ph_idx,
              }
              resolved = true
            }
          }
          None => ()
        }
      }
      None => ()
    }
    if not(resolved) {
      match master_opt {
        Some(master) => {
          match find_placeholder(master.shapes, shape.ph_type, shape.ph_idx) {
            Some(master_ph) => {
              if master_ph.transform.cx > 0 && master_ph.transform.cy > 0 {
                shapes[i] = {
                  kind: shape.kind,
                  transform: master_ph.transform,
                  fill: shape.fill,
                  grad_fill: shape.grad_fill,
                  blip_fill: shape.blip_fill,
                  patt_fill: shape.patt_fill,
                  stroke: shape.stroke,
                  stroke_w: shape.stroke_w,
                  paragraphs: shape.paragraphs,
                  body_props: if shape.body_props.anchor.is_empty() && not(master_ph.body_props.anchor.is_empty()) {
                    master_ph.body_props
                  } else {
                    shape.body_props
                  },
                  ph_type: shape.ph_type,
                  ph_idx: shape.ph_idx,
                }
              }
            }
            None => ()
          }
        }
        None => ()
      }
    }
    i = i + 1
  }
}

// ── Text style defaults application ─────────────────────────────────────────

///| Apply master text style defaults to a shape's paragraphs based on ph_type.
fn apply_text_defaults(shape : @ooxml.Shape, text_styles : @ooxml.MasterTextStyles) -> Unit {
  let ph = shape.ph_type
  if ph.is_empty() { return }
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
    let lvl_idx = if lvl >= 0 && lvl < @ooxml.max_outline_levels { lvl } else { 0 }
    let defaults = style.levels[lvl_idx]
    // Apply defaults to each run
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      // Apply defaults if any field is unset
      let need_font_size = run.font_size == 0 && defaults.font_size > 0
      let need_color = run.color.is_none() && not(defaults.color.is_none())
      let need_font_face = run.font_face.is_empty() && not(defaults.font_face.is_empty())
      let need_ea_font = run.ea_font.is_empty() && not(defaults.ea_font.is_empty())
      if need_font_size || need_color || need_font_face || need_ea_font {
        let new_run : @ooxml.TextRun = {
          ..run,
          font_size: if need_font_size { defaults.font_size } else { run.font_size },
          color: if need_color { defaults.color } else { run.color },
          font_face: if need_font_face { defaults.font_face } else { run.font_face },
          ea_font: if need_ea_font { defaults.ea_font } else { run.ea_font },
        }
        para.runs[ri] = new_run
      }
      ri = ri + 1
    }
    pi = pi + 1
  }
}

// ── Theme path resolution ────────────────────────────────────────────────────

///| Find the theme file path from presentation.xml.rels.
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

fn ends_with(s : String, pat : String) -> Bool { @xml.str_ends_with(s, pat) }

fn starts_with(s : String, pat : String) -> Bool { @xml.str_starts_with(s, pat) }

fn suffix(s : String, from : Int) -> String { @xml.str_suffix(s, from) }

// ── Integer → String conversion ───────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }

// ── Slide counting ────────────────────────────────────────────────────────────

fn count_slide_ids(xml : String) -> Int {
  count_occurrences(xml, "<p:sldId ") + count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") + count_occurrences(xml, "<p:sldId/>")
}

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

fn match_at(text : String, pattern : String, pos : Int) -> Bool {
  let plen = pattern.length()
  for j = 0; j < plen; j = j + 1 {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}
