///|
/// The placeholder in `shapes` matching by ph_idx, else by ph_type.
fn find_placeholder(
  shapes : Array[@ooxml.Shape],
  ph_type : String,
  ph_idx : Int32,
) -> @ooxml.Shape? {
  let n = shapes.length()
  let mut i = 0
  if ph_idx >= 0 {
    while i < n {
      if shapes[i].ph_idx == ph_idx {
        return Some(shapes[i])
      }
      i = i + 1
    }
  }
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
/// Give placeholders with a zero-size transform (or no body props) the layout's, else the master's.

///|
/// Try to inherit transform and body_props from a source placeholder.
/// Applies what it can and returns (got_transform, got_body_props).
fn try_inherit_from_source(
  shapes : Array[@ooxml.Shape],
  idx : Int32,
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
        new_shape = { ..new_shape, transform: ph.transform, }
      }
      if needs_body_props {
        new_shape = { ..new_shape, body_props: ph.body_props, }
      }
      shapes[idx] = new_shape
      (got_t, needs_body_props)
    }
  }
}

///|
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
    if !needs_transform && !needs_body_props {
      i = i + 1
      continue
    }
    if layout_opt is Some(layout) {
      let (got_t, got_bp) = try_inherit_from_source(
        shapes,
        i,
        shapes[i],
        layout.shapes,
        needs_transform,
        needs_body_props,
      )
      if got_t {
        needs_transform = false
      }
      if got_bp {
        needs_body_props = false
      }
    }
    if needs_transform || needs_body_props {
      if master_opt is Some(master) {
        ignore(
          try_inherit_from_source(
            shapes,
            i,
            shapes[i],
            master.shapes,
            needs_transform,
            needs_body_props,
          ),
        )
      }
    }
    i = i + 1
  }
}

///|
/// Apply the shape's own `a:lstStyle` levels to its paragraphs.
fn apply_inline_lst_style(shape : @ooxml.Shape) -> Unit {
  let style = shape.lst_style
  let nl = style.levels.length()
  if nl == 0 {
    return
  }
  apply_text_style_to_runs(shape.paragraphs, style)
  apply_para_alignment_from_style(shape.paragraphs, style)
  apply_para_spacing_from_style(shape.paragraphs, style)
}

///|
/// Apply text style defaults from a source (layout placeholder or master style)
/// to runs that are missing font_size, color, font_face, or ea_font.
fn apply_text_style_to_runs(
  paragraphs : Array[@ooxml.TextParagraph],
  style : @ooxml.TextStyleGroup,
) -> Unit {
  for para in paragraphs {
    let lvl = para.level
    let lvl_idx = if lvl >= 0 && lvl < @ooxml.max_outline_levels {
      lvl
    } else {
      0
    }
    let defaults = style.levels[lvl_idx]
    let nr = para.runs.length()
    for ri in 0..<nr {
      let run = para.runs[ri]
      let need_font_size = run.font_size == 0 && defaults.font_size > 0
      let need_color = run.color.is_none() && !defaults.color.is_none()
      let need_font_face = run.font_face.is_empty() &&
        !defaults.font_face.is_empty()
      let need_ea_font = run.ea_font.is_empty() && !defaults.ea_font.is_empty()
      let need_bold = !run.bold_explicit && !run.bold && defaults.bold
      if need_font_size ||
        need_color ||
        need_font_face ||
        need_ea_font ||
        need_bold {
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
    }
  }
}

///|
/// Apply the lst_style of the matching placeholder in `source_shapes` as a fallback.
fn apply_ph_lst_style_defaults(
  shape : @ooxml.Shape,
  source_shapes : Array[@ooxml.Shape],
) -> Unit {
  if find_placeholder(source_shapes, shape.ph_type, shape.ph_idx) is Some(ph) {
    let ph_style = ph.lst_style
    apply_text_style_to_runs(shape.paragraphs, ph_style)
    apply_para_alignment_from_style(shape.paragraphs, ph_style)
    apply_para_spacing_from_style(shape.paragraphs, ph_style)
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
  if layout_opt is Some(layout) {
    apply_ph_lst_style_defaults(shape, layout.shapes)
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
  for pi in 0..<np {
    let para = paragraphs[pi]
    if para.align.is_empty() {
      let lvl_idx = if para.level >= 0 && para.level < @ooxml.max_outline_levels {
        para.level
      } else {
        0
      }
      let align = style.levels[lvl_idx].align
      if !align.is_empty() {
        paragraphs[pi] = { ..para, align, }
      }
    }
  }
}

///|
/// Fill a paragraph's unset spacing, margins and bullet from a TextStyleGroup.
fn apply_para_spacing_from_style(
  paragraphs : Array[@ooxml.TextParagraph],
  style : @ooxml.TextStyleGroup,
) -> Unit {
  let np = paragraphs.length()
  for pi in 0..<np {
    let para = paragraphs[pi]
    let lvl_idx = if para.level >= 0 && para.level < @ooxml.max_outline_levels {
      para.level
    } else {
      0
    }
    let defaults = style.levels[lvl_idx]
    let mut changed = false
    let mut new_para = para
    if para.line_spacing == 0 && defaults.line_spacing != 0 {
      new_para = { ..new_para, line_spacing: defaults.line_spacing, }
      changed = true
    }
    if para.spc_before == 0 && defaults.spc_before != 0 {
      new_para = { ..new_para, spc_before: defaults.spc_before, }
      changed = true
    }
    if para.spc_after == 0 && defaults.spc_after != 0 {
      new_para = { ..new_para, spc_after: defaults.spc_after, }
      changed = true
    }
    if para.mar_l < 0 && defaults.mar_l >= 0 {
      new_para = { ..new_para, mar_l: defaults.mar_l, }
      changed = true
    }
    if para.indent == 0 && defaults.indent != 0 {
      new_para = { ..new_para, indent: defaults.indent, }
      changed = true
    }
    // Bullets are inherited only when the paragraph sets none of bullet / buAutoNum / buNone.
    let has_any_bullet_spec = !para.bullet.is_empty() ||
      !para.bullet_auto.is_empty() ||
      para.bullet_none
    if !has_any_bullet_spec {
      if defaults.bullet_none {
        new_para = { ..new_para, bullet_none: true, }
        changed = true
      } else if !defaults.bullet_auto.is_empty() {
        new_para = {
          ..new_para,
          bullet_auto: defaults.bullet_auto,
          bullet_font: if para.bullet_font.is_empty() {
            defaults.bullet_font
          } else {
            para.bullet_font
          },
          bullet_size: if para.bullet_size == 0 {
            defaults.bullet_size
          } else {
            para.bullet_size
          },
          bullet_color: if para.bullet_color.is_none() {
            defaults.bullet_color
          } else {
            para.bullet_color
          },
        }
        changed = true
      } else if !defaults.bullet.is_empty() {
        new_para = {
          ..new_para,
          bullet: defaults.bullet,
          bullet_font: if para.bullet_font.is_empty() {
            defaults.bullet_font
          } else {
            para.bullet_font
          },
          bullet_size: if para.bullet_size == 0 {
            defaults.bullet_size
          } else {
            para.bullet_size
          },
          bullet_color: if para.bullet_color.is_none() {
            defaults.bullet_color
          } else {
            para.bullet_color
          },
        }
        changed = true
      }
    }
    if changed {
      paragraphs[pi] = new_para
    }
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
  if master_opt is Some(master) {
    apply_ph_lst_style_defaults(shape, master.shapes)
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

///|
/// Last-resort run defaults from endParaRPr, after all placeholder inheritance: PowerPoint
/// carries this "remembered run state" across a free textbox's paragraphs.
fn apply_epr_fallbacks(shape : @ooxml.Shape) -> Unit {
  // Shape-wide fallback: the first endParaRPr value of each kind in any paragraph.
  let mut fb_fs = 0
  let mut fb_clr = @ooxml.Color::none()
  let mut fb_ff = ""
  let mut fb_ea = ""
  for para in shape.paragraphs {
    if fb_fs == 0 && para.epr_font_size > 0 {
      fb_fs = para.epr_font_size
    }
    if fb_clr.is_none() && !para.epr_color.is_none() {
      fb_clr = para.epr_color
    }
    if fb_ff.is_empty() && !para.epr_font_face.is_empty() {
      fb_ff = para.epr_font_face
    }
    if fb_ea.is_empty() && !para.epr_ea_font.is_empty() {
      fb_ea = para.epr_ea_font
    }
  }
  if fb_fs == 0 && fb_clr.is_none() && fb_ff.is_empty() && fb_ea.is_empty() {
    return
  }
  // Precedence: run value > the paragraph's epr_* > the shape-wide fallback.
  for para in shape.paragraphs {
    let eff_fs = if para.epr_font_size > 0 { para.epr_font_size } else { fb_fs }
    let eff_clr = if !para.epr_color.is_none() {
      para.epr_color
    } else {
      fb_clr
    }
    let eff_ff = if !para.epr_font_face.is_empty() {
      para.epr_font_face
    } else {
      fb_ff
    }
    let eff_ea = if !para.epr_ea_font.is_empty() {
      para.epr_ea_font
    } else {
      fb_ea
    }
    let nr = para.runs.length()
    for ri in 0..<nr {
      let run = para.runs[ri]
      let need_fs = eff_fs > 0 && run.font_size == 0
      let need_clr = !eff_clr.is_none() && run.color.is_none()
      let need_ff = !eff_ff.is_empty() && run.font_face.is_empty()
      let need_ea = !eff_ea.is_empty() && run.ea_font.is_empty()
      if need_fs || need_clr || need_ff || need_ea {
        para.runs[ri] = {
          ..run,
          font_size: if need_fs {
            eff_fs
          } else {
            run.font_size
          },
          color: if need_clr {
            eff_clr
          } else {
            run.color
          },
          font_face: if need_ff {
            eff_ff
          } else {
            run.font_face
          },
          ea_font: if need_ea {
            eff_ea
          } else {
            run.ea_font
          },
        }
      }
    }
  }
}

///|
/// Color still-uncolored runs with the shape's `p:style/a:fontRef` color, as PowerPoint
/// does (accent rectangles get white text through a fontRef to lt1).
fn apply_font_ref_fallback(shape : @ooxml.Shape) -> Unit {
  if shape.font_ref_color.is_none() {
    return
  }
  let fr_clr = shape.font_ref_color
  for para in shape.paragraphs {
    let nr = para.runs.length()
    for ri in 0..<nr {
      let run = para.runs[ri]
      if run.color.is_none() {
        para.runs[ri] = { ..run, color: fr_clr, }
      }
    }
  }
}

///|
/// Check if all paragraphs in a shape have no text.
fn has_no_text(paragraphs : Array[@ooxml.TextParagraph]) -> Bool {
  let np = paragraphs.length()
  if np == 0 {
    return true
  }
  for i in 0..<np {
    let nr = paragraphs[i].runs.length()
    for j in 0..<nr {
      if !paragraphs[i].runs[j].text.is_empty() {
        return false
      }
    }
  }
  true
}

///|
/// Fill date / footer / slide-number placeholders (`slide_num` 1-indexed): sldNum always
/// shows the real number, dt the host date (else the layout/master literal), and ftr
/// keeps its own text or inherits it. The existing field run's formatting is kept.
fn inject_placeholder_auto_content(
  shapes : Array[@ooxml.Shape],
  slide_num : Int32,
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    let ph = shape.ph_type
    if ph != "sldNum" && ph != "dt" && ph != "ftr" {
      i = i + 1
      continue
    }
    // Empty = keep the shape's existing content.
    let auto_text = if ph == "sldNum" {
      int_to_str(slide_num)
    } else if ph == "dt" {
      let host_date = @ffi.ffi_current_date()
      if !host_date.is_empty() {
        host_date
      } else {
        find_ph_text(ph, shape.ph_idx, layout_opt, master_opt)
      }
    } else if has_no_text(shape.paragraphs) {
      find_ph_text(ph, shape.ph_idx, layout_opt, master_opt)
    } else {
      ""
    }
    if !auto_text.is_empty() {
      shapes[i] = {
        ..shape,
        paragraphs: [replace_auto_content_text(shape.paragraphs, auto_text)],
      }
    }
    i = i + 1
  }
}

///|
/// The auto-content paragraph: the placeholder's first run with only its text replaced.
fn replace_auto_content_text(
  paragraphs : Array[@ooxml.TextParagraph],
  text : String,
) -> @ooxml.TextParagraph {
  for para in paragraphs {
    if !para.runs.is_empty() {
      return { ..para, runs: [{ ..para.runs[0], text, math_xml: "", }], }
    }
  }
  make_default_paragraph(make_default_run(text))
}

///|
/// The text of the matching placeholder in the layout, else the master.
fn find_ph_text(
  ph_type : String,
  ph_idx : Int32,
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
) -> String {
  if layout_opt is Some(layout) {
    if find_placeholder(layout.shapes, ph_type, ph_idx) is Some(ph) {
      let text = extract_text(ph.paragraphs)
      if !text.is_empty() {
        return text
      }
    }
  }
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
  let mut result = ""
  for paragraph in paragraphs {
    for run in paragraph.runs {
      result = result + run.text
    }
  }
  result
}
