// OOXML slide XML serializer — SlideData → OOXML XML
//
// Generates valid pptx slide XML from the internal SlideData model.
// Used for the round-trip pipeline: PPTX → SlideData → SVG → SlideData → PPTX.
//
// NOTE: No StringBuilder — all strings built via `+` (wasm:js-string "concat").

// ── Shared helpers (from @xml) ───────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }

// ── XML helpers ──────────────────────────────────────────────────────────────

fn xml_escape(s : String) -> String {
  let len = s.length()
  let mut needs = false
  let mut k = 0
  while k < len {
    let c = s.get_char(k).unwrap()
    if c == '<' || c == '>' || c == '&' || c == '"' || c == '\'' {
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
    } else if c == '\'' {
      result = result + "&apos;"
    } else {
      result = result + @ffi.ffi_char_code_to_str(Char::to_int(c))
    }
    i = i + 1
  }
  result
}

// ── Color serialization ─────────────────────────────────────────────────────

///| Serialize inner color element (a:srgbClr with optional a:alpha child).
///  No solidFill wrapper — used for gradient stops, pattFill fg/bg, etc.
fn serialize_color_inner(color : @ooxml.Color) -> String {
  if color.alpha < 255 {
    // alpha → OOXML 1/1000th percent: alpha * 100000 / 255
    let alpha_val = color.alpha * 100000 / 255
    "<a:srgbClr val=\"" + color.to_hex_rgb() + "\"><a:alpha val=\"" + int_to_str(alpha_val) + "\"/></a:srgbClr>"
  } else {
    "<a:srgbClr val=\"" + color.to_hex_rgb() + "\"/>"
  }
}

fn serialize_color(color : @ooxml.Color) -> String {
  if color.is_none() {
    "<a:noFill/>"
  } else {
    "<a:solidFill>" + serialize_color_inner(color) + "</a:solidFill>"
  }
}

// ── Gradient fill serialization ──────────────────────────────────────────────

fn serialize_gradient_fill(gf : @ooxml.GradientFill) -> String {
  let rot_attr = if gf.rot_with_shape { " rotWithShape=\"1\"" } else { " rotWithShape=\"0\"" }
  let flip_attr = if gf.tile_flip.is_empty() || gf.tile_flip == "none" { "" } else { " tileFlip=\"" + gf.tile_flip + "\"" }
  let mut out = "<a:gradFill" + rot_attr + flip_attr + ">"
  // Gradient stops
  out = out + "<a:gsLst>"
  let ns = gf.stops.length()
  let mut si = 0
  while si < ns {
    let stop = gf.stops[si]
    out = out + "<a:gs pos=\"" + int_to_str(stop.pos) + "\">" + serialize_color_inner(stop.color) + "</a:gs>"
    si = si + 1
  }
  out = out + "</a:gsLst>"
  // Direction
  if gf.angle >= 0 && gf.path_type.is_empty() {
    // Linear
    out = out + "<a:lin ang=\"" + int_to_str(gf.angle) + "\" scaled=\"1\"/>"
  } else if not(gf.path_type.is_empty()) {
    // Path gradient
    out = out + "<a:path path=\"" + gf.path_type + "\">"
    if gf.fill_to_l >= 0 || gf.fill_to_t >= 0 || gf.fill_to_r >= 0 || gf.fill_to_b >= 0 {
      out = out + "<a:fillToRect"
      if gf.fill_to_l >= 0 { out = out + " l=\"" + int_to_str(gf.fill_to_l) + "\"" }
      if gf.fill_to_t >= 0 { out = out + " t=\"" + int_to_str(gf.fill_to_t) + "\"" }
      if gf.fill_to_r >= 0 { out = out + " r=\"" + int_to_str(gf.fill_to_r) + "\"" }
      if gf.fill_to_b >= 0 { out = out + " b=\"" + int_to_str(gf.fill_to_b) + "\"" }
      out = out + "/>"
    }
    out = out + "</a:path>"
  }
  out + "</a:gradFill>"
}

// ── Blip fill serialization ──────────────────────────────────────────────────

fn serialize_blip_fill(bf : @ooxml.BlipFill) -> String {
  let mut out = "<a:blipFill>"
  out = out + "<a:blip r:embed=\"" + bf.rid + "\"/>"
  if bf.src_l > 0 || bf.src_t > 0 || bf.src_r > 0 || bf.src_b > 0 {
    out = out + "<a:srcRect"
    if bf.src_l > 0 { out = out + " l=\"" + int_to_str(bf.src_l) + "\"" }
    if bf.src_t > 0 { out = out + " t=\"" + int_to_str(bf.src_t) + "\"" }
    if bf.src_r > 0 { out = out + " r=\"" + int_to_str(bf.src_r) + "\"" }
    if bf.src_b > 0 { out = out + " b=\"" + int_to_str(bf.src_b) + "\"" }
    out = out + "/>"
  }
  if bf.stretch {
    out = out + "<a:stretch><a:fillRect/></a:stretch>"
  } else if bf.tile_sx > 0 || bf.tile_sy > 0 {
    out = out + "<a:tile"
    if bf.tile_tx != 0 { out = out + " tx=\"" + int_to_str(bf.tile_tx) + "\"" }
    if bf.tile_ty != 0 { out = out + " ty=\"" + int_to_str(bf.tile_ty) + "\"" }
    if bf.tile_sx != 0 { out = out + " sx=\"" + int_to_str(bf.tile_sx) + "\"" }
    if bf.tile_sy != 0 { out = out + " sy=\"" + int_to_str(bf.tile_sy) + "\"" }
    if not(bf.tile_flip.is_empty()) && bf.tile_flip != "none" { out = out + " flip=\"" + bf.tile_flip + "\"" }
    if not(bf.tile_algn.is_empty()) { out = out + " algn=\"" + bf.tile_algn + "\"" }
    out = out + "/>"
  }
  out + "</a:blipFill>"
}

// ── Pattern fill serialization ──────────────────────────────────────────────

fn serialize_pattern_fill(pf : @ooxml.PatternFill) -> String {
  let mut out = "<a:pattFill prst=\"" + pf.prst + "\">"
  out = out + "<a:fgClr>" + serialize_color_inner(pf.fg_color) + "</a:fgClr>"
  out = out + "<a:bgClr>" + serialize_color_inner(pf.bg_color) + "</a:bgClr>"
  out + "</a:pattFill>"
}

// ── Text body serialization ──────────────────────────────────────────────────

///| Serialize a:bodyPr element.
fn serialize_body_pr(bp : @ooxml.BodyProps) -> String {
  let has_attrs = not(bp.anchor.is_empty()) || bp.l_ins >= 0 || bp.t_ins >= 0 ||
    bp.r_ins >= 0 || bp.b_ins >= 0 || bp.wrap == "none" || bp.rot != 0 ||
    not(bp.vert.is_empty()) || bp.num_cols > 1 || bp.col_spacing > 0
  let has_children = bp.auto_fit || bp.font_scale >= 0
  if not(has_attrs) && not(has_children) {
    return "<a:bodyPr/>"
  }
  let mut out = "<a:bodyPr"
  if not(bp.anchor.is_empty()) {
    out = out + " anchor=\"" + bp.anchor + "\""
  }
  if bp.wrap == "none" {
    out = out + " wrap=\"none\""
  }
  if bp.rot != 0 {
    out = out + " rot=\"" + int_to_str(bp.rot) + "\""
  }
  if not(bp.vert.is_empty()) {
    out = out + " vert=\"" + bp.vert + "\""
  }
  if bp.num_cols > 1 {
    out = out + " numCol=\"" + int_to_str(bp.num_cols) + "\""
  }
  if bp.col_spacing > 0 {
    out = out + " spcCol=\"" + int_to_str(bp.col_spacing) + "\""
  }
  if bp.l_ins >= 0 {
    out = out + " lIns=\"" + int_to_str(bp.l_ins) + "\""
  }
  if bp.t_ins >= 0 {
    out = out + " tIns=\"" + int_to_str(bp.t_ins) + "\""
  }
  if bp.r_ins >= 0 {
    out = out + " rIns=\"" + int_to_str(bp.r_ins) + "\""
  }
  if bp.b_ins >= 0 {
    out = out + " bIns=\"" + int_to_str(bp.b_ins) + "\""
  }
  if bp.auto_fit {
    out + "><a:spAutoFit/></a:bodyPr>"
  } else if bp.font_scale >= 0 {
    let mut naf = "<a:normAutofit"
    if bp.font_scale != @ooxml.pct_100 {
      naf = naf + " fontScale=\"" + int_to_str(bp.font_scale) + "\""
    }
    if bp.ln_spc_reduction > 0 {
      naf = naf + " lnSpcReduction=\"" + int_to_str(bp.ln_spc_reduction) + "\""
    }
    naf = naf + "/>"
    out + ">" + naf + "</a:bodyPr>"
  } else {
    out + "/>"
  }
}

///| Serialize spacing as a:spcBef or a:spcAft child element.
fn serialize_spacing(tag : String, val : Int) -> String {
  if val == 0 {
    return ""
  }
  let mut out = "<" + tag + ">"
  if val < 0 {
    // Negative = percentage (stored as negated value)
    out = out + "<a:spcPct val=\"" + int_to_str(-val) + "\"/>"
  } else {
    // Positive = EMU → convert back to hundredths-of-a-point
    let hpt = val / @ooxml.emu_per_hundredth_pt
    out = out + "<a:spcPts val=\"" + int_to_str(hpt) + "\"/>"
  }
  out + "</" + tag + ">"
}

///| Serialize paragraphs as OOXML text body content (shared by shapes and table cells).
///  `wrapper` is the outer tag: "p:txBody" for shapes, "a:txBody" for table cells.
fn serialize_text_body_inner(
  paragraphs : Array[@ooxml.TextParagraph],
  wrapper : String,
  body_props : @ooxml.BodyProps
) -> String {
  let mut out = "<" + wrapper + ">"
  out = out + serialize_body_pr(body_props) + "<a:lstStyle/>"
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    out = out + "<a:p>"
    // Paragraph properties
    let has_align = not(para.align.is_empty())
    let has_level = para.level > 0
    let has_mar_l = para.mar_l > 0
    let has_indent = para.indent != 0
    let has_spc = para.spc_before != 0 || para.spc_after != 0
    let has_lnspc = para.line_spacing != 0
    let has_bullet = not(para.bullet.is_empty()) || not(para.bullet_auto.is_empty()) || para.bullet_none ||
      not(para.bullet_font.is_empty()) || para.bullet_size != 0 || not(para.bullet_color.is_none()) ||
      not(para.bullet_img_rid.is_empty())
    let has_tabs = para.tab_stops.length() > 0
    let has_rtl = para.rtl
    let has_ppr = has_align || has_level || has_mar_l || has_indent || has_spc || has_lnspc || has_bullet || has_tabs || has_rtl
    if has_ppr {
      let mut ppr = "<a:pPr"
      if has_align {
        ppr = ppr + " algn=\"" + para.align + "\""
      }
      if has_level {
        ppr = ppr + " lvl=\"" + int_to_str(para.level) + "\""
      }
      if has_mar_l {
        ppr = ppr + " marL=\"" + int_to_str(para.mar_l) + "\""
      }
      if has_indent {
        ppr = ppr + " indent=\"" + int_to_str(para.indent) + "\""
      }
      if has_rtl {
        ppr = ppr + " rtl=\"1\""
      }
      let has_children = has_spc || has_lnspc || has_bullet || has_tabs
      if has_children {
        ppr = ppr + ">"
        if has_lnspc {
          ppr = ppr + serialize_spacing("a:lnSpc", para.line_spacing)
        }
        ppr = ppr + serialize_spacing("a:spcBef", para.spc_before)
        ppr = ppr + serialize_spacing("a:spcAft", para.spc_after)
        // Bullet formatting
        if not(para.bullet_font.is_empty()) {
          ppr = ppr + "<a:buFont typeface=\"" + xml_escape(para.bullet_font) + "\"/>"
        }
        if para.bullet_size > 0 {
          ppr = ppr + "<a:buSzPct val=\"" + int_to_str(para.bullet_size) + "\"/>"
        } else if para.bullet_size < 0 {
          ppr = ppr + "<a:buSzPts val=\"" + int_to_str(-(para.bullet_size)) + "\"/>"
        }
        if not(para.bullet_color.is_none()) {
          ppr = ppr + "<a:buClr>" + serialize_color_inner(para.bullet_color) + "</a:buClr>"
        }
        if para.bullet_none {
          ppr = ppr + "<a:buNone/>"
        } else if not(para.bullet.is_empty()) {
          ppr = ppr + "<a:buChar char=\"" + xml_escape(para.bullet) + "\"/>"
        } else if not(para.bullet_auto.is_empty()) {
          ppr = ppr + "<a:buAutoNum type=\"" + para.bullet_auto + "\"/>"
        } else if not(para.bullet_img_rid.is_empty()) {
          ppr = ppr + "<a:buBlip><a:blip r:embed=\"" + para.bullet_img_rid + "\"/></a:buBlip>"
        }
        // Tab stops
        if has_tabs {
          ppr = ppr + "<a:tabLst>"
          let nt = para.tab_stops.length()
          let mut ti = 0
          while ti < nt {
            let ts = para.tab_stops[ti]
            ppr = ppr + "<a:tab pos=\"" + int_to_str(ts.pos) + "\""
            if not(ts.algn.is_empty()) {
              ppr = ppr + " algn=\"" + ts.algn + "\""
            }
            ppr = ppr + "/>"
            ti = ti + 1
          }
          ppr = ppr + "</a:tabLst>"
        }
        out = out + ppr + "</a:pPr>"
      } else {
        out = out + ppr + "/>"
      }
    }
    // Runs
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      out = out + "<a:r>"
      out = out + serialize_run_props(run)
      out = out + "<a:t>" + xml_escape(run.text) + "</a:t>"
      out = out + "</a:r>"
      ri = ri + 1
    }
    // If no runs, emit an empty run to keep the paragraph
    if nr == 0 {
      out = out + "<a:endParaRPr lang=\"en-US\"/>"
    }
    out = out + "</a:p>"
    pi = pi + 1
  }
  out + "</" + wrapper + ">"
}

///| Serialize a:rPr element for a text run.
fn serialize_run_props(run : @ooxml.TextRun) -> String {
  let mut rpr = "<a:rPr lang=\"en-US\" dirty=\"0\""
  if run.bold {
    rpr = rpr + " b=\"1\""
  }
  if run.italic {
    rpr = rpr + " i=\"1\""
  }
  if run.font_size > 0 {
    rpr = rpr + " sz=\"" + int_to_str(run.font_size) + "\""
  }
  if not(run.underline.is_empty()) {
    rpr = rpr + " u=\"" + run.underline + "\""
  }
  if not(run.strike.is_empty()) {
    rpr = rpr + " strike=\"" + run.strike + "\""
  }
  if run.baseline != 0 {
    rpr = rpr + " baseline=\"" + int_to_str(run.baseline) + "\""
  }
  if run.char_spacing != 0 {
    rpr = rpr + " spc=\"" + int_to_str(run.char_spacing) + "\""
  }
  if not(run.cap.is_empty()) {
    rpr = rpr + " cap=\"" + run.cap + "\""
  }
  if run.kern > 0 {
    rpr = rpr + " kern=\"" + int_to_str(run.kern) + "\""
  }
  let has_children = not(run.color.is_none()) || not(run.font_face.is_empty()) || not(run.ea_font.is_empty()) || not(run.cs_font.is_empty()) || not(run.sym_font.is_empty()) || not(run.hlink_rid.is_empty()) || not(run.hlink_mouse_over_rid.is_empty())
  if has_children {
    rpr = rpr + ">"
    if not(run.color.is_none()) {
      rpr = rpr + serialize_color(run.color)
    }
    if not(run.font_face.is_empty()) {
      rpr = rpr + "<a:latin typeface=\"" + xml_escape(run.font_face) + "\"/>"
    }
    if not(run.ea_font.is_empty()) {
      rpr = rpr + "<a:ea typeface=\"" + xml_escape(run.ea_font) + "\"/>"
    }
    if not(run.cs_font.is_empty()) {
      rpr = rpr + "<a:cs typeface=\"" + xml_escape(run.cs_font) + "\"/>"
    }
    if not(run.sym_font.is_empty()) {
      rpr = rpr + "<a:sym typeface=\"" + xml_escape(run.sym_font) + "\"/>"
    }
    if not(run.hlink_rid.is_empty()) {
      rpr = rpr + "<a:hlinkClick r:id=\"" + run.hlink_rid + "\"/>"
    }
    if not(run.hlink_mouse_over_rid.is_empty()) {
      rpr = rpr + "<a:hlinkHover r:id=\"" + run.hlink_mouse_over_rid + "\"/>"
    }
    rpr + "</a:rPr>"
  } else {
    rpr + "/>"
  }
}

fn serialize_text_body(paragraphs : Array[@ooxml.TextParagraph], body_props : @ooxml.BodyProps) -> String {
  if paragraphs.length() == 0 {
    return ""
  }
  serialize_text_body_inner(paragraphs, "p:txBody", body_props)
}

// ── Shape serialization ──────────────────────────────────────────────────────

fn serialize_transform(t : @ooxml.ShapeTransform) -> String {
  let mut xfrm = "<a:xfrm"
  if t.rot != 0 {
    xfrm = xfrm + " rot=\"" + int_to_str(t.rot) + "\""
  }
  if t.flip_h {
    xfrm = xfrm + " flipH=\"1\""
  }
  if t.flip_v {
    xfrm = xfrm + " flipV=\"1\""
  }
  xfrm = xfrm + ">"
  xfrm = xfrm + "<a:off x=\"" + int_to_str(t.x) + "\" y=\"" + int_to_str(t.y) + "\"/>"
  xfrm = xfrm + "<a:ext cx=\"" + int_to_str(t.cx) + "\" cy=\"" + int_to_str(t.cy) + "\"/>"
  xfrm + "</a:xfrm>"
}

fn serialize_autoshape(shape : @ooxml.Shape, geom : @ooxml.ShapeGeom) -> String {
  let mut out = "<p:sp>"
  // Non-visual properties
  out = out + "<p:nvSpPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvSpPr/>"
  if not(shape.ph_type.is_empty()) {
    let mut ph = "<p:nvPr><p:ph type=\"" + shape.ph_type + "\""
    if shape.ph_idx >= 0 {
      ph = ph + " idx=\"" + int_to_str(shape.ph_idx) + "\""
    }
    out = out + ph + "/></p:nvPr>"
  } else {
    out = out + "<p:nvPr/>"
  }
  out = out + "</p:nvSpPr>"
  // Shape properties
  out = out + "<p:spPr>"
  out = out + serialize_transform(shape.transform)
  out = out + "<a:prstGeom prst=\"" + geom.to_prst() + "\"><a:avLst/></a:prstGeom>"
  // Fill (priority: blipFill > pattFill > gradFill > solidFill)
  if not(shape.blip_fill.is_none()) {
    out = out + serialize_blip_fill(shape.blip_fill)
  } else if not(shape.patt_fill.is_none()) {
    out = out + serialize_pattern_fill(shape.patt_fill)
  } else if not(shape.grad_fill.is_none()) {
    out = out + serialize_gradient_fill(shape.grad_fill)
  } else if shape.fill.is_none() {
    out = out + "<a:noFill/>"
  } else {
    out = out + serialize_color(shape.fill)
  }
  // Stroke
  let has_stroke = not(shape.stroke.is_none()) || shape.stroke_no_fill ||
    not(shape.stroke_dash.is_empty()) || not(shape.stroke_head_type.is_empty()) ||
    not(shape.stroke_tail_type.is_empty())
  if has_stroke {
    let w = if shape.stroke_w > 0 { shape.stroke_w } else { @ooxml.default_stroke_w }
    let mut ln_attrs = " w=\"" + int_to_str(w) + "\""
    if not(shape.stroke_cap.is_empty()) {
      ln_attrs = ln_attrs + " cap=\"" + shape.stroke_cap + "\""
    }
    if not(shape.stroke_cmpd.is_empty()) {
      ln_attrs = ln_attrs + " cmpd=\"" + shape.stroke_cmpd + "\""
    }
    out = out + "<a:ln" + ln_attrs + ">"
    // Fill or noFill
    if shape.stroke_no_fill {
      out = out + "<a:noFill/>"
    } else if not(shape.stroke.is_none()) {
      out = out + serialize_color(shape.stroke)
    }
    // Preset dash
    if not(shape.stroke_dash.is_empty()) {
      out = out + "<a:prstDash val=\"" + shape.stroke_dash + "\"/>"
    }
    // Line join
    if shape.stroke_join == "round" {
      out = out + "<a:round/>"
    } else if shape.stroke_join == "bevel" {
      out = out + "<a:bevel/>"
    } else if shape.stroke_join == "miter" {
      if shape.stroke_miter_limit > 0 {
        out = out + "<a:miter lim=\"" + int_to_str(shape.stroke_miter_limit) + "\"/>"
      } else {
        out = out + "<a:miter/>"
      }
    }
    // Arrow head
    if not(shape.stroke_head_type.is_empty()) {
      let mut he_attrs = " type=\"" + shape.stroke_head_type + "\""
      if not(shape.stroke_head_w.is_empty()) {
        he_attrs = he_attrs + " w=\"" + shape.stroke_head_w + "\""
      }
      if not(shape.stroke_head_len.is_empty()) {
        he_attrs = he_attrs + " len=\"" + shape.stroke_head_len + "\""
      }
      out = out + "<a:headEnd" + he_attrs + "/>"
    }
    // Arrow tail
    if not(shape.stroke_tail_type.is_empty()) {
      let mut te_attrs = " type=\"" + shape.stroke_tail_type + "\""
      if not(shape.stroke_tail_w.is_empty()) {
        te_attrs = te_attrs + " w=\"" + shape.stroke_tail_w + "\""
      }
      if not(shape.stroke_tail_len.is_empty()) {
        te_attrs = te_attrs + " len=\"" + shape.stroke_tail_len + "\""
      }
      out = out + "<a:tailEnd" + te_attrs + "/>"
    }
    out = out + "</a:ln>"
  }
  out = out + "</p:spPr>"
  // Text body
  out = out + serialize_text_body(shape.paragraphs, shape.body_props)
  out + "</p:sp>"
}

fn serialize_picture(shape : @ooxml.Shape, rid : String) -> String {
  let mut out = "<p:pic>"
  // Non-visual properties
  out = out + "<p:nvPicPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvPicPr>"
  // Blip fill
  out = out + "<p:blipFill>"
  out = out + "<a:blip r:embed=\"" + rid + "\"/>"
  out = out + "<a:stretch><a:fillRect/></a:stretch>"
  out = out + "</p:blipFill>"
  // Shape properties
  out = out + "<p:spPr>"
  out = out + serialize_transform(shape.transform)
  out = out + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
  out = out + "</p:spPr>"
  out + "</p:pic>"
}

fn serialize_table_cell(cell : @ooxml.TableCell) -> String {
  let mut out = "<a:tc>"
  // Text body
  if cell.paragraphs.length() > 0 {
    out = out + serialize_table_text_body(cell.paragraphs)
  } else {
    out = out + "<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></a:txBody>"
  }
  // Cell properties
  out = out + "<a:tcPr>"
  if not(cell.grad_fill.is_none()) {
    out = out + serialize_gradient_fill(cell.grad_fill)
  } else if not(cell.fill.is_none()) {
    out = out + serialize_color(cell.fill)
  }
  out = out + "</a:tcPr>"
  out + "</a:tc>"
}

fn serialize_table_text_body(paragraphs : Array[@ooxml.TextParagraph]) -> String {
  serialize_text_body_inner(paragraphs, "a:txBody", @ooxml.BodyProps::default())
}

fn serialize_table(shape : @ooxml.Shape, table : @ooxml.TableData) -> String {
  let mut out = "<p:graphicFrame>"
  // Non-visual properties
  out = out + "<p:nvGraphicFramePr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGraphicFramePr>"
  // Transform (directly under graphicFrame)
  out = out + serialize_transform(shape.transform)
  // Table content
  out = out + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">"
  out = out + "<a:tbl>"
  out = out + "<a:tblPr firstRow=\"1\" bandRow=\"1\"/>"
  // Grid columns
  out = out + "<a:tblGrid>"
  let ncols = table.col_widths.length()
  let mut ci = 0
  while ci < ncols {
    out = out + "<a:gridCol w=\"" + int_to_str(table.col_widths[ci]) + "\"/>"
    ci = ci + 1
  }
  out = out + "</a:tblGrid>"
  // Rows
  let nrows = table.rows.length()
  let mut ri = 0
  while ri < nrows {
    let row = table.rows[ri]
    out = out + "<a:tr h=\"" + int_to_str(row.height) + "\">"
    let ncells = row.cells.length()
    let mut cj = 0
    while cj < ncells {
      out = out + serialize_table_cell(row.cells[cj])
      cj = cj + 1
    }
    out = out + "</a:tr>"
    ri = ri + 1
  }
  out = out + "</a:tbl>"
  out = out + "</a:graphicData></a:graphic>"
  out + "</p:graphicFrame>"
}

fn serialize_shape(shape : @ooxml.Shape) -> String {
  match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => serialize_autoshape(shape, geom)
    @ooxml.ShapeKind::Picture(rid) => serialize_picture(shape, rid)
    @ooxml.ShapeKind::TableShape(table) => serialize_table(shape, table)
    _ => ""
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

///| Serialize a SlideData into a complete OOXML slide XML string.
pub fn serialize_slide(slide : @ooxml.SlideData) -> String {
  let mut out = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  out = out + "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
  out = out + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
  out = out + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
  out = out + "<p:cSld>"
  // Background
  if not(slide.bg_grad.is_none()) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_gradient_fill(slide.bg_grad)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  } else if not(slide.background.is_none()) && not(
    slide.background.r == 255 && slide.background.g == 255 && slide.background.b == 255
  ) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_color(slide.background)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  }
  // Shape tree
  out = out + "<p:spTree>"
  out = out + "<p:nvGrpSpPr>"
  out = out + "<p:cNvPr id=\"1\" name=\"\"/>"
  out = out + "<p:cNvGrpSpPr/>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGrpSpPr>"
  out = out + "<p:grpSpPr>"
  out = out + "<a:xfrm>"
  out = out + "<a:off x=\"0\" y=\"0\"/>"
  out = out + "<a:ext cx=\"0\" cy=\"0\"/>"
  out = out + "<a:chOff x=\"0\" y=\"0\"/>"
  out = out + "<a:chExt cx=\"0\" cy=\"0\"/>"
  out = out + "</a:xfrm>"
  out = out + "</p:grpSpPr>"
  // Shapes
  let n = slide.shapes.length()
  let mut i = 0
  while i < n {
    out = out + serialize_shape(slide.shapes[i])
    i = i + 1
  }
  out = out + "</p:spTree>"
  out = out + "</p:cSld>"
  out = out + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
  out + "</p:sld>"
}
