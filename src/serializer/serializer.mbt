// OOXML slide XML serializer — SlideData → OOXML XML
//
// Generates valid pptx slide XML from the internal SlideData model.
// Used for the round-trip pipeline: PPTX → SlideData → SVG → SlideData → PPTX.
//
// NOTE: No StringBuilder — all strings built via `+` (wasm:js-string "concat").

// ── Non-visual property helpers ──────────────────────────────────────────────

///|
/// Serialize p:cNvPr with optional hyperlinks.
fn serialize_cnvpr(id : String, link_rid : String, hover_rid : String) -> String {
  if link_rid.is_empty() && hover_rid.is_empty() {
    "<p:cNvPr id=\"" + id + "\" name=\"\"/>"
  } else {
    let mut inner = ""
    if not(link_rid.is_empty()) {
      inner = inner + "<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"" + link_rid + "\"/>"
    }
    if not(hover_rid.is_empty()) {
      inner = inner + "<a:hlinkHover xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"" + hover_rid + "\"/>"
    }
    "<p:cNvPr id=\"" + id + "\" name=\"\">" + inner + "</p:cNvPr>"
  }
}

// ── Shared helpers (from @xml) ───────────────────────────────────────────────

///|
fn int_to_str(n : Int) -> String {
  @xml.int_to_str(n)
}

///|
/// Serialize an arrow end element (a:headEnd or a:tailEnd).
fn serialize_arrow_end(
  tag : String,
  type_ : String,
  w : String,
  len : String,
) -> String {
  if type_.is_empty() {
    return ""
  }
  let mut attrs = " type=\"" + type_ + "\""
  if not(w.is_empty()) {
    attrs = attrs + " w=\"" + w + "\""
  }
  if not(len.is_empty()) {
    attrs = attrs + " len=\"" + len + "\""
  }
  "<" + tag + attrs + "/>"
}

// ── XML helpers ──────────────────────────────────────────────────────────────

///|
fn xml_escape(s : String) -> String {
  @xml.xml_escape(s)
}

// ── Color serialization ─────────────────────────────────────────────────────

///|
/// Serialize inner color element (a:srgbClr with optional a:alpha child).
///  No solidFill wrapper — used for gradient stops, pattFill fg/bg, etc.
fn serialize_color_inner(color : @ooxml.Color) -> String {
  if color.alpha < 255 {
    // alpha → OOXML 1/1000th percent: alpha * pct_100 / 255
    let alpha_val = color.alpha * @ooxml.pct_100 / 255
    "<a:srgbClr val=\"" +
    color.to_hex_rgb() +
    "\"><a:alpha val=\"" +
    int_to_str(alpha_val) +
    "\"/></a:srgbClr>"
  } else {
    "<a:srgbClr val=\"" + color.to_hex_rgb() + "\"/>"
  }
}

///|
fn serialize_color(color : @ooxml.Color) -> String {
  if color.is_none() {
    "<a:noFill/>"
  } else {
    "<a:solidFill>" + serialize_color_inner(color) + "</a:solidFill>"
  }
}

// ── Gradient fill serialization ──────────────────────────────────────────────

///|
fn serialize_gradient_fill(gf : @ooxml.GradientFill) -> String {
  let rot_attr = if gf.rot_with_shape {
    " rotWithShape=\"1\""
  } else {
    " rotWithShape=\"0\""
  }
  let flip_attr = if gf.tile_flip.is_empty() || gf.tile_flip == "none" {
    ""
  } else {
    " tileFlip=\"" + gf.tile_flip + "\""
  }
  let mut out = "<a:gradFill" + rot_attr + flip_attr + ">"
  // Gradient stops
  out = out + "<a:gsLst>"
  let ns = gf.stops.length()
  let mut si = 0
  while si < ns {
    let stop = gf.stops[si]
    out = out +
      "<a:gs pos=\"" +
      int_to_str(stop.pos) +
      "\">" +
      serialize_color_inner(stop.color) +
      "</a:gs>"
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
    if gf.fill_to_l >= 0 ||
      gf.fill_to_t >= 0 ||
      gf.fill_to_r >= 0 ||
      gf.fill_to_b >= 0 {
      out = out + "<a:fillToRect"
      if gf.fill_to_l >= 0 {
        out = out + " l=\"" + int_to_str(gf.fill_to_l) + "\""
      }
      if gf.fill_to_t >= 0 {
        out = out + " t=\"" + int_to_str(gf.fill_to_t) + "\""
      }
      if gf.fill_to_r >= 0 {
        out = out + " r=\"" + int_to_str(gf.fill_to_r) + "\""
      }
      if gf.fill_to_b >= 0 {
        out = out + " b=\"" + int_to_str(gf.fill_to_b) + "\""
      }
      out = out + "/>"
    }
    out = out + "</a:path>"
  }
  out + "</a:gradFill>"
}

// ── Blip fill serialization ──────────────────────────────────────────────────

///|
fn serialize_blip_fill(bf : @ooxml.BlipFill) -> String {
  let mut out = "<a:blipFill>"
  // Build blip children (effects + SVG extension)
  let mut blip_inner = ""
  if bf.alpha > 0 && bf.alpha < @ooxml.pct_100 {
    blip_inner = blip_inner + "<a:alphaModFix amt=\"" + int_to_str(bf.alpha) + "\"/>"
  }
  if bf.bright != 0 || bf.contrast != 0 {
    blip_inner = blip_inner + "<a:lum"
    if bf.bright != 0 {
      blip_inner = blip_inner + " bright=\"" + int_to_str(bf.bright) + "\""
    }
    if bf.contrast != 0 {
      blip_inner = blip_inner + " contrast=\"" + int_to_str(bf.contrast) + "\""
    }
    blip_inner = blip_inner + "/>"
  }
  if not(bf.duotone_1.is_none()) && not(bf.duotone_2.is_none()) {
    blip_inner = blip_inner + "<a:duotone>" +
      "<a:srgbClr val=\"" + bf.duotone_1.to_hex_rgb() + "\"/>" +
      "<a:srgbClr val=\"" + bf.duotone_2.to_hex_rgb() + "\"/>" +
      "</a:duotone>"
  }
  if not(bf.clr_from.is_none()) && not(bf.clr_to.is_none()) {
    blip_inner = blip_inner + "<a:clrChange>" +
      "<a:clrFrom><a:srgbClr val=\"" + bf.clr_from.to_hex_rgb() + "\"/></a:clrFrom>" +
      "<a:clrTo><a:srgbClr val=\"" + bf.clr_to.to_hex_rgb() + "\"/></a:clrTo>" +
      "</a:clrChange>"
  }
  if not(bf.svg_rid.is_empty()) {
    blip_inner = blip_inner +
      "<a:extLst><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">" +
      "<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"" +
      bf.svg_rid + "\"/></a:ext></a:extLst>"
  }
  if blip_inner.is_empty() {
    out = out + "<a:blip r:embed=\"" + bf.rid + "\"/>"
  } else {
    out = out + "<a:blip r:embed=\"" + bf.rid + "\">" + blip_inner + "</a:blip>"
  }
  if bf.has_crop() {
    out = out + "<a:srcRect"
    if bf.src_l > 0 {
      out = out + " l=\"" + int_to_str(bf.src_l) + "\""
    }
    if bf.src_t > 0 {
      out = out + " t=\"" + int_to_str(bf.src_t) + "\""
    }
    if bf.src_r > 0 {
      out = out + " r=\"" + int_to_str(bf.src_r) + "\""
    }
    if bf.src_b > 0 {
      out = out + " b=\"" + int_to_str(bf.src_b) + "\""
    }
    out = out + "/>"
  }
  if bf.stretch {
    out = out + "<a:stretch><a:fillRect/></a:stretch>"
  } else if bf.tile_sx > 0 || bf.tile_sy > 0 {
    out = out + "<a:tile"
    if bf.tile_tx != 0 {
      out = out + " tx=\"" + int_to_str(bf.tile_tx) + "\""
    }
    if bf.tile_ty != 0 {
      out = out + " ty=\"" + int_to_str(bf.tile_ty) + "\""
    }
    if bf.tile_sx != 0 {
      out = out + " sx=\"" + int_to_str(bf.tile_sx) + "\""
    }
    if bf.tile_sy != 0 {
      out = out + " sy=\"" + int_to_str(bf.tile_sy) + "\""
    }
    if not(bf.tile_flip.is_empty()) && bf.tile_flip != "none" {
      out = out + " flip=\"" + bf.tile_flip + "\""
    }
    if not(bf.tile_algn.is_empty()) {
      out = out + " algn=\"" + bf.tile_algn + "\""
    }
    out = out + "/>"
  }
  out + "</a:blipFill>"
}

// ── Pattern fill serialization ──────────────────────────────────────────────

///|
fn serialize_pattern_fill(pf : @ooxml.PatternFill) -> String {
  let mut out = "<a:pattFill prst=\"" + pf.prst + "\">"
  out = out + "<a:fgClr>" + serialize_color_inner(pf.fg_color) + "</a:fgClr>"
  out = out + "<a:bgClr>" + serialize_color_inner(pf.bg_color) + "</a:bgClr>"
  out + "</a:pattFill>"
}

// ── Text body serialization ──────────────────────────────────────────────────

///|
/// Serialize a:bodyPr element.
fn serialize_body_pr(bp : @ooxml.BodyProps) -> String {
  let has_attrs = not(bp.anchor.is_empty()) ||
    bp.l_ins >= 0 ||
    bp.t_ins >= 0 ||
    bp.r_ins >= 0 ||
    bp.b_ins >= 0 ||
    bp.wrap == "none" ||
    bp.rot != 0 ||
    not(bp.vert.is_empty()) ||
    bp.num_cols > 1 ||
    bp.col_spacing > 0
  let has_warp = not(bp.warp_prst.is_empty())
  let has_children = bp.auto_fit || bp.font_scale >= 0 || has_warp
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
  if not(has_children) {
    return out + "/>"
  }
  out = out + ">"
  // a:prstTxWarp must come before a:normAutofit/a:spAutoFit per OOXML schema order
  if has_warp {
    let mut warp = "<a:prstTxWarp prst=\"" + bp.warp_prst + "\">"
    if bp.warp_av1 >= 0 || bp.warp_av2 >= 0 {
      warp = warp + "<a:avLst>"
      if bp.warp_av1 >= 0 {
        warp = warp +
          "<a:gd name=\"adj\" fmla=\"val " +
          int_to_str(bp.warp_av1) +
          "\"/>"
      }
      if bp.warp_av2 >= 0 {
        warp = warp +
          "<a:gd name=\"adj2\" fmla=\"val " +
          int_to_str(bp.warp_av2) +
          "\"/>"
      }
      warp = warp + "</a:avLst>"
    } else {
      warp = warp + "<a:avLst/>"
    }
    out = out + warp + "</a:prstTxWarp>"
  }
  if bp.auto_fit {
    out = out + "<a:spAutoFit/>"
  } else if bp.font_scale >= 0 {
    let mut naf = "<a:normAutofit"
    if bp.font_scale != @ooxml.pct_100 {
      naf = naf + " fontScale=\"" + int_to_str(bp.font_scale) + "\""
    }
    if bp.ln_spc_reduction > 0 {
      naf = naf + " lnSpcReduction=\"" + int_to_str(bp.ln_spc_reduction) + "\""
    }
    out = out + naf + "/>"
  }
  out + "</a:bodyPr>"
}

///|
/// Serialize spacing as a:spcBef or a:spcAft child element.
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

///|
/// Serialize paragraphs as OOXML text body content (shared by shapes and table cells).
///  `wrapper` is the outer tag: "p:txBody" for shapes, "a:txBody" for table cells.
fn serialize_text_body_inner(
  paragraphs : Array[@ooxml.TextParagraph],
  wrapper : String,
  body_props : @ooxml.BodyProps,
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
    let has_mar_l = para.mar_l >= 0
    let has_indent = para.indent != 0
    let has_spc = para.spc_before != 0 || para.spc_after != 0
    let has_lnspc = para.line_spacing != 0
    let has_bullet = not(para.bullet.is_empty()) ||
      not(para.bullet_auto.is_empty()) ||
      para.bullet_none ||
      not(para.bullet_font.is_empty()) ||
      para.bullet_size != 0 ||
      not(para.bullet_color.is_none()) ||
      not(para.bullet_img_rid.is_empty())
    let has_tabs = para.tab_stops.length() > 0
    let has_rtl = para.rtl
    let has_ppr = has_align ||
      has_level ||
      has_mar_l ||
      has_indent ||
      has_spc ||
      has_lnspc ||
      has_bullet ||
      has_tabs ||
      has_rtl
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
          ppr = ppr +
            "<a:buFont typeface=\"" +
            xml_escape(para.bullet_font) +
            "\"/>"
        }
        if para.bullet_size > 0 {
          ppr = ppr +
            "<a:buSzPct val=\"" +
            int_to_str(para.bullet_size) +
            "\"/>"
        } else if para.bullet_size < 0 {
          ppr = ppr +
            "<a:buSzPts val=\"" +
            int_to_str(-para.bullet_size) +
            "\"/>"
        }
        if not(para.bullet_color.is_none()) {
          ppr = ppr +
            "<a:buClr>" +
            serialize_color_inner(para.bullet_color) +
            "</a:buClr>"
        }
        if para.bullet_none {
          ppr = ppr + "<a:buNone/>"
        } else if not(para.bullet.is_empty()) {
          ppr = ppr + "<a:buChar char=\"" + xml_escape(para.bullet) + "\"/>"
        } else if not(para.bullet_auto.is_empty()) {
          ppr = ppr + "<a:buAutoNum type=\"" + para.bullet_auto + "\"/>"
        } else if not(para.bullet_img_rid.is_empty()) {
          ppr = ppr +
            "<a:buBlip><a:blip r:embed=\"" +
            para.bullet_img_rid +
            "\"/></a:buBlip>"
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

///|
/// Serialize a:rPr element for a text run.
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
  let has_children = not(run.color.is_none()) ||
    not(run.font_face.is_empty()) ||
    not(run.ea_font.is_empty()) ||
    not(run.cs_font.is_empty()) ||
    not(run.sym_font.is_empty()) ||
    not(run.hlink_rid.is_empty()) ||
    not(run.hlink_mouse_over_rid.is_empty()) ||
    not(run.outline_color.is_none()) ||
    not(run.text_grad_fill.is_none()) ||
    not(run.text_patt_fill.is_none())
  if has_children {
    rpr = rpr + ">"
    // Text outline: a:ln
    if not(run.outline_color.is_none()) {
      let ow = if run.outline_w > 0 { run.outline_w } else { @ooxml.default_outline_w }
      rpr = rpr + "<a:ln w=\"" + int_to_str(ow) + "\">"
      rpr = rpr + serialize_color(run.outline_color)
      rpr = rpr + "</a:ln>"
    }
    if not(run.color.is_none()) {
      rpr = rpr + serialize_color(run.color)
    }
    // Text gradient fill: a:gradFill on a:rPr
    if not(run.text_grad_fill.is_none()) {
      rpr = rpr + serialize_gradient_fill(run.text_grad_fill)
    }
    // Text pattern fill: a:pattFill on a:rPr
    if not(run.text_patt_fill.is_none()) {
      rpr = rpr + serialize_pattern_fill(run.text_patt_fill)
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
    if not(run.effects.is_none()) {
      rpr = rpr + serialize_effect_lst(run.effects)
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

///|
fn serialize_text_body(
  paragraphs : Array[@ooxml.TextParagraph],
  body_props : @ooxml.BodyProps,
) -> String {
  if paragraphs.length() == 0 {
    return ""
  }
  serialize_text_body_inner(paragraphs, "p:txBody", body_props)
}

// ── Shape serialization ──────────────────────────────────────────────────────

///|
/// Serialize a:xfrm with optional extra content (e.g. chOff/chExt for groups).
fn serialize_transform_with_extra(
  t : @ooxml.ShapeTransform,
  extra : String,
) -> String {
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
  xfrm = xfrm +
    "<a:off x=\"" +
    int_to_str(t.x) +
    "\" y=\"" +
    int_to_str(t.y) +
    "\"/>"
  xfrm = xfrm +
    "<a:ext cx=\"" +
    int_to_str(t.cx) +
    "\" cy=\"" +
    int_to_str(t.cy) +
    "\"/>"
  xfrm + extra + "</a:xfrm>"
}

///|
fn serialize_transform(t : @ooxml.ShapeTransform) -> String {
  serialize_transform_with_extra(t, "")
}

///|
fn serialize_prst_geom(geom : @ooxml.ShapeGeom) -> String {
  match geom {
    @ooxml.ShapeGeom::Custom(data) => serialize_custom_geom(data)
    _ => {
      let prst = geom.to_prst()
      let avs = match geom {
        @ooxml.ShapeGeom::Connector(_, a) => a
        @ooxml.ShapeGeom::Other(_, a) => a
        _ => []
      }
      if avs.length() == 0 {
        "<a:prstGeom prst=\"" + prst + "\"><a:avLst/></a:prstGeom>"
      } else {
        let mut out = "<a:prstGeom prst=\"" + prst + "\"><a:avLst>"
        let n = avs.length()
        let mut i = 0
        while i < n {
          let name = if i == 0 { "adj" } else { "adj" + int_to_str(i + 1) }
          out = out +
            "<a:gd name=\"" +
            name +
            "\" fmla=\"val " +
            int_to_str(avs[i]) +
            "\"/>"
          i = i + 1
        }
        out + "</a:avLst></a:prstGeom>"
      }
    }
  }
}

///|
/// Serialize custom geometry (a:custGeom) back to OOXML XML.
fn serialize_custom_geom(data : @ooxml.CustomGeomData) -> String {
  let mut out = "<a:custGeom>"
  // Serialize guide list → split into avLst and gdLst
  let mut avlst = "<a:avLst/>"
  let mut gdlst_out = ""
  if data.gdlst.length() > 0 {
    let mut av_entries = ""
    let mut gd_entries = ""
    // Split by ";" and separate avLst (val formulas) from gdLst (others)
    let len = data.gdlst.length()
    let mut start = 0
    let mut i = 0
    while i <= len {
      if i == len || data.gdlst.get_char(i).unwrap() == ';' {
        if i > start {
          let entry = @xml.str_substring(data.gdlst, start, i)
          // Find "=" separator
          let elen = entry.length()
          let mut eq = 0
          while eq < elen {
            if entry.get_char(eq).unwrap() == '=' {
              break
            }
            eq = eq + 1
          }
          if eq < elen {
            let name = @xml.str_substring(entry, 0, eq)
            let fmla = @xml.str_substring(entry, eq + 1, elen)
            if @xml.str_starts_with(fmla, "val ") {
              av_entries = av_entries +
                "<a:gd name=\"" +
                name +
                "\" fmla=\"" +
                fmla +
                "\"/>"
            } else {
              gd_entries = gd_entries +
                "<a:gd name=\"" +
                name +
                "\" fmla=\"" +
                fmla +
                "\"/>"
            }
          }
        }
        start = i + 1
      }
      i = i + 1
    }
    if av_entries.length() > 0 {
      avlst = "<a:avLst>" + av_entries + "</a:avLst>"
    }
    if gd_entries.length() > 0 {
      gdlst_out = "<a:gdLst>" + gd_entries + "</a:gdLst>"
    }
  }
  out = out + avlst + gdlst_out
  // Serialize paths
  out = out + "<a:pathLst>"
  // Split paths by "|" for multiple sub-paths
  let plen = data.paths.length()
  let mut pstart = 0
  let mut pi = 0
  while pi <= plen {
    if pi == plen || data.paths.get_char(pi).unwrap() == '|' {
      if pi > pstart {
        let seg = @xml.str_substring(data.paths, pstart, pi)
        out = out + serialize_path_segment(seg, data.path_w, data.path_h)
      }
      pstart = pi + 1
    }
    pi = pi + 1
  }
  out = out + "</a:pathLst>"
  // Serialize a:cxnLst if connection points are defined
  if not(data.cxn_lst.is_empty()) {
    out = out + "<a:cxnLst>"
    // Parse "x,y,ang;x2,y2,ang2;..."
    let clen = data.cxn_lst.length()
    let mut cstart = 0
    let mut ci = 0
    while ci <= clen {
      if ci == clen || data.cxn_lst.get_char(ci).unwrap() == ';' {
        if ci > cstart {
          let entry = @xml.str_substring(data.cxn_lst, cstart, ci)
          // Parse "x,y,ang"
          let mut c1 = 0
          let mut c2 = 0
          let elen = entry.length()
          let mut ei = 0
          let mut comma_count = 0
          while ei < elen {
            if entry.get_char(ei).unwrap() == ',' {
              if comma_count == 0 {
                c1 = ei
              } else {
                c2 = ei
              }
              comma_count = comma_count + 1
            }
            ei = ei + 1
          }
          if comma_count >= 2 {
            let cx = @xml.str_substring(entry, 0, c1)
            let cy = @xml.str_substring(entry, c1 + 1, c2)
            let ang = @xml.str_substring(entry, c2 + 1, elen)
            out = out +
              "<a:cxn ang=\"" +
              ang +
              "\"><a:pos x=\"" +
              cx +
              "\" y=\"" +
              cy +
              "\"/></a:cxn>"
          }
        }
        cstart = ci + 1
      }
      ci = ci + 1
    }
    out = out + "</a:cxnLst>"
  }
  // Serialize a:rect if text rect is defined
  if not(data.rect_l.is_empty()) || not(data.rect_t.is_empty()) ||
    not(data.rect_r.is_empty()) ||
    not(data.rect_b.is_empty()) {
    let rl = if data.rect_l.is_empty() { "l" } else { data.rect_l }
    let rt = if data.rect_t.is_empty() { "t" } else { data.rect_t }
    let rr = if data.rect_r.is_empty() { "r" } else { data.rect_r }
    let rb = if data.rect_b.is_empty() { "b" } else { data.rect_b }
    out = out +
      "<a:rect l=\"" +
      rl +
      "\" t=\"" +
      rt +
      "\" r=\"" +
      rr +
      "\" b=\"" +
      rb +
      "\"/>"
  }
  out + "</a:custGeom>"
}

///|
/// Serialize a single path segment from compact format to OOXML XML.
fn serialize_path_segment(seg : String, path_w : Int, path_h : Int) -> String {
  let mut out = "<a:path"
  if path_w > 0 {
    out = out + " w=\"" + int_to_str(path_w) + "\""
  }
  if path_h > 0 {
    out = out + " h=\"" + int_to_str(path_h) + "\""
  }
  out = out + ">"
  // Parse compact format tokens
  let tokens = @xml.split_by_space(seg)
  let nt = tokens.length()
  let mut i = 0
  while i < nt {
    let cmd = tokens[i]
    if cmd == "M" && i + 2 < nt {
      out = out +
        "<a:moveTo><a:pt x=\"" +
        tokens[i + 1] +
        "\" y=\"" +
        tokens[i + 2] +
        "\"/></a:moveTo>"
      i = i + 3
    } else if cmd == "L" && i + 2 < nt {
      out = out +
        "<a:lnTo><a:pt x=\"" +
        tokens[i + 1] +
        "\" y=\"" +
        tokens[i + 2] +
        "\"/></a:lnTo>"
      i = i + 3
    } else if cmd == "C" && i + 6 < nt {
      out = out + "<a:cubicBezTo>"
      for j in 0..<3 {
        out = out +
          "<a:pt x=\"" +
          tokens[i + 1 + j * 2] +
          "\" y=\"" +
          tokens[i + 2 + j * 2] +
          "\"/>"
      }
      out = out + "</a:cubicBezTo>"
      i = i + 7
    } else if cmd == "Q" && i + 4 < nt {
      out = out + "<a:quadBezTo>"
      for j in 0..<2 {
        out = out +
          "<a:pt x=\"" +
          tokens[i + 1 + j * 2] +
          "\" y=\"" +
          tokens[i + 2 + j * 2] +
          "\"/>"
      }
      out = out + "</a:quadBezTo>"
      i = i + 5
    } else if cmd == "A" && i + 4 < nt {
      out = out +
        "<a:arcTo wR=\"" +
        tokens[i + 1] +
        "\" hR=\"" +
        tokens[i + 2] +
        "\" stAng=\"" +
        tokens[i + 3] +
        "\" swAng=\"" +
        tokens[i + 4] +
        "\"/>"
      i = i + 5
    } else if cmd == "Z" {
      out = out + "<a:close/>"
      i = i + 1
    } else {
      i = i + 1
    }
  }
  out + "</a:path>"
}


///|
fn serialize_autoshape(shape : @ooxml.Shape, geom : @ooxml.ShapeGeom) -> String {
  let mut out = "<p:sp>"
  // Non-visual properties
  out = out + "<p:nvSpPr>"
  out = out + serialize_cnvpr("0", shape.sh_link_rid, shape.sh_link_hover_rid)
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
  out = out + serialize_prst_geom(geom)
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
  let has_stroke = not(shape.stroke.is_none()) ||
    shape.stroke_no_fill ||
    not(shape.stroke_dash.is_empty()) ||
    not(shape.stroke_head_type.is_empty()) ||
    not(shape.stroke_tail_type.is_empty())
  if has_stroke {
    out = out + serialize_stroke(shape)
  }
  out = out + serialize_effect_lst(shape.effects)
  out = out + serialize_scene_3d(shape.scene_3d)
  out = out + serialize_sp_3d(shape.sp_3d)
  out = out + "</p:spPr>"
  // Text body
  out = out + serialize_text_body(shape.paragraphs, shape.body_props)
  out + "</p:sp>"
}

///|
/// Conditionally emit an XML integer attribute (non-zero only).
fn xml_int_attr(name : String, val : Int) -> String {
  if val > 0 { " " + name + "=\"" + int_to_str(val) + "\"" } else { "" }
}

///|
/// Conditionally emit an XML integer attribute when it differs from a default.
fn xml_int_attr_ne(name : String, val : Int, default : Int) -> String {
  if val != default { " " + name + "=\"" + int_to_str(val) + "\"" } else { "" }
}

///|
fn serialize_effect_lst(eff : @ooxml.EffectList) -> String {
  if eff.is_none() {
    return ""
  }
  let mut inner = ""
  if not(eff.outer_shadow.is_none()) {
    let os = eff.outer_shadow
    let attrs = xml_int_attr("blurRad", os.blur_rad) +
      xml_int_attr("dist", os.dist) +
      xml_int_attr("dir", os.dir) +
      xml_int_attr_ne("sx", os.sx, 100000) +
      xml_int_attr_ne("sy", os.sy, 100000) +
      (if not(os.algn.is_empty()) { " algn=\"" + os.algn + "\"" } else { "" }) +
      (if not(os.rot_with_shape) { " rotWithShape=\"0\"" } else { "" })
    inner = inner + "<a:outerShdw" + attrs + ">" + serialize_eff_color(os.color) + "</a:outerShdw>"
  }
  if not(eff.inner_shadow.is_none()) {
    let ishdw = eff.inner_shadow
    let attrs = xml_int_attr("blurRad", ishdw.blur_rad) +
      xml_int_attr("dist", ishdw.dist) +
      xml_int_attr("dir", ishdw.dir)
    inner = inner + "<a:innerShdw" + attrs + ">" + serialize_eff_color(ishdw.color) + "</a:innerShdw>"
  }
  if not(eff.glow.is_none()) {
    inner = inner + "<a:glow" + xml_int_attr("rad", eff.glow.rad) + ">" + serialize_eff_color(eff.glow.color) + "</a:glow>"
  }
  if not(eff.soft_edge.is_none()) {
    inner = inner + "<a:softEdge rad=\"" + int_to_str(eff.soft_edge.rad) + "\"/>"
  }
  if not(eff.reflection.is_none()) {
    let r = eff.reflection
    let attrs = xml_int_attr("blurRad", r.blur_rad) +
      xml_int_attr("dist", r.dist) +
      xml_int_attr("dir", r.dir) +
      xml_int_attr_ne("stA", r.st_alpha, 100000) +
      xml_int_attr("endA", r.end_alpha) +
      xml_int_attr("fadeDir", r.fade_dir) +
      xml_int_attr_ne("sx", r.sx, 100000) +
      xml_int_attr_ne("sy", r.sy, 100000) +
      (if not(r.algn.is_empty()) { " algn=\"" + r.algn + "\"" } else { "" }) +
      (if not(r.rot_with_shape) { " rotWithShape=\"0\"" } else { "" })
    inner = inner + "<a:reflection" + attrs + "/>"
  }
  "<a:effectLst>" + inner + "</a:effectLst>"
}

///|
fn serialize_scene_3d(s : @ooxml.Scene3d) -> String {
  if s.is_none() { return "" }
  let cam = "<a:camera prst=\"" + s.camera_prst + "\"/>"
  let lr = if not(s.light_rig.is_empty()) {
    "<a:lightRig rig=\"" + s.light_rig + "\"" +
      (if not(s.light_dir.is_empty()) { " dir=\"" + s.light_dir + "\"" } else { "" }) +
      "/>"
  } else {
    ""
  }
  "<a:scene3d>" + cam + lr + "</a:scene3d>"
}

///|
fn serialize_bevel(tag : String, b : @ooxml.Bevel) -> String {
  if b.is_none() { return "" }
  let mut attrs = ""
  if b.w > 0 { attrs = attrs + " w=\"" + int_to_str(b.w) + "\"" }
  if b.h > 0 { attrs = attrs + " h=\"" + int_to_str(b.h) + "\"" }
  if not(b.prst.is_empty()) { attrs = attrs + " prst=\"" + b.prst + "\"" }
  "<" + tag + attrs + "/>"
}

///|
fn serialize_sp_3d(s : @ooxml.Shape3d) -> String {
  if s.is_none() { return "" }
  let mut attrs = ""
  if s.extrusion_h > 0 { attrs = attrs + " extrusionH=\"" + int_to_str(s.extrusion_h) + "\"" }
  if s.contour_w > 0 { attrs = attrs + " contourW=\"" + int_to_str(s.contour_w) + "\"" }
  if not(s.prst_material.is_empty()) { attrs = attrs + " prstMaterial=\"" + s.prst_material + "\"" }
  if s.z != 0 { attrs = attrs + " z=\"" + int_to_str(s.z) + "\"" }
  let inner = serialize_bevel("a:bevelT", s.bevel_t) +
    serialize_bevel("a:bevelB", s.bevel_b) +
    (if not(s.extrusion_clr.is_none()) { "<a:extrusionClr>" + serialize_eff_color(s.extrusion_clr) + "</a:extrusionClr>" } else { "" }) +
    (if not(s.contour_clr.is_none()) { "<a:contourClr>" + serialize_eff_color(s.contour_clr) + "</a:contourClr>" } else { "" })
  "<a:sp3d" + attrs + ">" + inner + "</a:sp3d>"
}

///|
fn serialize_eff_color(c : @ooxml.Color) -> String {
  if c.is_none() {
    ""
  } else {
    let hex = c.to_hex()
    if c.alpha < 255 {
      "<a:srgbClr val=\"" + hex + "\"><a:alpha val=\"" + int_to_str(c.alpha * 100000 / 255) + "\"/></a:srgbClr>"
    } else {
      "<a:srgbClr val=\"" + hex + "\"/>"
    }
  }
}

///|
fn serialize_stroke(shape : @ooxml.Shape) -> String {
  let w = if shape.stroke_w > 0 {
    shape.stroke_w
  } else {
    @ooxml.default_stroke_w
  }
  let mut ln_attrs = " w=\"" + int_to_str(w) + "\""
  if not(shape.stroke_cap.is_empty()) {
    ln_attrs = ln_attrs + " cap=\"" + shape.stroke_cap + "\""
  }
  if not(shape.stroke_cmpd.is_empty()) {
    ln_attrs = ln_attrs + " cmpd=\"" + shape.stroke_cmpd + "\""
  }
  let mut out = "<a:ln" + ln_attrs + ">"
  // Fill or noFill
  if shape.stroke_no_fill {
    out = out + "<a:noFill/>"
  } else if not(shape.stroke_grad_fill.is_none()) {
    out = out + serialize_gradient_fill(shape.stroke_grad_fill)
  } else if not(shape.stroke_patt_fill.is_none()) {
    out = out + serialize_pattern_fill(shape.stroke_patt_fill)
  } else if not(shape.stroke.is_none()) {
    out = out + serialize_color(shape.stroke)
  }
  // Preset dash or custom dash
  if @xml.str_starts_with(shape.stroke_dash, @ooxml.cust_dash_prefix) {
    let data = @xml.str_suffix(shape.stroke_dash, @ooxml.cust_dash_prefix_len)
    out = out + "<a:custDash>"
    let data_len = data.length()
    let mut pos = 0
    while pos < data_len {
      let mut d_end = pos
      while d_end < data_len && data.get_char(d_end).unwrap() != ',' {
        d_end = d_end + 1
      }
      let d_val = @xml.parse_int(@xml.str_substring(data, pos, d_end))
      let sp_start = d_end + 1
      let mut sp_end = sp_start
      while sp_end < data_len && data.get_char(sp_end).unwrap() != ';' {
        sp_end = sp_end + 1
      }
      let sp_val = @xml.parse_int(@xml.str_substring(data, sp_start, sp_end))
      out = out +
        "<a:ds d=\"" +
        int_to_str(d_val) +
        "\" sp=\"" +
        int_to_str(sp_val) +
        "\"/>"
      pos = sp_end + 1
    }
    out = out + "</a:custDash>"
  } else if not(shape.stroke_dash.is_empty()) {
    out = out + "<a:prstDash val=\"" + shape.stroke_dash + "\"/>"
  }
  // Line join
  if shape.stroke_join == "round" {
    out = out + "<a:round/>"
  } else if shape.stroke_join == "bevel" {
    out = out + "<a:bevel/>"
  } else if shape.stroke_join == "miter" {
    if shape.stroke_miter_limit > 0 {
      out = out +
        "<a:miter lim=\"" +
        int_to_str(shape.stroke_miter_limit) +
        "\"/>"
    } else {
      out = out + "<a:miter/>"
    }
  }
  // Arrow head/tail
  out = out +
    serialize_arrow_end(
      "a:headEnd",
      shape.stroke_head_type,
      shape.stroke_head_w,
      shape.stroke_head_len,
    )
  out = out +
    serialize_arrow_end(
      "a:tailEnd",
      shape.stroke_tail_type,
      shape.stroke_tail_w,
      shape.stroke_tail_len,
    )
  out + "</a:ln>"
}

///|
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

///|
fn serialize_tc_border(tag : String, w : Int, color : @ooxml.Color) -> String {
  if w == 0 && color.is_none() {
    return ""
  }
  let mut out = "<" + tag
  if w == -1 {
    out = out + ">"
    out = out + "<a:noFill/>"
    out = out + "</" + tag + ">"
    return out
  }
  if w > 0 {
    out = out + " w=\"" + int_to_str(w) + "\""
  }
  out = out + ">"
  if not(color.is_none()) {
    out = out + serialize_color(color)
  }
  out + "</" + tag + ">"
}

///|
fn serialize_table_cell(cell : @ooxml.TableCell) -> String {
  let mut out = "<a:tc"
  if cell.grid_span > 1 {
    out = out + " gridSpan=\"" + int_to_str(cell.grid_span) + "\""
  }
  if cell.row_span > 1 {
    out = out + " rowSpan=\"" + int_to_str(cell.row_span) + "\""
  }
  if cell.v_merge {
    out = out + " vMerge=\"1\""
  }
  if cell.h_merge {
    out = out + " hMerge=\"1\""
  }
  out = out + ">"
  // Text body
  if cell.paragraphs.length() > 0 {
    out = out + serialize_table_text_body(cell.paragraphs)
  } else {
    out = out +
      "<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></a:txBody>"
  }
  // Cell properties
  out = out + "<a:tcPr"
  if cell.mar_l >= 0 {
    out = out + " marL=\"" + int_to_str(cell.mar_l) + "\""
  }
  if cell.mar_r >= 0 {
    out = out + " marR=\"" + int_to_str(cell.mar_r) + "\""
  }
  if cell.mar_t >= 0 {
    out = out + " marT=\"" + int_to_str(cell.mar_t) + "\""
  }
  if cell.mar_b >= 0 {
    out = out + " marB=\"" + int_to_str(cell.mar_b) + "\""
  }
  if cell.anchor != "" {
    out = out + " anchor=\"" + cell.anchor + "\""
  }
  out = out + ">"
  // Borders
  out = out + serialize_tc_border("a:lnL", cell.bdr_l_w, cell.bdr_l_color)
  out = out + serialize_tc_border("a:lnR", cell.bdr_r_w, cell.bdr_r_color)
  out = out + serialize_tc_border("a:lnT", cell.bdr_t_w, cell.bdr_t_color)
  out = out + serialize_tc_border("a:lnB", cell.bdr_b_w, cell.bdr_b_color)
  out = out + serialize_tc_border("a:lnTlToBr", cell.bdr_tl_br_w, cell.bdr_tl_br_color)
  out = out + serialize_tc_border("a:lnBlToTr", cell.bdr_bl_tr_w, cell.bdr_bl_tr_color)
  // Fill
  if not(cell.grad_fill.is_none()) {
    out = out + serialize_gradient_fill(cell.grad_fill)
  } else if not(cell.fill.is_none()) {
    out = out + serialize_color(cell.fill)
  }
  out = out + "</a:tcPr>"
  out + "</a:tc>"
}

///|
fn serialize_table_text_body(
  paragraphs : Array[@ooxml.TextParagraph],
) -> String {
  serialize_text_body_inner(paragraphs, "a:txBody", @ooxml.BodyProps::default())
}

///|
fn serialize_table(shape : @ooxml.Shape, table : @ooxml.TableData) -> String {
  let mut out = "<p:graphicFrame>"
  // Non-visual properties
  out = out + "<p:nvGraphicFramePr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out +
    "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGraphicFramePr>"
  // Transform (directly under graphicFrame)
  out = out + serialize_transform(shape.transform)
  // Table content
  out = out +
    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">"
  out = out + "<a:tbl>"
  // Table properties
  out = out + "<a:tblPr"
  if not(table.style_id.is_empty()) {
    out = out + " tblStyleId=\"" + table.style_id + "\""
  }
  if table.first_row { out = out + " firstRow=\"1\"" }
  if table.last_row { out = out + " lastRow=\"1\"" }
  if table.first_col { out = out + " firstCol=\"1\"" }
  if table.last_col { out = out + " lastCol=\"1\"" }
  if table.band_row { out = out + " bandRow=\"1\"" }
  if table.band_col { out = out + " bandCol=\"1\"" }
  out = out + "/>"
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

///|
fn serialize_group(shape : @ooxml.Shape, grp : @ooxml.GroupShapeData) -> String {
  let mut out = "<p:grpSp>"
  out = out + "<p:nvGrpSpPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvGrpSpPr/>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGrpSpPr>"
  out = out + "<p:grpSpPr>"
  // Group xfrm includes chOff/chExt in addition to normal off/ext
  let ch_extra = "<a:chOff x=\"" +
    int_to_str(grp.ch_off_x) +
    "\" y=\"" +
    int_to_str(grp.ch_off_y) +
    "\"/>" +
    "<a:chExt cx=\"" +
    int_to_str(grp.ch_ext_cx) +
    "\" cy=\"" +
    int_to_str(grp.ch_ext_cy) +
    "\"/>"
  out = out + serialize_transform_with_extra(shape.transform, ch_extra)
  out = out + "</p:grpSpPr>"
  // Serialize child shapes
  let nc = grp.children.length()
  let mut i = 0
  while i < nc {
    out = out + serialize_shape(grp.children[i])
    i = i + 1
  }
  out + "</p:grpSp>"
}

///|
fn serialize_shape_inner(shape : @ooxml.Shape) -> String {
  match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => {
      // Connectors with connection refs serialize as p:cxnSp
      let is_cxn = match geom {
        @ooxml.ShapeGeom::Connector(_, _) => true
        @ooxml.ShapeGeom::Line => true
        _ => false
      }
      if is_cxn && (shape.st_cxn_id >= 0 || shape.end_cxn_id >= 0) {
        serialize_connector(shape, geom)
      } else {
        serialize_autoshape(shape, geom)
      }
    }
    @ooxml.ShapeKind::Picture(rid) => serialize_picture(shape, rid)
    @ooxml.ShapeKind::TableShape(table) => serialize_table(shape, table)
    @ooxml.ShapeKind::GroupShape(grp) => serialize_group(shape, grp)
    @ooxml.ShapeKind::ChartShape(_) => "" // Charts preserved via chart_xml round-trip
    _ => ""
  }
}

///|
fn serialize_shape(shape : @ooxml.Shape) -> String {
  let inner = serialize_shape_inner(shape)
  if shape.mc_choice_xml.length() == 0 {
    inner
  } else {
    "<mc:AlternateContent>" +
    shape.mc_choice_xml +
    "<mc:Fallback>" +
    inner +
    "</mc:Fallback></mc:AlternateContent>"
  }
}

///|
fn serialize_connector(shape : @ooxml.Shape, geom : @ooxml.ShapeGeom) -> String {
  let mut out = "<p:cxnSp>"
  // Non-visual properties
  out = out + "<p:nvCxnSpPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvCxnSpPr>"
  if shape.st_cxn_id >= 0 {
    out = out +
      "<a:stCxn id=\"" +
      int_to_str(shape.st_cxn_id) +
      "\" idx=\"" +
      int_to_str(if shape.st_cxn_idx >= 0 { shape.st_cxn_idx } else { 0 }) +
      "\"/>"
  }
  if shape.end_cxn_id >= 0 {
    out = out +
      "<a:endCxn id=\"" +
      int_to_str(shape.end_cxn_id) +
      "\" idx=\"" +
      int_to_str(if shape.end_cxn_idx >= 0 { shape.end_cxn_idx } else { 0 }) +
      "\"/>"
  }
  out = out + "</p:cNvCxnSpPr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvCxnSpPr>"
  // Shape properties
  out = out + "<p:spPr>"
  out = out + serialize_transform(shape.transform)
  out = out + serialize_prst_geom(geom)
  // Stroke
  let has_stroke = not(shape.stroke.is_none()) ||
    shape.stroke_no_fill ||
    not(shape.stroke_dash.is_empty()) ||
    not(shape.stroke_head_type.is_empty()) ||
    not(shape.stroke_tail_type.is_empty())
  if has_stroke {
    out = out + serialize_stroke(shape)
  }
  out = out + serialize_effect_lst(shape.effects)
  out = out + serialize_scene_3d(shape.scene_3d)
  out = out + serialize_sp_3d(shape.sp_3d)
  out = out + "</p:spPr>"
  out + "</p:cxnSp>"
}

// ── Public API ────────────────────────────────────────────────────────────────

///|
/// Serialize a SlideData into a complete OOXML slide XML string.
pub fn serialize_slide(slide : @ooxml.SlideData) -> String {
  let mut out = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  out = out +
    "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
  out = out +
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
  out = out +
    " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"" +
    " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
  out = out + "<p:cSld>"
  // Background
  if not(slide.bg_blip_fill.is_none()) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_blip_fill(slide.bg_blip_fill)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  } else if not(slide.bg_patt_fill.is_none()) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_pattern_fill(slide.bg_patt_fill)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  } else if not(slide.bg_grad.is_none()) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_gradient_fill(slide.bg_grad)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  } else if not(slide.background.is_none()) &&
    not(
      slide.background.r == 255 &&
      slide.background.g == 255 &&
      slide.background.b == 255,
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
