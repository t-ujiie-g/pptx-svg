// OOXML Text body parsing — paragraphs, runs, bodyPr
//
// Parses <a:bodyPr>, <a:p>, <a:r>, <a:rPr>, <a:pPr> elements into
// TextParagraph/TextRun/BodyProps structs. Handles bullets (char/auto/image),
// spacing, indentation, tabs, RTL, run-level effects, and font resolution.
//
// Same package as ooxml.mbt — all types and helpers are shared.

// ── Spacing parsing ──────────────────────────────────────────────────────────

///|
/// Parse spacing value from <a:spcBef>/<a:spcAft>/<a:lnSpc> children.
///  Returns EMU. Supports <a:spcPts val="N"> (N in hundredths of a point).
fn parse_spacing(parent : @xml.XmlNode) -> Int {
  let ch = @xml.get_children(parent)
  match @xml.find_child(ch, "a:spcPts") {
    Some(pts) => {
      // val is in hundredths of a point → convert to EMU (1 pt = 12700 EMU)
      let v = attr_int(pts, "val")
      v * emu_per_hundredth_pt
    }
    None =>
      // a:spcPct (percentage of line height) — store raw for now as negative sentinel
      match @xml.find_child(ch, "a:spcPct") {
        Some(pct) => {
          let v = attr_int(pct, "val") // e.g. 100000 = 100%
          -v
        } // negative = percentage marker
        None => 0
      }
  }
}

// ── Body properties parsing ──────────────────────────────────────────────────

///|
/// Parse <a:bodyPr> element into BodyProps.
fn parse_body_props(ch : Array[@xml.XmlNode]) -> BodyProps {
  match @xml.find_child(ch, "a:bodyPr") {
    None => BodyProps::default()
    Some(bpr) => {
      let anchor = match @xml.get_attr(bpr, "anchor") {
        Some(v) => v
        None => ""
      }
      let l_ins = match @xml.get_attr(bpr, "lIns") {
        Some(v) => @xml.parse_int(v)
        None => -1
      }
      let t_ins = match @xml.get_attr(bpr, "tIns") {
        Some(v) => @xml.parse_int(v)
        None => -1
      }
      let r_ins = match @xml.get_attr(bpr, "rIns") {
        Some(v) => @xml.parse_int(v)
        None => -1
      }
      let b_ins = match @xml.get_attr(bpr, "bIns") {
        Some(v) => @xml.parse_int(v)
        None => -1
      }
      let bpr_ch = @xml.get_children(bpr)
      let auto_fit = match @xml.find_child(bpr_ch, "a:spAutoFit") {
        Some(_) => true
        None => false
      }
      // Parse <a:normAutofit fontScale="..." lnSpcReduction="..."/>
      let (font_scale, ln_spc_reduction) = match
        @xml.find_child(bpr_ch, "a:normAutofit") {
        Some(naf) => {
          let fs = match @xml.get_attr(naf, "fontScale") {
            Some(v) => @xml.parse_int(v)
            None => pct_100
          }
          let lr = match @xml.get_attr(naf, "lnSpcReduction") {
            Some(v) => @xml.parse_int(v)
            None => 0
          }
          (fs, lr)
        }
        None => (-1, -1)
      }
      let wrap = match @xml.get_attr(bpr, "wrap") {
        Some(v) => v
        None => ""
      }
      let rot = attr_int(bpr, "rot")
      let vert = match @xml.get_attr(bpr, "vert") {
        Some(v) => v
        None => ""
      }
      let num_cols = match @xml.get_attr(bpr, "numCol") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let col_spacing = match @xml.get_attr(bpr, "spcCol") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      // Parse <a:prstTxWarp prst="..."> with optional <a:avLst>
      let (warp_prst, warp_av1, warp_av2) = match
        @xml.find_child(bpr_ch, "a:prstTxWarp") {
        None => ("", -1, -1)
        Some(w) => {
          let prst = match @xml.get_attr(w, "prst") {
            Some(v) => v
            None => ""
          }
          let w_ch = @xml.get_children(w)
          let (av1, av2) = match @xml.find_child(w_ch, "a:avLst") {
            None => (-1, -1)
            Some(avl) => {
              let gds = @xml.find_children(@xml.get_children(avl), "a:gd")
              let v1 = if gds.length() > 0 {
                // a:gd name="adj" fmla="val N" — extract N
                let fmla = match @xml.get_attr(gds[0], "fmla") {
                  Some(f) => f
                  None => ""
                }
                if @xml.str_starts_with(fmla, "val ") {
                  @xml.parse_int(@xml.str_suffix(fmla, 4))
                } else {
                  -1
                }
              } else {
                -1
              }
              let v2 = if gds.length() > 1 {
                let fmla2 = match @xml.get_attr(gds[1], "fmla") {
                  Some(f) => f
                  None => ""
                }
                if @xml.str_starts_with(fmla2, "val ") {
                  @xml.parse_int(@xml.str_suffix(fmla2, 4))
                } else {
                  -1
                }
              } else {
                -1
              }
              (v1, v2)
            }
          }
          (prst, av1, av2)
        }
      }
      {
        anchor,
        l_ins,
        t_ins,
        r_ins,
        b_ins,
        auto_fit,
        font_scale,
        ln_spc_reduction,
        wrap,
        rot,
        vert,
        num_cols,
        col_spacing,
        warp_prst,
        warp_av1,
        warp_av2,
      }
    }
  }
}

// ── Font theme resolution ────────────────────────────────────────────────────

///|
/// Resolve font theme references for Latin fonts.
fn resolve_font_theme(typeface : String, theme : ThemeData) -> String {
  if typeface == "+mj-lt" {
    theme.major_font
  } else if typeface == "+mn-lt" {
    theme.minor_font
  } else if typeface == "+mj-ea" {
    theme.major_ea_font
  } else if typeface == "+mn-ea" {
    theme.minor_ea_font
  } else {
    typeface
  }
}

///|
/// Resolve font theme references for East Asian fonts.
fn resolve_ea_font_theme(typeface : String, theme : ThemeData) -> String {
  if typeface == "+mj-ea" {
    theme.major_ea_font
  } else if typeface == "+mn-ea" {
    theme.minor_ea_font
  } else {
    typeface
  }
}

// ── Paragraph properties parsing ─────────────────────────────────────────────

///|
/// Parse paragraph-level properties from <a:pPr>.
fn parse_para_props(
  ppr : @xml.XmlNode,
  theme : ThemeData,
) -> (
  String,
  Int,
  Int,
  Int,
  Int,
  Int,
  Int,
  String,
  String,
  Bool,
  String,
  Int,
  Color,
  String,
  Array[TabStop],
  Bool,
) {
  let a = match @xml.get_attr(ppr, "algn") {
    Some(v) => v
    None => ""
  }
  let l = attr_int(ppr, "lvl")
  let ml = attr_int(ppr, "marL")
  let ind = attr_int(ppr, "indent")
  let rtl = match @xml.get_attr(ppr, "rtl") {
    Some(v) => v == "1" || v == "true"
    None => false
  }
  let ppr_ch = @xml.get_children(ppr)
  let sb = match @xml.find_child(ppr_ch, "a:spcBef") {
    None => 0
    Some(n) => parse_spacing(n)
  }
  let sa = match @xml.find_child(ppr_ch, "a:spcAft") {
    None => 0
    Some(n) => parse_spacing(n)
  }
  let lnspc = match @xml.find_child(ppr_ch, "a:lnSpc") {
    None => 0
    Some(n) => parse_spacing(n)
  }
  let bu_char = match @xml.find_child(ppr_ch, "a:buChar") {
    None => ""
    Some(bc) =>
      match @xml.get_attr(bc, "char") {
        Some(v) => v
        None => ""
      }
  }
  let bu_auto = match @xml.find_child(ppr_ch, "a:buAutoNum") {
    None => ""
    Some(ba) =>
      match @xml.get_attr(ba, "type") {
        Some(v) => v
        None => ""
      }
  }
  let bu_none = match @xml.find_child(ppr_ch, "a:buNone") {
    Some(_) => true
    None => false
  }
  // Bullet formatting
  let bu_font = match @xml.find_child(ppr_ch, "a:buFont") {
    None => ""
    Some(bf) =>
      match @xml.get_attr(bf, "typeface") {
        Some(v) => v
        None => ""
      }
  }
  let bu_size = match @xml.find_child(ppr_ch, "a:buSzPct") {
    Some(bp) =>
      match @xml.get_attr(bp, "val") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
    None =>
      match @xml.find_child(ppr_ch, "a:buSzPts") {
        Some(bp) =>
          -(match @xml.get_attr(bp, "val") {
            Some(v) => @xml.parse_int(v)
            None => 0
          })
        None => 0
      }
  }
  let bu_color = match @xml.find_child(ppr_ch, "a:buClr") {
    None => Color::none()
    Some(bc) => {
      let bc_ch = @xml.get_children(bc)
      parse_color_from_children(bc_ch, theme)
    }
  }
  // Image bullet: <a:buBlip><a:blip r:embed="rIdN"/></a:buBlip>
  let bu_img_rid = match @xml.find_child(ppr_ch, "a:buBlip") {
    None => ""
    Some(bub) => {
      let bub_ch = @xml.get_children(bub)
      match @xml.find_child(bub_ch, "a:blip") {
        None => ""
        Some(blip) =>
          match @xml.get_attr(blip, "r:embed") {
            Some(v) => v
            None => ""
          }
      }
    }
  }
  // Tab stops: <a:tabLst><a:tab pos="N" algn="l"/></a:tabLst>
  let tab_stops : Array[TabStop] = []
  match @xml.find_child(ppr_ch, "a:tabLst") {
    None => ()
    Some(tl) => {
      let tabs = @xml.find_children(@xml.get_children(tl), "a:tab")
      let nt = tabs.length()
      let mut ti = 0
      while ti < nt {
        let pos = attr_int(tabs[ti], "pos")
        let algn = match @xml.get_attr(tabs[ti], "algn") {
          Some(v) => v
          None => ""
        }
        tab_stops.push({ pos, algn })
        ti = ti + 1
      }
    }
  }
  (
    a, l, sb, sa, lnspc, ml, ind, bu_char, bu_auto, bu_none, bu_font, bu_size, bu_color,
    bu_img_rid, tab_stops, rtl,
  )
}

// ── Default run properties ───────────────────────────────────────────────────

///|
/// Parse <a:defRPr> defaults from a <a:pPr> node.
fn parse_def_rpr(
  ppr : @xml.XmlNode,
  theme : ThemeData,
) -> (Bool, Bool, Int, Color, String, String) {
  let ppr_ch = @xml.get_children(ppr)
  match @xml.find_child(ppr_ch, "a:defRPr") {
    None => (false, false, 0, Color::none(), "", "")
    Some(drpr) => {
      let db = match @xml.get_attr(drpr, "b") {
        Some(v) => v == "1" || v == "true"
        None => false
      }
      let di = match @xml.get_attr(drpr, "i") {
        Some(v) => v == "1" || v == "true"
        None => false
      }
      let dsz = attr_int(drpr, "sz")
      let drpr_ch = @xml.get_children(drpr)
      let dclr = parse_solid_fill(drpr_ch, theme)
      let dff_raw = child_attr(drpr_ch, "a:latin", "typeface")
      let dff = resolve_font_theme(dff_raw, theme)
      let dea_raw = child_attr(drpr_ch, "a:ea", "typeface")
      let dea = resolve_ea_font_theme(dea_raw, theme)
      (db, di, dsz, dclr, dff, dea)
    }
  }
}

// ── Run properties parsing ───────────────────────────────────────────────────

///|
/// Create a TextRun with default properties (no explicit formatting).
fn default_text_run(
  text : String,
  def_bold : Bool,
  def_italic : Bool,
  def_font_size : Int,
  def_color : Color,
  def_font_face : String,
  def_ea_font : String,
) -> TextRun {
  {
    text,
    bold: def_bold,
    italic: def_italic,
    font_size: def_font_size,
    color: def_color,
    font_face: def_font_face,
    ea_font: def_ea_font,
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
    effects: EffectList::none(),
    outline_color: Color::none(),
    outline_w: 0,
    text_grad_fill: GradientFill::none(),
    text_patt_fill: PatternFill::none(),
    bold_explicit: false,
  }
}

///|
/// Parse run properties from <a:rPr> with defRPr fallback. Returns a TextRun (text field set to "").
fn parse_run_props(
  rpr : @xml.XmlNode,
  def_bold : Bool,
  def_italic : Bool,
  def_font_size : Int,
  def_color : Color,
  def_font_face : String,
  def_ea_font : String,
  theme : ThemeData,
) -> TextRun {
  let bold_attr = @xml.get_attr(rpr, "b")
  let bold_explicit = match bold_attr {
    Some(_) => true
    None => false
  }
  let bold = match bold_attr {
    Some(v) => v == "1" || v == "true"
    None => def_bold
  }
  let italic = match @xml.get_attr(rpr, "i") {
    Some(v) => v == "1" || v == "true"
    None => def_italic
  }
  let sz_raw = attr_int(rpr, "sz")
  let font_size = if sz_raw > 0 { sz_raw } else { def_font_size }
  let underline = match @xml.get_attr(rpr, "u") {
    Some(v) => v
    None => ""
  }
  let strike = match @xml.get_attr(rpr, "strike") {
    Some(v) => v
    None => ""
  }
  let baseline = attr_int(rpr, "baseline")
  let char_spacing = attr_int(rpr, "spc")
  let kern = attr_int(rpr, "kern")
  let rpr_ch = @xml.get_children(rpr)
  let clr_raw = parse_solid_fill(rpr_ch, theme)
  let color = if clr_raw.is_none() { def_color } else { clr_raw }
  let ff_raw = child_attr(rpr_ch, "a:latin", "typeface")
  let ff_resolved = resolve_font_theme(ff_raw, theme)
  let font_face = if ff_resolved.is_empty() {
    def_font_face
  } else {
    ff_resolved
  }
  let ea_raw = child_attr(rpr_ch, "a:ea", "typeface")
  let ea_resolved = resolve_ea_font_theme(ea_raw, theme)
  let ea_font = if ea_resolved.is_empty() { def_ea_font } else { ea_resolved }
  let cs_font = child_attr(rpr_ch, "a:cs", "typeface")
  let sym_font = child_attr(rpr_ch, "a:sym", "typeface")
  let cap = match @xml.get_attr(rpr, "cap") {
    Some(v) => v
    None => ""
  }
  // Hyperlink: <a:hlinkClick r:id="rIdN"/>
  let hlink_rid = match @xml.find_child(rpr_ch, "a:hlinkClick") {
    None => ""
    Some(hl) =>
      match @xml.get_attr(hl, "r:id") {
        Some(v) => v
        None => ""
      }
  }
  let hlink_mouse_over_rid = match @xml.find_child(rpr_ch, "a:hlinkHover") {
    None => ""
    Some(hl) =>
      match @xml.get_attr(hl, "r:id") {
        Some(v) => v
        None => ""
      }
  }
  // Text outline: a:ln on a:rPr
  let (outline_color, outline_w) = match @xml.find_child(rpr_ch, "a:ln") {
    None => (Color::none(), 0)
    Some(ln) => {
      let ow = attr_int(ln, "w")
      let ln_ch = @xml.get_children(ln)
      let oc = parse_solid_fill(ln_ch, theme)
      (if oc.is_none() { Color::black() } else { oc }, ow)
    }
  }
  // Text gradient fill: a:gradFill on a:rPr
  let text_grad_fill = parse_gradient_fill(rpr_ch, theme)
  // Text pattern fill: a:pattFill on a:rPr
  let text_patt_fill = parse_pattern_fill(rpr_ch, theme)
  {
    text: "",
    bold,
    bold_explicit,
    italic,
    font_size,
    color,
    font_face,
    ea_font,
    cs_font,
    sym_font,
    underline,
    strike,
    baseline,
    char_spacing,
    kern,
    cap,
    hlink_rid,
    hlink_mouse_over_rid,
    effects: parse_run_effect_lst(rpr_ch, theme),
    outline_color,
    outline_w,
    text_grad_fill,
    text_patt_fill,
  }
}

///|
fn parse_run_effect_lst(rpr_ch : Array[@xml.XmlNode], theme : ThemeData) -> EffectList {
  // a:effectLst can appear as a child of a:rPr
  match @xml.find_child(rpr_ch, "a:effectLst") {
    None => EffectList::none()
    Some(eff) => {
      let eff_ch = @xml.get_children(eff)
      // Reuse the shape-level effect parsing (only need outerShdw for text typically)
      let os = match @xml.find_child(eff_ch, "a:outerShdw") {
        None => OuterShadow::none()
        Some(osn) => parse_outer_shadow(osn, theme)
      }
      let ishdw = match @xml.find_child(eff_ch, "a:innerShdw") {
        None => InnerShadow::none()
        Some(isn) => parse_inner_shadow(isn, theme)
      }
      let gl = match @xml.find_child(eff_ch, "a:glow") {
        None => Glow::none()
        Some(gn) => parse_glow(gn, theme)
      }
      {
        outer_shadow: os,
        inner_shadow: ishdw,
        glow: gl,
        soft_edge: SoftEdge::none(),
        reflection: Reflection::none(),
      }
    }
  }
}

// ── Text body parsing ────────────────────────────────────────────────────────

///|
fn parse_text_body(
  txbody : @xml.XmlNode,
  theme : ThemeData,
) -> (Array[TextParagraph], BodyProps, TextStyleGroup) {
  let paras : Array[TextParagraph] = []
  let ch = @xml.get_children(txbody)
  let body_props = parse_body_props(ch)
  // Parse a:lstStyle for shape-specific level defaults
  let lst_style = match @xml.find_child(ch, "a:lstStyle") {
    None => TextStyleGroup::empty()
    Some(ls) => parse_text_style_group(ls, theme)
  }
  let pnodes = @xml.find_children(ch, "a:p")
  let np = pnodes.length()
  let mut pi = 0
  while pi < np {
    let p = pnodes[pi]
    let p_ch = @xml.get_children(p)
    let (
      align,
      level,
      spc_before,
      spc_after,
      line_spacing,
      mar_l,
      indent,
      bullet,
      bullet_auto,
      bullet_none,
      bullet_font,
      bullet_size,
      bullet_color,
      bullet_img_rid,
      tab_stops,
      rtl,
    ) = match @xml.find_child(p_ch, "a:pPr") {
      None =>
        (
          "",
          0,
          0,
          0,
          0,
          0,
          0,
          "",
          "",
          false,
          "",
          0,
          Color::none(),
          "",
          ([] : Array[TabStop]),
          false,
        )
      Some(ppr) => parse_para_props(ppr, theme)
    }
    let (
      def_bold,
      def_italic,
      def_font_size,
      def_color,
      def_font_face,
      def_ea_font,
    ) = match @xml.find_child(p_ch, "a:pPr") {
      None => (false, false, 0, Color::none(), "", "")
      Some(ppr) => parse_def_rpr(ppr, theme)
    }
    // Apply lstStyle level defaults as fallback
    let lvl_idx = if level >= 0 && level < max_outline_levels {
      level
    } else {
      0
    }
    let lst_def = lst_style.levels[lvl_idx]
    let def_font_size = if def_font_size == 0 && lst_def.font_size > 0 {
      lst_def.font_size
    } else {
      def_font_size
    }
    let def_color = if def_color.is_none() && not(lst_def.color.is_none()) {
      lst_def.color
    } else {
      def_color
    }
    let def_font_face = if def_font_face.is_empty() &&
      not(lst_def.font_face.is_empty()) {
      lst_def.font_face
    } else {
      def_font_face
    }
    let def_ea_font = if def_ea_font.is_empty() &&
      not(lst_def.ea_font.is_empty()) {
      lst_def.ea_font
    } else {
      def_ea_font
    }
    // Iterate over all children to handle a:r (runs), a:fld (fields), and a:br (breaks)
    let runs : Array[TextRun] = []
    let p_ch_len = p_ch.length()
    let mut ri = 0
    while ri < p_ch_len {
      let child = p_ch[ri]
      let child_tag = @xml.get_tag(child)
      // a:r (run) and a:fld (field, e.g. slide number) share the same structure
      if child_tag == "a:r" || child_tag == "a:fld" {
        let r_ch = @xml.get_children(child)
        let text = match @xml.find_child(r_ch, "a:t") {
          None => ""
          Some(t_node) => @xml.get_text_content(t_node)
        }
        if not(text.is_empty()) {
          let run = match @xml.find_child(r_ch, "a:rPr") {
            None =>
              default_text_run(
                text, def_bold, def_italic, def_font_size, def_color,
                def_font_face, def_ea_font,
              )
            Some(rpr) => {
              let rp = parse_run_props(
                rpr, def_bold, def_italic, def_font_size, def_color,
                def_font_face, def_ea_font, theme,
              )
              { ..rp, text }
            }
          }
          runs.push(run)
        }
      } else if child_tag == "a:br" {
        // Line break — insert a newline run
        runs.push(
          default_text_run(
            "\n", def_bold, def_italic, def_font_size, def_color,
            def_font_face, def_ea_font,
          ),
        )
      }
      ri = ri + 1
    }
    paras.push({
      runs,
      align,
      level,
      spc_before,
      spc_after,
      line_spacing,
      mar_l,
      indent,
      bullet,
      bullet_auto,
      bullet_none,
      bullet_font,
      bullet_size,
      bullet_color,
      bullet_img_rid,
      tab_stops,
      rtl,
    })
    pi = pi + 1
  }
  // Apply lstStyle fallbacks to paragraphs missing explicit values
  let npara = paras.length()
  let mut pai = 0
  while pai < npara {
    let para = paras[pai]
    let lvl_idx = if para.level >= 0 && para.level < max_outline_levels {
      para.level
    } else {
      0
    }
    let lst_def = lst_style.levels[lvl_idx]
    let mut changed = false
    let mut new_para = para
    // Alignment fallback
    if para.align.is_empty() && not(lst_def.align.is_empty()) {
      new_para = { ..new_para, align: lst_def.align }
      changed = true
    }
    // Bullet fallback: inherit buChar from lstStyle if paragraph has no explicit bullet
    if para.bullet.is_empty() && para.bullet_auto.is_empty() &&
      not(para.bullet_none) &&
      para.bullet_img_rid.is_empty() &&
      not(lst_def.bullet.is_empty()) {
      new_para = { ..new_para, bullet: lst_def.bullet }
      changed = true
      // Also inherit bullet_font/size/color from lstStyle if not set on paragraph
      if new_para.bullet_font.is_empty() && not(lst_def.bullet_font.is_empty()) {
        new_para = { ..new_para, bullet_font: lst_def.bullet_font }
      }
      if new_para.bullet_size == 0 && lst_def.bullet_size != 0 {
        new_para = { ..new_para, bullet_size: lst_def.bullet_size }
      }
      if new_para.bullet_color.is_none() && not(lst_def.bullet_color.is_none()) {
        new_para = { ..new_para, bullet_color: lst_def.bullet_color }
      }
    } else if not(para.bullet.is_empty()) || not(para.bullet_auto.is_empty()) {
      // Paragraph has explicit bullet, but may still need font/size/color from lstStyle
      if para.bullet_font.is_empty() && not(lst_def.bullet_font.is_empty()) {
        new_para = { ..new_para, bullet_font: lst_def.bullet_font }
        changed = true
      }
      if para.bullet_size == 0 && lst_def.bullet_size != 0 {
        new_para = { ..new_para, bullet_size: lst_def.bullet_size }
        changed = true
      }
      if para.bullet_color.is_none() && not(lst_def.bullet_color.is_none()) {
        new_para = { ..new_para, bullet_color: lst_def.bullet_color }
        changed = true
      }
    }
    // buNone fallback
    if not(para.bullet_none) && lst_def.bullet_none &&
      para.bullet.is_empty() &&
      para.bullet_auto.is_empty() &&
      para.bullet_img_rid.is_empty() {
      new_para = { ..new_para, bullet_none: true }
      changed = true
    }
    // Line spacing fallback
    if para.line_spacing == 0 && lst_def.line_spacing != 0 {
      new_para = { ..new_para, line_spacing: lst_def.line_spacing }
      changed = true
    }
    // Space before/after fallback
    if para.spc_before == 0 && lst_def.spc_before != 0 {
      new_para = { ..new_para, spc_before: lst_def.spc_before }
      changed = true
    }
    if para.spc_after == 0 && lst_def.spc_after != 0 {
      new_para = { ..new_para, spc_after: lst_def.spc_after }
      changed = true
    }
    if changed {
      paras[pai] = new_para
    }
    pai = pai + 1
  }
  (paras, body_props, lst_style)
}
