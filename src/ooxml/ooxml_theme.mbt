// OOXML Theme / ColorMap / Master-Layout parsing
//
// Same package as ooxml.mbt — all types and helpers are shared.

// ── Theme color node parser ──────────────────────────────────────────────────

///|
/// Parse the first color child of a clrScheme element (sysClr or srgbClr).
fn parse_theme_color_node(nodes : Array[@xml.XmlNode], tag : String) -> Color? {
  match @xml.find_child(nodes, tag) {
    None => None
    Some(node) => {
      let ch = @xml.get_children(node)
      // Try <a:srgbClr val="RRGGBB">
      match @xml.find_child(ch, "a:srgbClr") {
        Some(clr) =>
          match @xml.get_attr(clr, "val") {
            Some(v) => Some(parse_hex_color(v))
            None => None
          }
        None =>
          // Try <a:sysClr lastClr="RRGGBB">
          match @xml.find_child(ch, "a:sysClr") {
            Some(sys) =>
              match @xml.get_attr(sys, "lastClr") {
                Some(v) => Some(parse_hex_color(v))
                None => None
              }
            None => None
          }
      }
    }
  }
}

// ── Theme parser ────────────────────────────────────────────────────────────

///|
/// Parse theme1.xml into ThemeData.
pub fn parse_theme(xml_str : String) -> ThemeData {
  let def = ThemeData::default()
  let nodes = @xml.parse_xml(xml_str)
  // Navigate: a:theme → a:themeElements
  let theme_el = match @xml.find_child(nodes, "a:theme") {
    None => return def
    Some(t) =>
      match @xml.find_child(@xml.get_children(t), "a:themeElements") {
        None => return def
        Some(te) => te
      }
  }
  let te_ch = @xml.get_children(theme_el)
  // ── Color scheme ──
  let slot_tags = [
    "a:dk1", "a:lt1", "a:dk2", "a:lt2", "a:accent1", "a:accent2", "a:accent3", "a:accent4",
    "a:accent5", "a:accent6", "a:hlink", "a:folHlink",
  ]
  let slot_colors = [
    def.dk1,
    def.lt1,
    def.dk2,
    def.lt2,
    def.accent1,
    def.accent2,
    def.accent3,
    def.accent4,
    def.accent5,
    def.accent6,
    def.hlink,
    def.fol_hlink,
  ]
  match @xml.find_child(te_ch, "a:clrScheme") {
    None => ()
    Some(cs) => {
      let cs_ch = @xml.get_children(cs)
      let mut si = 0
      while si < slot_tags.length() {
        match parse_theme_color_node(cs_ch, slot_tags[si]) {
          Some(c) => slot_colors[si] = c
          None => ()
        }
        si = si + 1
      }
    }
  }
  let dk1 = slot_colors[0]
  let lt1 = slot_colors[1]
  let dk2 = slot_colors[2]
  let lt2 = slot_colors[3]
  let accent1 = slot_colors[4]
  let accent2 = slot_colors[5]
  let accent3 = slot_colors[6]
  let accent4 = slot_colors[7]
  let accent5 = slot_colors[8]
  let accent6 = slot_colors[9]
  let hlink = slot_colors[10]
  let fol_hlink = slot_colors[11]
  // ── Font scheme ──
  let mut major_font = def.major_font
  let mut minor_font = def.minor_font
  let mut major_ea_font = def.major_ea_font
  let mut minor_ea_font = def.minor_ea_font
  match @xml.find_child(te_ch, "a:fontScheme") {
    None => ()
    Some(fs) => {
      let fs_ch = @xml.get_children(fs)
      match @xml.find_child(fs_ch, "a:majorFont") {
        None => ()
        Some(mj) => {
          let mj_ch = @xml.get_children(mj)
          match @xml.find_child(mj_ch, "a:latin") {
            None => ()
            Some(latin) =>
              match @xml.get_attr(latin, "typeface") {
                Some(v) => major_font = v
                None => ()
              }
          }
          match @xml.find_child(mj_ch, "a:ea") {
            None => ()
            Some(ea) =>
              match @xml.get_attr(ea, "typeface") {
                Some(v) => major_ea_font = v
                None => ()
              }
          }
        }
      }
      match @xml.find_child(fs_ch, "a:minorFont") {
        None => ()
        Some(mn) => {
          let mn_ch = @xml.get_children(mn)
          match @xml.find_child(mn_ch, "a:latin") {
            None => ()
            Some(latin) =>
              match @xml.get_attr(latin, "typeface") {
                Some(v) => minor_font = v
                None => ()
              }
          }
          match @xml.find_child(mn_ch, "a:ea") {
            None => ()
            Some(ea) =>
              match @xml.get_attr(ea, "typeface") {
                Some(v) => minor_ea_font = v
                None => ()
              }
          }
        }
      }
    }
  }
  {
    dk1,
    lt1,
    dk2,
    lt2,
    accent1,
    accent2,
    accent3,
    accent4,
    accent5,
    accent6,
    hlink,
    fol_hlink,
    major_font,
    minor_font,
    major_ea_font,
    minor_ea_font,
  }
}

// ── Color map helpers ────────────────────────────────────────────────────────

///|
/// Get an XML attribute with a fallback default value.
fn get_attr_or(node : @xml.XmlNode, name : String, default : String) -> String {
  match @xml.get_attr(node, name) {
    Some(v) => v
    None => default
  }
}

///|
/// Parse a ColorMap from XML node attributes (e.g. `<p:clrMap bg1="dk1" .../>`).
pub fn parse_clr_map_from_node(node : @xml.XmlNode) -> ColorMap {
  let d = ColorMap::default()
  {
    bg1: get_attr_or(node, "bg1", d.bg1),
    tx1: get_attr_or(node, "tx1", d.tx1),
    bg2: get_attr_or(node, "bg2", d.bg2),
    tx2: get_attr_or(node, "tx2", d.tx2),
    accent1: get_attr_or(node, "accent1", d.accent1),
    accent2: get_attr_or(node, "accent2", d.accent2),
    accent3: get_attr_or(node, "accent3", d.accent3),
    accent4: get_attr_or(node, "accent4", d.accent4),
    accent5: get_attr_or(node, "accent5", d.accent5),
    accent6: get_attr_or(node, "accent6", d.accent6),
    hlink: get_attr_or(node, "hlink", d.hlink),
    fol_hlink: get_attr_or(node, "folHlink", d.fol_hlink),
  }
}

///|
/// Resolve a direct theme slot name (dk1, lt1, accent1, etc.) to a Color.
///  Unlike `resolve_scheme_color`, this does NOT apply bg1→lt1 / tx1→dk1 aliasing.
fn resolve_slot(name : String, theme : ThemeData) -> Color {
  if name == "dk1" {
    theme.dk1
  } else if name == "lt1" {
    theme.lt1
  } else if name == "dk2" {
    theme.dk2
  } else if name == "lt2" {
    theme.lt2
  } else if name == "accent1" {
    theme.accent1
  } else if name == "accent2" {
    theme.accent2
  } else if name == "accent3" {
    theme.accent3
  } else if name == "accent4" {
    theme.accent4
  } else if name == "accent5" {
    theme.accent5
  } else if name == "accent6" {
    theme.accent6
  } else if name == "hlink" {
    theme.hlink
  } else if name == "folHlink" {
    theme.fol_hlink
  } else {
    Color::black()
  }
}

///|
/// Apply a ColorMap to a ThemeData, producing a remapped ThemeData.
///
///  `resolve_scheme_color` hard-codes: bg1/lt1→theme.lt1, tx1/dk1→theme.dk1, etc.
///  This function sets `new.lt1 = resolve_slot(cmap.bg1, original)` so that
///  `resolve_scheme_color("bg1")` on the remapped theme yields the correct color.
pub fn apply_color_map(theme : ThemeData, cmap : ColorMap) -> ThemeData {
  {
    dk1: resolve_slot(cmap.tx1, theme),
    lt1: resolve_slot(cmap.bg1, theme),
    dk2: resolve_slot(cmap.tx2, theme),
    lt2: resolve_slot(cmap.bg2, theme),
    accent1: resolve_slot(cmap.accent1, theme),
    accent2: resolve_slot(cmap.accent2, theme),
    accent3: resolve_slot(cmap.accent3, theme),
    accent4: resolve_slot(cmap.accent4, theme),
    accent5: resolve_slot(cmap.accent5, theme),
    accent6: resolve_slot(cmap.accent6, theme),
    hlink: resolve_slot(cmap.hlink, theme),
    fol_hlink: resolve_slot(cmap.fol_hlink, theme),
    major_font: theme.major_font,
    minor_font: theme.minor_font,
    major_ea_font: theme.major_ea_font,
    minor_ea_font: theme.minor_ea_font,
  }
}

///|
/// Extract `p:clrMapOvr` from already-parsed XML nodes.
///  Returns `None` if not present or if `<a:masterClrMapping/>` is used (= inherit master).
pub fn extract_clr_map_ovr_from_nodes(
  root_ch : Array[@xml.XmlNode],
) -> ColorMap? {
  match @xml.find_child(root_ch, "p:clrMapOvr") {
    None => None
    Some(ovr_node) => {
      let ovr_ch = @xml.get_children(ovr_node)
      match @xml.find_child(ovr_ch, "a:overrideClrMapping") {
        Some(mapping_node) => Some(parse_clr_map_from_node(mapping_node))
        None => None
      }
    }
  }
}

// ── Level text defaults parsing ──────────────────────────────────────────────

///|
/// Parse a level paragraph properties node (e.g. a:lvl1pPr) into LevelTextDefaults.
fn parse_level_defaults(
  ppr : @xml.XmlNode,
  theme : ThemeData,
) -> LevelTextDefaults {
  let align = match @xml.get_attr(ppr, "algn") {
    Some(v) => v
    None => ""
  }
  let mar_l = match @xml.get_attr(ppr, "marL") {
    Some(v) => @xml.parse_int(v)
    None => -1
  }
  let indent = attr_int(ppr, "indent")
  let ppr_ch = @xml.get_children(ppr)
  // Bullet
  let bu_char = match @xml.find_child(ppr_ch, "a:buChar") {
    None => ""
    Some(bc) =>
      match @xml.get_attr(bc, "char") {
        Some(v) => v
        None => ""
      }
  }
  let bu_none = match @xml.find_child(ppr_ch, "a:buNone") {
    Some(_) => true
    None => false
  }
  let bu_font = match @xml.find_child(ppr_ch, "a:buFont") {
    None => ""
    Some(bf) =>
      match @xml.get_attr(bf, "typeface") {
        Some(v) => v
        None => ""
      }
  }
  let bu_size = match @xml.find_child(ppr_ch, "a:buSzPct") {
    Some(bsp) =>
      match @xml.get_attr(bsp, "val") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
    None =>
      match @xml.find_child(ppr_ch, "a:buSzPts") {
        Some(bsp) =>
          match @xml.get_attr(bsp, "val") {
            Some(v) => -@xml.parse_int(v)
            None => 0
          }
        None => 0
      }
  }
  let bu_color = match @xml.find_child(ppr_ch, "a:buClr") {
    None => Color::none()
    Some(bc) => {
      let bc_ch = @xml.get_children(bc)
      parse_solid_fill(bc_ch, theme)
    }
  }
  // Line spacing & spacing before/after
  let lnspc = match @xml.find_child(ppr_ch, "a:lnSpc") {
    Some(n) => parse_spacing(n)
    None => 0
  }
  let spc_bef = match @xml.find_child(ppr_ch, "a:spcBef") {
    Some(n) => parse_spacing(n)
    None => 0
  }
  let spc_aft = match @xml.find_child(ppr_ch, "a:spcAft") {
    Some(n) => parse_spacing(n)
    None => 0
  }
  // Default run properties from a:defRPr
  let (bold, italic, font_size, color, font_face, ea_font) = match
    @xml.find_child(ppr_ch, "a:defRPr") {
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
  {
    font_size,
    bold,
    italic,
    color,
    font_face,
    ea_font,
    align,
    mar_l,
    indent,
    bullet: bu_char,
    bullet_none: bu_none,
    bullet_font: bu_font,
    bullet_size: bu_size,
    bullet_color: bu_color,
    line_spacing: lnspc,
    spc_before: spc_bef,
    spc_after: spc_aft,
  }
}

///|
/// Parse a text style group (e.g. p:titleStyle) into TextStyleGroup.
pub fn parse_text_style_group(
  node : @xml.XmlNode,
  theme : ThemeData,
) -> TextStyleGroup {
  let ch = @xml.get_children(node)
  let levels : Array[LevelTextDefaults] = []
  let mut i = 0
  while i < max_outline_levels {
    let tag = "a:lvl" + @xml.int_to_str(i + 1) + "pPr"
    let lvl = match @xml.find_child(ch, tag) {
      None => LevelTextDefaults::empty()
      Some(ppr) => parse_level_defaults(ppr, theme)
    }
    levels.push(lvl)
    i = i + 1
  }
  { levels, }
}

///|
/// Parse p:txStyles from slide master.
fn parse_tx_styles(node : @xml.XmlNode, theme : ThemeData) -> MasterTextStyles {
  let ch = @xml.get_children(node)
  let title_style = match @xml.find_child(ch, "p:titleStyle") {
    None => TextStyleGroup::empty()
    Some(ts) => parse_text_style_group(ts, theme)
  }
  let body_style = match @xml.find_child(ch, "p:bodyStyle") {
    None => TextStyleGroup::empty()
    Some(bs) => parse_text_style_group(bs, theme)
  }
  let other_style = match @xml.find_child(ch, "p:otherStyle") {
    None => TextStyleGroup::empty()
    Some(os) => parse_text_style_group(os, theme)
  }
  { title_style, body_style, other_style }
}

// ── Slide master parser ──────────────────────────────────────────────────────

///|
/// Parse a slide master XML string into SlideMasterData.
pub fn parse_slide_master(
  xml_str : String,
  theme : ThemeData,
) -> SlideMasterData {
  let nodes = @xml.parse_xml(xml_str)
  // Parse p:clrMap and apply to theme
  let clr_map = match @xml.find_child(nodes, "p:sldMaster") {
    None => ColorMap::default()
    Some(master_node) => {
      let master_ch = @xml.get_children(master_node)
      match @xml.find_child(master_ch, "p:clrMap") {
        None => ColorMap::default()
        Some(cm_node) => parse_clr_map_from_node(cm_node)
      }
    }
  }
  let effective_theme = apply_color_map(theme, clr_map)
  let csld = @xml.find_path(nodes, ["p:sldMaster", "p:cSld"])
  let (background, bg_grad, bg_blip_fill, bg_patt_fill) = match csld {
    None =>
      (
        Color::none(),
        GradientFill::none(),
        BlipFill::none(),
        PatternFill::none(),
      )
    Some(csld_node) =>
      parse_bg_fills_from_csld(@xml.get_children(csld_node), effective_theme)
  }
  let shapes : Array[Shape] = match csld {
    None => []
    Some(csld_node) => {
      let csld_ch = @xml.get_children(csld_node)
      match @xml.find_child(csld_ch, "p:spTree") {
        None => []
        Some(tree) => parse_shapes_from_tree(tree, effective_theme)
      }
    }
  }
  // Parse p:txStyles
  let text_styles = match @xml.find_child(nodes, "p:sldMaster") {
    None => MasterTextStyles::empty()
    Some(master_node) => {
      let master_ch = @xml.get_children(master_node)
      match @xml.find_child(master_ch, "p:txStyles") {
        None => MasterTextStyles::empty()
        Some(txs) => parse_tx_styles(txs, effective_theme)
      }
    }
  }
  {
    background,
    bg_grad,
    bg_blip_fill,
    bg_patt_fill,
    shapes,
    text_styles,
    clr_map,
  }
}

// ── Slide layout parser ──────────────────────────────────────────────────────

///|
/// Parse a slide layout XML string into SlideLayoutData.
pub fn parse_slide_layout(
  xml_str : String,
  theme : ThemeData,
) -> SlideLayoutData {
  let nodes = @xml.parse_xml(xml_str)
  let root_ch = match @xml.find_child(nodes, "p:sldLayout") {
    Some(n) => @xml.get_children(n)
    None => []
  }
  let csld = @xml.find_child(root_ch, "p:cSld")
  let (background, bg_grad, bg_blip_fill, bg_patt_fill) = match csld {
    None =>
      (
        Color::none(),
        GradientFill::none(),
        BlipFill::none(),
        PatternFill::none(),
      )
    Some(csld_node) =>
      parse_bg_fills_from_csld(@xml.get_children(csld_node), theme)
  }
  let shapes : Array[Shape] = match csld {
    None => []
    Some(csld_node) => {
      let csld_ch = @xml.get_children(csld_node)
      match @xml.find_child(csld_ch, "p:spTree") {
        None => []
        Some(tree) => parse_shapes_from_tree(tree, theme)
      }
    }
  }
  let clr_map_ovr = extract_clr_map_ovr_from_nodes(root_ch)
  { background, bg_grad, bg_blip_fill, bg_patt_fill, shapes, clr_map_ovr }
}
