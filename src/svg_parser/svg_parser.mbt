// SVG → SlideData parser for round-trip editing
//
// Reads data-ooxml-* attributes embedded by the renderer to reconstruct
// a SlideData from edited SVG. All values in OOXML native units (EMU, etc.)
//
// NOTE: No StringBuilder — all strings built via `+` (wasm:js-string "concat").

// ── Integer parsing ──────────────────────────────────────────────────────────

// ── Color parsing ────────────────────────────────────────────────────────────

///|
fn parse_hex_color(s : String) -> @ooxml.Color {
  if s == "none" {
    @ooxml.Color::none()
  } else {
    @ooxml.parse_hex_color(s)
  }
}

///|
/// Parse a hex color from a data attribute, returning Color::none() if empty.
fn color_or_none(node : @xml.XmlNode, attr : String) -> @ooxml.Color {
  let v = get_data_attr(node, attr)
  if v.is_empty() { @ooxml.Color::none() } else { parse_hex_color(v) }
}

// ── Attribute helpers ────────────────────────────────────────────────────────

///|
fn get_data_attr(node : @xml.XmlNode, name : String) -> String {
  match @xml.get_attr(node, name) {
    Some(v) => v
    None => ""
  }
}

///|
fn get_data_int(node : @xml.XmlNode, name : String) -> Int {
  @xml.parse_int(get_data_attr(node, name))
}

///|
/// Shorthand for reading data-ooxml-* attributes.
fn get_da(node : @xml.XmlNode, name : String) -> String {
  get_data_attr(node, "data-ooxml-" + name)
}

///|
fn get_dai(node : @xml.XmlNode, name : String) -> Int {
  get_data_int(node, "data-ooxml-" + name)
}

// ── Effect parsing from SVG data attributes ─────────────────────────────────

///|
/// Parse effects from data-ooxml-* attributes on a shape element.
fn parse_effects_from_da(node : @xml.XmlNode) -> @ooxml.EffectList {
  let os_blur = get_dai(node, "eff-os-blur")
  let os_dist = get_dai(node, "eff-os-dist")
  let os_dir = get_dai(node, "eff-os-dir")
  let os_clr = get_da(node, "eff-os-clr")
  let os = if os_clr.is_empty() {
    @ooxml.OuterShadow::none()
  } else {
    {
      blur_rad: os_blur,
      dist: os_dist,
      dir: os_dir,
      color: parse_hex_color(os_clr),
      sx: get_dai(node, "eff-os-sx"),
      sy: get_dai(node, "eff-os-sy"),
      algn: get_da(node, "eff-os-algn"),
      rot_with_shape: get_da(node, "eff-os-rot") != "0",
    }
  }
  let is_clr = get_da(node, "eff-is-clr")
  let ishdw = if is_clr.is_empty() {
    @ooxml.InnerShadow::none()
  } else {
    {
      blur_rad: get_dai(node, "eff-is-blur"),
      dist: get_dai(node, "eff-is-dist"),
      dir: get_dai(node, "eff-is-dir"),
      color: parse_hex_color(is_clr),
    }
  }
  let gl_rad = get_dai(node, "eff-gl-rad")
  let gl_clr = get_da(node, "eff-gl-clr")
  let gl = if gl_rad == 0 && gl_clr.is_empty() {
    @ooxml.Glow::none()
  } else {
    {
      rad: gl_rad,
      color: if gl_clr.is_empty() {
        @ooxml.Color::none()
      } else {
        parse_hex_color(gl_clr)
      },
    }
  }
  let se_rad = get_dai(node, "eff-se-rad")
  let se : @ooxml.SoftEdge = { rad: se_rad }
  let refl_sta = get_dai(node, "eff-refl-sta")
  let refl = if refl_sta == 0 {
    @ooxml.Reflection::none()
  } else {
    {
      blur_rad: get_dai(node, "eff-refl-blur"),
      dist: get_dai(node, "eff-refl-dist"),
      dir: get_dai(node, "eff-refl-dir"),
      st_alpha: refl_sta,
      end_alpha: get_dai(node, "eff-refl-enda"),
      fade_dir: get_dai(node, "eff-refl-fadedir"),
      sx: get_dai(node, "eff-refl-sx"),
      sy: get_dai(node, "eff-refl-sy"),
      algn: get_da(node, "eff-refl-algn"),
      rot_with_shape: get_da(node, "eff-refl-rot") != "0",
    }
  }
  {
    outer_shadow: os,
    inner_shadow: ishdw,
    glow: gl,
    soft_edge: se,
    reflection: refl,
  }
}

///|
/// Parse 3D scene from data attributes.
fn parse_scene_3d_from_da(node : @xml.XmlNode) -> @ooxml.Scene3d {
  let cam = get_data_attr(node, "3d-cam")
  let rig = get_data_attr(node, "3d-rig")
  let rig_dir = get_data_attr(node, "3d-rig-dir")
  if cam.is_empty() && rig.is_empty() {
    @ooxml.Scene3d::none()
  } else {
    { camera_prst: cam, light_rig: rig, light_dir: rig_dir }
  }
}

///|
/// Parse 3D shape properties from data attributes.
fn parse_sp_3d_from_da(node : @xml.XmlNode) -> @ooxml.Shape3d {
  let bt_w = get_dai(node, "3d-bt-w")
  let bt_h = get_dai(node, "3d-bt-h")
  let bt_prst = get_data_attr(node, "3d-bt-prst")
  let bb_w = get_dai(node, "3d-bb-w")
  let bb_h = get_dai(node, "3d-bb-h")
  let bb_prst = get_data_attr(node, "3d-bb-prst")
  let ext_h = get_dai(node, "3d-ext-h")
  let cnt_w = get_dai(node, "3d-cnt-w")
  let mat = get_data_attr(node, "3d-mat")
  let z_val = get_dai(node, "3d-z")
  if bt_w == 0 && bt_h == 0 && bb_w == 0 && bb_h == 0 && ext_h == 0 && cnt_w == 0 {
    @ooxml.Shape3d::none()
  } else {
    {
      bevel_t: { w: bt_w, h: bt_h, prst: bt_prst },
      bevel_b: { w: bb_w, h: bb_h, prst: bb_prst },
      extrusion_h: ext_h,
      contour_w: cnt_w,
      extrusion_clr: color_or_none(node, "3d-ext-clr"),
      contour_clr: color_or_none(node, "3d-cnt-clr"),
      prst_material: mat,
      z: z_val,
    }
  }
}

// ── Gradient parsing from SVG data attributes ───────────────────────────────

///|
/// Parse gradient stops from "pos,RRGGBB;pos,RRGGBB;..." string.
fn parse_grad_stops(s : String) -> Array[@ooxml.GradientStop] {
  let stops : Array[@ooxml.GradientStop] = []
  if s.is_empty() {
    return stops
  }
  let len = s.length()
  let mut start = 0
  let mut i = 0
  while i <= len {
    let at_end = i == len
    let is_semi = if at_end { false } else { s.get_char(i).unwrap() == ';' }
    if is_semi || at_end {
      if i > start {
        let mut ci = start
        let mut comma = -1
        while ci < i {
          if s.get_char(ci).unwrap() == ',' {
            comma = ci
          }
          ci = ci + 1
        }
        if comma > start {
          let pos_str = @xml.str_substring(s, start, comma)
          let hex_str = @xml.str_substring(s, comma + 1, i)
          let pos = @xml.parse_int(pos_str)
          let color = @ooxml.parse_hex_color(hex_str)
          stops.push({ pos, color })
        }
      }
      start = i + 1
    }
    i = i + 1
  }
  stops
}

///|
/// Parse GradientFill from data-ooxml-grad-* attributes on a <g> node.
fn parse_grad_fill_from_g(g_node : @xml.XmlNode) -> @ooxml.GradientFill {
  let stops_str = get_da(g_node, "grad-stops")
  if stops_str.is_empty() {
    return @ooxml.GradientFill::none()
  }
  let stops = parse_grad_stops(stops_str)
  if stops.length() == 0 {
    return @ooxml.GradientFill::none()
  }
  let angle = get_dai(g_node, "grad-angle")
  let path_type = get_da(g_node, "grad-path")
  let rot_s = get_da(g_node, "grad-rot")
  let rot_with_shape = rot_s != "0"
  let fill_to_l = get_dai(g_node, "grad-ftl")
  let fill_to_t = get_dai(g_node, "grad-ftt")
  let fill_to_r = get_dai(g_node, "grad-ftr")
  let fill_to_b = get_dai(g_node, "grad-ftb")
  let tile_flip = get_da(g_node, "grad-tile-flip")
  {
    stops,
    angle,
    path_type,
    rot_with_shape,
    fill_to_l: if fill_to_l == 0 {
      -1
    } else {
      fill_to_l
    },
    fill_to_t: if fill_to_t == 0 {
      -1
    } else {
      fill_to_t
    },
    fill_to_r: if fill_to_r == 0 {
      -1
    } else {
      fill_to_r
    },
    fill_to_b: if fill_to_b == 0 {
      -1
    } else {
      fill_to_b
    },
    tile_flip,
  }
}

// ── Text run fill helpers ────────────────────────────────────────────────────

///|
/// Parse text gradient fill from tgrad-* data attributes on a run tspan.
fn parse_text_grad_fill(node : @xml.XmlNode) -> @ooxml.GradientFill {
  let stops_str = get_da(node, "tgrad-stops")
  if stops_str.is_empty() {
    return @ooxml.GradientFill::none()
  }
  let stops = parse_grad_stops(stops_str)
  if stops.length() == 0 {
    return @ooxml.GradientFill::none()
  }
  let angle = get_dai(node, "tgrad-angle")
  let path_type = get_da(node, "tgrad-path")
  let rot_s = get_da(node, "tgrad-rot")
  let rot_with_shape = rot_s != "0"
  {
    stops,
    angle,
    path_type,
    rot_with_shape,
    fill_to_l: -1,
    fill_to_t: -1,
    fill_to_r: -1,
    fill_to_b: -1,
    tile_flip: "",
  }
}

///|
/// Parse text pattern fill from tpatt-* data attributes on a run tspan.
fn parse_text_patt_fill(node : @xml.XmlNode) -> @ooxml.PatternFill {
  let prst = get_da(node, "tpatt-prst")
  if prst.is_empty() {
    return @ooxml.PatternFill::none()
  }
  let fg_hex = get_da(node, "tpatt-fg")
  let bg_hex = get_da(node, "tpatt-bg")
  let fg = if fg_hex.is_empty() {
    @ooxml.Color::black()
  } else {
    parse_hex_color(fg_hex)
  }
  let bg = if bg_hex.is_empty() {
    @ooxml.Color::white()
  } else {
    parse_hex_color(bg_hex)
  }
  { prst, fg_color: fg, bg_color: bg }
}

// ── Text parsing from SVG ────────────────────────────────────────────────────

///|
fn parse_text_from_svg(text_node : @xml.XmlNode) -> Array[@ooxml.TextParagraph] {
  let paras : Array[@ooxml.TextParagraph] = []
  let children = @xml.get_children(text_node)
  let n = children.length()
  let mut i = 0
  while i < n {
    let child = children[i]
    match child {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "tspan" {
          // This is a paragraph-level tspan
          let align = get_da(child, "para-align")
          let level = get_dai(child, "para-lvl")
          let runs : Array[@ooxml.TextRun] = []
          let tspan_children = @xml.get_children(child)
          let nc = tspan_children.length()
          let mut j = 0
          while j < nc {
            match tspan_children[j] {
              @xml.XmlNode::Element(rtag, _, _) =>
                if rtag == "tspan" {
                  // Run-level tspan
                  let run_node = tspan_children[j]
                  let text = @xml.get_text_content(run_node)
                  let bold_s = get_da(run_node, "bold")
                  let bold = bold_s == "1"
                  let font_size = get_dai(run_node, "font-size")
                  let color_hex = get_da(run_node, "color")
                  let color = if color_hex.is_empty() {
                    @ooxml.Color::none()
                  } else {
                    parse_hex_color(color_hex)
                  }
                  // Check italic from font-style attribute
                  let italic_s = get_data_attr(run_node, "font-style")
                  let italic = italic_s == "italic"
                  let font_face = get_da(run_node, "run-font")
                  let ea_font = get_da(run_node, "ea-font")
                  let cs_font = get_da(run_node, "cs-font")
                  let sym_font = get_da(run_node, "sym-font")
                  let char_spacing = get_dai(run_node, "char-spacing")
                  let kern = get_dai(run_node, "kern")
                  let cap = get_da(run_node, "cap")
                  let hlink_rid = get_da(run_node, "hlink-rid")
                  let hlink_mouse_over_rid = get_da(run_node, "hlink-hover-rid")
                  let underline = get_da(run_node, "underline")
                  let strike = get_da(run_node, "strike")
                  let baseline = get_dai(run_node, "baseline")
                  runs.push({
                    text,
                    bold,
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
                    effects: @ooxml.EffectList::none(),
                    outline_color: color_or_none(run_node, "outline-color"),
                    outline_w: get_dai(run_node, "outline-w"),
                    text_grad_fill: parse_text_grad_fill(run_node),
                    text_patt_fill: parse_text_patt_fill(run_node),
                    bold_explicit: bold,
                    math_xml: @xml.xml_unescape(get_da(run_node, "math-xml")),
                  })
                }
              @xml.XmlNode::Text(text) =>
                // Direct text content (no run tspan wrapper)
                if not(text.is_empty()) {
                  runs.push({
                    text,
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
                    math_xml: "",
                  })
                }
            }
            j = j + 1
          }
          let bu_font = get_da(child, "bu-font")
          let bu_size = get_dai(child, "bu-size")
          let bu_color_hex = get_da(child, "bu-color")
          let bu_color = if bu_color_hex.is_empty() {
            @ooxml.Color::none()
          } else {
            parse_hex_color(bu_color_hex)
          }
          let rtl = get_da(child, "rtl") == "1"
          paras.push({
            runs,
            align,
            level,
            spc_before: 0,
            spc_after: 0,
            line_spacing: 0,
            mar_l: -1,
            indent: 0,
            bullet: "",
            bullet_auto: "",
            bullet_none: false,
            bullet_font: bu_font,
            bullet_size: bu_size,
            bullet_color: bu_color,
            bullet_img_rid: "",
            tab_stops: [],
            rtl,
          })
        }
      _ => ()
    }
    i = i + 1
  }
  paras
}

// ── Shape parsing from SVG <g> ───────────────────────────────────────────────

///|
fn parse_shape_from_g(g_node : @xml.XmlNode) -> @ooxml.Shape {
  let shape_type = get_da(g_node, "shape-type")
  let geom_str = get_da(g_node, "geom")
  let x = get_dai(g_node, "x")
  let y = get_dai(g_node, "y")
  let cx = get_dai(g_node, "cx")
  let cy = get_dai(g_node, "cy")
  let rot = get_dai(g_node, "rot")
  let fill_hex = get_da(g_node, "fill")
  let stroke_hex = get_da(g_node, "stroke")
  let stroke_w = get_dai(g_node, "stroke-w")
  let stroke_dash = get_da(g_node, "stroke-dash")
  let stroke_cap = get_da(g_node, "stroke-cap")
  let stroke_join = get_da(g_node, "stroke-join")
  let stroke_miter_limit = get_dai(g_node, "stroke-miter-lim")
  let stroke_head_type = get_da(g_node, "stroke-head-type")
  let stroke_head_w = get_da(g_node, "stroke-head-w")
  let stroke_head_len = get_da(g_node, "stroke-head-len")
  let stroke_tail_type = get_da(g_node, "stroke-tail-type")
  let stroke_tail_w = get_da(g_node, "stroke-tail-w")
  let stroke_tail_len = get_da(g_node, "stroke-tail-len")
  let stroke_cmpd = get_da(g_node, "stroke-cmpd")
  let stroke_no_fill = get_da(g_node, "stroke-no-fill") == "1"

  let fill = parse_hex_color(fill_hex)
  let stroke = parse_hex_color(stroke_hex)
  let flip_h = get_da(g_node, "flip-h") == "1"
  let flip_v = get_da(g_node, "flip-v") == "1"
  let transform : @ooxml.ShapeTransform = { x, y, cx, cy, rot, flip_h, flip_v }

  // Determine shape kind
  let kind : @ooxml.ShapeKind = if shape_type == "picture" {
    // For pictures, we don't have the relationship ID in SVG
    // Use an empty string — the original rels will be preserved
    @ooxml.ShapeKind::Picture("")
  } else if shape_type == "chart" {
    // Chart shapes preserved via chart_xml round-trip
    @ooxml.ShapeKind::ChartShape(@ooxml.ChartData::empty())
  } else if shape_type == "table" {
    // Table shapes carry no inline data in SVG for now — use empty table
    let table_data : @ooxml.TableData = {
      col_widths: [], rows: [], style_id: "",
      first_row: false, last_row: false, first_col: false,
      last_col: false, band_row: false, band_col: false,
    }
    @ooxml.ShapeKind::TableShape(table_data)
  } else if shape_type == "group" {
    // Parse group children from inner <g> containing child shape <g> elements
    let grp_children = parse_group_children_from_g(g_node)
    let grp_data : @ooxml.GroupShapeData = {
      ch_off_x: get_dai(g_node, "grp-ch-off-x"),
      ch_off_y: get_dai(g_node, "grp-ch-off-y"),
      ch_ext_cx: get_dai(g_node, "grp-ch-ext-cx"),
      ch_ext_cy: get_dai(g_node, "grp-ch-ext-cy"),
      children: grp_children,
    }
    @ooxml.ShapeKind::GroupShape(grp_data)
  } else {
    let geom = if geom_str == "custGeom" {
      // Custom geometry round-trip
      let cust_gdlst = get_da(g_node, "cust-gdlst")
      let cust_paths = get_da(g_node, "cust-paths")
      let cust_pw = @xml.parse_int(get_da(g_node, "cust-pw"))
      let cust_ph = @xml.parse_int(get_da(g_node, "cust-ph"))
      @ooxml.ShapeGeom::Custom({
        gdlst: cust_gdlst,
        paths: cust_paths,
        path_w: cust_pw,
        path_h: cust_ph,
        rect_l: get_da(g_node, "cust-rl"),
        rect_t: get_da(g_node, "cust-rt"),
        rect_r: get_da(g_node, "cust-rr"),
        rect_b: get_da(g_node, "cust-rb"),
        cxn_lst: get_da(g_node, "cust-cxn"),
      })
    } else {
      let adj_str = get_da(g_node, "cxn-adj")
      if is_connector_prst(geom_str) {
        parse_connector_geom_from_svg(geom_str, adj_str)
      } else if not(adj_str.is_empty()) {
        let avs = parse_adj_string(adj_str)
        @ooxml.parse_geom(geom_str, avs)
      } else {
        @ooxml.parse_geom(geom_str, [])
      }
    }
    @ooxml.ShapeKind::AutoShape(geom)
  }

  // Parse text from <text> children
  let children = @xml.get_children(g_node)
  let mut paragraphs : Array[@ooxml.TextParagraph] = []
  let nc = children.length()
  let mut i = 0
  while i < nc {
    match children[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "text" {
          paragraphs = parse_text_from_svg(children[i])
        }
      _ => ()
    }
    i = i + 1
  }

  // Parse body props from data attributes
  let bp_anchor = get_da(g_node, "anchor")
  let bp_l_ins = get_dai(g_node, "l-ins")
  let bp_t_ins = get_dai(g_node, "t-ins")
  let bp_r_ins = get_dai(g_node, "r-ins")
  let bp_b_ins = get_dai(g_node, "b-ins")
  let bp_auto_fit = get_da(g_node, "auto-fit") == "1"
  // normAutofit: get_dai returns 0 for missing attrs; 0 maps to -1 (inactive).
  // This is safe because font_scale=0 (0%) never occurs in practice, and
  // ln_spc_reduction=0 is functionally identical to -1 (no reduction).
  let bp_font_scale = get_dai(g_node, "font-scale")
  let bp_ln_spc_reduction = get_dai(g_node, "ln-spc-reduction")
  let bp_wrap = get_da(g_node, "wrap")
  let bp_rot = get_dai(g_node, "bp-rot")
  let bp_vert = get_da(g_node, "bp-vert")
  let bp_num_cols = get_dai(g_node, "bp-num-cols")
  let bp_col_spacing = get_dai(g_node, "bp-col-spacing")
  let body_props : @ooxml.BodyProps = {
    anchor: bp_anchor,
    l_ins: if bp_l_ins == 0 {
      -1
    } else {
      bp_l_ins
    },
    t_ins: if bp_t_ins == 0 {
      -1
    } else {
      bp_t_ins
    },
    r_ins: if bp_r_ins == 0 {
      -1
    } else {
      bp_r_ins
    },
    b_ins: if bp_b_ins == 0 {
      -1
    } else {
      bp_b_ins
    },
    auto_fit: bp_auto_fit,
    font_scale: if bp_font_scale == 0 {
      -1
    } else {
      bp_font_scale
    },
    ln_spc_reduction: if bp_ln_spc_reduction == 0 {
      -1
    } else {
      bp_ln_spc_reduction
    },
    wrap: bp_wrap,
    rot: bp_rot,
    vert: bp_vert,
    num_cols: bp_num_cols,
    col_spacing: bp_col_spacing,
    warp_prst: get_da(g_node, "warp-prst"),
    warp_av1: {
      let v = get_dai(g_node, "warp-av1")
      if v == 0 { -1 } else { v }
    },
    warp_av2: {
      let v = get_dai(g_node, "warp-av2")
      if v == 0 { -1 } else { v }
    },
  }
  let ph_type = get_da(g_node, "ph-type")
  let ph_idx = get_dai(g_node, "ph-idx")
  let grad_fill = parse_grad_fill_from_g(g_node)
  // BlipFill from data attributes
  let blip_rid = get_da(g_node, "blip-rid")
  let blip_fill : @ooxml.BlipFill = if blip_rid.is_empty() {
    @ooxml.BlipFill::none()
  } else {
    let blip_stretch = get_da(g_node, "blip-stretch") != "0"
    let blip_tile_tx = get_dai(g_node, "blip-tile-tx")
    let blip_tile_ty = get_dai(g_node, "blip-tile-ty")
    let blip_tile_sx = get_dai(g_node, "blip-tile-sx")
    let blip_tile_sy = get_dai(g_node, "blip-tile-sy")
    let blip_tile_flip = get_da(g_node, "blip-tile-flip")
    let blip_tile_algn = get_da(g_node, "blip-tile-algn")
    let blip_src_l = get_dai(g_node, "blip-src-l")
    let blip_src_t = get_dai(g_node, "blip-src-t")
    let blip_src_r = get_dai(g_node, "blip-src-r")
    let blip_src_b = get_dai(g_node, "blip-src-b")
    {
      rid: blip_rid,
      stretch: blip_stretch,
      src_l: blip_src_l,
      src_t: blip_src_t,
      src_r: blip_src_r,
      src_b: blip_src_b,
      tile_tx: blip_tile_tx,
      tile_ty: blip_tile_ty,
      tile_sx: blip_tile_sx,
      tile_sy: blip_tile_sy,
      tile_flip: blip_tile_flip,
      tile_algn: blip_tile_algn,
      alpha: get_dai(g_node, "blip-alpha"),
      svg_rid: get_da(g_node, "blip-svg-rid"),
      bright: get_dai(g_node, "blip-bright"),
      contrast: get_dai(g_node, "blip-contrast"),
      duotone_1: color_or_none(g_node, "blip-duotone1"),
      duotone_2: color_or_none(g_node, "blip-duotone2"),
      clr_from: color_or_none(g_node, "blip-clr-from"),
      clr_to: color_or_none(g_node, "blip-clr-to"),
    }
  }
  // PatternFill from data attributes
  let patt_prst = get_da(g_node, "patt-prst")
  let patt_fill : @ooxml.PatternFill = if patt_prst.is_empty() {
    @ooxml.PatternFill::none()
  } else {
    let patt_fg_hex = get_da(g_node, "patt-fg")
    let patt_bg_hex = get_da(g_node, "patt-bg")
    let fg = if patt_fg_hex.is_empty() {
      @ooxml.Color::black()
    } else {
      parse_hex_color(patt_fg_hex)
    }
    let bg = if patt_bg_hex.is_empty() {
      @ooxml.Color::white()
    } else {
      parse_hex_color(patt_bg_hex)
    }
    { prst: patt_prst, fg_color: fg, bg_color: bg }
  }
  {
    kind,
    transform,
    fill,
    grad_fill,
    blip_fill,
    patt_fill,
    stroke,
    stroke_w,
    stroke_dash,
    stroke_cap,
    stroke_join,
    stroke_miter_limit,
    stroke_head_type,
    stroke_head_w,
    stroke_head_len,
    stroke_tail_type,
    stroke_tail_w,
    stroke_tail_len,
    stroke_cmpd,
    stroke_no_fill,
    stroke_grad_fill: @ooxml.GradientFill::none(),
    stroke_patt_fill: @ooxml.PatternFill::none(),
    paragraphs,
    body_props,
    ph_type,
    ph_idx,
    st_cxn_id: get_dai(g_node, "st-cxn-id"),
    st_cxn_idx: get_dai(g_node, "st-cxn-idx"),
    end_cxn_id: get_dai(g_node, "end-cxn-id"),
    end_cxn_idx: get_dai(g_node, "end-cxn-idx"),
    sh_link_rid: get_da(g_node, "sh-link-rid"),
    sh_link_hover_rid: get_da(g_node, "sh-link-hover-rid"),
    mc_choice_xml: @xml.xml_unescape(get_da(g_node, "mc-choice")),
    ole_xml: @xml.xml_unescape(get_da(g_node, "ole")),
    effects: parse_effects_from_da(g_node),
    scene_3d: parse_scene_3d_from_da(g_node),
    sp_3d: parse_sp_3d_from_da(g_node),
    lst_style: @ooxml.TextStyleGroup::empty(),
  }
}

///|
/// Parse group children from the inner transform <g> of a group shape.
fn parse_group_children_from_g(group_g : @xml.XmlNode) -> Array[@ooxml.Shape] {
  let children : Array[@ooxml.Shape] = []
  let g_children = @xml.get_children(group_g)
  let n = g_children.length()
  let mut i = 0
  while i < n {
    match g_children[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "g" {
          // Look inside: could be the transform <g> or a child shape <g>
          let st = get_da(g_children[i], "shape-type")
          if not(st.is_empty()) {
            children.push(parse_shape_from_g(g_children[i]))
          } else {
            // This might be the inner transform <g> wrapping children
            let inner_ch = @xml.get_children(g_children[i])
            let ni = inner_ch.length()
            let mut j = 0
            while j < ni {
              match inner_ch[j] {
                @xml.XmlNode::Element(tag2, _, _) =>
                  if tag2 == "g" {
                    let st2 = get_da(inner_ch[j], "shape-type")
                    if not(st2.is_empty()) {
                      children.push(parse_shape_from_g(inner_ch[j]))
                    }
                  }
                _ => ()
              }
              j = j + 1
            }
          }
        }
      _ => ()
    }
    i = i + 1
  }
  children
}

///|
fn is_connector_prst(prst : String) -> Bool {
  let len = prst.length()
  (len >= 14 && @xml.str_substring(prst, 0, 4) == "bent") ||
  (len >= 16 && @xml.str_substring(prst, 0, 6) == "curved")
}

///|
/// Parse connector geometry from geom string + adjustment data attribute.
fn parse_connector_geom_from_svg(
  geom_str : String,
  adj_str : String,
) -> @ooxml.ShapeGeom {
  if adj_str.is_empty() {
    @ooxml.parse_connector_geom(geom_str, [])
  } else {
    let avs = parse_adj_string(adj_str)
    @ooxml.parse_connector_geom(geom_str, avs)
  }
}

///|
/// Parse comma-separated adjustment values.
fn parse_adj_string(s : String) -> Array[Int] {
  let result : Array[Int] = []
  let mut start = 0
  let len = s.length()
  let mut i = 0
  while i < len {
    if s.get_char(i).unwrap() == ',' {
      if i > start {
        result.push(@xml.parse_int(@xml.str_substring(s, start, i)))
      }
      start = i + 1
    }
    i = i + 1
  }
  if start < len {
    result.push(@xml.parse_int(@xml.str_substring(s, start, len)))
  }
  result
}

// ── Public API ────────────────────────────────────────────────────────────────

///|
/// Parse an SVG string (with data-ooxml-* attributes) back into SlideData.
pub fn parse_svg_to_slide(svg_string : String) -> @ooxml.SlideData {
  let nodes = @xml.parse_xml(svg_string)
  // Find the <svg> root
  let svg_node = match @xml.find_child(nodes, "svg") {
    Some(n) => n
    None =>
      return {
        slide_size: @ooxml.SlideSize::default(),
        background: @ooxml.Color::white(),
        bg_grad: @ooxml.GradientFill::none(), bg_blip_fill: @ooxml.BlipFill::none(), bg_patt_fill: @ooxml.PatternFill::none(),
        shapes: [],
        transition_xml: "", timing_xml: "", hidden: false,
      }
  }
  // Read slide metadata from SVG root
  let slide_cx = get_dai(svg_node, "slide-cx")
  let slide_cy = get_dai(svg_node, "slide-cy")
  let bg_hex = get_da(svg_node, "bg")
  let hidden = get_da(svg_node, "hidden") == "1"
  let transition_xml = get_da(svg_node, "transition-xml")
  let timing_xml = get_da(svg_node, "timing-xml")

  let slide_size : @ooxml.SlideSize = if slide_cx > 0 && slide_cy > 0 {
    { cx: slide_cx, cy: slide_cy }
  } else {
    @ooxml.SlideSize::default()
  }

  let background = if bg_hex.is_empty() {
    @ooxml.Color::white()
  } else {
    parse_hex_color(bg_hex)
  }

  // Parse shapes from <g data-ooxml-shape-type="..."> elements
  let shapes : Array[@ooxml.Shape] = []
  let svg_children = @xml.get_children(svg_node)
  let n = svg_children.length()
  let mut i = 0
  while i < n {
    match svg_children[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "g" {
          let st = get_da(svg_children[i], "shape-type")
          if not(st.is_empty()) {
            shapes.push(parse_shape_from_g(svg_children[i]))
          }
        }
      _ => ()
    }
    i = i + 1
  }

  { slide_size, background, bg_grad: @ooxml.GradientFill::none(), bg_blip_fill: @ooxml.BlipFill::none(), bg_patt_fill: @ooxml.PatternFill::none(), shapes, transition_xml, timing_xml, hidden }
}
