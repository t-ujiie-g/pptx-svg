// OOXML Shape / Slide / Fill parsing
//
// Same package as ooxml.mbt — all types and helpers are shared.

// ── Attribute helpers ─────────────────────────────────────────────────────────

///|
fn attr_int(node : @xml.XmlNode, name : String) -> Int {
  match @xml.get_attr(node, name) {
    Some(v) => @xml.parse_int(v)
    None => 0
  }
}

///|
fn child_attr(
  nodes : Array[@xml.XmlNode],
  tag : String,
  attr : String,
) -> String {
  match @xml.find_child(nodes, tag) {
    None => ""
    Some(node) =>
      match @xml.get_attr(node, attr) {
        Some(v) => v
        None => ""
      }
  }
}

// ── Color parsing ─────────────────────────────────────────────────────────────

///|
/// Parse a color from direct children (a:srgbClr or a:schemeClr), without a:solidFill wrapper.
///  Used for e.g. a:buClr which contains color child directly.
fn parse_color_from_children(
  children : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Color {
  match @xml.find_child(children, "a:srgbClr") {
    Some(clr) =>
      match @xml.get_attr(clr, "val") {
        Some(v) => {
          let base = parse_hex_color(v)
          apply_color_modifiers(base, @xml.get_children(clr))
        }
        None => Color::none()
      }
    None =>
      match @xml.find_child(children, "a:schemeClr") {
        Some(clr) =>
          match @xml.get_attr(clr, "val") {
            Some(v) => {
              let base = resolve_scheme_color(v, theme)
              apply_color_modifiers(base, @xml.get_children(clr))
            }
            None => Color::none()
          }
        None => Color::none()
      }
  }
}

///|
fn parse_solid_fill(nodes : Array[@xml.XmlNode], theme : ThemeData) -> Color {
  match @xml.find_child(nodes, "a:solidFill") {
    None => Color::none()
    Some(fill_node) =>
      parse_color_from_children(@xml.get_children(fill_node), theme)
  }
}

// ── Gradient fill parsing ────────────────────────────────────────────────────

///|
/// Parse <a:gradFill> from child nodes. Returns GradientFill::none() if absent.
fn parse_gradient_fill(
  nodes : Array[@xml.XmlNode],
  theme : ThemeData,
) -> GradientFill {
  match @xml.find_child(nodes, "a:gradFill") {
    None => GradientFill::none()
    Some(gf_node) => {
      let rot_with_shape = match @xml.get_attr(gf_node, "rotWithShape") {
        Some(v) => v != "0"
        None => true
      }
      let gf_ch = @xml.get_children(gf_node)
      // Parse gradient stops
      let stops : Array[GradientStop] = []
      match @xml.find_child(gf_ch, "a:gsLst") {
        None => ()
        Some(gs_lst) => {
          let gs_nodes = @xml.find_children(@xml.get_children(gs_lst), "a:gs")
          let ngs = gs_nodes.length()
          let mut gi = 0
          while gi < ngs {
            let pos = attr_int(gs_nodes[gi], "pos")
            let color = parse_color_from_children(
              @xml.get_children(gs_nodes[gi]),
              theme,
            )
            stops.push({ pos, color })
            gi = gi + 1
          }
        }
      }
      if stops.length() == 0 {
        return GradientFill::none()
      }
      // Parse linear gradient: <a:lin ang="N" scaled="0|1"/>
      let mut angle = 0
      let mut path_type = ""
      let mut fill_to_l = -1
      let mut fill_to_t = -1
      let mut fill_to_r = -1
      let mut fill_to_b = -1
      match @xml.find_child(gf_ch, "a:lin") {
        Some(lin) => angle = attr_int(lin, "ang")
        None =>
          // Parse path gradient: <a:path path="circle|rect|shape">
          match @xml.find_child(gf_ch, "a:path") {
            Some(path_node) => {
              angle = -1
              path_type = match @xml.get_attr(path_node, "path") {
                Some(v) => v
                None => "circle"
              }
              // Parse <a:fillToRect l="N" t="N" r="N" b="N"/>
              match
                @xml.find_child(@xml.get_children(path_node), "a:fillToRect") {
                Some(ftr) => {
                  fill_to_l = attr_int(ftr, "l")
                  fill_to_t = attr_int(ftr, "t")
                  fill_to_r = attr_int(ftr, "r")
                  fill_to_b = attr_int(ftr, "b")
                }
                None => ()
              }
            }
            None => ()
          }
      }
      let tile_flip = match @xml.get_attr(gf_node, "tileFlip") {
        Some(v) => v
        None => ""
      }
      {
        stops,
        angle,
        path_type,
        rot_with_shape,
        fill_to_l,
        fill_to_t,
        fill_to_r,
        fill_to_b,
        tile_flip,
      }
    }
  }
}

// ── Transform parsing ─────────────────────────────────────────────────────────

///|
/// Parse transform from an xfrm element node (either a:xfrm or p:xfrm).
fn parse_xfrm_node(xfrm : @xml.XmlNode) -> ShapeTransform {
  let rot = attr_int(xfrm, "rot")
  let flip_h_s = match @xml.get_attr(xfrm, "flipH") {
    Some(v) => v == "1" || v == "true"
    None => false
  }
  let flip_v_s = match @xml.get_attr(xfrm, "flipV") {
    Some(v) => v == "1" || v == "true"
    None => false
  }
  let ch = @xml.get_children(xfrm)
  let x = @xml.parse_int(child_attr(ch, "a:off", "x"))
  let y = @xml.parse_int(child_attr(ch, "a:off", "y"))
  let cx = @xml.parse_int(child_attr(ch, "a:ext", "cx"))
  let cy = @xml.parse_int(child_attr(ch, "a:ext", "cy"))
  { x, y, cx, cy, rot, flip_h: flip_h_s, flip_v: flip_v_s }
}

///|
/// Parse transform from a:xfrm child within spPr children.
fn parse_xfrm(sppr_children : Array[@xml.XmlNode]) -> ShapeTransform {
  match @xml.find_child(sppr_children, "a:xfrm") {
    None => ShapeTransform::default()
    Some(xfrm) => parse_xfrm_node(xfrm)
  }
}

// ── Stroke parsing ────────────────────────────────────────────────────────────

///|
fn parse_stroke(
  sppr_children : Array[@xml.XmlNode],
  theme : ThemeData,
) -> StrokeProps {
  match @xml.find_child(sppr_children, "a:ln") {
    None => StrokeProps::default()
    Some(ln) => {
      let w = attr_int(ln, "w")
      let cap = match @xml.get_attr(ln, "cap") {
        Some(v) => v
        None => ""
      }
      let cmpd = match @xml.get_attr(ln, "cmpd") {
        Some(v) => v
        None => ""
      }
      let ln_ch = @xml.get_children(ln)
      let no_fill = match @xml.find_child(ln_ch, "a:noFill") {
        Some(_) => true
        None => false
      }
      let ln_gf_early = parse_gradient_fill(ln_ch, theme)
      let ln_pf_early = parse_pattern_fill(ln_ch, theme)
      let has_alt_fill = not(ln_gf_early.is_none()) || not(ln_pf_early.is_none())
      let color = if no_fill {
        Color::none()
      } else {
        let clr = parse_solid_fill(ln_ch, theme)
        if clr.is_none() {
          if has_alt_fill { Color::none() } else { Color::black() }
        } else {
          clr
        }
      }
      // Preset dash or custom dash
      let mut dash = match @xml.find_child(ln_ch, "a:prstDash") {
        None => ""
        Some(d) =>
          match @xml.get_attr(d, "val") {
            Some(v) => v
            None => ""
          }
      }
      if dash.is_empty() {
        match @xml.find_child(ln_ch, "a:custDash") {
          None => ()
          Some(cd) => {
            // a:custDash contains a:ds children with d (dash) and sp (space) percentages
            // Format: "custDash:d1,sp1;d2,sp2;..." (percentages of line width * 100)
            let ds_nodes = @xml.find_children(@xml.get_children(cd), "a:ds")
            let nds = ds_nodes.length()
            let mut cd_str = cust_dash_prefix
            let mut di = 0
            while di < nds {
              if di > 0 {
                cd_str = cd_str + ";"
              }
              let d_val = attr_int(ds_nodes[di], "d")
              let sp_val = attr_int(ds_nodes[di], "sp")
              cd_str = cd_str +
                @xml.int_to_str(d_val) +
                "," +
                @xml.int_to_str(sp_val)
              di = di + 1
            }
            if nds > 0 {
              dash = cd_str
            }
          }
        }
      }
      // Line join
      let mut join = ""
      let mut miter_limit = 0
      match @xml.find_child(ln_ch, "a:round") {
        Some(_) => join = "round"
        None => ()
      }
      if join.is_empty() {
        match @xml.find_child(ln_ch, "a:bevel") {
          Some(_) => join = "bevel"
          None => ()
        }
      }
      if join.is_empty() {
        match @xml.find_child(ln_ch, "a:miter") {
          Some(m) => {
            join = "miter"
            miter_limit = attr_int(m, "lim")
          }
          None => ()
        }
      }
      // Arrow head
      let mut head_type = ""
      let mut head_w = ""
      let mut head_len = ""
      match @xml.find_child(ln_ch, "a:headEnd") {
        Some(he) => {
          head_type = match @xml.get_attr(he, "type") {
            Some(v) => v
            None => ""
          }
          head_w = match @xml.get_attr(he, "w") {
            Some(v) => v
            None => ""
          }
          head_len = match @xml.get_attr(he, "len") {
            Some(v) => v
            None => ""
          }
        }
        None => ()
      }
      // Arrow tail
      let mut tail_type = ""
      let mut tail_w = ""
      let mut tail_len = ""
      match @xml.find_child(ln_ch, "a:tailEnd") {
        Some(te) => {
          tail_type = match @xml.get_attr(te, "type") {
            Some(v) => v
            None => ""
          }
          tail_w = match @xml.get_attr(te, "w") {
            Some(v) => v
            None => ""
          }
          tail_len = match @xml.get_attr(te, "len") {
            Some(v) => v
            None => ""
          }
        }
        None => ()
      }
      {
        color,
        width: w,
        dash,
        cap,
        join,
        miter_limit,
        head_type,
        head_w,
        head_len,
        tail_type,
        tail_w,
        tail_len,
        cmpd,
        no_fill,
        grad_fill: ln_gf_early,
        patt_fill: ln_pf_early,
      }
    }
  }
}

// ── Effect list parsing ──────────────────────────────────────────────────────

///|
/// Parse a:effectLst from shape property children.
///|
fn parse_outer_shadow(node : @xml.XmlNode, theme : ThemeData) -> OuterShadow {
  let ch = @xml.get_children(node)
  let color = parse_color_from_children(ch, theme)
  {
    blur_rad: attr_int(node, "blurRad"),
    dist: attr_int(node, "dist"),
    dir: attr_int(node, "dir"),
    color,
    sx: attr_int(node, "sx"),
    sy: attr_int(node, "sy"),
    algn: match @xml.get_attr(node, "algn") {
      Some(v) => v
      None => ""
    },
    rot_with_shape: match @xml.get_attr(node, "rotWithShape") {
      Some(v) => v != "0"
      None => true
    },
  }
}

///|
fn parse_inner_shadow(node : @xml.XmlNode, theme : ThemeData) -> InnerShadow {
  let ch = @xml.get_children(node)
  let color = parse_color_from_children(ch, theme)
  {
    blur_rad: attr_int(node, "blurRad"),
    dist: attr_int(node, "dist"),
    dir: attr_int(node, "dir"),
    color,
  }
}

///|
fn parse_glow(node : @xml.XmlNode, theme : ThemeData) -> Glow {
  let ch = @xml.get_children(node)
  let color = parse_color_from_children(ch, theme)
  { rad: attr_int(node, "rad"), color }
}

fn parse_effect_lst(
  sppr_ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> EffectList {
  match @xml.find_child(sppr_ch, "a:effectLst") {
    None => EffectList::none()
    Some(eff) => {
      let eff_ch = @xml.get_children(eff)
      let os = match @xml.find_child(eff_ch, "a:outerShdw") {
        None => OuterShadow::none()
        Some(node) => parse_outer_shadow(node, theme)
      }
      let ishdw = match @xml.find_child(eff_ch, "a:innerShdw") {
        None => InnerShadow::none()
        Some(node) => parse_inner_shadow(node, theme)
      }
      let gl = match @xml.find_child(eff_ch, "a:glow") {
        None => Glow::none()
        Some(node) => parse_glow(node, theme)
      }
      let se = match @xml.find_child(eff_ch, "a:softEdge") {
        None => SoftEdge::none()
        Some(node) => { rad: attr_int(node, "rad") }
      }
      let refl = match @xml.find_child(eff_ch, "a:reflection") {
        None => Reflection::none()
        Some(node) =>
          {
            blur_rad: attr_int(node, "blurRad"),
            dist: attr_int(node, "dist"),
            dir: attr_int(node, "dir"),
            st_alpha: attr_int(node, "stA"),
            end_alpha: attr_int(node, "endA"),
            fade_dir: attr_int(node, "fadeDir"),
            sx: attr_int(node, "sx"),
            sy: attr_int(node, "sy"),
            algn: match @xml.get_attr(node, "algn") {
              Some(v) => v
              None => ""
            },
            rot_with_shape: match @xml.get_attr(node, "rotWithShape") {
              Some(v) => v != "0"
              None => true
            },
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
  }
}

// ── 3D parsing ──────────────────────────────────────────────────────────────

///|
fn parse_scene_3d(sppr_ch : Array[@xml.XmlNode]) -> Scene3d {
  match @xml.find_child(sppr_ch, "a:scene3d") {
    None => Scene3d::none()
    Some(scene) => {
      let sch = @xml.get_children(scene)
      let (cam_prst, rig, rig_dir) = match @xml.find_child(sch, "a:camera") {
        None => ("", "", "")
        Some(cam) => {
          let cp = match @xml.get_attr(cam, "prst") {
            Some(v) => v
            None => ""
          }
          let lr = match @xml.find_child(sch, "a:lightRig") {
            None => ("", "")
            Some(lr_node) => {
              let r = match @xml.get_attr(lr_node, "rig") {
                Some(v) => v
                None => ""
              }
              let d = match @xml.get_attr(lr_node, "dir") {
                Some(v) => v
                None => ""
              }
              (r, d)
            }
          }
          (cp, lr.0, lr.1)
        }
      }
      { camera_prst: cam_prst, light_rig: rig, light_dir: rig_dir }
    }
  }
}

///|
fn parse_bevel(parent_ch : Array[@xml.XmlNode], tag : String) -> Bevel {
  match @xml.find_child(parent_ch, tag) {
    None => Bevel::none()
    Some(b) => {
      let w = match @xml.get_attr(b, "w") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let h = match @xml.get_attr(b, "h") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let prst = match @xml.get_attr(b, "prst") {
        Some(v) => v
        None => ""
      }
      { w, h, prst }
    }
  }
}

///|
fn parse_sp_3d(sppr_ch : Array[@xml.XmlNode]) -> Shape3d {
  match @xml.find_child(sppr_ch, "a:sp3d") {
    None => Shape3d::none()
    Some(sp3d) => {
      let sp_ch = @xml.get_children(sp3d)
      let ext_h = match @xml.get_attr(sp3d, "extrusionH") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let cnt_w = match @xml.get_attr(sp3d, "contourW") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let z_val = match @xml.get_attr(sp3d, "z") {
        Some(v) => @xml.parse_int(v)
        None => 0
      }
      let mat = match @xml.get_attr(sp3d, "prstMaterial") {
        Some(v) => v
        None => ""
      }
      let ext_clr = match @xml.find_child(sp_ch, "a:extrusionClr") {
        None => Color::none()
        Some(ec) => parse_first_color(@xml.get_children(ec))
      }
      let cnt_clr = match @xml.find_child(sp_ch, "a:contourClr") {
        None => Color::none()
        Some(cc) => parse_first_color(@xml.get_children(cc))
      }
      {
        bevel_t: parse_bevel(sp_ch, "a:bevelT"),
        bevel_b: parse_bevel(sp_ch, "a:bevelB"),
        extrusion_h: ext_h,
        contour_w: cnt_w,
        extrusion_clr: ext_clr,
        contour_clr: cnt_clr,
        prst_material: mat,
        z: z_val,
      }
    }
  }
}

// ── Blip color helpers ───────────────────────────────────────────────────────

///|
/// Parse a single color from children (a:srgbClr only, no theme needed).
fn parse_first_color(children : Array[@xml.XmlNode]) -> Color {
  match @xml.find_child(children, "a:srgbClr") {
    Some(clr) =>
      match @xml.get_attr(clr, "val") {
        Some(v) => parse_hex_color(v)
        None => Color::none()
      }
    None => Color::none()
  }
}

///|
/// Parse a list of colors from children (for a:duotone which has 2 color elements).
fn parse_color_list(children : Array[@xml.XmlNode]) -> Array[Color] {
  let colors : Array[Color] = []
  let n = children.length()
  let mut i = 0
  while i < n {
    match children[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "a:srgbClr" {
          match @xml.get_attr(children[i], "val") {
            Some(v) => colors.push(parse_hex_color(v))
            None => ()
          }
        } else if tag == "a:schemeClr" {
          // For duotone with scheme colors, use a placeholder (rare case)
          colors.push(Color::none())
        }
      _ => ()
    }
    i = i + 1
  }
  colors
}

// ── Blip fill parser ──────────────────────────────────────────────────────────

///|
/// Parse blipFill node internals → BlipFill (shared by a:blipFill and p:blipFill).
fn parse_blip_fill_node(bf_node : @xml.XmlNode) -> BlipFill {
  let bf_ch = @xml.get_children(bf_node)
  let mut rid = ""
  let mut alpha = 0
  let mut svg_rid = ""
  let mut bright = 0
  let mut contrast = 0
  let mut duotone_1 = Color::none()
  let mut duotone_2 = Color::none()
  let mut clr_from = Color::none()
  let mut clr_to = Color::none()
  match @xml.find_child(bf_ch, "a:blip") {
    None => ()
    Some(blip) => {
      rid = match @xml.get_attr(blip, "r:embed") {
        Some(v) => v
        None => ""
      }
      let blip_ch = @xml.get_children(blip)
      // Parse a:alphaModFix
      match @xml.find_child(blip_ch, "a:alphaModFix") {
        Some(amf) => alpha = attr_int(amf, "amt")
        None => ()
      }
      // Parse a:lum (brightness/contrast)
      match @xml.find_child(blip_ch, "a:lum") {
        Some(lum) => {
          bright = attr_int(lum, "bright")
          contrast = attr_int(lum, "contrast")
        }
        None => ()
      }
      // Parse a:duotone — contains 2 color children
      match @xml.find_child(blip_ch, "a:duotone") {
        Some(dt) => {
          let dt_ch = @xml.get_children(dt)
          let colors = parse_color_list(dt_ch)
          if colors.length() >= 2 {
            duotone_1 = colors[0]
            duotone_2 = colors[1]
          }
        }
        None => ()
      }
      // Parse a:clrChange — from/to color
      match @xml.find_child(blip_ch, "a:clrChange") {
        Some(cc) => {
          let cc_ch = @xml.get_children(cc)
          match @xml.find_child(cc_ch, "a:clrFrom") {
            Some(from_node) =>
              clr_from = parse_first_color(@xml.get_children(from_node))
            None => ()
          }
          match @xml.find_child(cc_ch, "a:clrTo") {
            Some(to_node) =>
              clr_to = parse_first_color(@xml.get_children(to_node))
            None => ()
          }
        }
        None => ()
      }
      // Parse asvg:svgBlip from a:extLst
      match @xml.find_child(blip_ch, "a:extLst") {
        Some(ext_lst) => {
          let exts = @xml.get_children(ext_lst)
          let n = exts.length()
          let mut i = 0
          while i < n {
            let ext_ch = @xml.get_children(exts[i])
            // Look for asvg:svgBlip or any svgBlip tag
            let m = ext_ch.length()
            let mut j = 0
            while j < m {
              match ext_ch[j] {
                @xml.XmlNode::Element(tag, _, _) =>
                  if tag == "asvg:svgBlip" || tag == "a16:svgBlip" {
                    svg_rid = match @xml.get_attr(ext_ch[j], "r:embed") {
                      Some(v) => v
                      None => ""
                    }
                  }
                _ => ()
              }
              j = j + 1
            }
            i = i + 1
          }
        }
        None => ()
      }
    }
  }
  if rid.is_empty() {
    return BlipFill::none()
  }
  let stretch = match @xml.find_child(bf_ch, "a:stretch") {
    Some(_) => true
    None => false
  }
  let mut src_l = 0
  let mut src_t = 0
  let mut src_r = 0
  let mut src_b = 0
  match @xml.find_child(bf_ch, "a:srcRect") {
    Some(sr) => {
      src_l = attr_int(sr, "l")
      src_t = attr_int(sr, "t")
      src_r = attr_int(sr, "r")
      src_b = attr_int(sr, "b")
    }
    None => ()
  }
  // Parse a:tile element
  let mut tile_tx = 0
  let mut tile_ty = 0
  let mut tile_sx = 0
  let mut tile_sy = 0
  let mut tile_flip = ""
  let mut tile_algn = ""
  match @xml.find_child(bf_ch, "a:tile") {
    Some(tile) => {
      tile_tx = attr_int(tile, "tx")
      tile_ty = attr_int(tile, "ty")
      tile_sx = attr_int(tile, "sx")
      tile_sy = attr_int(tile, "sy")
      tile_flip = match @xml.get_attr(tile, "flip") {
        Some(v) => v
        None => ""
      }
      tile_algn = match @xml.get_attr(tile, "algn") {
        Some(v) => v
        None => ""
      }
    }
    None => ()
  }
  {
    rid,
    stretch,
    src_l,
    src_t,
    src_r,
    src_b,
    tile_tx,
    tile_ty,
    tile_sx,
    tile_sy,
    tile_flip,
    tile_algn,
    alpha,
    svg_rid,
    bright,
    contrast,
    duotone_1,
    duotone_2,
    clr_from,
    clr_to,
  }
}

///|
/// Parse a:blipFill from spPr children → BlipFill.
fn parse_blip_fill(nodes : Array[@xml.XmlNode]) -> BlipFill {
  match @xml.find_child(nodes, "a:blipFill") {
    None => BlipFill::none()
    Some(bf_node) => parse_blip_fill_node(bf_node)
  }
}

// ── Pattern fill parser ──────────────────────────────────────────────────────

///|
/// Parse a:pattFill from spPr children → PatternFill.
fn parse_pattern_fill(
  nodes : Array[@xml.XmlNode],
  theme : ThemeData,
) -> PatternFill {
  match @xml.find_child(nodes, "a:pattFill") {
    None => PatternFill::none()
    Some(pf_node) => {
      let prst = match @xml.get_attr(pf_node, "prst") {
        Some(v) => v
        None => ""
      }
      if prst.is_empty() {
        return PatternFill::none()
      }
      let pf_ch = @xml.get_children(pf_node)
      let fg_color = match @xml.find_child(pf_ch, "a:fgClr") {
        None => Color::black()
        Some(fg) => parse_color_from_children(@xml.get_children(fg), theme)
      }
      let bg_color = match @xml.find_child(pf_ch, "a:bgClr") {
        None => Color::white()
        Some(bg) => parse_color_from_children(@xml.get_children(bg), theme)
      }
      { prst, fg_color, bg_color }
    }
  }
}

// ── Shape link helpers ────────────────────────────────────────────────────────

///|
/// Parse hyperlinks from p:cNvPr element (a:hlinkClick, a:hlinkHover).
/// Returns (click_rid, hover_rid).
fn parse_cnvpr_links(nv_ch : Array[@xml.XmlNode]) -> (String, String) {
  let mut click_rid = ""
  let mut hover_rid = ""
  match @xml.find_child(nv_ch, "p:cNvPr") {
    None => ()
    Some(cnvpr) => {
      let cnvpr_ch = @xml.get_children(cnvpr)
      match @xml.find_child(cnvpr_ch, "a:hlinkClick") {
        None => ()
        Some(hc) =>
          click_rid = match @xml.get_attr(hc, "r:id") {
            Some(v) => v
            None => ""
          }
      }
      match @xml.find_child(cnvpr_ch, "a:hlinkHover") {
        None => ()
        Some(hh) =>
          hover_rid = match @xml.get_attr(hh, "r:id") {
            Some(v) => v
            None => ""
          }
      }
    }
  }
  (click_rid, hover_rid)
}

// ── Shape parsers ─────────────────────────────────────────────────────────────

///|
fn parse_sp(sp : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(sp)
  // Extract placeholder info from p:nvSpPr/p:nvPr/p:ph
  let mut ph_type = ""
  let mut ph_idx = -1
  let mut sh_link_rid = ""
  let mut sh_link_hover_rid = ""
  match @xml.find_child(ch, "p:nvSpPr") {
    None => ()
    Some(nvsp) => {
      let nvsp_ch = @xml.get_children(nvsp)
      // Parse shape hyperlinks from p:cNvPr
      let (lr, hr) = parse_cnvpr_links(nvsp_ch)
      sh_link_rid = lr
      sh_link_hover_rid = hr
      match @xml.find_child(nvsp_ch, "p:nvPr") {
        None => ()
        Some(nvpr) => {
          let nvpr_ch = @xml.get_children(nvpr)
          match @xml.find_child(nvpr_ch, "p:ph") {
            None => ()
            Some(ph) => {
              ph_type = match @xml.get_attr(ph, "type") {
                Some(v) => v
                None => "body"
              }
              ph_idx = match @xml.get_attr(ph, "idx") {
                Some(v) => @xml.parse_int(v)
                None => -1
              }
            }
          }
        }
      }
    }
  }
  let sppr = match @xml.find_child(ch, "p:spPr") {
    None => return empty_shape()
    Some(n) => n
  }
  let sppr_ch = @xml.get_children(sppr)
  let transform = parse_xfrm(sppr_ch)
  // Geometry
  let geom = match @xml.find_child(sppr_ch, "a:prstGeom") {
    Some(g) =>
      match @xml.get_attr(g, "prst") {
        Some(v) => parse_geom(v, parse_avlst(g))
        None => ShapeGeom::Rect
      }
    None =>
      match @xml.find_child(sppr_ch, "a:custGeom") {
        Some(g) => parse_custom_geom(g)
        None => ShapeGeom::Rect
      }
  }
  // Fill
  let fill = match @xml.find_child(sppr_ch, "a:noFill") {
    Some(_) => Color::none()
    None => {
      let c = parse_solid_fill(sppr_ch, theme)
      if c.is_none() {
        Color::none()
      } else {
        c
      }
    }
  }
  let grad_fill = parse_gradient_fill(sppr_ch, theme)
  let sp = parse_stroke(sppr_ch, theme)
  // Blip fill on AutoShape (a:blipFill inside p:spPr)
  let blip_fill = parse_blip_fill(sppr_ch)
  // Pattern fill
  let patt_fill = parse_pattern_fill(sppr_ch, theme)
  // Text
  let (paras, body_props, txt_lst_style) = match @xml.find_child(ch, "p:txBody") {
    None => ([], BodyProps::default(), TextStyleGroup::empty())
    Some(txb) => parse_text_body(txb, theme)
  }
  {
    kind: ShapeKind::AutoShape(geom),
    transform,
    fill,
    grad_fill,
    blip_fill,
    patt_fill,
    stroke: sp.color,
    stroke_w: sp.width,
    stroke_dash: sp.dash,
    stroke_cap: sp.cap,
    stroke_join: sp.join,
    stroke_miter_limit: sp.miter_limit,
    stroke_head_type: sp.head_type,
    stroke_head_w: sp.head_w,
    stroke_head_len: sp.head_len,
    stroke_tail_type: sp.tail_type,
    stroke_tail_w: sp.tail_w,
    stroke_tail_len: sp.tail_len,
    stroke_cmpd: sp.cmpd,
    stroke_no_fill: sp.no_fill,
    stroke_grad_fill: sp.grad_fill,
    stroke_patt_fill: sp.patt_fill,
    paragraphs: paras,
    body_props,
    ph_type,
    ph_idx,
    st_cxn_id: -1,
    st_cxn_idx: -1,
    end_cxn_id: -1,
    end_cxn_idx: -1,
    sh_link_rid,
    sh_link_hover_rid,
    effects: parse_effect_lst(sppr_ch, theme),
    scene_3d: parse_scene_3d(sppr_ch),
    sp_3d: parse_sp_3d(sppr_ch),
    lst_style: txt_lst_style,
  }
}

///|
fn parse_pic(pic : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(pic)
  let sppr = match @xml.find_child(ch, "p:spPr") {
    None => return empty_shape()
    Some(n) => n
  }
  let sppr_ch = @xml.get_children(sppr)
  let transform = parse_xfrm(sppr_ch)
  let sp = parse_stroke(sppr_ch, theme)
  // Parse p:blipFill (same structure as a:blipFill) for rid, crop, alpha
  let bf = match @xml.find_child(ch, "p:blipFill") {
    None => BlipFill::none()
    Some(bf_node) => parse_blip_fill_node(bf_node)
  }
  // Parse shape hyperlinks from p:nvPicPr/p:cNvPr
  let (sh_link_rid, sh_link_hover_rid) = match @xml.find_child(ch, "p:nvPicPr") {
    None => ("", "")
    Some(nvpic) => parse_cnvpr_links(@xml.get_children(nvpic))
  }
  let mut s = Shape::default_with_blip(ShapeKind::Picture(bf.rid), transform, bf)
  // Apply stroke properties
  s = {
    ..s,
    stroke: sp.color,
    stroke_w: sp.width,
    stroke_dash: sp.dash,
    stroke_cap: sp.cap,
    stroke_join: sp.join,
    stroke_miter_limit: sp.miter_limit,
    stroke_no_fill: sp.no_fill,
  }
  if not(sh_link_rid.is_empty()) || not(sh_link_hover_rid.is_empty()) {
    s = { ..s, sh_link_rid, sh_link_hover_rid }
  }
  s
}

///|
fn parse_tc_border(tcpr_ch : Array[@xml.XmlNode], tag : String, theme : ThemeData) -> (Int, Color) {
  match @xml.find_child(tcpr_ch, tag) {
    None => (0, Color::none())
    Some(ln_node) => {
      let w = attr_int(ln_node, "w")
      let ln_ch = @xml.get_children(ln_node)
      // Check for <a:noFill>
      match @xml.find_child(ln_ch, "a:noFill") {
        Some(_) => (-1, Color::none())
        None => {
          let color = parse_solid_fill(ln_ch, theme)
          (w, color)
        }
      }
    }
  }
}

///|
fn parse_table_cell(tc : @xml.XmlNode, theme : ThemeData) -> TableCell {
  let ch = @xml.get_children(tc)
  // Parse cell merge attributes
  let grid_span = { let v = attr_int(tc, "gridSpan"); if v > 0 { v } else { 1 } }
  let row_span = { let v = attr_int(tc, "rowSpan"); if v > 0 { v } else { 1 } }
  let v_merge = match @xml.get_attr(tc, "vMerge") {
    Some("1") => true
    _ => false
  }
  // Parse cell text from <a:txBody>
  let (paragraphs, _, _) = match @xml.find_child(ch, "a:txBody") {
    None => ([], BodyProps::default(), TextStyleGroup::empty())
    Some(txb) => parse_text_body(txb, theme)
  }
  // Parse cell properties from <a:tcPr>
  let mut fill = Color::none()
  let mut grad_fill = GradientFill::none()
  let mut mar_l = -1
  let mut mar_r = -1
  let mut mar_t = -1
  let mut mar_b = -1
  let mut anchor = ""
  let mut bdr_l_w = 0
  let mut bdr_r_w = 0
  let mut bdr_t_w = 0
  let mut bdr_b_w = 0
  let mut bdr_l_color = Color::none()
  let mut bdr_r_color = Color::none()
  let mut bdr_t_color = Color::none()
  let mut bdr_b_color = Color::none()
  let mut bdr_tl_br_w = 0
  let mut bdr_tl_br_color = Color::none()
  let mut bdr_bl_tr_w = 0
  let mut bdr_bl_tr_color = Color::none()
  match @xml.find_child(ch, "a:tcPr") {
    None => ()
    Some(tcpr) => {
      let tcpr_ch = @xml.get_children(tcpr)
      // Fill
      let gf = parse_gradient_fill(tcpr_ch, theme)
      if not(gf.is_none()) {
        grad_fill = gf
      } else {
        fill = parse_solid_fill(tcpr_ch, theme)
      }
      // Margins
      let ml = attr_int(tcpr, "marL")
      let mr = attr_int(tcpr, "marR")
      let mt = attr_int(tcpr, "marT")
      let mb = attr_int(tcpr, "marB")
      if ml > 0 { mar_l = ml }
      if mr > 0 { mar_r = mr }
      if mt > 0 { mar_t = mt }
      if mb > 0 { mar_b = mb }
      // Anchor
      anchor = match @xml.get_attr(tcpr, "anchor") {
        Some(v) => v
        None => ""
      }
      // Borders (a:lnL, a:lnR, a:lnT, a:lnB)
      let (lw, lc) = parse_tc_border(tcpr_ch, "a:lnL", theme)
      bdr_l_w = lw
      bdr_l_color = lc
      let (rw, rc) = parse_tc_border(tcpr_ch, "a:lnR", theme)
      bdr_r_w = rw
      bdr_r_color = rc
      let (tw, tc_) = parse_tc_border(tcpr_ch, "a:lnT", theme)
      bdr_t_w = tw
      bdr_t_color = tc_
      let (bw, bc) = parse_tc_border(tcpr_ch, "a:lnB", theme)
      bdr_b_w = bw
      bdr_b_color = bc
      // Diagonal borders
      let (tlbr_w, tlbr_c) = parse_tc_border(tcpr_ch, "a:lnTlToBr", theme)
      bdr_tl_br_w = tlbr_w
      bdr_tl_br_color = tlbr_c
      let (bltr_w, bltr_c) = parse_tc_border(tcpr_ch, "a:lnBlToTr", theme)
      bdr_bl_tr_w = bltr_w
      bdr_bl_tr_color = bltr_c
    }
  }
  {
    paragraphs, fill, grad_fill,
    grid_span, row_span, v_merge,
    bdr_l_w, bdr_r_w, bdr_t_w, bdr_b_w,
    bdr_l_color, bdr_r_color, bdr_t_color, bdr_b_color,
    bdr_tl_br_w, bdr_tl_br_color, bdr_bl_tr_w, bdr_bl_tr_color,
    mar_l, mar_r, mar_t, mar_b, anchor,
  }
}

///|
/// Parse transform from either a:xfrm or p:xfrm (graphicFrame uses p:xfrm).
fn parse_xfrm_any(nodes : Array[@xml.XmlNode]) -> ShapeTransform {
  match @xml.find_child(nodes, "a:xfrm") {
    Some(xfrm) => parse_xfrm_node(xfrm)
    None =>
      match @xml.find_child(nodes, "p:xfrm") {
        Some(xfrm) => parse_xfrm_node(xfrm)
        None => ShapeTransform::default()
      }
  }
}

///|
fn parse_graphic_frame(gf : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(gf)
  // Transform from <p:xfrm> (direct child of graphicFrame, not inside spPr)
  let transform = parse_xfrm_any(ch)
  // Navigate: <a:graphic> → <a:graphicData>
  let gdata = match @xml.find_child(ch, "a:graphic") {
    None => None
    Some(graphic) =>
      @xml.find_child(@xml.get_children(graphic), "a:graphicData")
  }
  let gdata_node = match gdata {
    None => return empty_shape()
    Some(n) => n
  }
  let gdata_ch = @xml.get_children(gdata_node)
  // Check for chart (c:chart) — external file via relationship
  // Return a placeholder ChartShape with empty data; main.mbt resolves the chart XML.
  match @xml.find_child(gdata_ch, "c:chart") {
    Some(cn) => {
      let rid = match @xml.get_attr(cn, "r:id") {
        Some(v) => v
        None => ""
      }
      if not(rid.is_empty()) {
        // Store rid in chart_xml field temporarily; main.mbt will resolve it.
        let chart_data : ChartData = { ..ChartData::empty(), chart_xml: rid }
        return Shape::default_with(
          ShapeKind::ChartShape(chart_data),
          transform,
        )
      }
    }
    None => ()
  }
  // Check for table (a:tbl)
  let tbl = @xml.find_child(gdata_ch, "a:tbl")
  match tbl {
    None => empty_shape()
    Some(tbl_node) => {
      let tbl_ch = @xml.get_children(tbl_node)
      // Parse table properties from <a:tblPr>
      let mut style_id = ""
      let mut first_row = false
      let mut last_row = false
      let mut first_col = false
      let mut last_col = false
      let mut band_row = false
      let mut band_col = false
      match @xml.find_child(tbl_ch, "a:tblPr") {
        None => ()
        Some(tbl_pr) => {
          style_id = match @xml.get_attr(tbl_pr, "tblStyleId") {
            Some(v) => v
            None => ""
          }
          first_row = attr_int(tbl_pr, "firstRow") == 1
          last_row = attr_int(tbl_pr, "lastRow") == 1
          first_col = attr_int(tbl_pr, "firstCol") == 1
          last_col = attr_int(tbl_pr, "lastCol") == 1
          band_row = attr_int(tbl_pr, "bandRow") == 1
          band_col = attr_int(tbl_pr, "bandCol") == 1
        }
      }
      // Parse column widths from <a:tblGrid>/<a:gridCol w="...">
      let col_widths : Array[Int] = []
      match @xml.find_child(tbl_ch, "a:tblGrid") {
        None => ()
        Some(grid) => {
          let gcols = @xml.find_children(@xml.get_children(grid), "a:gridCol")
          let ng = gcols.length()
          let mut gi = 0
          while gi < ng {
            col_widths.push(attr_int(gcols[gi], "w"))
            gi = gi + 1
          }
        }
      }
      // Parse rows from <a:tr>
      let tr_nodes = @xml.find_children(tbl_ch, "a:tr")
      let rows : Array[TableRow] = []
      let ntr = tr_nodes.length()
      let mut tri = 0
      while tri < ntr {
        let tr = tr_nodes[tri]
        let row_h = attr_int(tr, "h")
        let tc_nodes = @xml.find_children(@xml.get_children(tr), "a:tc")
        let cells : Array[TableCell] = []
        let ntc = tc_nodes.length()
        let mut tci = 0
        while tci < ntc {
          cells.push(parse_table_cell(tc_nodes[tci], theme))
          tci = tci + 1
        }
        rows.push({ height: row_h, cells })
        tri = tri + 1
      }
      let table_data : TableData = {
        col_widths, rows, style_id,
        first_row, last_row, first_col, last_col, band_row, band_col,
      }
      Shape::default_with(ShapeKind::TableShape(table_data), transform)
    }
  }
}

///|
/// Create a Shape with given kind, transform, and stroke properties.
fn shape_with_stroke(
  kind : ShapeKind,
  transform : ShapeTransform,
  sp : StrokeProps,
) -> Shape {
  {
    ..Shape::default_with(kind, transform),
    stroke: sp.color,
    stroke_w: sp.width,
    stroke_dash: sp.dash,
    stroke_cap: sp.cap,
    stroke_join: sp.join,
    stroke_miter_limit: sp.miter_limit,
    stroke_head_type: sp.head_type,
    stroke_head_w: sp.head_w,
    stroke_head_len: sp.head_len,
    stroke_tail_type: sp.tail_type,
    stroke_tail_w: sp.tail_w,
    stroke_tail_len: sp.tail_len,
    stroke_cmpd: sp.cmpd,
    stroke_no_fill: sp.no_fill,
    stroke_grad_fill: sp.grad_fill,
    stroke_patt_fill: sp.patt_fill,
  }
}

///|
/// Parse adjustment values from a:avLst inside a:prstGeom.
fn parse_avlst(geom_node : @xml.XmlNode) -> Array[Int] {
  let avs : Array[Int] = []
  match @xml.find_child(@xml.get_children(geom_node), "a:avLst") {
    None => ()
    Some(avlst) => {
      let gds = @xml.find_children(@xml.get_children(avlst), "a:gd")
      let n = gds.length()
      let mut i = 0
      while i < n {
        // <a:gd name="adj1" fmla="val 50000"/>
        let fmla = match @xml.get_attr(gds[i], "fmla") {
          Some(v) => v
          None => ""
        }
        // Extract value from "val NNNNN"
        if fmla.length() > 4 {
          let val_str = @xml.str_suffix(fmla, 4)
          avs.push(@xml.parse_int(val_str))
        }
        i = i + 1
      }
    }
  }
  avs
}

///|
/// Parse a:custGeom into CustomGeomData.
fn parse_custom_geom(node : @xml.XmlNode) -> ShapeGeom {
  let ch = @xml.get_children(node)
  let mut gdlst = ""
  // Parse a:avLst → guide formulas with "val" operator
  match @xml.find_child(ch, "a:avLst") {
    None => ()
    Some(avlst_node) => {
      let gds = @xml.find_children(@xml.get_children(avlst_node), "a:gd")
      for i in 0..<gds.length() {
        let name = match @xml.get_attr(gds[i], "name") {
          Some(v) => v
          None => continue
        }
        let fmla = match @xml.get_attr(gds[i], "fmla") {
          Some(v) => v
          None => "val 0"
        }
        if gdlst.length() > 0 {
          gdlst = gdlst + ";"
        }
        gdlst = gdlst + name + "=" + fmla
      }
    }
  }
  // Parse a:gdLst → guide formulas
  match @xml.find_child(ch, "a:gdLst") {
    None => ()
    Some(gdlst_node) => {
      let gds = @xml.find_children(@xml.get_children(gdlst_node), "a:gd")
      for i in 0..<gds.length() {
        let name = match @xml.get_attr(gds[i], "name") {
          Some(v) => v
          None => continue
        }
        let fmla = match @xml.get_attr(gds[i], "fmla") {
          Some(v) => v
          None => continue
        }
        if gdlst.length() > 0 {
          gdlst = gdlst + ";"
        }
        gdlst = gdlst + name + "=" + fmla
      }
    }
  }
  // Parse a:pathLst → compact path format
  let mut paths = ""
  let mut path_w = 0
  let mut path_h = 0
  match @xml.find_child(ch, "a:pathLst") {
    None => ()
    Some(pathlst_node) => {
      let path_nodes = @xml.find_children(
        @xml.get_children(pathlst_node),
        "a:path",
      )
      for pi in 0..<path_nodes.length() {
        let pn = path_nodes[pi]
        // Get path coordinate space (first path's values used)
        if pi == 0 {
          path_w = match @xml.get_attr(pn, "w") {
            Some(v) => @xml.parse_int(v)
            None => 0
          }
          path_h = match @xml.get_attr(pn, "h") {
            Some(v) => @xml.parse_int(v)
            None => 0
          }
        }
        let seg = parse_custom_path_segment(pn)
        if seg.length() > 0 {
          if paths.length() > 0 {
            paths = paths + "|"
          }
          paths = paths + seg
        }
      }
    }
  }
  // Parse a:rect → text rectangle guide variable references
  let mut rect_l = ""
  let mut rect_t = ""
  let mut rect_r = ""
  let mut rect_b = ""
  match @xml.find_child(ch, "a:rect") {
    None => ()
    Some(rect_node) => {
      rect_l = match @xml.get_attr(rect_node, "l") {
        Some(v) => v
        None => ""
      }
      rect_t = match @xml.get_attr(rect_node, "t") {
        Some(v) => v
        None => ""
      }
      rect_r = match @xml.get_attr(rect_node, "r") {
        Some(v) => v
        None => ""
      }
      rect_b = match @xml.get_attr(rect_node, "b") {
        Some(v) => v
        None => ""
      }
    }
  }
  // Parse a:cxnLst → connection points: "x,y,ang;x2,y2,ang2;..."
  let mut cxn_lst = ""
  match @xml.find_child(ch, "a:cxnLst") {
    None => ()
    Some(cxn_node) => {
      let cxns = @xml.find_children(@xml.get_children(cxn_node), "a:cxn")
      for ci in 0..<cxns.length() {
        let ang = match @xml.get_attr(cxns[ci], "ang") {
          Some(v) => v
          None => "0"
        }
        let cxn_ch = @xml.get_children(cxns[ci])
        let mut cx = "0"
        let mut cy = "0"
        match @xml.find_child(cxn_ch, "a:pos") {
          None => ()
          Some(pos) => {
            cx = match @xml.get_attr(pos, "x") {
              Some(v) => v
              None => "0"
            }
            cy = match @xml.get_attr(pos, "y") {
              Some(v) => v
              None => "0"
            }
          }
        }
        if cxn_lst.length() > 0 {
          cxn_lst = cxn_lst + ";"
        }
        cxn_lst = cxn_lst + cx + "," + cy + "," + ang
      }
    }
  }
  ShapeGeom::Custom(
    { gdlst, paths, path_w, path_h, rect_l, rect_t, rect_r, rect_b, cxn_lst },
  )
}

///|
/// Parse a single a:path element into compact path format.
fn parse_custom_path_segment(path_node : @xml.XmlNode) -> String {
  let ch = @xml.get_children(path_node)
  let mut d = ""
  for i in 0..<ch.length() {
    let tag = @xml.get_tag(ch[i])
    if tag == "a:moveTo" {
      let pts = @xml.find_children(@xml.get_children(ch[i]), "a:pt")
      if pts.length() >= 1 {
        let x = match @xml.get_attr(pts[0], "x") {
          Some(v) => v
          None => "0"
        }
        let y = match @xml.get_attr(pts[0], "y") {
          Some(v) => v
          None => "0"
        }
        if d.length() > 0 {
          d = d + " "
        }
        d = d + "M " + x + " " + y
      }
    } else if tag == "a:lnTo" {
      let pts = @xml.find_children(@xml.get_children(ch[i]), "a:pt")
      if pts.length() >= 1 {
        let x = match @xml.get_attr(pts[0], "x") {
          Some(v) => v
          None => "0"
        }
        let y = match @xml.get_attr(pts[0], "y") {
          Some(v) => v
          None => "0"
        }
        if d.length() > 0 {
          d = d + " "
        }
        d = d + "L " + x + " " + y
      }
    } else if tag == "a:cubicBezTo" {
      let pts = @xml.find_children(@xml.get_children(ch[i]), "a:pt")
      if pts.length() >= 3 {
        if d.length() > 0 {
          d = d + " "
        }
        d = d + "C"
        for j in 0..<3 {
          let x = match @xml.get_attr(pts[j], "x") {
            Some(v) => v
            None => "0"
          }
          let y = match @xml.get_attr(pts[j], "y") {
            Some(v) => v
            None => "0"
          }
          d = d + " " + x + " " + y
        }
      }
    } else if tag == "a:quadBezTo" {
      let pts = @xml.find_children(@xml.get_children(ch[i]), "a:pt")
      if pts.length() >= 2 {
        if d.length() > 0 {
          d = d + " "
        }
        d = d + "Q"
        for j in 0..<2 {
          let x = match @xml.get_attr(pts[j], "x") {
            Some(v) => v
            None => "0"
          }
          let y = match @xml.get_attr(pts[j], "y") {
            Some(v) => v
            None => "0"
          }
          d = d + " " + x + " " + y
        }
      }
    } else if tag == "a:arcTo" {
      let wr = match @xml.get_attr(ch[i], "wR") {
        Some(v) => v
        None => "0"
      }
      let hr = match @xml.get_attr(ch[i], "hR") {
        Some(v) => v
        None => "0"
      }
      let st = match @xml.get_attr(ch[i], "stAng") {
        Some(v) => v
        None => "0"
      }
      let sw = match @xml.get_attr(ch[i], "swAng") {
        Some(v) => v
        None => "0"
      }
      if d.length() > 0 {
        d = d + " "
      }
      d = d + "A " + wr + " " + hr + " " + st + " " + sw
    } else if tag == "a:close" {
      if d.length() > 0 {
        d = d + " "
      }
      d = d + "Z"
    }
  }
  d
}

///|
/// Parse connector shape — rendered as a line or path with stroke properties.
fn parse_cxn_sp(cxn : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(cxn)
  let sppr = match @xml.find_child(ch, "p:spPr") {
    None => return empty_shape()
    Some(n) => n
  }
  let sppr_ch = @xml.get_children(sppr)
  let transform = parse_xfrm(sppr_ch)
  let sp = parse_stroke(sppr_ch, theme)
  let geom = match @xml.find_child(sppr_ch, "a:prstGeom") {
    None => ShapeGeom::Line
    Some(g) => {
      let prst = match @xml.get_attr(g, "prst") {
        Some(v) => v
        None => "line"
      }
      let avs = parse_avlst(g)
      parse_connector_geom(prst, avs)
    }
  }
  // Parse connection references from p:nvCxnSpPr/p:cNvCxnSpPr
  let mut st_id = -1
  let mut st_idx = -1
  let mut end_id = -1
  let mut end_idx = -1
  match @xml.find_child(ch, "p:nvCxnSpPr") {
    None => ()
    Some(nv) => {
      let nv_ch = @xml.get_children(nv)
      match @xml.find_child(nv_ch, "p:cNvCxnSpPr") {
        None => ()
        Some(cnv) => {
          let cnv_ch = @xml.get_children(cnv)
          match @xml.find_child(cnv_ch, "a:stCxn") {
            None => ()
            Some(st) => {
              st_id = @xml.parse_int(
                match @xml.get_attr(st, "id") {
                  Some(v) => v
                  None => "-1"
                },
              )
              st_idx = @xml.parse_int(
                match @xml.get_attr(st, "idx") {
                  Some(v) => v
                  None => "-1"
                },
              )
            }
          }
          match @xml.find_child(cnv_ch, "a:endCxn") {
            None => ()
            Some(en) => {
              end_id = @xml.parse_int(
                match @xml.get_attr(en, "id") {
                  Some(v) => v
                  None => "-1"
                },
              )
              end_idx = @xml.parse_int(
                match @xml.get_attr(en, "idx") {
                  Some(v) => v
                  None => "-1"
                },
              )
            }
          }
        }
      }
    }
  }
  let base = shape_with_stroke(ShapeKind::AutoShape(geom), transform, sp)
  {
    ..base,
    st_cxn_id: st_id,
    st_cxn_idx: st_idx,
    end_cxn_id: end_id,
    end_cxn_idx: end_idx,
    sh_link_rid: "",
    sh_link_hover_rid: "",
  }
}

///|
fn parse_grp_sp(grp : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(grp)
  // Group shape properties
  let grp_sp_pr = match @xml.find_child(ch, "p:grpSpPr") {
    None => return empty_shape()
    Some(n) => n
  }
  let pr_ch = @xml.get_children(grp_sp_pr)
  // Parse group xfrm: has a:off, a:ext, a:chOff, a:chExt
  let xfrm = match @xml.find_child(pr_ch, "a:xfrm") {
    None => return empty_shape()
    Some(n) => n
  }
  let transform = parse_xfrm_node(xfrm)
  let xfrm_ch = @xml.get_children(xfrm)
  let ch_off_x = @xml.parse_int(child_attr(xfrm_ch, "a:chOff", "x"))
  let ch_off_y = @xml.parse_int(child_attr(xfrm_ch, "a:chOff", "y"))
  let ch_ext_cx = @xml.parse_int(child_attr(xfrm_ch, "a:chExt", "cx"))
  let ch_ext_cy = @xml.parse_int(child_attr(xfrm_ch, "a:chExt", "cy"))
  // Parse child shapes recursively
  let children = parse_shapes_from_children(ch, theme)
  let grp_data : GroupShapeData = {
    ch_off_x,
    ch_off_y,
    ch_ext_cx,
    ch_ext_cy,
    children,
  }
  Shape::default_with(ShapeKind::GroupShape(grp_data), transform)
}

///|
fn empty_shape() -> Shape {
  Shape::default_with(ShapeKind::Other, ShapeTransform::default())
}

// ── Relationship parsing ──────────────────────────────────────────────────────

///|
/// Parse a .rels XML file into a map of relationship ID → target path.
pub fn parse_rels(xml_str : String) -> Array[Relationship] {
  let nodes = @xml.parse_xml(xml_str)
  let rels : Array[Relationship] = []
  // Find <Relationships> root
  match @xml.find_child(nodes, "Relationships") {
    None => ()
    Some(root) => {
      let ch = @xml.get_children(root)
      let rel_nodes = @xml.find_children(ch, "Relationship")
      let n = rel_nodes.length()
      let mut i = 0
      while i < n {
        let rn = rel_nodes[i]
        let id = match @xml.get_attr(rn, "Id") {
          Some(v) => v
          None => ""
        }
        let target = match @xml.get_attr(rn, "Target") {
          Some(v) => v
          None => ""
        }
        let type_ = match @xml.get_attr(rn, "Type") {
          Some(v) => v
          None => ""
        }
        let target_mode = match @xml.get_attr(rn, "TargetMode") {
          Some(v) => v
          None => ""
        }
        if not(id.is_empty()) {
          rels.push({ id, target, type_, target_mode })
        }
        i = i + 1
      }
    }
  }
  rels
}

///|
/// Look up a relationship target by ID.
pub fn find_rel_target(rels : Array[Relationship], id : String) -> String {
  let n = rels.length()
  let mut i = 0
  while i < n {
    if rels[i].id == id {
      return rels[i].target
    }
    i = i + 1
  }
  ""
}

///|
/// Look up a full Relationship by ID.
pub fn find_rel(rels : Array[Relationship], id : String) -> Relationship? {
  let n = rels.length()
  let mut i = 0
  while i < n {
    if rels[i].id == id {
      return Some(rels[i])
    }
    i = i + 1
  }
  None
}

// ── Slide size ────────────────────────────────────────────────────────────────

///|
/// Extract slide size from ppt/presentation.xml.
pub fn parse_slide_size(prs_xml : String) -> SlideSize {
  let nodes = @xml.parse_xml(prs_xml)
  match @xml.find_child(nodes, "p:presentation") {
    None => SlideSize::default()
    Some(prs) => {
      let prs_ch = @xml.get_children(prs)
      match @xml.find_child(prs_ch, "p:sldSz") {
        None => SlideSize::default()
        Some(sz) => {
          let cx = attr_int(sz, "cx")
          let cy = attr_int(sz, "cy")
          if cx > 0 && cy > 0 {
            { cx, cy }
          } else {
            SlideSize::default()
          }
        }
      }
    }
  }
}

// ── Main slide parser ─────────────────────────────────────────────────────────

///|
/// Parse a slide XML string into a SlideData.
pub fn parse_slide(
  slide_xml : String,
  slide_size : SlideSize,
  theme : ThemeData,
) -> SlideData {
  let nodes = @xml.parse_xml(slide_xml)
  let shapes : Array[Shape] = []
  // Navigate p:sld → p:cSld
  let csld = @xml.find_path(nodes, ["p:sld", "p:cSld"])
  match csld {
    None =>
      {
        slide_size,
        background: Color::none(),
        bg_grad: GradientFill::none(),
        bg_blip_fill: BlipFill::none(),
        bg_patt_fill: PatternFill::none(),
        shapes,
      }
    Some(csld_node) => {
      let csld_ch = @xml.get_children(csld_node)
      // Parse background: p:bg → p:bgPr → solidFill / gradFill
      // Color::none() = no explicit bg (inherit from layout/master)
      let (bg, bg_grad, bg_bf, bg_pf) = match @xml.find_child(csld_ch, "p:bg") {
        None =>
          (Color::none(), GradientFill::none(), BlipFill::none(), PatternFill::none())
        Some(bg_node) => {
          let bg_ch = @xml.get_children(bg_node)
          match @xml.find_child(bg_ch, "p:bgPr") {
            None =>
              (
                Color::none(),
                GradientFill::none(),
                BlipFill::none(),
                PatternFill::none(),
              )
            Some(bgpr) => {
              let bgpr_ch = @xml.get_children(bgpr)
              let c = parse_solid_fill(bgpr_ch, theme)
              let gf = parse_gradient_fill(bgpr_ch, theme)
              let bf = parse_blip_fill(bgpr_ch)
              let pf = parse_pattern_fill(bgpr_ch, theme)
              (c, gf, bf, pf)
            }
          }
        }
      }
      // Navigate to p:spTree
      match @xml.find_child(csld_ch, "p:spTree") {
        None =>
          {
            slide_size,
            background: bg,
            bg_grad,
            bg_blip_fill: bg_bf,
            bg_patt_fill: bg_pf,
            shapes,
          }
        Some(tree) => {
          let parsed = parse_shapes_from_children(
            @xml.get_children(tree),
            theme,
          )
          {
            slide_size,
            background: bg,
            bg_grad,
            bg_blip_fill: bg_bf,
            bg_patt_fill: bg_pf,
            shapes: parsed,
          }
        }
      }
    }
  }
}

// ── Background parsing (shared) ──────────────────────────────────────────────

///|
/// Parse background color from p:cSld children. Returns Color::none() if no bg.
///|
/// Parse all background fills from p:cSld children.
pub fn parse_bg_fills_from_csld(
  csld_ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> (Color, GradientFill, BlipFill, PatternFill) {
  match @xml.find_child(csld_ch, "p:bg") {
    None =>
      (Color::none(), GradientFill::none(), BlipFill::none(), PatternFill::none())
    Some(bg_node) => {
      let bg_ch = @xml.get_children(bg_node)
      match @xml.find_child(bg_ch, "p:bgPr") {
        None =>
          (
            Color::none(),
            GradientFill::none(),
            BlipFill::none(),
            PatternFill::none(),
          )
        Some(bgpr) => {
          let bgpr_ch = @xml.get_children(bgpr)
          let c = parse_solid_fill(bgpr_ch, theme)
          let gf = parse_gradient_fill(bgpr_ch, theme)
          let bf = parse_blip_fill(bgpr_ch)
          let pf = parse_pattern_fill(bgpr_ch, theme)
          (c, gf, bf, pf)
        }
      }
    }
  }
}

///|
/// Parse shapes from a list of child nodes (shared by spTree and grpSp).
fn parse_shapes_from_children(
  children : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Array[Shape] {
  let shapes : Array[Shape] = []
  let n = children.length()
  let mut i = 0
  while i < n {
    match children[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "p:sp" {
          shapes.push(parse_sp(children[i], theme))
        } else if tag == "p:pic" {
          shapes.push(parse_pic(children[i], theme))
        } else if tag == "p:graphicFrame" {
          shapes.push(parse_graphic_frame(children[i], theme))
        } else if tag == "p:grpSp" {
          shapes.push(parse_grp_sp(children[i], theme))
        } else if tag == "p:cxnSp" {
          shapes.push(parse_cxn_sp(children[i], theme))
        }
      _ => ()
    }
    i = i + 1
  }
  shapes
}

///|
/// Parse shapes from p:spTree children.
fn parse_shapes_from_tree(
  tree : @xml.XmlNode,
  theme : ThemeData,
) -> Array[Shape] {
  parse_shapes_from_children(@xml.get_children(tree), theme)
}

// ── Table style parsing ─────────────────────────────────────────────────────

///|
fn parse_table_style_cell(node : @xml.XmlNode, theme : ThemeData) -> TableStyleCell {
  let ch = @xml.get_children(node)
  let mut fill = Color::none()
  let mut grad_fill = GradientFill::none()
  let mut bold = false
  let mut italic = false
  let mut font_color = Color::none()
  let mut bdr_l_w = 0
  let mut bdr_r_w = 0
  let mut bdr_t_w = 0
  let mut bdr_b_w = 0
  let mut bdr_l_color = Color::none()
  let mut bdr_r_color = Color::none()
  let mut bdr_t_color = Color::none()
  let mut bdr_b_color = Color::none()
  // Parse a:tcPr (cell properties within style)
  match @xml.find_child(ch, "a:tcPr") {
    None => ()
    Some(tcpr) => {
      let tcpr_ch = @xml.get_children(tcpr)
      let gf = parse_gradient_fill(tcpr_ch, theme)
      if not(gf.is_none()) {
        grad_fill = gf
      } else {
        fill = parse_solid_fill(tcpr_ch, theme)
      }
      let (lw, lc) = parse_tc_border(tcpr_ch, "a:lnL", theme)
      bdr_l_w = lw
      bdr_l_color = lc
      let (rw, rc) = parse_tc_border(tcpr_ch, "a:lnR", theme)
      bdr_r_w = rw
      bdr_r_color = rc
      let (tw, tc_) = parse_tc_border(tcpr_ch, "a:lnT", theme)
      bdr_t_w = tw
      bdr_t_color = tc_
      let (bw, bc) = parse_tc_border(tcpr_ch, "a:lnB", theme)
      bdr_b_w = bw
      bdr_b_color = bc
    }
  }
  // Parse a:tcTxBody → a:lstStyle / a:defRPr for font style
  match @xml.find_child(ch, "a:tcTxBody") {
    None => ()
    Some(txb) => {
      let txb_ch = @xml.get_children(txb)
      // Look for a:lstStyle → a:defPPr → a:defRPr or direct a:p → a:pPr → a:defRPr
      match @xml.find_child(txb_ch, "a:lstStyle") {
        None => ()
        Some(lst) => {
          let lst_ch = @xml.get_children(lst)
          // Check lvl1pPr
          match @xml.find_child(lst_ch, "a:lvl1pPr") {
            None => ()
            Some(ppr) => {
              match @xml.find_child(@xml.get_children(ppr), "a:defRPr") {
                None => ()
                Some(rpr) => {
                  bold = attr_int(rpr, "b") == 1
                  italic = attr_int(rpr, "i") == 1
                  font_color = parse_solid_fill(@xml.get_children(rpr), theme)
                }
              }
            }
          }
        }
      }
    }
  }
  {
    fill, grad_fill,
    bdr_l_w, bdr_r_w, bdr_t_w, bdr_b_w,
    bdr_l_color, bdr_r_color, bdr_t_color, bdr_b_color,
    bold, italic, font_color,
  }
}

///|
pub fn parse_table_styles(xml_str : String, theme : ThemeData) -> Array[TableStyleDef] {
  let styles : Array[TableStyleDef] = []
  let nodes = @xml.parse_xml(xml_str)
  // root = <a:tblStyleLst>
  let lst_node = match @xml.find_child(nodes, "a:tblStyleLst") {
    Some(n) => n
    None => return styles // no table styles
  }
  let style_nodes = @xml.find_children(@xml.get_children(lst_node), "a:tblStyle")
  let ns = style_nodes.length()
  let mut si = 0
  while si < ns {
    let sn = style_nodes[si]
    let id = match @xml.get_attr(sn, "styleId") {
      Some(v) => v
      None => ""
    }
    let sn_ch = @xml.get_children(sn)
    let parse = fn(tag : String) -> TableStyleCell {
      match @xml.find_child(sn_ch, tag) {
        Some(n) => parse_table_style_cell(n, theme)
        None => TableStyleCell::none()
      }
    }
    styles.push({
      id,
      whole_tbl: parse("a:wholeTbl"),
      band1_h: parse("a:band1H"),
      band2_h: parse("a:band2H"),
      band1_v: parse("a:band1V"),
      band2_v: parse("a:band2V"),
      first_row: parse("a:firstRow"),
      last_row: parse("a:lastRow"),
      first_col: parse("a:firstCol"),
      last_col: parse("a:lastCol"),
    })
    si = si + 1
  }
  styles
}
