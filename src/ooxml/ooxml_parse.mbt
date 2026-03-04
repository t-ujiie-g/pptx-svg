// OOXML Shape / Slide / Fill parsing
//
// Same package as ooxml.mbt — all types and helpers are shared.

// ── Attribute helpers ─────────────────────────────────────────────────────────

fn attr_int(node : @xml.XmlNode, name : String) -> Int {
  match @xml.get_attr(node, name) {
    Some(v) => @xml.parse_int(v)
    None => 0
  }
}

fn child_attr(
  nodes : Array[@xml.XmlNode],
  tag : String,
  attr : String
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

///| Parse a color from direct children (a:srgbClr or a:schemeClr), without a:solidFill wrapper.
///  Used for e.g. a:buClr which contains color child directly.
fn parse_color_from_children(children : Array[@xml.XmlNode], theme : ThemeData) -> Color {
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

fn parse_solid_fill(nodes : Array[@xml.XmlNode], theme : ThemeData) -> Color {
  match @xml.find_child(nodes, "a:solidFill") {
    None => Color::none()
    Some(fill_node) => parse_color_from_children(@xml.get_children(fill_node), theme)
  }
}

// ── Gradient fill parsing ────────────────────────────────────────────────────

///| Parse <a:gradFill> from child nodes. Returns GradientFill::none() if absent.
fn parse_gradient_fill(nodes : Array[@xml.XmlNode], theme : ThemeData) -> GradientFill {
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
            let color = parse_color_from_children(@xml.get_children(gs_nodes[gi]), theme)
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
        Some(lin) => {
          angle = attr_int(lin, "ang")
        }
        None => {
          // Parse path gradient: <a:path path="circle|rect|shape">
          match @xml.find_child(gf_ch, "a:path") {
            Some(path_node) => {
              angle = -1
              path_type = match @xml.get_attr(path_node, "path") {
                Some(v) => v
                None => "circle"
              }
              // Parse <a:fillToRect l="N" t="N" r="N" b="N"/>
              match @xml.find_child(@xml.get_children(path_node), "a:fillToRect") {
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
      }
      let tile_flip = match @xml.get_attr(gf_node, "tileFlip") {
        Some(v) => v
        None => ""
      }
      { stops, angle, path_type, rot_with_shape, fill_to_l, fill_to_t, fill_to_r, fill_to_b, tile_flip }
    }
  }
}

// ── Transform parsing ─────────────────────────────────────────────────────────

///| Parse transform from an xfrm element node (either a:xfrm or p:xfrm).
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

///| Parse transform from a:xfrm child within spPr children.
fn parse_xfrm(sppr_children : Array[@xml.XmlNode]) -> ShapeTransform {
  match @xml.find_child(sppr_children, "a:xfrm") {
    None => ShapeTransform::default()
    Some(xfrm) => parse_xfrm_node(xfrm)
  }
}

// ── Stroke parsing ────────────────────────────────────────────────────────────

fn parse_stroke(sppr_children : Array[@xml.XmlNode], theme : ThemeData) -> (Color, Int) {
  match @xml.find_child(sppr_children, "a:ln") {
    None => (Color::none(), 0)
    Some(ln) => {
      let w = attr_int(ln, "w")
      let ln_ch = @xml.get_children(ln)
      match @xml.find_child(ln_ch, "a:noFill") {
        Some(_) => (Color::none(), w)
        None => {
          let clr = parse_solid_fill(ln_ch, theme)
          if clr.is_none() {
            (Color::black(), w)
          } else {
            (clr, w)
          }
        }
      }
    }
  }
}

// ── Blip fill parser ──────────────────────────────────────────────────────────

///| Parse a:blipFill from spPr children → BlipFill.
fn parse_blip_fill(nodes : Array[@xml.XmlNode]) -> BlipFill {
  match @xml.find_child(nodes, "a:blipFill") {
    None => BlipFill::none()
    Some(bf_node) => {
      let bf_ch = @xml.get_children(bf_node)
      let rid = match @xml.find_child(bf_ch, "a:blip") {
        None => ""
        Some(blip) =>
          match @xml.get_attr(blip, "r:embed") {
            Some(v) => v
            None => ""
          }
      }
      if rid.is_empty() { return BlipFill::none() }
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
      { rid, stretch, src_l, src_t, src_r, src_b, tile_tx, tile_ty, tile_sx, tile_sy, tile_flip, tile_algn }
    }
  }
}

// ── Pattern fill parser ──────────────────────────────────────────────────────

///| Parse a:pattFill from spPr children → PatternFill.
fn parse_pattern_fill(nodes : Array[@xml.XmlNode], theme : ThemeData) -> PatternFill {
  match @xml.find_child(nodes, "a:pattFill") {
    None => PatternFill::none()
    Some(pf_node) => {
      let prst = match @xml.get_attr(pf_node, "prst") {
        Some(v) => v
        None => ""
      }
      if prst.is_empty() { return PatternFill::none() }
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

// ── Shape parsers ─────────────────────────────────────────────────────────────

fn parse_sp(sp : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(sp)
  // Extract placeholder info from p:nvSpPr/p:nvPr/p:ph
  let mut ph_type = ""
  let mut ph_idx = -1
  match @xml.find_child(ch, "p:nvSpPr") {
    None => ()
    Some(nvsp) => {
      let nvsp_ch = @xml.get_children(nvsp)
      match @xml.find_child(nvsp_ch, "p:nvPr") {
        None => ()
        Some(nvpr) => {
          let nvpr_ch = @xml.get_children(nvpr)
          match @xml.find_child(nvpr_ch, "p:ph") {
            None => ()
            Some(ph) => {
              ph_type = match @xml.get_attr(ph, "type") { Some(v) => v; None => "body" }
              ph_idx = match @xml.get_attr(ph, "idx") { Some(v) => @xml.parse_int(v); None => -1 }
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
    None => ShapeGeom::Rect
    Some(g) =>
      match @xml.get_attr(g, "prst") {
        Some(v) => parse_geom(v)
        None => ShapeGeom::Rect
      }
  }
  // Fill
  let fill = match @xml.find_child(sppr_ch, "a:noFill") {
    Some(_) => Color::none()
    None => {
      let c = parse_solid_fill(sppr_ch, theme)
      if c.is_none() { Color::none() } else { c }
    }
  }
  let grad_fill = parse_gradient_fill(sppr_ch, theme)
  let (stroke, stroke_w) = parse_stroke(sppr_ch, theme)
  // Blip fill on AutoShape (a:blipFill inside p:spPr)
  let blip_fill = parse_blip_fill(sppr_ch)
  // Pattern fill
  let patt_fill = parse_pattern_fill(sppr_ch, theme)
  // Text
  let (paras, body_props) = match @xml.find_child(ch, "p:txBody") {
    None => ([], BodyProps::default())
    Some(txb) => parse_text_body(txb, theme)
  }
  {
    kind: ShapeKind::AutoShape(geom),
    transform,
    fill,
    grad_fill,
    blip_fill,
    patt_fill,
    stroke,
    stroke_w,
    paragraphs: paras,
    body_props,
    ph_type,
    ph_idx,
  }
}

fn parse_pic(pic : @xml.XmlNode) -> Shape {
  let ch = @xml.get_children(pic)
  let sppr = match @xml.find_child(ch, "p:spPr") {
    None => return empty_shape()
    Some(n) => n
  }
  let transform = parse_xfrm(@xml.get_children(sppr))
  // Image relationship ID from <p:blipFill><a:blip r:embed="rIdN"/>
  let rid = match @xml.find_child(ch, "p:blipFill") {
    None => ""
    Some(bf) =>
      match @xml.find_child(@xml.get_children(bf), "a:blip") {
        None => ""
        Some(blip) =>
          match @xml.get_attr(blip, "r:embed") {
            Some(v) => v
            None => ""
          }
      }
  }
  {
    kind: ShapeKind::Picture(rid),
    transform,
    fill: Color::none(),
    grad_fill: GradientFill::none(),
    blip_fill: BlipFill::none(),
    patt_fill: PatternFill::none(),
    stroke: Color::none(),
    stroke_w: 0,
    paragraphs: [],
    body_props: BodyProps::default(),
    ph_type: "",
    ph_idx: -1,
  }
}

fn parse_table_cell(tc : @xml.XmlNode, theme : ThemeData) -> TableCell {
  let ch = @xml.get_children(tc)
  // Parse cell text from <a:txBody>
  let (paragraphs, _) = match @xml.find_child(ch, "a:txBody") {
    None => ([], BodyProps::default())
    Some(txb) => parse_text_body(txb, theme)
  }
  // Parse cell fill from <a:tcPr> → solidFill
  let fill = match @xml.find_child(ch, "a:tcPr") {
    None => Color::none()
    Some(tcpr) => {
      let c = parse_solid_fill(@xml.get_children(tcpr), theme)
      c
    }
  }
  { paragraphs, fill, grad_fill: GradientFill::none() }
}

///| Parse transform from either a:xfrm or p:xfrm (graphicFrame uses p:xfrm).
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

fn parse_graphic_frame(gf : @xml.XmlNode, theme : ThemeData) -> Shape {
  let ch = @xml.get_children(gf)
  // Transform from <p:xfrm> (direct child of graphicFrame, not inside spPr)
  let transform = parse_xfrm_any(ch)
  // Navigate: <a:graphic> → <a:graphicData> → <a:tbl>
  let tbl = match @xml.find_child(ch, "a:graphic") {
    None => None
    Some(graphic) =>
      match @xml.find_child(@xml.get_children(graphic), "a:graphicData") {
        None => None
        Some(gdata) => @xml.find_child(@xml.get_children(gdata), "a:tbl")
      }
  }
  match tbl {
    None => empty_shape()
    Some(tbl_node) => {
      let tbl_ch = @xml.get_children(tbl_node)
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
      let table_data : TableData = { col_widths, rows }
      {
        kind: ShapeKind::TableShape(table_data),
        transform,
        fill: Color::none(),
        grad_fill: GradientFill::none(),
        blip_fill: BlipFill::none(),
        patt_fill: PatternFill::none(),
        stroke: Color::none(),
        stroke_w: 0,
        paragraphs: [],
        body_props: BodyProps::default(),
        ph_type: "",
        ph_idx: -1,
      }
    }
  }
}

fn empty_shape() -> Shape {
  {
    kind: ShapeKind::Other,
    transform: ShapeTransform::default(),
    fill: Color::none(),
    grad_fill: GradientFill::none(),
    blip_fill: BlipFill::none(),
    patt_fill: PatternFill::none(),
    stroke: Color::none(),
    stroke_w: 0,
    paragraphs: [],
    body_props: BodyProps::default(),
    ph_type: "",
    ph_idx: -1,
  }
}

// ── Relationship parsing ──────────────────────────────────────────────────────

///| Parse a .rels XML file into a map of relationship ID → target path.
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
        if not(id.is_empty()) {
          rels.push({ id, target, type_ })
        }
        i = i + 1
      }
    }
  }
  rels
}

///| Look up a relationship target by ID.
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

// ── Slide size ────────────────────────────────────────────────────────────────

///| Extract slide size from ppt/presentation.xml.
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
          if cx > 0 && cy > 0 { { cx, cy } } else { SlideSize::default() }
        }
      }
    }
  }
}

// ── Main slide parser ─────────────────────────────────────────────────────────

///| Parse a slide XML string into a SlideData.
pub fn parse_slide(
  slide_xml : String,
  slide_size : SlideSize,
  theme : ThemeData
) -> SlideData {
  let nodes = @xml.parse_xml(slide_xml)
  let shapes : Array[Shape] = []
  // Navigate p:sld → p:cSld
  let csld = @xml.find_path(nodes, ["p:sld", "p:cSld"])
  match csld {
    None => { slide_size, background: Color::none(), bg_grad: GradientFill::none(), shapes }
    Some(csld_node) => {
      let csld_ch = @xml.get_children(csld_node)
      // Parse background: p:bg → p:bgPr → solidFill / gradFill
      // Color::none() = no explicit bg (inherit from layout/master)
      let (bg, bg_grad) = match @xml.find_child(csld_ch, "p:bg") {
        None => (Color::none(), GradientFill::none())
        Some(bg_node) => {
          let bg_ch = @xml.get_children(bg_node)
          match @xml.find_child(bg_ch, "p:bgPr") {
            None => (Color::none(), GradientFill::none())
            Some(bgpr) => {
              let bgpr_ch = @xml.get_children(bgpr)
              let c = parse_solid_fill(bgpr_ch, theme)
              let gf = parse_gradient_fill(bgpr_ch, theme)
              (c, gf)
            }
          }
        }
      }
      // Navigate to p:spTree
      match @xml.find_child(csld_ch, "p:spTree") {
        None => { slide_size, background: bg, bg_grad, shapes }
        Some(tree) => {
          let tree_ch = @xml.get_children(tree)
          let n = tree_ch.length()
          let mut i = 0
          while i < n {
            match tree_ch[i] {
              @xml.XmlNode::Element(tag, _, _) => {
                if tag == "p:sp" {
                  shapes.push(parse_sp(tree_ch[i], theme))
                } else if tag == "p:pic" {
                  shapes.push(parse_pic(tree_ch[i]))
                } else if tag == "p:graphicFrame" {
                  shapes.push(parse_graphic_frame(tree_ch[i], theme))
                }
              }
              _ => ()
            }
            i = i + 1
          }
          { slide_size, background: bg, bg_grad, shapes }
        }
      }
    }
  }
}

// ── Background parsing (shared) ──────────────────────────────────────────────

///| Parse background color from p:cSld children. Returns Color::none() if no bg.
fn parse_bg_from_csld(csld_ch : Array[@xml.XmlNode], theme : ThemeData) -> Color {
  match @xml.find_child(csld_ch, "p:bg") {
    None => Color::none()
    Some(bg_node) => {
      let bg_ch = @xml.get_children(bg_node)
      match @xml.find_child(bg_ch, "p:bgPr") {
        None => Color::none()
        Some(bgpr) => parse_solid_fill(@xml.get_children(bgpr), theme)
      }
    }
  }
}

///| Parse shapes from p:spTree children.
fn parse_shapes_from_tree(tree : @xml.XmlNode, theme : ThemeData) -> Array[Shape] {
  let shapes : Array[Shape] = []
  let tree_ch = @xml.get_children(tree)
  let n = tree_ch.length()
  let mut i = 0
  while i < n {
    match tree_ch[i] {
      @xml.XmlNode::Element(tag, _, _) => {
        if tag == "p:sp" {
          shapes.push(parse_sp(tree_ch[i], theme))
        } else if tag == "p:pic" {
          shapes.push(parse_pic(tree_ch[i]))
        } else if tag == "p:graphicFrame" {
          shapes.push(parse_graphic_frame(tree_ch[i], theme))
        }
      }
      _ => ()
    }
    i = i + 1
  }
  shapes
}

// ── Integer→string ───────────────────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }
