// SVG → SlideData parser for round-trip editing
//
// Reads data-ooxml-* attributes embedded by the renderer to reconstruct
// a SlideData from edited SVG. All values in OOXML native units (EMU, etc.)
//
// NOTE: No StringBuilder — all strings built via `+` (wasm:js-string "concat").

// ── Integer parsing ──────────────────────────────────────────────────────────

fn parse_int(s : String) -> Int {
  @xml.parse_int(s)
}

// ── Color parsing ────────────────────────────────────────────────────────────

fn parse_hex_color(s : String) -> @ooxml.Color {
  if s == "none" {
    @ooxml.Color::none()
  } else {
    @ooxml.parse_hex_color(s)
  }
}

// ── Attribute helpers ────────────────────────────────────────────────────────

fn get_data_attr(node : @xml.XmlNode, name : String) -> String {
  match @xml.get_attr(node, name) {
    Some(v) => v
    None => ""
  }
}

fn get_data_int(node : @xml.XmlNode, name : String) -> Int {
  parse_int(get_data_attr(node, name))
}

// ── Text parsing from SVG ────────────────────────────────────────────────────

fn parse_text_from_svg(text_node : @xml.XmlNode) -> Array[@ooxml.TextParagraph] {
  let paras : Array[@ooxml.TextParagraph] = []
  let children = @xml.get_children(text_node)
  let n = children.length()
  let mut i = 0
  while i < n {
    let child = children[i]
    match child {
      @xml.XmlNode::Element(tag, _, _) => {
        if tag == "tspan" {
          // This is a paragraph-level tspan
          let align = get_data_attr(child, "data-ooxml-para-align")
          let level = get_data_int(child, "data-ooxml-para-lvl")
          let runs : Array[@ooxml.TextRun] = []
          let tspan_children = @xml.get_children(child)
          let nc = tspan_children.length()
          let mut j = 0
          while j < nc {
            match tspan_children[j] {
              @xml.XmlNode::Element(rtag, _, _) => {
                if rtag == "tspan" {
                  // Run-level tspan
                  let run_node = tspan_children[j]
                  let text = @xml.get_text_content(run_node)
                  let bold_s = get_data_attr(run_node, "data-ooxml-bold")
                  let bold = bold_s == "1"
                  let font_size = get_data_int(run_node, "data-ooxml-font-size")
                  let color_hex = get_data_attr(run_node, "data-ooxml-color")
                  let color = if color_hex.is_empty() {
                    @ooxml.Color::none()
                  } else {
                    parse_hex_color(color_hex)
                  }
                  // Check italic from font-style attribute
                  let italic_s = get_data_attr(run_node, "font-style")
                  let italic = italic_s == "italic"
                  let font_face = get_data_attr(run_node, "data-ooxml-run-font")
                  let ea_font = get_data_attr(run_node, "data-ooxml-ea-font")
                  let char_spacing = get_data_int(run_node, "data-ooxml-char-spacing")
                  runs.push({ text, bold, italic, font_size, color, font_face, ea_font, underline: "", strike: "", baseline: 0, char_spacing })
                }
              }
              @xml.XmlNode::Text(text) => {
                // Direct text content (no run tspan wrapper)
                if not(text.is_empty()) {
                  runs.push(
                    {
                      text,
                      bold: false,
                      italic: false,
                      font_size: 0,
                      color: @ooxml.Color::none(),
                      font_face: "",
                      ea_font: "",
                      underline: "",
                      strike: "",
                      baseline: 0,
                      char_spacing: 0,
                    },
                  )
                }
              }
            }
            j = j + 1
          }
          paras.push({ runs, align, level, spc_before: 0, spc_after: 0, line_spacing: 0, mar_l: 0, indent: 0, bullet: "", bullet_auto: "", bullet_none: false })
        }
      }
      _ => ()
    }
    i = i + 1
  }
  paras
}

// ── Shape parsing from SVG <g> ───────────────────────────────────────────────

fn parse_shape_from_g(g_node : @xml.XmlNode) -> @ooxml.Shape {
  let shape_type = get_data_attr(g_node, "data-ooxml-shape-type")
  let geom_str = get_data_attr(g_node, "data-ooxml-geom")
  let x = get_data_int(g_node, "data-ooxml-x")
  let y = get_data_int(g_node, "data-ooxml-y")
  let cx = get_data_int(g_node, "data-ooxml-cx")
  let cy = get_data_int(g_node, "data-ooxml-cy")
  let rot = get_data_int(g_node, "data-ooxml-rot")
  let fill_hex = get_data_attr(g_node, "data-ooxml-fill")
  let stroke_hex = get_data_attr(g_node, "data-ooxml-stroke")
  let stroke_w = get_data_int(g_node, "data-ooxml-stroke-w")

  let fill = parse_hex_color(fill_hex)
  let stroke = parse_hex_color(stroke_hex)
  let transform : @ooxml.ShapeTransform = {
    x, y, cx, cy, rot, flip_h: false, flip_v: false,
  }

  // Determine shape kind
  let kind : @ooxml.ShapeKind = if shape_type == "picture" {
    // For pictures, we don't have the relationship ID in SVG
    // Use an empty string — the original rels will be preserved
    @ooxml.ShapeKind::Picture("")
  } else if shape_type == "table" {
    // Table shapes carry no inline data in SVG for now — use empty table
    let table_data : @ooxml.TableData = { col_widths: [], rows: [] }
    @ooxml.ShapeKind::TableShape(table_data)
  } else {
    let geom = parse_geom(geom_str)
    @ooxml.ShapeKind::AutoShape(geom)
  }

  // Parse text from <text> children
  let children = @xml.get_children(g_node)
  let mut paragraphs : Array[@ooxml.TextParagraph] = []
  let nc = children.length()
  let mut i = 0
  while i < nc {
    match children[i] {
      @xml.XmlNode::Element(tag, _, _) => {
        if tag == "text" {
          paragraphs = parse_text_from_svg(children[i])
        }
      }
      _ => ()
    }
    i = i + 1
  }

  // Parse body props from data attributes
  let bp_anchor = get_data_attr(g_node, "data-ooxml-anchor")
  let bp_l_ins = get_data_int(g_node, "data-ooxml-l-ins")
  let bp_t_ins = get_data_int(g_node, "data-ooxml-t-ins")
  let bp_r_ins = get_data_int(g_node, "data-ooxml-r-ins")
  let bp_b_ins = get_data_int(g_node, "data-ooxml-b-ins")
  let bp_auto_fit = get_data_attr(g_node, "data-ooxml-auto-fit") == "1"
  // normAutofit: get_data_int returns 0 for missing attrs; 0 maps to -1 (inactive).
  // This is safe because font_scale=0 (0%) never occurs in practice, and
  // ln_spc_reduction=0 is functionally identical to -1 (no reduction).
  let bp_font_scale = get_data_int(g_node, "data-ooxml-font-scale")
  let bp_ln_spc_reduction = get_data_int(g_node, "data-ooxml-ln-spc-reduction")
  let body_props : @ooxml.BodyProps = {
    anchor: bp_anchor,
    l_ins: if bp_l_ins == 0 { -1 } else { bp_l_ins },
    t_ins: if bp_t_ins == 0 { -1 } else { bp_t_ins },
    r_ins: if bp_r_ins == 0 { -1 } else { bp_r_ins },
    b_ins: if bp_b_ins == 0 { -1 } else { bp_b_ins },
    auto_fit: bp_auto_fit,
    font_scale: if bp_font_scale == 0 { -1 } else { bp_font_scale },
    ln_spc_reduction: if bp_ln_spc_reduction == 0 { -1 } else { bp_ln_spc_reduction },
  }
  let ph_type = get_data_attr(g_node, "data-ooxml-ph-type")
  let ph_idx = get_data_int(g_node, "data-ooxml-ph-idx")
  { kind, transform, fill, stroke, stroke_w, paragraphs, body_props, ph_type, ph_idx }
}

fn parse_geom(prst : String) -> @ooxml.ShapeGeom {
  if prst == "rect" {
    @ooxml.ShapeGeom::Rect
  } else if prst == "ellipse" {
    @ooxml.ShapeGeom::Ellipse
  } else if prst == "roundRect" {
    @ooxml.ShapeGeom::RoundRect
  } else if prst == "line" {
    @ooxml.ShapeGeom::Line
  } else {
    @ooxml.ShapeGeom::Other(prst)
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

///| Parse an SVG string (with data-ooxml-* attributes) back into SlideData.
pub fn parse_svg_to_slide(svg_string : String) -> @ooxml.SlideData {
  let nodes = @xml.parse_xml(svg_string)
  // Find the <svg> root
  let svg_node = match @xml.find_child(nodes, "svg") {
    Some(n) => n
    None => {
      return {
        slide_size: @ooxml.SlideSize::default(),
        background: @ooxml.Color::white(),
        shapes: [],
      }
    }
  }
  // Read slide metadata from SVG root
  let slide_cx = get_data_int(svg_node, "data-ooxml-slide-cx")
  let slide_cy = get_data_int(svg_node, "data-ooxml-slide-cy")
  let bg_hex = get_data_attr(svg_node, "data-ooxml-bg")

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
      @xml.XmlNode::Element(tag, _, _) => {
        if tag == "g" {
          let st = get_data_attr(svg_children[i], "data-ooxml-shape-type")
          if not(st.is_empty()) {
            shapes.push(parse_shape_from_g(svg_children[i]))
          }
        }
      }
      _ => ()
    }
    i = i + 1
  }

  { slide_size, background, shapes }
}
