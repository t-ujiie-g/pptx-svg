// OOXML / PPTX parser — Phase 2
//
// Extracts shape data from slide XML and .rels files.
// All positions/sizes are in EMUs (English Metric Units).
//   1 inch = 914400 EMU,  1 pt = 12700 EMU,  1 cm = 360000 EMU
// Standard slide: 9144000 x 6858000 EMU (10in x 7.5in)

// ── Color ─────────────────────────────────────────────────────────────────────

pub struct Color {
  r : Int
  g : Int
  b : Int
}

pub fn Color::white() -> Color {
  { r: 255, g: 255, b: 255 }
}

pub fn Color::black() -> Color {
  { r: 0, g: 0, b: 0 }
}

pub fn Color::none() -> Color {
  { r: -1, g: -1, b: -1 }
}

pub fn Color::is_none(self : Color) -> Bool {
  self.r < 0
}

// Parse a 6-char hex color string "RRGGBB" → Color
fn parse_hex_color(s : String) -> Color {
  if s.length() < 6 {
    return Color::black()
  }
  { r: parse_hex2(s, 0), g: parse_hex2(s, 2), b: parse_hex2(s, 4) }
}

fn hex_digit(c : Char) -> Int {
  let v = Char::to_int(c)
  if v >= 48 && v <= 57 {
    v - 48
  } else if v >= 65 && v <= 70 {
    v - 55
  } else if v >= 97 && v <= 102 {
    v - 87
  } else {
    0
  }
}

fn parse_hex2(s : String, offset : Int) -> Int {
  hex_digit(s.get_char(offset).unwrap()) * 16 +
  hex_digit(s.get_char(offset + 1).unwrap())
}

pub fn Color::to_css(self : Color) -> String {
  if self.is_none() {
    "none"
  } else {
    "rgb(" + int_to_str(self.r) + "," + int_to_str(self.g) + "," + int_to_str(self.b) + ")"
  }
}

// ── Text model ────────────────────────────────────────────────────────────────

pub struct TextRun {
  text      : String
  bold      : Bool
  italic    : Bool
  font_size : Int   // in half-points (sz attr); 0 = inherit
  color     : Color
}

pub struct TextParagraph {
  runs      : Array[TextRun]
  align     : String  // "l", "c", "r", "just" or ""
}

// ── Shape geometry ────────────────────────────────────────────────────────────

pub enum ShapeGeom {
  Rect
  Ellipse
  RoundRect
  Line
  Other(String)
}

fn parse_geom(prst : String) -> ShapeGeom {
  if prst == "rect" || prst == "snip1Rect" || prst == "snip2SameRect" {
    ShapeGeom::Rect
  } else if prst == "ellipse" || prst == "oval" {
    ShapeGeom::Ellipse
  } else if prst == "roundRect" || prst == "round1Rect" || prst == "round2SameRect" {
    ShapeGeom::RoundRect
  } else if prst == "line" || prst == "straightConnector1" {
    ShapeGeom::Line
  } else {
    ShapeGeom::Other(prst)
  }
}

// ── Shape ─────────────────────────────────────────────────────────────────────

pub struct ShapeTransform {
  x  : Int
  y  : Int
  cx : Int
  cy : Int
  rot : Int  // rotation in 1/60000 degree units
  flip_h : Bool
  flip_v : Bool
}

pub fn ShapeTransform::default() -> ShapeTransform {
  { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flip_h: false, flip_v: false }
}

pub enum ShapeKind {
  AutoShape(ShapeGeom)
  Picture(String)  // relationship ID (r:embed)
  Table
  GroupShape
  Other
}

pub struct Shape {
  kind        : ShapeKind
  transform   : ShapeTransform
  fill        : Color          // Color::none() = no fill
  stroke      : Color          // Color::none() = no stroke
  stroke_w    : Int            // stroke width in EMU (0 = default ~12700)
  paragraphs  : Array[TextParagraph]
}

// ── Slide data ────────────────────────────────────────────────────────────────

pub struct SlideSize {
  cx : Int
  cy : Int
}

pub fn SlideSize::default() -> SlideSize {
  // Standard widescreen 10in x 7.5in
  { cx: 9144000, cy: 6858000 }
}

pub struct SlideData {
  slide_size : SlideSize
  background : Color
  shapes     : Array[Shape]
}

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

// Extract a solid fill color from <a:solidFill> child of `nodes`.
fn parse_solid_fill(nodes : Array[@xml.XmlNode]) -> Color {
  match @xml.find_child(nodes, "a:solidFill") {
    None => Color::none()
    Some(fill_node) => {
      let fill_ch = @xml.get_children(fill_node)
      // Try <a:srgbClr val="RRGGBB">
      match @xml.find_child(fill_ch, "a:srgbClr") {
        Some(clr) =>
          match @xml.get_attr(clr, "val") {
            Some(v) => parse_hex_color(v)
            None => Color::black()
          }
        None =>
          // Try <a:schemeClr val="..."> — map common scheme colors
          match @xml.find_child(fill_ch, "a:schemeClr") {
            Some(clr) =>
              match @xml.get_attr(clr, "val") {
                Some(v) => scheme_color(v)
                None => Color::black()
              }
            None => Color::black()
          }
      }
    }
  }
}

fn scheme_color(name : String) -> Color {
  if name == "lt1" || name == "bg1" {
    Color::white()
  } else if name == "dk1" || name == "tx1" {
    Color::black()
  } else if name == "accent1" {
    { r: 68, g: 114, b: 196 }
  } else if name == "accent2" {
    { r: 237, g: 125, b: 49 }
  } else if name == "accent3" {
    { r: 165, g: 165, b: 165 }
  } else if name == "accent4" {
    { r: 255, g: 192, b: 0 }
  } else if name == "accent5" {
    { r: 91, g: 155, b: 213 }
  } else if name == "accent6" {
    { r: 112, g: 173, b: 71 }
  } else {
    Color::black()
  }
}

// ── Transform parsing ─────────────────────────────────────────────────────────

fn parse_xfrm(sppr_children : Array[@xml.XmlNode]) -> ShapeTransform {
  match @xml.find_child(sppr_children, "a:xfrm") {
    None => ShapeTransform::default()
    Some(xfrm) => {
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
  }
}

// ── Stroke parsing ────────────────────────────────────────────────────────────

fn parse_stroke(sppr_children : Array[@xml.XmlNode]) -> (Color, Int) {
  match @xml.find_child(sppr_children, "a:ln") {
    None => (Color::none(), 0)
    Some(ln) => {
      let w = attr_int(ln, "w")
      let ln_ch = @xml.get_children(ln)
      match @xml.find_child(ln_ch, "a:noFill") {
        Some(_) => (Color::none(), w)
        None => {
          let clr = parse_solid_fill(ln_ch)
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

// ── Text body parsing ─────────────────────────────────────────────────────────

fn parse_text_body(txbody : @xml.XmlNode) -> Array[TextParagraph] {
  let paras : Array[TextParagraph] = []
  let ch = @xml.get_children(txbody)
  let pnodes = @xml.find_children(ch, "a:p")
  let np = pnodes.length()
  let mut pi = 0
  while pi < np {
    let p = pnodes[pi]
    let p_ch = @xml.get_children(p)
    // Paragraph-level alignment from <a:pPr algn="...">
    let align = match @xml.find_child(p_ch, "a:pPr") {
      None => ""
      Some(ppr) =>
        match @xml.get_attr(ppr, "algn") {
          Some(v) => v
          None => ""
        }
    }
    // Collect runs from <a:r> elements
    let rnodes = @xml.find_children(p_ch, "a:r")
    let runs : Array[TextRun] = []
    let nr = rnodes.length()
    let mut ri = 0
    while ri < nr {
      let r = rnodes[ri]
      let r_ch = @xml.get_children(r)
      let text = match @xml.find_child(r_ch, "a:t") {
        None => ""
        Some(t_node) => @xml.get_text_content(t_node)
      }
      let (bold, italic, font_size, run_color) = match @xml.find_child(
        r_ch,
        "a:rPr",
      ) {
        None => (false, false, 0, Color::none())
        Some(rpr) => {
          let b = match @xml.get_attr(rpr, "b") {
            Some(v) => v == "1" || v == "true"
            None => false
          }
          let it = match @xml.get_attr(rpr, "i") {
            Some(v) => v == "1" || v == "true"
            None => false
          }
          let sz = attr_int(rpr, "sz") // half-points
          let clr = parse_solid_fill(@xml.get_children(rpr))
          (b, it, sz, clr)
        }
      }
      if not(text.is_empty()) {
        runs.push(
          { text, bold, italic, font_size, color: run_color },
        )
      }
      ri = ri + 1
    }
    paras.push({ runs, align })
    pi = pi + 1
  }
  paras
}

// ── Shape parsers ─────────────────────────────────────────────────────────────

fn parse_sp(sp : @xml.XmlNode) -> Shape {
  let ch = @xml.get_children(sp)
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
      let c = parse_solid_fill(sppr_ch)
      if c.is_none() { Color::none() } else { c }
    }
  }
  let (stroke, stroke_w) = parse_stroke(sppr_ch)
  // Text
  let paras = match @xml.find_child(ch, "p:txBody") {
    None => []
    Some(txb) => parse_text_body(txb)
  }
  {
    kind: ShapeKind::AutoShape(geom),
    transform,
    fill,
    stroke,
    stroke_w,
    paragraphs: paras,
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
    stroke: Color::none(),
    stroke_w: 0,
    paragraphs: [],
  }
}

fn empty_shape() -> Shape {
  {
    kind: ShapeKind::Other,
    transform: ShapeTransform::default(),
    fill: Color::none(),
    stroke: Color::none(),
    stroke_w: 0,
    paragraphs: [],
  }
}

// ── Relationship parsing ──────────────────────────────────────────────────────

pub struct Relationship {
  id     : String
  target : String
  type_  : String
}

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
  slide_size : SlideSize
) -> SlideData {
  let nodes = @xml.parse_xml(slide_xml)
  let shapes : Array[Shape] = []
  let bg = Color::white()
  // Navigate p:sld → p:cSld → p:spTree
  let sptree = @xml.find_path(nodes, ["p:sld", "p:cSld", "p:spTree"])
  match sptree {
    None => { slide_size, background: bg, shapes }
    Some(tree) => {
      let tree_ch = @xml.get_children(tree)
      let n = tree_ch.length()
      let mut i = 0
      while i < n {
        match tree_ch[i] {
          @xml.XmlNode::Element(tag, _, _) => {
            if tag == "p:sp" {
              shapes.push(parse_sp(tree_ch[i]))
            } else if tag == "p:pic" {
              shapes.push(parse_pic(tree_ch[i]))
            }
          }
          _ => ()
        }
        i = i + 1
      }
      { slide_size, background: bg, shapes }
    }
  }
}

// ── Integer→string (needed here too) ─────────────────────────────────────────

fn int_digit_str(d : Int) -> String {
  match d {
    0 => "0"
    1 => "1"
    2 => "2"
    3 => "3"
    4 => "4"
    5 => "5"
    6 => "6"
    7 => "7"
    8 => "8"
    _ => "9"
  }
}

fn int_to_str(n : Int) -> String {
  if n < 0 {
    "-" + int_to_str(-n)
  } else if n < 10 {
    int_digit_str(n)
  } else {
    int_to_str(n / 10) + int_digit_str(n % 10)
  }
}
