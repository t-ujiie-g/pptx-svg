// OOXML / PPTX parser — Phase 2
//
// Extracts shape data from slide XML and .rels files.
// All positions/sizes are in EMUs (English Metric Units).
//   1 inch = 914400 EMU,  1 pt = 12700 EMU,  1 cm = 360000 EMU
// Standard slide: 9144000 x 6858000 EMU (10in x 7.5in)

// ── Shared constants ─────────────────────────────────────────────────────────

///| Default stroke width: 1 pt in EMU.
pub let default_stroke_w : Int = 12700

// ── Color ─────────────────────────────────────────────────────────────────────

pub(all) struct Color {
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

///| Parse a 6-char hex color string "RRGGBB" → Color.
pub fn parse_hex_color(s : String) -> Color {
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

///| Convert Color to 6-char hex string "RRGGBB" for data attributes.
pub fn Color::to_hex(self : Color) -> String {
  if self.is_none() {
    "none"
  } else {
    hex_byte(self.r) + hex_byte(self.g) + hex_byte(self.b)
  }
}

fn hex_byte(v : Int) -> String {
  hex_nibble(v / 16) + hex_nibble(v % 16)
}

fn hex_nibble(n : Int) -> String {
  match n {
    0 => "0"; 1 => "1"; 2 => "2"; 3 => "3"; 4 => "4"
    5 => "5"; 6 => "6"; 7 => "7"; 8 => "8"; 9 => "9"
    10 => "A"; 11 => "B"; 12 => "C"; 13 => "D"; 14 => "E"
    _ => "F"
  }
}

// ── Theme data ───────────────────────────────────────────────────────────────

pub(all) struct ThemeData {
  dk1 : Color
  lt1 : Color
  dk2 : Color
  lt2 : Color
  accent1 : Color
  accent2 : Color
  accent3 : Color
  accent4 : Color
  accent5 : Color
  accent6 : Color
  hlink : Color
  fol_hlink : Color
  major_font : String
  minor_font : String
}

pub fn ThemeData::default() -> ThemeData {
  {
    dk1: Color::black(),
    lt1: Color::white(),
    dk2: { r: 68, g: 84, b: 106 },
    lt2: { r: 231, g: 230, b: 230 },
    accent1: { r: 68, g: 114, b: 196 },
    accent2: { r: 237, g: 125, b: 49 },
    accent3: { r: 165, g: 165, b: 165 },
    accent4: { r: 255, g: 192, b: 0 },
    accent5: { r: 91, g: 155, b: 213 },
    accent6: { r: 112, g: 173, b: 71 },
    hlink: { r: 5, g: 99, b: 193 },
    fol_hlink: { r: 149, g: 79, b: 163 },
    major_font: "Calibri Light",
    minor_font: "Calibri",
  }
}

///| Resolve a scheme color name to a Color from the theme.
pub fn resolve_scheme_color(name : String, theme : ThemeData) -> Color {
  if name == "dk1" || name == "tx1" {
    theme.dk1
  } else if name == "lt1" || name == "bg1" {
    theme.lt1
  } else if name == "dk2" || name == "tx2" {
    theme.dk2
  } else if name == "lt2" || name == "bg2" {
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

// ── RGB ↔ HSL conversion ────────────────────────────────────────────────────

fn min3(a : Int, b : Int, c : Int) -> Int {
  let m = if a < b { a } else { b }
  if m < c { m } else { c }
}

fn max3(a : Int, b : Int, c : Int) -> Int {
  let m = if a > b { a } else { b }
  if m > c { m } else { c }
}

///| Convert RGB (0-255) to HSL (h: 0-3600, s: 0-1000, l: 0-1000).
///  Uses integer arithmetic scaled by 1000 to avoid floating point.
fn rgb_to_hsl(r : Int, g : Int, b : Int) -> (Int, Int, Int) {
  let max_c = max3(r, g, b)
  let min_c = min3(r, g, b)
  let l = (max_c + min_c) * 1000 / 510 // 0-1000
  if max_c == min_c {
    return (0, 0, l)
  }
  let d = max_c - min_c
  let s = if l > 500 {
    d * 1000 / (510 - max_c - min_c)
  } else {
    d * 1000 / (max_c + min_c)
  }
  let h = if r == max_c {
    (g - b) * 600 / d
  } else if g == max_c {
    1200 + (b - r) * 600 / d
  } else {
    2400 + (r - g) * 600 / d
  }
  let h2 = if h < 0 { h + 3600 } else { h }
  (h2, s, l)
}

fn hue_to_rgb(p : Int, q : Int, t_in : Int) -> Int {
  let t = if t_in < 0 {
    t_in + 3600
  } else if t_in > 3600 {
    t_in - 3600
  } else {
    t_in
  }
  if t < 600 {
    (p + (q - p) * t / 600) * 255 / 1000
  } else if t < 1800 {
    q * 255 / 1000
  } else if t < 2400 {
    (p + (q - p) * (2400 - t) / 600) * 255 / 1000
  } else {
    p * 255 / 1000
  }
}

///| Convert HSL (h: 0-3600, s: 0-1000, l: 0-1000) to RGB Color.
fn hsl_to_rgb(h : Int, s : Int, l : Int) -> Color {
  if s == 0 {
    let v = l * 255 / 1000
    return { r: v, g: v, b: v }
  }
  let q = if l < 500 {
    l * (1000 + s) / 1000
  } else {
    l + s - l * s / 1000
  }
  let p = 2 * l - q
  let r = hue_to_rgb(p, q, h + 1200)
  let g = hue_to_rgb(p, q, h)
  let b = hue_to_rgb(p, q, h - 1200)
  // Clamp 0-255
  let clamp = fn(v : Int) -> Int {
    if v < 0 { 0 } else if v > 255 { 255 } else { v }
  }
  { r: clamp(r), g: clamp(g), b: clamp(b) }
}

// ── Color modifiers ─────────────────────────────────────────────────────────

///| Apply OOXML color modifier children (lumMod, lumOff, shade, tint, etc.)
///  to a base color. modifier_nodes are children of <a:schemeClr> or <a:srgbClr>.
pub fn apply_color_modifiers(
  color : Color,
  modifier_nodes : Array[@xml.XmlNode]
) -> Color {
  let n = modifier_nodes.length()
  if n == 0 {
    return color
  }
  let mut r = color.r
  let mut g = color.g
  let mut b = color.b
  let mut lum_mod = -1
  let mut lum_off = 0
  let mut i = 0
  while i < n {
    match modifier_nodes[i] {
      @xml.XmlNode::Element(tag, _, _) => {
        let val = attr_int(modifier_nodes[i], "val")
        if tag == "a:shade" {
          // shade: multiply each RGB component by val/100000
          r = r * val / 100000
          g = g * val / 100000
          b = b * val / 100000
        } else if tag == "a:tint" {
          // tint: move each component toward 255
          r = r + (255 - r) * val / 100000
          g = g + (255 - g) * val / 100000
          b = b + (255 - b) * val / 100000
        } else if tag == "a:lumMod" {
          lum_mod = val
        } else if tag == "a:lumOff" {
          lum_off = val
        } else if tag == "a:satMod" {
          // Saturation modification via HSL
          let (h, s, l) = rgb_to_hsl(r, g, b)
          let new_s = s * val / 100000
          let ns = if new_s > 1000 { 1000 } else { new_s }
          let c2 = hsl_to_rgb(h, ns, l)
          r = c2.r
          g = c2.g
          b = c2.b
        } else if tag == "a:alpha" {
          // alpha does not affect RGB, skip
          ()
        }
      }
      _ => ()
    }
    i = i + 1
  }
  // Apply lumMod/lumOff together (they work as a pair)
  if lum_mod >= 0 {
    let (h, s, l) = rgb_to_hsl(r, g, b)
    let mut new_l = l * lum_mod / 100000 + lum_off * 1000 / 100000
    if new_l < 0 { new_l = 0 }
    if new_l > 1000 { new_l = 1000 }
    let c2 = hsl_to_rgb(h, s, new_l)
    r = c2.r
    g = c2.g
    b = c2.b
  }
  // Clamp
  if r < 0 { r = 0 } else if r > 255 { r = 255 }
  if g < 0 { g = 0 } else if g > 255 { g = 255 }
  if b < 0 { b = 0 } else if b > 255 { b = 255 }
  { r, g, b }
}

// ── Theme parser ────────────────────────────────────────────────────────────

///| Parse the first color child of a clrScheme element (sysClr or srgbClr).
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

///| Parse theme1.xml into ThemeData.
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
  let mut dk1 = def.dk1
  let mut lt1 = def.lt1
  let mut dk2 = def.dk2
  let mut lt2 = def.lt2
  let mut accent1 = def.accent1
  let mut accent2 = def.accent2
  let mut accent3 = def.accent3
  let mut accent4 = def.accent4
  let mut accent5 = def.accent5
  let mut accent6 = def.accent6
  let mut hlink = def.hlink
  let mut fol_hlink = def.fol_hlink
  match @xml.find_child(te_ch, "a:clrScheme") {
    None => ()
    Some(cs) => {
      let cs_ch = @xml.get_children(cs)
      match parse_theme_color_node(cs_ch, "a:dk1") {
        Some(c) => dk1 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:lt1") {
        Some(c) => lt1 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:dk2") {
        Some(c) => dk2 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:lt2") {
        Some(c) => lt2 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent1") {
        Some(c) => accent1 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent2") {
        Some(c) => accent2 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent3") {
        Some(c) => accent3 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent4") {
        Some(c) => accent4 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent5") {
        Some(c) => accent5 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:accent6") {
        Some(c) => accent6 = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:hlink") {
        Some(c) => hlink = c
        None => ()
      }
      match parse_theme_color_node(cs_ch, "a:folHlink") {
        Some(c) => fol_hlink = c
        None => ()
      }
    }
  }
  // ── Font scheme ──
  let mut major_font = def.major_font
  let mut minor_font = def.minor_font
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
        }
      }
    }
  }
  {
    dk1, lt1, dk2, lt2,
    accent1, accent2, accent3, accent4, accent5, accent6,
    hlink, fol_hlink,
    major_font, minor_font,
  }
}

// ── Master / Layout types ────────────────────────────────────────────────────

pub(all) struct LevelTextDefaults {
  font_size   : Int      // hundredths-of-pt (sz); 0 = unset
  bold        : Bool
  italic      : Bool
  color       : Color    // Color::none() = unset
  font_face   : String   // "" = inherit
  align       : String   // "" = inherit
  mar_l       : Int      // EMU; -1 = unset
  indent      : Int      // EMU; 0 = unset
  bullet      : String   // "" = none
  bullet_none : Bool
}

pub fn LevelTextDefaults::empty() -> LevelTextDefaults {
  { font_size: 0, bold: false, italic: false, color: Color::none(), font_face: "", align: "", mar_l: -1, indent: 0, bullet: "", bullet_none: false }
}

pub(all) struct TextStyleGroup {
  levels : Array[LevelTextDefaults]  // 9 elements (lvl1pPr–lvl9pPr)
}

pub fn TextStyleGroup::empty() -> TextStyleGroup {
  let levels : Array[LevelTextDefaults] = []
  let mut i = 0
  while i < 9 {
    levels.push(LevelTextDefaults::empty())
    i = i + 1
  }
  { levels, }
}

pub(all) struct MasterTextStyles {
  title_style : TextStyleGroup
  body_style  : TextStyleGroup
  other_style : TextStyleGroup
}

pub fn MasterTextStyles::empty() -> MasterTextStyles {
  { title_style: TextStyleGroup::empty(), body_style: TextStyleGroup::empty(), other_style: TextStyleGroup::empty() }
}

pub(all) struct SlideMasterData {
  background  : Color
  shapes      : Array[Shape]
  text_styles : MasterTextStyles
}

pub(all) struct SlideLayoutData {
  background  : Color   // Color::none() = inherit from master
  shapes      : Array[Shape]
}

// ── Text model ────────────────────────────────────────────────────────────────

pub(all) struct TextRun {
  text      : String
  bold      : Bool
  italic    : Bool
  font_size : Int      // in hundredths-of-a-point (sz attr); 0 = inherit
  color     : Color
  font_face : String   // explicit typeface from <a:latin>; "" = use theme font
  underline : String   // "sng", "dbl", "heavy", etc.; "" = none
  strike    : String   // "sngStrike", "dblStrike"; "" = none
  baseline  : Int      // superscript (+30000 = 30%) / subscript (-25000); 0 = normal
}

pub(all) struct TextParagraph {
  runs       : Array[TextRun]
  align      : String  // "l", "c", "r", "just" or ""
  level      : Int     // indent level (0 = top-level, from a:pPr lvl="N")
  spc_before : Int     // spacing before in EMU (from a:spcBef); 0 = default
  spc_after  : Int     // spacing after in EMU (from a:spcAft); 0 = default
  mar_l      : Int     // left margin in EMU (from marL attr in a:pPr); 0 = default
  indent     : Int     // first-line indent in EMU (from indent attr in a:pPr); 0 = default
  bullet     : String  // bullet char (from a:buChar); "" = none/inherit
  bullet_auto : String // auto-number type (from a:buAutoNum); "" = none
  bullet_none : Bool   // explicitly no bullet (a:buNone)
}

// ── Text body properties ────────────────────────────────────────────────────

pub(all) struct BodyProps {
  anchor  : String  // vertical alignment: "t", "ctr", "b", "" = default (top)
  l_ins   : Int     // left inset in EMU (-1 = default 91440 ≈ 0.1in)
  t_ins   : Int     // top inset in EMU (-1 = default 45720 ≈ 0.05in)
  r_ins   : Int     // right inset in EMU (-1 = default 91440)
  b_ins   : Int     // bottom inset in EMU (-1 = default 45720)
  auto_fit : Bool   // <a:spAutoFit/> — shrink shape to fit text
}

pub fn BodyProps::default() -> BodyProps {
  { anchor: "", l_ins: -1, t_ins: -1, r_ins: -1, b_ins: -1, auto_fit: false }
}

// ── Table model ──────────────────────────────────────────────────────────────

pub(all) struct TableCell {
  paragraphs : Array[TextParagraph]
  fill : Color
}

pub(all) struct TableRow {
  height : Int  // EMU
  cells : Array[TableCell]
}

pub(all) struct TableData {
  col_widths : Array[Int]  // EMU per column
  rows : Array[TableRow]
}

// ── Shape geometry ────────────────────────────────────────────────────────────

pub(all) enum ShapeGeom {
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

///| Convert ShapeGeom back to OOXML preset geometry name.
pub fn ShapeGeom::to_prst(self : ShapeGeom) -> String {
  match self {
    ShapeGeom::Rect => "rect"
    ShapeGeom::Ellipse => "ellipse"
    ShapeGeom::RoundRect => "roundRect"
    ShapeGeom::Line => "line"
    ShapeGeom::Other(s) => s
  }
}

// ── Shape ─────────────────────────────────────────────────────────────────────

pub(all) struct ShapeTransform {
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

pub(all) enum ShapeKind {
  AutoShape(ShapeGeom)
  Picture(String)  // relationship ID (r:embed)
  TableShape(TableData)
  GroupShape
  Other
}

pub(all) struct Shape {
  kind        : ShapeKind
  transform   : ShapeTransform
  fill        : Color          // Color::none() = no fill
  stroke      : Color          // Color::none() = no stroke
  stroke_w    : Int            // stroke width in EMU (0 = default ~12700)
  paragraphs  : Array[TextParagraph]
  body_props  : BodyProps      // text body properties (anchor, insets)
  ph_type     : String         // "title","body","ctrTitle","subTitle","dt","ftr","sldNum","" (non-PH)
  ph_idx      : Int            // -1 = none
}

// ── Slide data ────────────────────────────────────────────────────────────────

pub(all) struct SlideSize {
  cx : Int
  cy : Int
}

pub fn SlideSize::default() -> SlideSize {
  // Standard widescreen 10in x 7.5in
  { cx: 9144000, cy: 6858000 }
}

pub(all) struct SlideData {
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
fn parse_solid_fill(nodes : Array[@xml.XmlNode], theme : ThemeData) -> Color {
  match @xml.find_child(nodes, "a:solidFill") {
    None => Color::none()
    Some(fill_node) => {
      let fill_ch = @xml.get_children(fill_node)
      // Try <a:srgbClr val="RRGGBB">
      match @xml.find_child(fill_ch, "a:srgbClr") {
        Some(clr) =>
          match @xml.get_attr(clr, "val") {
            Some(v) => {
              let base = parse_hex_color(v)
              apply_color_modifiers(base, @xml.get_children(clr))
            }
            None => Color::black()
          }
        None =>
          // Try <a:schemeClr val="..."> — resolve from theme + apply modifiers
          match @xml.find_child(fill_ch, "a:schemeClr") {
            Some(clr) =>
              match @xml.get_attr(clr, "val") {
                Some(v) => {
                  let base = resolve_scheme_color(v, theme)
                  apply_color_modifiers(base, @xml.get_children(clr))
                }
                None => Color::black()
              }
            None => Color::black()
          }
      }
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

// ── Text body parsing ─────────────────────────────────────────────────────────

///| Parse spacing value from <a:spcBef>/<a:spcAft>/<a:lnSpc> children.
///  Returns EMU. Supports <a:spcPts val="N"> (N in hundredths of a point).
fn parse_spacing(parent : @xml.XmlNode) -> Int {
  let ch = @xml.get_children(parent)
  match @xml.find_child(ch, "a:spcPts") {
    Some(pts) => {
      // val is in hundredths of a point → convert to EMU (1 pt = 12700 EMU)
      let v = attr_int(pts, "val")
      v * 127 // hundredths-of-pt * 12700/100 = * 127
    }
    None =>
      // a:spcPct (percentage of line height) — store raw for now as negative sentinel
      match @xml.find_child(ch, "a:spcPct") {
        Some(pct) => {
          let v = attr_int(pct, "val") // e.g. 100000 = 100%
          -(v) // negative = percentage marker
        }
        None => 0
      }
  }
}

///| Parse <a:bodyPr> element into BodyProps.
fn parse_body_props(ch : Array[@xml.XmlNode]) -> BodyProps {
  match @xml.find_child(ch, "a:bodyPr") {
    None => BodyProps::default()
    Some(bpr) => {
      let anchor = match @xml.get_attr(bpr, "anchor") { Some(v) => v; None => "" }
      let l_ins = match @xml.get_attr(bpr, "lIns") { Some(v) => @xml.parse_int(v); None => -1 }
      let t_ins = match @xml.get_attr(bpr, "tIns") { Some(v) => @xml.parse_int(v); None => -1 }
      let r_ins = match @xml.get_attr(bpr, "rIns") { Some(v) => @xml.parse_int(v); None => -1 }
      let b_ins = match @xml.get_attr(bpr, "bIns") { Some(v) => @xml.parse_int(v); None => -1 }
      let bpr_ch = @xml.get_children(bpr)
      let auto_fit = match @xml.find_child(bpr_ch, "a:spAutoFit") { Some(_) => true; None => false }
      { anchor, l_ins, t_ins, r_ins, b_ins, auto_fit }
    }
  }
}

///| Parse paragraph-level properties from <a:pPr>.
fn parse_para_props(
  ppr : @xml.XmlNode
) -> (String, Int, Int, Int, Int, Int, String, String, Bool) {
  let a = match @xml.get_attr(ppr, "algn") { Some(v) => v; None => "" }
  let l = attr_int(ppr, "lvl")
  let ml = attr_int(ppr, "marL")
  let ind = attr_int(ppr, "indent")
  let ppr_ch = @xml.get_children(ppr)
  let sb = match @xml.find_child(ppr_ch, "a:spcBef") { None => 0; Some(n) => parse_spacing(n) }
  let sa = match @xml.find_child(ppr_ch, "a:spcAft") { None => 0; Some(n) => parse_spacing(n) }
  let bu_char = match @xml.find_child(ppr_ch, "a:buChar") {
    None => ""
    Some(bc) => match @xml.get_attr(bc, "char") { Some(v) => v; None => "" }
  }
  let bu_auto = match @xml.find_child(ppr_ch, "a:buAutoNum") {
    None => ""
    Some(ba) => match @xml.get_attr(ba, "type") { Some(v) => v; None => "" }
  }
  let bu_none = match @xml.find_child(ppr_ch, "a:buNone") { Some(_) => true; None => false }
  (a, l, sb, sa, ml, ind, bu_char, bu_auto, bu_none)
}

///| Parse <a:defRPr> defaults from a <a:pPr> node.
fn parse_def_rpr(
  ppr : @xml.XmlNode,
  theme : ThemeData
) -> (Bool, Bool, Int, Color, String) {
  let ppr_ch = @xml.get_children(ppr)
  match @xml.find_child(ppr_ch, "a:defRPr") {
    None => (false, false, 0, Color::none(), "")
    Some(drpr) => {
      let db = match @xml.get_attr(drpr, "b") { Some(v) => v == "1" || v == "true"; None => false }
      let di = match @xml.get_attr(drpr, "i") { Some(v) => v == "1" || v == "true"; None => false }
      let dsz = attr_int(drpr, "sz")
      let drpr_ch = @xml.get_children(drpr)
      let dclr = parse_solid_fill(drpr_ch, theme)
      let dff = child_attr(drpr_ch, "a:latin", "typeface")
      (db, di, dsz, dclr, dff)
    }
  }
}

///| Parse run properties from <a:rPr> with defRPr fallback.
fn parse_run_props(
  rpr : @xml.XmlNode,
  def_bold : Bool, def_italic : Bool, def_font_size : Int,
  def_color : Color, def_font_face : String,
  theme : ThemeData
) -> (Bool, Bool, Int, Color, String, String, String, Int) {
  let b = match @xml.get_attr(rpr, "b") { Some(v) => v == "1" || v == "true"; None => def_bold }
  let it = match @xml.get_attr(rpr, "i") { Some(v) => v == "1" || v == "true"; None => def_italic }
  let sz_raw = attr_int(rpr, "sz")
  let sz = if sz_raw > 0 { sz_raw } else { def_font_size }
  let u = match @xml.get_attr(rpr, "u") { Some(v) => v; None => "" }
  let st = match @xml.get_attr(rpr, "strike") { Some(v) => v; None => "" }
  let bl = attr_int(rpr, "baseline")
  let rpr_ch = @xml.get_children(rpr)
  let clr_raw = parse_solid_fill(rpr_ch, theme)
  let clr = if clr_raw.is_none() { def_color } else { clr_raw }
  let ff_raw = child_attr(rpr_ch, "a:latin", "typeface")
  let ff = if ff_raw.is_empty() { def_font_face } else { ff_raw }
  (b, it, sz, clr, ff, u, st, bl)
}

fn parse_text_body(txbody : @xml.XmlNode, theme : ThemeData) -> (Array[TextParagraph], BodyProps) {
  let paras : Array[TextParagraph] = []
  let ch = @xml.get_children(txbody)
  let body_props = parse_body_props(ch)
  let pnodes = @xml.find_children(ch, "a:p")
  let np = pnodes.length()
  let mut pi = 0
  while pi < np {
    let p = pnodes[pi]
    let p_ch = @xml.get_children(p)
    let (align, level, spc_before, spc_after, mar_l, indent, bullet, bullet_auto, bullet_none) = match @xml.find_child(p_ch, "a:pPr") {
      None => ("", 0, 0, 0, 0, 0, "", "", false)
      Some(ppr) => parse_para_props(ppr)
    }
    let (def_bold, def_italic, def_font_size, def_color, def_font_face) = match @xml.find_child(p_ch, "a:pPr") {
      None => (false, false, 0, Color::none(), "")
      Some(ppr) => parse_def_rpr(ppr, theme)
    }
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
      let (bold, italic, font_size, run_color, font_face, underline, strike, baseline) = match @xml.find_child(r_ch, "a:rPr") {
        None => (def_bold, def_italic, def_font_size, def_color, def_font_face, "", "", 0)
        Some(rpr) => parse_run_props(rpr, def_bold, def_italic, def_font_size, def_color, def_font_face, theme)
      }
      if not(text.is_empty()) {
        runs.push({ text, bold, italic, font_size, color: run_color, font_face, underline, strike, baseline })
      }
      ri = ri + 1
    }
    paras.push({ runs, align, level, spc_before, spc_after, mar_l, indent, bullet, bullet_auto, bullet_none })
    pi = pi + 1
  }
  (paras, body_props)
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
  let (stroke, stroke_w) = parse_stroke(sppr_ch, theme)
  // Text
  let (paras, body_props) = match @xml.find_child(ch, "p:txBody") {
    None => ([], BodyProps::default())
    Some(txb) => parse_text_body(txb, theme)
  }
  {
    kind: ShapeKind::AutoShape(geom),
    transform,
    fill,
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
  { paragraphs, fill }
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
    stroke: Color::none(),
    stroke_w: 0,
    paragraphs: [],
    body_props: BodyProps::default(),
    ph_type: "",
    ph_idx: -1,
  }
}

// ── Relationship parsing ──────────────────────────────────────────────────────

pub(all) struct Relationship {
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
  slide_size : SlideSize,
  theme : ThemeData
) -> SlideData {
  let nodes = @xml.parse_xml(slide_xml)
  let shapes : Array[Shape] = []
  // Navigate p:sld → p:cSld
  let csld = @xml.find_path(nodes, ["p:sld", "p:cSld"])
  match csld {
    None => { slide_size, background: Color::none(), shapes }
    Some(csld_node) => {
      let csld_ch = @xml.get_children(csld_node)
      // Parse background: p:bg → p:bgPr → solidFill
      // Color::none() = no explicit bg (inherit from layout/master)
      let bg = match @xml.find_child(csld_ch, "p:bg") {
        None => Color::none()
        Some(bg_node) => {
          let bg_ch = @xml.get_children(bg_node)
          match @xml.find_child(bg_ch, "p:bgPr") {
            None => Color::none()
            Some(bgpr) => {
              let c = parse_solid_fill(@xml.get_children(bgpr), theme)
              c  // may be Color::none() if no solidFill found
            }
          }
        }
      }
      // Navigate to p:spTree
      match @xml.find_child(csld_ch, "p:spTree") {
        None => { slide_size, background: bg, shapes }
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
          { slide_size, background: bg, shapes }
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

// ── Level text defaults parsing ──────────────────────────────────────────────

///| Parse a level paragraph properties node (e.g. a:lvl1pPr) into LevelTextDefaults.
fn parse_level_defaults(ppr : @xml.XmlNode, theme : ThemeData) -> LevelTextDefaults {
  let align = match @xml.get_attr(ppr, "algn") { Some(v) => v; None => "" }
  let mar_l = match @xml.get_attr(ppr, "marL") { Some(v) => @xml.parse_int(v); None => -1 }
  let indent = attr_int(ppr, "indent")
  let ppr_ch = @xml.get_children(ppr)
  // Bullet
  let bu_char = match @xml.find_child(ppr_ch, "a:buChar") {
    None => ""
    Some(bc) => match @xml.get_attr(bc, "char") { Some(v) => v; None => "" }
  }
  let bu_none = match @xml.find_child(ppr_ch, "a:buNone") { Some(_) => true; None => false }
  // Default run properties from a:defRPr
  let (bold, italic, font_size, color, font_face) = match @xml.find_child(ppr_ch, "a:defRPr") {
    None => (false, false, 0, Color::none(), "")
    Some(drpr) => {
      let db = match @xml.get_attr(drpr, "b") { Some(v) => v == "1" || v == "true"; None => false }
      let di = match @xml.get_attr(drpr, "i") { Some(v) => v == "1" || v == "true"; None => false }
      let dsz = attr_int(drpr, "sz")
      let drpr_ch = @xml.get_children(drpr)
      let dclr = parse_solid_fill(drpr_ch, theme)
      let dff = child_attr(drpr_ch, "a:latin", "typeface")
      (db, di, dsz, dclr, dff)
    }
  }
  { font_size, bold, italic, color, font_face, align, mar_l, indent, bullet: bu_char, bullet_none: bu_none }
}

///| Parse a text style group (e.g. p:titleStyle) into TextStyleGroup.
fn parse_text_style_group(node : @xml.XmlNode, theme : ThemeData) -> TextStyleGroup {
  let ch = @xml.get_children(node)
  let levels : Array[LevelTextDefaults] = []
  let mut i = 0
  while i < 9 {
    let tag = "a:lvl" + int_to_str(i + 1) + "pPr"
    let lvl = match @xml.find_child(ch, tag) {
      None => LevelTextDefaults::empty()
      Some(ppr) => parse_level_defaults(ppr, theme)
    }
    levels.push(lvl)
    i = i + 1
  }
  { levels, }
}

///| Parse p:txStyles from slide master.
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

///| Parse a slide master XML string into SlideMasterData.
pub fn parse_slide_master(xml_str : String, slide_size : SlideSize, theme : ThemeData) -> SlideMasterData {
  let nodes = @xml.parse_xml(xml_str)
  let csld = @xml.find_path(nodes, ["p:sldMaster", "p:cSld"])
  let background = match csld {
    None => Color::none()
    Some(csld_node) => parse_bg_from_csld(@xml.get_children(csld_node), theme)
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
  // Parse p:txStyles
  let text_styles = match @xml.find_child(nodes, "p:sldMaster") {
    None => MasterTextStyles::empty()
    Some(master_node) => {
      let master_ch = @xml.get_children(master_node)
      match @xml.find_child(master_ch, "p:txStyles") {
        None => MasterTextStyles::empty()
        Some(txs) => parse_tx_styles(txs, theme)
      }
    }
  }
  // ignore slide_size parameter (just used for shape parsing context)
  ignore(slide_size)
  { background, shapes, text_styles }
}

// ── Slide layout parser ──────────────────────────────────────────────────────

///| Parse a slide layout XML string into SlideLayoutData.
pub fn parse_slide_layout(xml_str : String, slide_size : SlideSize, theme : ThemeData) -> SlideLayoutData {
  let nodes = @xml.parse_xml(xml_str)
  let csld = @xml.find_path(nodes, ["p:sldLayout", "p:cSld"])
  let background = match csld {
    None => Color::none()
    Some(csld_node) => parse_bg_from_csld(@xml.get_children(csld_node), theme)
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
  ignore(slide_size)
  { background, shapes }
}

// ── Integer→string (needed here too) ─────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }
