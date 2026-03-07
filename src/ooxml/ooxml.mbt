// OOXML / PPTX parser — Phase 2
//
// Extracts shape data from slide XML and .rels files.
// All positions/sizes are in EMUs (English Metric Units).
//   1 inch = 914400 EMU,  1 pt = 12700 EMU,  1 cm = 360000 EMU
// Standard slide: 9144000 x 6858000 EMU (10in x 7.5in)

// ── Shared constants ─────────────────────────────────────────────────────────

///|
/// 1 pt = 12700 EMU.
pub let emu_per_pt : Int = 12700

///|
/// Default stroke width: 1 pt in EMU.
pub let default_stroke_w : Int = 12700

///|
/// EMU per hundredth-of-a-point (12700 / 100).
pub let emu_per_hundredth_pt : Int = 127

///|
/// Rotation units: 60000ths of a degree per centidegree.
pub let rot_to_centideg : Int = 600

///|
/// Centidegrees per degree.
pub let centideg_per_deg : Int = 100

///|
/// Hundredths-of-a-point per EMU for char spacing (100ths-pt → px at 96 DPI).
pub let char_spc_dpi : Int = 96

///|
/// Hundredths-of-a-point per inch (for char spacing conversion).
pub let hundredth_pt_per_inch : Int = 7200

// ── Color ─────────────────────────────────────────────────────────────────────

///|
pub(all) struct Color {
  r : Int
  g : Int
  b : Int
  alpha : Int // 0-255, 255 = fully opaque
}

///|
pub fn Color::white() -> Color {
  { r: 255, g: 255, b: 255, alpha: 255 }
}

///|
pub fn Color::black() -> Color {
  { r: 0, g: 0, b: 0, alpha: 255 }
}

///|
pub fn Color::none() -> Color {
  { r: -1, g: -1, b: -1, alpha: 255 }
}

///|
pub fn Color::is_none(self : Color) -> Bool {
  self.r < 0
}

///|
/// Parse a 6-char "RRGGBB" or 8-char "AARRGGBB" hex color string → Color.
pub fn parse_hex_color(s : String) -> Color {
  if s.length() >= 8 {
    let a = parse_hex2(s, 0)
    { r: parse_hex2(s, 2), g: parse_hex2(s, 4), b: parse_hex2(s, 6), alpha: a }
  } else if s.length() >= 6 {
    {
      r: parse_hex2(s, 0),
      g: parse_hex2(s, 2),
      b: parse_hex2(s, 4),
      alpha: 255,
    }
  } else {
    Color::black()
  }
}

// ASCII code points for hex digit parsing

///|
let ascii_0 : Int = 48 // '0'

///|
let ascii_9 : Int = 57 // '9'

///|
let ascii_A : Int = 65 // 'A'

///|
let ascii_F : Int = 70 // 'F'

///|
let ascii_a : Int = 97 // 'a'

///|
let ascii_f : Int = 102 // 'f'

///|
fn hex_digit(c : Char) -> Int {
  let v = Char::to_int(c)
  if v >= ascii_0 && v <= ascii_9 {
    v - ascii_0
  } else if v >= ascii_A && v <= ascii_F {
    v - (ascii_A - 10)
  } else if v >= ascii_a && v <= ascii_f {
    v - (ascii_a - 10)
  } else {
    0
  }
}

///|
fn parse_hex2(s : String, offset : Int) -> Int {
  hex_digit(s.get_char(offset).unwrap()) * 16 +
  hex_digit(s.get_char(offset + 1).unwrap())
}

///|
pub fn Color::to_css(self : Color) -> String {
  if self.is_none() {
    "none"
  } else if self.alpha < 255 {
    "rgba(" +
    @xml.int_to_str(self.r) +
    "," +
    @xml.int_to_str(self.g) +
    "," +
    @xml.int_to_str(self.b) +
    "," +
    alpha_to_css_str(self.alpha) +
    ")"
  } else {
    "rgb(" +
    @xml.int_to_str(self.r) +
    "," +
    @xml.int_to_str(self.g) +
    "," +
    @xml.int_to_str(self.b) +
    ")"
  }
}

///|
/// Convert Color to hex string: 6-char "RRGGBB" if alpha==255, else 8-char "AARRGGBB".
pub fn Color::to_hex(self : Color) -> String {
  if self.is_none() {
    "none"
  } else if self.alpha < 255 {
    hex_byte(self.alpha) +
    hex_byte(self.r) +
    hex_byte(self.g) +
    hex_byte(self.b)
  } else {
    hex_byte(self.r) + hex_byte(self.g) + hex_byte(self.b)
  }
}

///|
/// Convert Color to 6-char "RRGGBB" (always, ignoring alpha). For OOXML val attributes.
pub fn Color::to_hex_rgb(self : Color) -> String {
  if self.is_none() {
    "none"
  } else {
    hex_byte(self.r) + hex_byte(self.g) + hex_byte(self.b)
  }
}

///|
fn hex_byte(v : Int) -> String {
  hex_nibble(v / 16) + hex_nibble(v % 16)
}

///|
/// Convert alpha (0-255) to CSS opacity string: "0", "0.XX", or "1".
pub fn alpha_to_css_str(alpha : Int) -> String {
  let a100 = alpha * 100 / 255
  if a100 <= 0 {
    "0"
  } else if a100 >= 100 {
    "1"
  } else {
    "0." +
    (if a100 < 10 {
      "0" + @xml.int_to_str(a100)
    } else {
      @xml.int_to_str(a100)
    })
  }
}

///|
fn hex_nibble(n : Int) -> String {
  match n {
    0 => "0"
    1 => "1"
    2 => "2"
    3 => "3"
    4 => "4"
    5 => "5"
    6 => "6"
    7 => "7"
    8 => "8"
    9 => "9"
    10 => "A"
    11 => "B"
    12 => "C"
    13 => "D"
    14 => "E"
    _ => "F"
  }
}

// ── Gradient fill ────────────────────────────────────────────────────────────

///|
pub(all) struct GradientStop {
  pos : Int // 0-100000 (1/1000%)
  color : Color // resolved RGB
}

///|
pub(all) struct GradientFill {
  stops : Array[GradientStop] // empty = none
  angle : Int // lin: 60000ths of degree; -1 = path gradient
  path_type : String // "circle","rect","shape"; "" = linear
  rot_with_shape : Bool
  fill_to_l : Int // path gradient fillToRect; -1 = unset
  fill_to_t : Int
  fill_to_r : Int
  fill_to_b : Int
  tile_flip : String // "none","x","y","xy"; "" = unset
}

///|
pub fn GradientFill::none() -> GradientFill {
  {
    stops: [],
    angle: 0,
    path_type: "",
    rot_with_shape: true,
    fill_to_l: -1,
    fill_to_t: -1,
    fill_to_r: -1,
    fill_to_b: -1,
    tile_flip: "",
  }
}

///|
pub fn GradientFill::is_none(self : GradientFill) -> Bool {
  self.stops.length() == 0
}

// ── Blip (image) fill ───────────────────────────────────────────────────

///|
pub(all) struct BlipFill {
  rid : String // r:embed relationship ID; "" = none
  stretch : Bool // true = stretch fill, false = tile
  src_l : Int // a:srcRect crop left (0-100000 千分率); 0 = no crop
  src_t : Int // crop top
  src_r : Int // crop right
  src_b : Int // crop bottom
  tile_tx : Int // a:tile tx offset (EMU); 0 = default
  tile_ty : Int // a:tile ty offset (EMU)
  tile_sx : Int // a:tile sx scale (100000 = 100%); 0 = unset
  tile_sy : Int // a:tile sy scale
  tile_flip : String // "none","x","y","xy"; "" = unset
  tile_algn : String // alignment: "tl","t","tr","l","ctr","r","bl","b","br"; "" = unset
}

///|
pub fn BlipFill::none() -> BlipFill {
  {
    rid: "",
    stretch: true,
    src_l: 0,
    src_t: 0,
    src_r: 0,
    src_b: 0,
    tile_tx: 0,
    tile_ty: 0,
    tile_sx: 0,
    tile_sy: 0,
    tile_flip: "",
    tile_algn: "",
  }
}

///|
pub fn BlipFill::is_none(self : BlipFill) -> Bool {
  self.rid.is_empty()
}

// ── Pattern fill ────────────────────────────────────────────────────────

///|
pub(all) struct PatternFill {
  prst : String // preset pattern name; "" = none
  fg_color : Color
  bg_color : Color
}

///|
pub fn PatternFill::none() -> PatternFill {
  { prst: "", fg_color: Color::none(), bg_color: Color::none() }
}

///|
pub fn PatternFill::is_none(self : PatternFill) -> Bool {
  self.prst.is_empty()
}

// ── Theme data ───────────────────────────────────────────────────────────────

///|
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
  major_ea_font : String // East Asian major font (from a:majorFont/a:ea)
  minor_ea_font : String // East Asian minor font (from a:minorFont/a:ea)
}

///|
pub fn ThemeData::default() -> ThemeData {
  {
    dk1: Color::black(),
    lt1: Color::white(),
    dk2: { r: 68, g: 84, b: 106, alpha: 255 },
    lt2: { r: 231, g: 230, b: 230, alpha: 255 },
    accent1: { r: 68, g: 114, b: 196, alpha: 255 },
    accent2: { r: 237, g: 125, b: 49, alpha: 255 },
    accent3: { r: 165, g: 165, b: 165, alpha: 255 },
    accent4: { r: 255, g: 192, b: 0, alpha: 255 },
    accent5: { r: 91, g: 155, b: 213, alpha: 255 },
    accent6: { r: 112, g: 173, b: 71, alpha: 255 },
    hlink: { r: 5, g: 99, b: 193, alpha: 255 },
    fol_hlink: { r: 149, g: 79, b: 163, alpha: 255 },
    major_font: "Calibri Light",
    minor_font: "Calibri",
    major_ea_font: "",
    minor_ea_font: "",
  }
}

///|
/// Resolve a scheme color name to a Color from the theme.
///  Handles aliasing: bg1→lt1, tx1→dk1, bg2→lt2, tx2→dk2, then delegates to resolve_slot.
pub fn resolve_scheme_color(name : String, theme : ThemeData) -> Color {
  let resolved = if name == "bg1" {
    "lt1"
  } else if name == "tx1" {
    "dk1"
  } else if name == "bg2" {
    "lt2"
  } else if name == "tx2" {
    "dk2"
  } else {
    name
  }
  resolve_slot(resolved, theme)
}

// ── RGB ↔ HSL conversion ────────────────────────────────────────────────────

///|
fn min3(a : Int, b : Int, c : Int) -> Int {
  let m = if a < b { a } else { b }
  if m < c {
    m
  } else {
    c
  }
}

///|
fn max3(a : Int, b : Int, c : Int) -> Int {
  let m = if a > b { a } else { b }
  if m > c {
    m
  } else {
    c
  }
}

// ── HSL color space constants (integer-scaled) ──────────────────────────────
// Hue: 0-3600 (tenths of a degree), Saturation/Lightness: 0-1000 (tenths of %)

///|
let rgb_max : Int = 255 // max RGB component value

///|
let hsl_scale : Int = 1000 // saturation/lightness precision (0.1% units)

///|
let hsl_half : Int = 500 // hsl_scale / 2

///|
let hue_full : Int = 3600 // full hue circle in tenths of degree

///|
let hue_sextant : Int = 600 // hue_full / 6 — one sixth of hue circle

///|
let hue_third : Int = 1200 // hue_full / 3

///|
let hue_two_thirds : Int = 2400 // hue_full * 2 / 3

///|
let rgb_sum_max : Int = 510 // rgb_max * 2

///|
/// Convert RGB (0-255) to HSL (h: 0-3600, s: 0-1000, l: 0-1000).
///  Uses integer arithmetic scaled by 1000 to avoid floating point.
fn rgb_to_hsl(r : Int, g : Int, b : Int) -> (Int, Int, Int) {
  let max_c = max3(r, g, b)
  let min_c = min3(r, g, b)
  let l = (max_c + min_c) * hsl_scale / rgb_sum_max
  if max_c == min_c {
    return (0, 0, l)
  }
  let d = max_c - min_c
  let s = if l > hsl_half {
    d * hsl_scale / (rgb_sum_max - max_c - min_c)
  } else {
    d * hsl_scale / (max_c + min_c)
  }
  let h = if r == max_c {
    (g - b) * hue_sextant / d
  } else if g == max_c {
    hue_third + (b - r) * hue_sextant / d
  } else {
    hue_two_thirds + (r - g) * hue_sextant / d
  }
  let h2 = if h < 0 { h + hue_full } else { h }
  (h2, s, l)
}

///|
fn hue_to_rgb(p : Int, q : Int, t_in : Int) -> Int {
  let t = if t_in < 0 {
    t_in + hue_full
  } else if t_in > hue_full {
    t_in - hue_full
  } else {
    t_in
  }
  if t < hue_sextant {
    (p + (q - p) * t / hue_sextant) * rgb_max / hsl_scale
  } else if t < hue_full / 2 {
    q * rgb_max / hsl_scale
  } else if t < hue_two_thirds {
    (p + (q - p) * (hue_two_thirds - t) / hue_sextant) * rgb_max / hsl_scale
  } else {
    p * rgb_max / hsl_scale
  }
}

///|
/// Convert HSL (h: 0-3600, s: 0-1000, l: 0-1000) to RGB Color.
fn hsl_to_rgb(h : Int, s : Int, l : Int) -> Color {
  if s == 0 {
    let v = l * rgb_max / hsl_scale
    return { r: v, g: v, b: v, alpha: 255 }
  }
  let q = if l < hsl_half {
    l * (hsl_scale + s) / hsl_scale
  } else {
    l + s - l * s / hsl_scale
  }
  let p = 2 * l - q
  let r = hue_to_rgb(p, q, h + hue_third)
  let g = hue_to_rgb(p, q, h)
  let b = hue_to_rgb(p, q, h - hue_third)
  let clamp = fn(v : Int) -> Int {
    if v < 0 {
      0
    } else if v > rgb_max {
      rgb_max
    } else {
      v
    }
  }
  { r: clamp(r), g: clamp(g), b: clamp(b), alpha: 255 }
}

// ── Color modifiers ─────────────────────────────────────────────────────────

///|
/// Apply OOXML color modifier children (lumMod, lumOff, shade, tint, etc.)
///  to a base color. modifier_nodes are children of <a:schemeClr> or <a:srgbClr>.
pub fn apply_color_modifiers(
  color : Color,
  modifier_nodes : Array[@xml.XmlNode],
) -> Color {
  let n = modifier_nodes.length()
  if n == 0 {
    return color
  }
  let mut r = color.r
  let mut g = color.g
  let mut b = color.b
  let mut alpha_val = color.alpha
  let mut lum_mod = -1
  let mut lum_off = 0
  let mut i = 0
  while i < n {
    match modifier_nodes[i] {
      @xml.XmlNode::Element(tag, _, _) => {
        let val = attr_int(modifier_nodes[i], "val")
        if tag == "a:shade" {
          // shade: multiply each RGB component by val/pct_100
          r = r * val / pct_100
          g = g * val / pct_100
          b = b * val / pct_100
        } else if tag == "a:tint" {
          // tint: move each component toward rgb_max
          r = r + (rgb_max - r) * val / pct_100
          g = g + (rgb_max - g) * val / pct_100
          b = b + (rgb_max - b) * val / pct_100
        } else if tag == "a:lumMod" {
          lum_mod = val
        } else if tag == "a:lumOff" {
          lum_off = val
        } else if tag == "a:satMod" {
          // Saturation modification via HSL
          let (h, s, l) = rgb_to_hsl(r, g, b)
          let new_s = s * val / pct_100
          let ns = if new_s > hsl_scale { hsl_scale } else { new_s }
          let c2 = hsl_to_rgb(h, ns, l)
          r = c2.r
          g = c2.g
          b = c2.b
        } else if tag == "a:alpha" {
          // alpha: val is in 1/1000th percent (0-100000)
          alpha_val = val * 255 / pct_100
        } else if tag == "a:alphaModFix" {
          // Fixed alpha modulation
          alpha_val = alpha_val * val / pct_100
        }
      }
      _ => ()
    }
    i = i + 1
  }
  // Apply lumMod/lumOff together (they work as a pair)
  if lum_mod >= 0 {
    let (h, s, l) = rgb_to_hsl(r, g, b)
    let mut new_l = l * lum_mod / pct_100 + lum_off * hsl_scale / pct_100
    if new_l < 0 {
      new_l = 0
    }
    if new_l > hsl_scale {
      new_l = hsl_scale
    }
    let c2 = hsl_to_rgb(h, s, new_l)
    r = c2.r
    g = c2.g
    b = c2.b
  }
  // Clamp
  if r < 0 {
    r = 0
  } else if r > rgb_max {
    r = rgb_max
  }
  if g < 0 {
    g = 0
  } else if g > rgb_max {
    g = rgb_max
  }
  if b < 0 {
    b = 0
  } else if b > rgb_max {
    b = rgb_max
  }
  if alpha_val < 0 {
    alpha_val = 0
  } else if alpha_val > 255 {
    alpha_val = 255
  }
  { r, g, b, alpha: alpha_val }
}

// ── Master / Layout types ────────────────────────────────────────────────────

///|
pub(all) struct LevelTextDefaults {
  font_size : Int // hundredths-of-pt (sz); 0 = unset
  bold : Bool
  italic : Bool
  color : Color // Color::none() = unset
  font_face : String // "" = inherit
  ea_font : String // "" = inherit
  align : String // "" = inherit
  mar_l : Int // EMU; -1 = unset
  indent : Int // EMU; 0 = unset
  bullet : String // "" = none
  bullet_none : Bool
}

///|
pub fn LevelTextDefaults::empty() -> LevelTextDefaults {
  {
    font_size: 0,
    bold: false,
    italic: false,
    color: Color::none(),
    font_face: "",
    ea_font: "",
    align: "",
    mar_l: -1,
    indent: 0,
    bullet: "",
    bullet_none: false,
  }
}

///|
pub(all) struct TextStyleGroup {
  levels : Array[LevelTextDefaults] // max_outline_levels elements (lvl1pPr–lvl9pPr)
}

///|
pub fn TextStyleGroup::empty() -> TextStyleGroup {
  let levels : Array[LevelTextDefaults] = []
  let mut i = 0
  while i < max_outline_levels {
    levels.push(LevelTextDefaults::empty())
    i = i + 1
  }
  { levels, }
}

///|
pub(all) struct MasterTextStyles {
  title_style : TextStyleGroup
  body_style : TextStyleGroup
  other_style : TextStyleGroup
}

///|
pub fn MasterTextStyles::empty() -> MasterTextStyles {
  {
    title_style: TextStyleGroup::empty(),
    body_style: TextStyleGroup::empty(),
    other_style: TextStyleGroup::empty(),
  }
}

// ── Color map (scheme color remapping) ───────────────────────────────────────

///|
pub(all) struct ColorMap {
  bg1 : String
  tx1 : String
  bg2 : String
  tx2 : String
  accent1 : String
  accent2 : String
  accent3 : String
  accent4 : String
  accent5 : String
  accent6 : String
  hlink : String
  fol_hlink : String
}

///|
/// Default color map: bg1→lt1, tx1→dk1, bg2→lt2, tx2→dk2, etc.
pub fn ColorMap::default() -> ColorMap {
  {
    bg1: "lt1",
    tx1: "dk1",
    bg2: "lt2",
    tx2: "dk2",
    accent1: "accent1",
    accent2: "accent2",
    accent3: "accent3",
    accent4: "accent4",
    accent5: "accent5",
    accent6: "accent6",
    hlink: "hlink",
    fol_hlink: "folHlink",
  }
}

///|
pub(all) struct SlideMasterData {
  background : Color
  shapes : Array[Shape]
  text_styles : MasterTextStyles
  clr_map : ColorMap
}

///|
pub(all) struct SlideLayoutData {
  background : Color // Color::none() = inherit from master
  shapes : Array[Shape]
  clr_map_ovr : ColorMap? // None = inherit master's color map
}

// ── Text model ────────────────────────────────────────────────────────────────

///|
pub(all) struct TextRun {
  text : String
  bold : Bool
  italic : Bool
  font_size : Int // in hundredths-of-a-point (sz attr); 0 = inherit
  color : Color
  font_face : String // explicit typeface from <a:latin>; "" = use theme font
  ea_font : String // East Asian typeface from <a:ea>; "" = use theme ea font
  cs_font : String // Complex Script typeface from <a:cs>; "" = use theme cs font
  sym_font : String // Symbol typeface from <a:sym>; "" = none
  underline : String // "sng", "dbl", "heavy", etc.; "" = none
  strike : String // "sngStrike", "dblStrike"; "" = none
  baseline : Int // superscript (+30000 = 30%) / subscript (-25000); 0 = normal
  char_spacing : Int // character spacing in hundredths-of-a-point (spc attr); 0 = default
  kern : Int // kerning in hundredths-of-a-point (kern attr); 0 = default
  cap : String // "all", "small"; "" = none
  hlink_rid : String // hyperlink relationship ID from <a:hlinkClick r:id="...">; "" = none
  hlink_mouse_over_rid : String // hover hyperlink from <a:hlinkHover r:id="...">; "" = none
}

///|
pub(all) struct TabStop {
  pos : Int // position in EMU
  algn : String // "l", "ctr", "r", "dec"; "" = left (default)
}

///|
pub(all) struct TextParagraph {
  runs : Array[TextRun]
  align : String // "l", "c", "r", "just" or ""
  level : Int // indent level (0 = top-level, from a:pPr lvl="N")
  spc_before : Int // spacing before in EMU (from a:spcBef); 0 = default
  spc_after : Int // spacing after in EMU (from a:spcAft); 0 = default
  line_spacing : Int // line spacing: >0 = EMU, <0 = percentage (-100000=100%), 0 = default
  mar_l : Int // left margin in EMU (from marL attr in a:pPr); 0 = default
  indent : Int // first-line indent in EMU (from indent attr in a:pPr); 0 = default
  bullet : String // bullet char (from a:buChar); "" = none/inherit
  bullet_auto : String // auto-number type (from a:buAutoNum); "" = none
  bullet_none : Bool // explicitly no bullet (a:buNone)
  bullet_font : String // a:buFont typeface; "" = run font
  bullet_size : Int // a:buSzPct (percentage 100000=100%) or negative for a:buSzPts; 0 = inherit
  bullet_color : Color // a:buClr; Color::none() = inherit
  bullet_img_rid : String // a:buBlip r:embed; "" = none
  tab_stops : Array[TabStop] // a:tabLst/a:tab
  rtl : Bool // right-to-left paragraph direction (a:pPr rtl="1")
}

// ── Text body properties ────────────────────────────────────────────────────

///|
/// 100000 = 100% in OOXML percentage units (fontScale, lnSpcReduction, etc.)
pub let pct_100 : Int = 100000

///|
/// Prefix for custom dash stroke_dash strings. Format: "custDash:d1,sp1;d2,sp2;..."
pub let cust_dash_prefix : String = "custDash:"

///|
pub let cust_dash_prefix_len : Int = 9

///|
/// Maximum outline levels in OOXML (lvl1pPr through lvl9pPr).
pub let max_outline_levels : Int = 9

///|
pub(all) struct BodyProps {
  anchor : String // vertical alignment: "t", "ctr", "b", "" = default (top)
  l_ins : Int // left inset in EMU (-1 = default 91440 ≈ 0.1in)
  t_ins : Int // top inset in EMU (-1 = default 45720 ≈ 0.05in)
  r_ins : Int // right inset in EMU (-1 = default 91440)
  b_ins : Int // bottom inset in EMU (-1 = default 45720)
  auto_fit : Bool // <a:spAutoFit/> — expand shape to fit text
  font_scale : Int // <a:normAutofit fontScale="..."/> -1=unset, 0-100000=percent
  ln_spc_reduction : Int // <a:normAutofit lnSpcReduction="..."/> -1=unset, 0-100000=percent
  wrap : String // "square" (default=wrap) or "none"; "" = default (square)
  rot : Int // text rotation in 60000ths of a degree; 0 = default
  vert : String // vertical text: "vert"/"eaVert"/"vert270"/"wordArtVert"; "" = horizontal
  num_cols : Int // number of text columns; 0 = default (1)
  col_spacing : Int // spacing between columns in EMU; 0 = default
}

///|
pub fn BodyProps::default() -> BodyProps {
  {
    anchor: "",
    l_ins: -1,
    t_ins: -1,
    r_ins: -1,
    b_ins: -1,
    auto_fit: false,
    font_scale: -1,
    ln_spc_reduction: -1,
    wrap: "",
    rot: 0,
    vert: "",
    num_cols: 0,
    col_spacing: 0,
  }
}

// ── Table model ──────────────────────────────────────────────────────────────

///|
pub(all) struct TableCell {
  paragraphs : Array[TextParagraph]
  fill : Color
  grad_fill : GradientFill
  // Cell merge
  grid_span : Int // horizontal merge span (1 = no merge)
  row_span : Int // vertical merge span (1 = no merge)
  v_merge : Bool // true = this cell is merged into the cell above
  // Cell borders (line width in EMU, 0 = default, -1 = none)
  bdr_l_w : Int
  bdr_r_w : Int
  bdr_t_w : Int
  bdr_b_w : Int
  bdr_l_color : Color
  bdr_r_color : Color
  bdr_t_color : Color
  bdr_b_color : Color
  // Diagonal borders
  bdr_tl_br_w : Int // top-left to bottom-right
  bdr_tl_br_color : Color
  bdr_bl_tr_w : Int // bottom-left to top-right
  bdr_bl_tr_color : Color
  // Cell margins in EMU (-1 = use default)
  mar_l : Int
  mar_r : Int
  mar_t : Int
  mar_b : Int
  // Cell anchor: "t" (top, default), "ctr" (center), "b" (bottom)
  anchor : String
}

///|
pub(all) struct TableRow {
  height : Int // EMU
  cells : Array[TableCell]
}

///|
pub(all) struct TableData {
  col_widths : Array[Int] // EMU per column
  rows : Array[TableRow]
  // Table properties from a:tblPr
  style_id : String // tblStyleId UUID (empty = none)
  first_row : Bool
  last_row : Bool
  first_col : Bool
  last_col : Bool
  band_row : Bool
  band_col : Bool
}

// ── Table style model ────────────────────────────────────────────────────────

///|
pub(all) struct TableStyleCell {
  fill : Color
  grad_fill : GradientFill
  bdr_l_w : Int
  bdr_r_w : Int
  bdr_t_w : Int
  bdr_b_w : Int
  bdr_l_color : Color
  bdr_r_color : Color
  bdr_t_color : Color
  bdr_b_color : Color
  bold : Bool
  italic : Bool
  font_color : Color
}

///|
pub(all) struct TableStyleDef {
  id : String
  whole_tbl : TableStyleCell
  band1_h : TableStyleCell
  band2_h : TableStyleCell
  band1_v : TableStyleCell
  band2_v : TableStyleCell
  first_row : TableStyleCell
  last_row : TableStyleCell
  first_col : TableStyleCell
  last_col : TableStyleCell
}

///|
pub fn TableStyleCell::none() -> TableStyleCell {
  {
    fill: Color::none(), grad_fill: GradientFill::none(),
    bdr_l_w: 0, bdr_r_w: 0, bdr_t_w: 0, bdr_b_w: 0,
    bdr_l_color: Color::none(), bdr_r_color: Color::none(),
    bdr_t_color: Color::none(), bdr_b_color: Color::none(),
    bold: false, italic: false, font_color: Color::none(),
  }
}

///|
pub fn TableStyleCell::is_none(self : TableStyleCell) -> Bool {
  self.fill.is_none() && self.font_color.is_none() && not(self.bold) && not(self.italic) &&
    self.bdr_l_w == 0 && self.bdr_r_w == 0 && self.bdr_t_w == 0 && self.bdr_b_w == 0
}

///|
pub fn TableStyleDef::none() -> TableStyleDef {
  {
    id: "", whole_tbl: TableStyleCell::none(),
    band1_h: TableStyleCell::none(), band2_h: TableStyleCell::none(),
    band1_v: TableStyleCell::none(), band2_v: TableStyleCell::none(),
    first_row: TableStyleCell::none(), last_row: TableStyleCell::none(),
    first_col: TableStyleCell::none(), last_col: TableStyleCell::none(),
  }
}

// ── Shape geometry ────────────────────────────────────────────────────────────

///|
pub(all) enum ShapeGeom {
  Rect
  Ellipse
  RoundRect
  Line
  Connector(String, Array[Int]) // preset name, adjustment values (0-100000)
  Other(String, Array[Int]) // preset name, adjustment values
  Custom(CustomGeomData) // a:custGeom freeform/custom geometry
}

///|
pub(all) struct CustomGeomData {
  gdlst : String // guide formulas: "name1=fmla1;name2=fmla2"
  paths : String // compact path: "M x y L x y Z" ("|" separates sub-paths)
  path_w : Int // path coordinate space width (0 = use shape width)
  path_h : Int // path coordinate space height (0 = use shape height)
  rect_l : String // text rect left guide name (empty = use shape bound)
  rect_t : String // text rect top guide name
  rect_r : String // text rect right guide name
  rect_b : String // text rect bottom guide name
  cxn_lst : String // connection points: "x1,y1,ang1;x2,y2,ang2;..." (guide names)
}

///|
pub fn parse_geom(prst : String, avlst : Array[Int]) -> ShapeGeom {
  if prst == "rect" {
    ShapeGeom::Rect
  } else if prst == "ellipse" || prst == "oval" {
    ShapeGeom::Ellipse
  } else if prst == "roundRect" {
    ShapeGeom::RoundRect
  } else if prst == "line" || prst == "straightConnector1" {
    ShapeGeom::Line
  } else {
    ShapeGeom::Other(prst, avlst)
  }
}

///|
/// Parse connector geometry with adjustment values.
pub fn parse_connector_geom(prst : String, avlst : Array[Int]) -> ShapeGeom {
  if prst == "straightConnector1" || prst == "line" {
    ShapeGeom::Line
  } else {
    ShapeGeom::Connector(prst, avlst)
  }
}

///|
/// Convert ShapeGeom back to OOXML preset geometry name.
pub fn ShapeGeom::to_prst(self : ShapeGeom) -> String {
  match self {
    ShapeGeom::Rect => "rect"
    ShapeGeom::Ellipse => "ellipse"
    ShapeGeom::RoundRect => "roundRect"
    ShapeGeom::Line => "line"
    ShapeGeom::Connector(prst, _) => prst
    ShapeGeom::Other(s, _) => s
    ShapeGeom::Custom(_) => "custGeom"
  }
}

// ── Shape ─────────────────────────────────────────────────────────────────────

///|
pub(all) struct ShapeTransform {
  x : Int
  y : Int
  cx : Int
  cy : Int
  rot : Int // rotation in 1/60000 degree units
  flip_h : Bool
  flip_v : Bool
}

///|
pub fn ShapeTransform::default() -> ShapeTransform {
  { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flip_h: false, flip_v: false }
}

///|
pub(all) struct GroupShapeData {
  ch_off_x : Int // child offset x (EMU)
  ch_off_y : Int // child offset y (EMU)
  ch_ext_cx : Int // child extent cx (EMU)
  ch_ext_cy : Int // child extent cy (EMU)
  children : Array[Shape]
}

///|
pub(all) enum ShapeKind {
  AutoShape(ShapeGeom)
  Picture(String) // relationship ID (r:embed)
  TableShape(TableData)
  GroupShape(GroupShapeData)
  Other
}

///|
/// Parsed stroke properties from a:ln.
pub(all) struct StrokeProps {
  color : Color
  width : Int
  dash : String
  cap : String
  join : String
  miter_limit : Int
  head_type : String
  head_w : String
  head_len : String
  tail_type : String
  tail_w : String
  tail_len : String
  cmpd : String
  no_fill : Bool
}

///|
pub fn StrokeProps::default() -> StrokeProps {
  {
    color: Color::none(),
    width: 0,
    dash: "",
    cap: "",
    join: "",
    miter_limit: 0,
    head_type: "",
    head_w: "",
    head_len: "",
    tail_type: "",
    tail_w: "",
    tail_len: "",
    cmpd: "",
    no_fill: false,
  }
}

///|
pub(all) struct Shape {
  kind : ShapeKind
  transform : ShapeTransform
  fill : Color // Color::none() = no fill
  grad_fill : GradientFill // GradientFill::none() = no gradient
  blip_fill : BlipFill // BlipFill::none() = no image fill
  patt_fill : PatternFill // PatternFill::none() = no pattern fill
  stroke : Color // Color::none() = no stroke
  stroke_w : Int // stroke width in EMU (0 = default ~12700)
  stroke_dash : String // preset dash: dash/dot/dashDot/lgDash/sysDot/sysDash/etc. "" = solid
  stroke_cap : String // line cap: flat/rnd/sq. "" = default (flat)
  stroke_join : String // line join: round/bevel/miter. "" = default
  stroke_miter_limit : Int // miter limit (EMU percentage, e.g. 800000). 0 = default
  stroke_head_type : String // arrow head type: triangle/stealth/diamond/oval/arrow. "" = none
  stroke_head_w : String // arrow head width: sm/med/lg. "" = default (med)
  stroke_head_len : String // arrow head length: sm/med/lg. "" = default (med)
  stroke_tail_type : String // arrow tail type
  stroke_tail_w : String // arrow tail width
  stroke_tail_len : String // arrow tail length
  stroke_cmpd : String // compound line: sng/dbl/thickThin/thinThick/tri. "" = default (sng)
  stroke_no_fill : Bool // explicit a:noFill on a:ln (line exists but invisible)
  paragraphs : Array[TextParagraph]
  body_props : BodyProps // text body properties (anchor, insets)
  ph_type : String // "title","body","ctrTitle","subTitle","dt","ftr","sldNum","" (non-PH)
  ph_idx : Int // -1 = none
  st_cxn_id : Int // start connection shape id (-1 = none)
  st_cxn_idx : Int // start connection point index (-1 = none)
  end_cxn_id : Int // end connection shape id (-1 = none)
  end_cxn_idx : Int // end connection point index (-1 = none)
}

///|
/// Create a Shape with default values (no fill, no stroke, no text).
pub fn Shape::default_with(
  kind : ShapeKind,
  transform : ShapeTransform,
) -> Shape {
  {
    kind,
    transform,
    fill: Color::none(),
    grad_fill: GradientFill::none(),
    blip_fill: BlipFill::none(),
    patt_fill: PatternFill::none(),
    stroke: Color::none(),
    stroke_w: 0,
    stroke_dash: "",
    stroke_cap: "",
    stroke_join: "",
    stroke_miter_limit: 0,
    stroke_head_type: "",
    stroke_head_w: "",
    stroke_head_len: "",
    stroke_tail_type: "",
    stroke_tail_w: "",
    stroke_tail_len: "",
    stroke_cmpd: "",
    stroke_no_fill: false,
    paragraphs: [],
    body_props: BodyProps::default(),
    ph_type: "",
    ph_idx: -1,
    st_cxn_id: -1,
    st_cxn_idx: -1,
    end_cxn_id: -1,
    end_cxn_idx: -1,
  }
}

// ── Slide data ────────────────────────────────────────────────────────────────

///|
pub(all) struct SlideSize {
  cx : Int
  cy : Int
}

///|
pub fn SlideSize::default() -> SlideSize {
  // Standard widescreen 10in x 7.5in
  { cx: 9144000, cy: 6858000 }
}

///|
pub(all) struct SlideData {
  slide_size : SlideSize
  background : Color
  bg_grad : GradientFill // background gradient; GradientFill::none() = solid
  shapes : Array[Shape]
}

// ── Relationship type ─────────────────────────────────────────────────────────

///|
pub(all) struct Relationship {
  id : String
  target : String
  type_ : String
}
