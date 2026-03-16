// SVG renderer for PPTX slides — Round-trip capable
//
// All OOXML coordinates are in EMU (English Metric Units).
// This renderer converts EMU → pixels at a fixed target width of 960 px,
// producing an SVG with explicit pixel dimensions and NO viewBox transform.
//
// Round-trip support: each SVG element carries data-ooxml-* attributes
// preserving original OOXML values (EMU, hundredths-of-a-point, etc.)
// so the SVG can be parsed back to SlideData losslessly.
//
// NOTE: No StringBuilder.  All strings built via `+` (wasm:js-string "concat"),
// which is polyfillable in Tier-2/3 browsers.

// ── EMU constants ────────────────────────────────────────────────────────────
// 1 pt = 12700 EMU, 1 inch = 914400 EMU

///|
let default_font_size_emu_ : Int = 228600 // 18 pt

///|
let default_table_font_size_emu : Int = 127000 // 10 pt — built-in table styles typically default to 10pt

///|
let default_l_inset : Int = 91440 // 0.1 in (left / right)

///|
let default_t_inset : Int = 45720 // 0.05 in (top / bottom)

///|
let level_indent_emu : Int = 342900 // indent per bullet level

///|
let svg_target_width : Int = 960 // output SVG width in px

///|
let default_emu_per_px : Int = 9525 // EMU per pixel at 96 DPI

///|
let line_height_numer : Int = 12 // default line-height = font-size * 12/10

///|
let line_height_denom : Int = 10

///|
let pct_line_height_numer : Int = 14 // for %-based spacing: font-size * 14/10 ≈ 1.4 (CJK font metrics)

///|
let pct_line_height_denom : Int = 10


///|
let default_text_fill : String = "black"

///|
let image_placeholder_fill : String = "#dddddd"

///|
let table_border_color : String = "#999999" // fallback when no border color specified

///|
let round_rect_radius_divisor : Int = 20 // corner radius = width / 20

///|
let group_child_idx_multiplier : Int = 1000 // child_idx = shape_idx * 1000 + ci

///|
let baseline_shift_extra_numer : Int = 2 // extra spacing for super/subscript = font_px * numer / denom

///|
let baseline_shift_extra_denom : Int = 5

// ── OOXML unit constants ─────────────────────────────────────────────────────

///|
let ooxml_pct_denom : Int = @ooxml.pct_100 // OOXML percentage denominator (100000 = 100%)

///|
let ooxml_angle_unit : Int = 60000 // 1 degree = 60000 units in OOXML

///|
let ooxml_stop_pct : Int = 1000 // stop pos 0-100000 → percentage: divide by 1000

///|
let ooxml_fill_to_pct : Int = 2000 // fillToRect 0-100000 → center%: divide by 2000

// ── SVG element ID builders ─────────────────────────────────────────────────
// Each SVG <defs> element (gradient, pattern, filter, mask) needs a unique ID.
// All IDs follow the pattern: prefix + sep + slide_num + sep + shape_idx.

///|
/// Core ID builder. All SVG element IDs are built through this function.
fn svg_id(prefix : String, sep : String, slide_num : Int, shape_idx : Int) -> String {
  prefix + sep + int_to_str(slide_num) + sep + int_to_str(shape_idx)
}

///|
/// Build a gradient definition ID. Negative shape_idx uses "-n" prefix.
fn grad_id(slide_num : Int, shape_idx : Int) -> String {
  if shape_idx >= 0 {
    svg_id("grad-s", "-", slide_num, shape_idx)
  } else {
    "grad-s" + int_to_str(slide_num) + "-n" + int_to_str(-shape_idx)
  }
}

///|
fn patt_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("patt-s", "-", slide_num, shape_idx)
}

///|
fn stroke_grad_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("sgrad-s", "-", slide_num, shape_idx)
}

///|
fn stroke_patt_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("spatt-s", "-", slide_num, shape_idx)
}

///|
fn effect_filter_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("eff_s", "_", slide_num, shape_idx)
}

///|
fn refl_shape_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("rshp_s", "_", slide_num, shape_idx)
}

///|
fn refl_mask_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("rmsk_s", "_", slide_num, shape_idx)
}

///|
fn blip_filter_id(slide_num : Int, shape_idx : Int) -> String {
  svg_id("blipfx_s", "_", slide_num, shape_idx)
}

///|
fn blip_clip_id(prefix : String, slide_num : Int, shape_idx : Int) -> String {
  prefix + "_s" + int_to_str(slide_num) + "_" + int_to_str(shape_idx)
}

// ── Shared helpers (from @xml) ───────────────────────────────────────────────

///|
fn int_to_str(n : Int) -> String {
  @xml.int_to_str(n)
}

// ── SVG attribute helpers ─────────────────────────────────────────────────────

///|
fn a(name : String, value : String) -> String {
  " " + name + "=\"" + value + "\""
}

///|
fn ai(name : String, value : Int) -> String {
  a(name, int_to_str(value))
}

///|
/// Shorthand for data-ooxml-* attributes.
fn da(name : String, value : String) -> String {
  " data-ooxml-" + name + "=\"" + value + "\""
}

///|
fn dai(name : String, value : Int) -> String {
  da(name, int_to_str(value))
}

///|
fn bool_str(b : Bool) -> String {
  if b {
    "1"
  } else {
    "0"
  }
}

// ── Stroke rendering helpers ────────────────────────────────────────────────

///|
/// Convert OOXML preset dash name to SVG stroke-dasharray attribute.
fn stroke_dasharray(dash : String, sw : Int) -> String {
  if dash.is_empty() {
    return ""
  }
  // Dash/gap values scaled relative to stroke width
  let s = int_to_str(sw)
  let s2 = int_to_str(sw * 2)
  let s4 = int_to_str(sw * 4)
  let s8 = int_to_str(sw * 8)
  let pattern = if dash == "dash" {
    s4 + "," + s2
  } else if dash == "dot" {
    s + "," + s2
  } else if dash == "dashDot" {
    s4 + "," + s2 + "," + s + "," + s2
  } else if dash == "lgDash" {
    s8 + "," + s2
  } else if dash == "sysDot" {
    s + "," + s
  } else if dash == "sysDash" {
    s2 + "," + s
  } else if dash == "sysDashDot" {
    s2 + "," + s + "," + s + "," + s
  } else if dash == "sysDashDotDot" {
    s2 + "," + s + "," + s + "," + s + "," + s + "," + s
  } else if dash == "dashDotDot" {
    s4 + "," + s2 + "," + s + "," + s2 + "," + s + "," + s2
  } else if dash == "lgDashDot" {
    s8 + "," + s2 + "," + s + "," + s2
  } else if dash == "lgDashDotDot" {
    s8 + "," + s2 + "," + s + "," + s2 + "," + s + "," + s2
  } else if @xml.str_starts_with(dash, @ooxml.cust_dash_prefix) {
    // Custom dash format: "custDash:d1,sp1;d2,sp2;..." (percentages * 100 of line width)
    // Convert to SVG stroke-dasharray by scaling with stroke width
    let data = @xml.str_suffix(dash, @ooxml.cust_dash_prefix_len)
    let mut result = ""
    let data_len = data.length()
    let mut pos = 0
    while pos < data_len {
      // Parse next number (before , or ;)
      let mut num_end = pos
      while num_end < data_len {
        let c = data.get_char(num_end).unwrap()
        if c == ',' || c == ';' {
          break
        }
        num_end = num_end + 1
      }
      if num_end > pos {
        let val = @xml.parse_int(@xml.str_substring(data, pos, num_end))
        // val is in percentage * 100 (e.g. 300000 = 300%). Scale: val * sw / ooxml_pct_denom
        let px_val = val * sw / ooxml_pct_denom
        let clamped = if px_val < 1 { 1 } else { px_val }
        if not(result.is_empty()) {
          result = result + ","
        }
        result = result + int_to_str(clamped)
      }
      pos = num_end + 1
    }
    result
  } else {
    return ""
  }
  a("stroke-dasharray", pattern)
}

///|
/// Convert OOXML line cap to SVG stroke-linecap attribute.
fn stroke_linecap(cap : String) -> String {
  if cap == "flat" {
    a("stroke-linecap", "butt")
  } else if cap == "rnd" {
    a("stroke-linecap", "round")
  } else if cap == "sq" {
    a("stroke-linecap", "square")
  } else {
    ""
  }
}

///|
/// Convert OOXML line join to SVG stroke-linejoin + miterlimit attributes.
fn stroke_linejoin(join : String, miter_limit : Int) -> String {
  if join == "round" {
    a("stroke-linejoin", "round")
  } else if join == "bevel" {
    a("stroke-linejoin", "bevel")
  } else if join == "miter" {
    let ml = if miter_limit > 0 {
      // OOXML miter limit is percentage * 1000 (e.g. 800000 = 800%)
      // SVG miterlimit is a plain ratio
      let ratio = miter_limit / ooxml_pct_denom
      if ratio > 1 {
        a("stroke-miterlimit", int_to_str(ratio))
      } else {
        ""
      }
    } else {
      ""
    }
    a("stroke-linejoin", "miter") + ml
  } else {
    ""
  }
}

///|
/// Check if an arrow type string represents an active arrow.
fn has_arrow(arrow_type : String) -> Bool {
  not(arrow_type.is_empty()) && arrow_type != "none"
}

///|
/// Generate arrow marker ID: "ah-s{slide}-{shape}-h" or "ah-s{slide}-{shape}-t".
fn arrow_marker_id(slide_num : Int, shape_idx : Int, is_head : Bool) -> String {
  "ah-s" +
  int_to_str(slide_num) +
  "-" +
  int_to_str(shape_idx) +
  (if is_head { "-h" } else { "-t" })
}

///|
/// Arrow marker size from sm/med/lg.
fn arrow_size(sz : String) -> Int {
  if sz == "sm" {
    3
  } else if sz == "lg" {
    7
  } else {
    5
  }
}

///|
/// Render arrow marker SVG definition.
fn render_arrow_marker(
  id : String,
  arrow_type : String,
  w_str : String,
  len_str : String,
  color : @ooxml.Color,
) -> String {
  let w = arrow_size(w_str)
  let h = arrow_size(len_str)
  let ws = int_to_str(w)
  let hs = int_to_str(h)
  let fill_color = color.to_css()
  let path = if arrow_type == "triangle" {
    "<path d=\"M 0 0 L 10 5 L 0 10 z\"" + a("fill", fill_color) + "/>"
  } else if arrow_type == "stealth" {
    "<path d=\"M 0 0 L 10 5 L 0 10 L 3 5 z\"" + a("fill", fill_color) + "/>"
  } else if arrow_type == "diamond" {
    "<path d=\"M 0 5 L 5 0 L 10 5 L 5 10 z\"" + a("fill", fill_color) + "/>"
  } else if arrow_type == "oval" {
    "<ellipse cx=\"5\" cy=\"5\" rx=\"5\" ry=\"5\"" +
    a("fill", fill_color) +
    "/>"
  } else if arrow_type == "arrow" {
    "<path d=\"M 0 0 L 10 5 L 0 10\" fill=\"none\"" +
    a("stroke", fill_color) +
    " stroke-width=\"2\"/>"
  } else {
    return ""
  }
  "<marker id=\"" +
  id +
  "\" viewBox=\"0 0 10 10\" refX=\"10\" refY=\"5\" orient=\"auto-start-reverse\"" +
  a("markerWidth", ws) +
  a("markerHeight", hs) +
  ">" +
  path +
  "</marker>"
}

///|
/// Generate marker-start/marker-end attributes for arrows.
fn stroke_markers(
  shape : @ooxml.Shape,
  shape_idx : Int,
  slide_num : Int,
) -> String {
  let mut result = ""
  if has_arrow(shape.stroke_head_type) {
    result = result +
      " marker-start=\"url(#" +
      arrow_marker_id(slide_num, shape_idx, true) +
      ")\""
  }
  if has_arrow(shape.stroke_tail_type) {
    result = result +
      " marker-end=\"url(#" +
      arrow_marker_id(slide_num, shape_idx, false) +
      ")\""
  }
  result
}

// Escape text for SVG content — delegates to shared xml_escape in xml package.

///|
fn svg_escape(s : String) -> String {
  @xml.xml_escape(s)
}

// ── EMU → pixel conversion ────────────────────────────────────────────────────

///|
fn px(emu : Int, scale : Int) -> Int {
  if scale <= 0 {
    0
  } else {
    emu / scale
  }
}

///|
let min_stroke_width_str : String = "0.3" // minimum visible stroke (sub-pixel)

///|
/// Convert EMU to fractional pixel string (e.g. "0.3", "1.5", "2").
/// Used for stroke-width where sub-pixel precision matters.
fn px_frac(emu : Int, scale : Int) -> String {
  if scale <= 0 || emu <= 0 {
    return min_stroke_width_str
  }
  let whole = emu / scale
  let remainder = emu - whole * scale
  let tenths = remainder * 10 / scale
  if whole == 0 && tenths == 0 {
    min_stroke_width_str
  } else if tenths == 0 {
    int_to_str(whole)
  } else {
    int_to_str(whole) + "." + int_to_str(tenths)
  }
}

// ── Font size ─────────────────────────────────────────────────────────────────

///|
fn hundredth_pt_to_emu(sz : Int) -> Int {
  sz * @ooxml.emu_per_hundredth_pt
}

///|
fn default_font_size_emu() -> Int {
  default_font_size_emu_
}

// ── Rotation transform ────────────────────────────────────────────────────────

///|
/// Format rotation in 60000ths of a degree as "D.DD" string (e.g. "45.00").
/// Handles negative rotations correctly.
fn format_rotation_deg(rot_60k : Int) -> String {
  let deg_x100 = rot_60k / @ooxml.rot_to_centideg
  let deg_int = deg_x100 / @ooxml.centideg_per_deg
  let deg_frac_raw = deg_x100 % @ooxml.centideg_per_deg
  let deg_frac = if deg_frac_raw < 0 { -deg_frac_raw } else { deg_frac_raw }
  let frac_str = if deg_frac < 10 {
    "0" + int_to_str(deg_frac)
  } else {
    int_to_str(deg_frac)
  }
  int_to_str(deg_int) + "." + frac_str
}

///|
/// Build SVG rotate() transform attribute string centered on (cx, cy).
fn rotation_transform(rot_60k : Int, cx : Int, cy : Int) -> String {
  if rot_60k == 0 {
    return ""
  }
  " transform=\"rotate(" +
  format_rotation_deg(rot_60k) +
  "," +
  int_to_str(cx) +
  "," +
  int_to_str(cy) +
  ")\""
}

///|
fn rotation_attr(t : @ooxml.ShapeTransform, scale : Int) -> String {
  let cx_p = px(t.x + t.cx / 2, scale)
  let cy_p = px(t.y + t.cy / 2, scale)
  rotation_transform(t.rot, cx_p, cy_p)
}

// ── Image helpers ─────────────────────────────────────────────────────────────

///|
fn resolve_image_path(target : String) -> String {
  if target.length() > 3 &&
    target.get_char(0).unwrap() == '.' &&
    target.get_char(1).unwrap() == '.' &&
    target.get_char(2).unwrap() == '/' {
    "ppt/" + @xml.str_suffix(target, 3)
  } else {
    target
  }
}

///|
fn mime_from_ext(path : String) -> String {
  let len = path.length()
  let mut dot = len - 1
  while dot >= 0 {
    if path.get_char(dot).unwrap() == '.' {
      break
    }
    dot = dot - 1
  }
  if dot < 0 {
    return "image/png"
  }
  let ext = @xml.str_suffix(path, dot + 1)
  if ext == "png" {
    "image/png"
  } else if ext == "jpg" || ext == "jpeg" {
    "image/jpeg"
  } else if ext == "gif" {
    "image/gif"
  } else if ext == "svg" || ext == "svgz" {
    "image/svg+xml"
  } else if ext == "bmp" {
    "image/bmp"
  } else if ext == "tif" || ext == "tiff" {
    "image/tiff"
  } else if ext == "webp" {
    "image/webp"
  } else if ext == "emf" || ext == "wmf" {
    // EMF/WMF are Windows-only vector formats; browsers cannot render them
    ""
  } else {
    "image/png"
  }
}

///|
/// Resolve image relationship to an href string (external URL or base64 data URI).
/// Returns "" if the image cannot be resolved.
fn resolve_image_href(
  rels : Array[@ooxml.Relationship],
  rid : String,
) -> String {
  let rel = @ooxml.find_rel(rels, rid)
  let target = match rel {
    Some(r) => r.target
    None => ""
  }
  if target.is_empty() {
    return ""
  }
  match rel {
    Some(r) =>
      if r.target_mode == "External" {
        return target
      }
    None => ()
  }
  let img_path = resolve_image_path(target)
  let mime = mime_from_ext(img_path)
  if mime.is_empty() {
    // EMF/WMF: try converting to SVG via FFI
    let svg_uri = @ffi.ffi_convert_emf(img_path)
    return svg_uri
  }
  let b64 = @ffi.ffi_get_file_base64(img_path)
  if b64.is_empty() {
    ""
  } else {
    "data:" + mime + ";base64," + b64
  }
}

///|
/// Generate SVG opacity attribute from BlipFill alpha value.
/// Returns "" if fully opaque or unset.
fn blip_opacity_attr(bf : @ooxml.BlipFill) -> String {
  if bf.alpha > 0 && bf.alpha < ooxml_pct_denom {
    " opacity=\"" + format_pct_decimal(bf.alpha) + "\""
  } else {
    ""
  }
}

///|
/// Render a cropped image with optional filter attribute.
fn render_cropped_image_with_filter(
  bf : @ooxml.BlipFill,
  href : String,
  x_p : Int,
  y_p : Int,
  cx_p : Int,
  cy_p : Int,
  clip_id : String,
  opacity : String,
  rot : String,
  filter_attr : String,
) -> String {
  let vis_w = ooxml_pct_denom - bf.src_l - bf.src_r
  let vis_h = ooxml_pct_denom - bf.src_t - bf.src_b
  let img_w = if vis_w > 0 { cx_p * ooxml_pct_denom / vis_w } else { cx_p }
  let img_h = if vis_h > 0 { cy_p * ooxml_pct_denom / vis_h } else { cy_p }
  let img_x = x_p - img_w * bf.src_l / ooxml_pct_denom
  let img_y = y_p - img_h * bf.src_t / ooxml_pct_denom
  "<defs><clipPath id=\"" + clip_id + "\">" +
  "<rect" + ai("x", x_p) + ai("y", y_p) +
  ai("width", cx_p) + ai("height", cy_p) + "/>" +
  "</clipPath></defs>" +
  "<image" +
  ai("x", img_x) + ai("y", img_y) +
  ai("width", img_w) + ai("height", img_h) +
  " preserveAspectRatio=\"none\"" +
  a("href", href) +
  " clip-path=\"url(#" + clip_id + ")\"" +
  opacity + filter_attr + rot + "/>"
}

///|
/// Resolve image href with SVG preference: if svg_rid is set and resolves, use SVG.
fn resolve_blip_href(
  rels : Array[@ooxml.Relationship],
  bf : @ooxml.BlipFill,
) -> String {
  // Prefer SVG image if available
  if not(bf.svg_rid.is_empty()) {
    let svg_href = resolve_image_href(rels, bf.svg_rid)
    if not(svg_href.is_empty()) {
      return svg_href
    }
  }
  resolve_image_href(rels, bf.rid)
}

///|
/// Check if BlipFill has any image effects that need SVG filters.
fn blip_has_effects(bf : @ooxml.BlipFill) -> Bool {
  // Note: clrChange is NOT included — it can't be rendered via SVG filters
  // (requires exact pixel color matching). Data is preserved in round-trip only.
  bf.bright != 0 ||
  bf.contrast != 0 ||
  not(bf.duotone_1.is_none())
}

///|
/// Generate SVG filter definition for blip image effects.
/// Returns (filter_defs, filter_attr) — defs SVG string and filter="url(#id)" attribute.
fn render_blip_filter(
  bf : @ooxml.BlipFill,
  filter_id : String,
) -> (String, String) {
  if not(blip_has_effects(bf)) {
    return ("", "")
  }
  let mut filters = ""
  let mut result_name = "SourceGraphic"
  // Brightness / contrast (a:lum)
  if bf.bright != 0 || bf.contrast != 0 {
    // bright: -100000..100000 → offset -1..1 (additive)
    // contrast: -100000..100000 → slope 0..2 (multiplicative)
    // SVG feComponentTransfer: slope * value + intercept
    // slope = 1 + contrast/100000 (range: 0..2)
    let slope = if bf.contrast != 0 {
      format_pct_decimal(ooxml_pct_denom + bf.contrast)
    } else {
      "1"
    }
    let intercept = if bf.bright != 0 {
      format_pct_decimal(bf.bright)
    } else {
      "0"
    }
    filters = filters +
      "<feComponentTransfer result=\"lum\">" +
      "<feFuncR type=\"linear\" slope=\"" + slope + "\" intercept=\"" + intercept + "\"/>" +
      "<feFuncG type=\"linear\" slope=\"" + slope + "\" intercept=\"" + intercept + "\"/>" +
      "<feFuncB type=\"linear\" slope=\"" + slope + "\" intercept=\"" + intercept + "\"/>" +
      "</feComponentTransfer>"
    result_name = "lum"
  }
  // Duotone (a:duotone)
  if not(bf.duotone_1.is_none()) && not(bf.duotone_2.is_none()) {
    let c1 = bf.duotone_1
    let c2 = bf.duotone_2
    // Step 1: Convert to grayscale
    filters = filters +
      "<feColorMatrix type=\"saturate\" values=\"0\" in=\"" + result_name + "\" result=\"gray\"/>"
    // Step 2: Map grayscale to duotone via component transfer
    // dark pixels → color1, light pixels → color2
    // slope = (c2 - c1) / 255, intercept = c1 / 255
    let r1 = format_ratio(c1.r, 255)
    let g1 = format_ratio(c1.g, 255)
    let b1 = format_ratio(c1.b, 255)
    let r_slope = format_ratio(c2.r - c1.r, 255)
    let g_slope = format_ratio(c2.g - c1.g, 255)
    let b_slope = format_ratio(c2.b - c1.b, 255)
    filters = filters +
      "<feComponentTransfer in=\"gray\" result=\"duotone\">" +
      "<feFuncR type=\"linear\" slope=\"" + r_slope + "\" intercept=\"" + r1 + "\"/>" +
      "<feFuncG type=\"linear\" slope=\"" + g_slope + "\" intercept=\"" + g1 + "\"/>" +
      "<feFuncB type=\"linear\" slope=\"" + b_slope + "\" intercept=\"" + b1 + "\"/>" +
      "</feComponentTransfer>"
    result_name = "duotone"
  }
  // Color change (a:clrChange) — not renderable via SVG filters
  // (requires exact pixel color matching which SVG filters can't do)
  // Data is preserved in round-trip via data-ooxml-* attributes
  ignore(result_name)
  let defs = "<defs><filter id=\"" + filter_id + "\">" + filters + "</filter></defs>"
  let attr = " filter=\"url(#" + filter_id + ")\""
  (defs, attr)
}

// ── Scale ratio formatting (for group transforms) ──────────────────────────

///|
/// Used for SVG scale() transform values.
fn format_ratio(num : Int, den : Int) -> String {
  if den == 0 || num == den {
    return "1"
  }
  // Compute ratio * 10000 without 32-bit overflow.
  // Split: whole = num/den, then fractional digits via remainder*100 steps.
  let q1 = num / den
  let r1 = num - q1 * den
  let q2 = r1 * 100 / den
  let r2 = r1 * 100 - q2 * den
  let q3 = r2 * 100 / den
  let scaled = q1 * 10000 + q2 * 100 + q3
  let whole = scaled / 10000
  let frac = scaled - whole * 10000
  let abs_frac = if frac < 0 { -frac } else { frac }
  if abs_frac == 0 {
    return int_to_str(whole)
  }
  // Build fractional part, trimming trailing zeros
  let d1 = abs_frac / 1000
  let d2 = abs_frac / 100 % 10
  let d3 = abs_frac / 10 % 10
  let d4 = abs_frac % 10
  let f = if d4 != 0 {
    int_to_str(d1) + int_to_str(d2) + int_to_str(d3) + int_to_str(d4)
  } else if d3 != 0 {
    int_to_str(d1) + int_to_str(d2) + int_to_str(d3)
  } else if d2 != 0 {
    int_to_str(d1) + int_to_str(d2)
  } else {
    int_to_str(d1)
  }
  int_to_str(whole) + "." + f
}

///|
/// Format OOXML percentage (0-100000) as decimal string (e.g. 50000 → "0.5").
fn format_pct_decimal(val : Int) -> String {
  format_ratio(val, ooxml_pct_denom)
}

///|
/// Encode adjustment values as comma-separated string for data attribute.
fn encode_adj_values(avs : Array[Int]) -> String {
  let n = avs.length()
  let mut out = ""
  let mut i = 0
  while i < n {
    if i > 0 {
      out = out + ","
    }
    out = out + int_to_str(avs[i])
    i = i + 1
  }
  out
}

// ── Line / Connector rendering helpers ──────────────────────────────────────

///|
/// Compute line endpoints considering flipH/flipV.
fn line_endpoints(
  x : Int,
  y : Int,
  w : Int,
  h : Int,
  flip_h : Bool,
  flip_v : Bool,
) -> (Int, Int, Int, Int) {
  if flip_h && flip_v {
    (x + w, y + h, x, y)
  } else if flip_h {
    (x + w, y, x, y + h)
  } else if flip_v {
    (x, y + h, x + w, y)
  } else {
    (x, y, x + w, y + h)
  }
}

///|
fn render_line_geom(
  x : Int,
  y : Int,
  w : Int,
  h : Int,
  flip_h : Bool,
  flip_v : Bool,
  stroke : String,
  rot : String,
) -> String {
  let (x1, y1, x2, y2) = line_endpoints(x, y, w, h, flip_h, flip_v)
  "<line" +
  ai("x1", x1) +
  ai("y1", y1) +
  ai("x2", x2) +
  ai("y2", y2) +
  stroke +
  rot +
  "/>"
}

///|
/// Get adjustment value with default fallback (50000 = 50%).
fn adj_val(avs : Array[Int], idx : Int) -> Int {
  if idx < avs.length() {
    avs[idx]
  } else {
    50000
  }
}

///|
/// Render bent/curved connector as SVG <path>.
fn render_connector_path(
  prst : String,
  avs : Array[Int],
  x : Int,
  y : Int,
  w : Int,
  h : Int,
  flip_h : Bool,
  flip_v : Bool,
  stroke : String,
  rot : String,
) -> String {
  let (x1, y1, x2, y2) = line_endpoints(x, y, w, h, flip_h, flip_v)
  let d = if is_bent_connector(prst) {
    render_bent_path(prst, avs, x1, y1, x2, y2)
  } else if is_curved_connector(prst) {
    render_curved_path(prst, avs, x1, y1, x2, y2)
  } else {
    // Fallback: straight line
    "M" +
    int_to_str(x1) +
    " " +
    int_to_str(y1) +
    "L" +
    int_to_str(x2) +
    " " +
    int_to_str(y2)
  }
  "<path d=\"" + d + "\" fill=\"none\"" + stroke + rot + "/>"
}

///|
fn is_bent_connector(prst : String) -> Bool {
  prst == "bentConnector2" ||
  prst == "bentConnector3" ||
  prst == "bentConnector4" ||
  prst == "bentConnector5"
}

///|
fn is_curved_connector(prst : String) -> Bool {
  prst == "curvedConnector2" ||
  prst == "curvedConnector3" ||
  prst == "curvedConnector4" ||
  prst == "curvedConnector5"
}

///|
/// Render bent connector SVG path (right-angle polyline).
fn render_bent_path(
  prst : String,
  avs : Array[Int],
  x1 : Int,
  y1 : Int,
  x2 : Int,
  y2 : Int,
) -> String {
  let s = " "
  if prst == "bentConnector2" {
    // L-shape: horizontal then vertical
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else if prst == "bentConnector3" {
    // Z-shape: adj1 controls the horizontal split point
    let adj1 = adj_val(avs, 0)
    let mx = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx) +
    s +
    int_to_str(y2) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else if prst == "bentConnector4" {
    // 4-segment: adj1 (horizontal split), adj2 (vertical split)
    let adj1 = adj_val(avs, 0)
    let adj2 = adj_val(avs, 1)
    let mx = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    let my = y1 + (y2 - y1) * adj2 / ooxml_pct_denom
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx) +
    s +
    int_to_str(my) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(my) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else {
    // bentConnector5: 5-segment
    let adj1 = adj_val(avs, 0)
    let adj2 = adj_val(avs, 1)
    let adj3 = adj_val(avs, 2)
    let mx1 = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    let my = y1 + (y2 - y1) * adj2 / ooxml_pct_denom
    let mx2 = x1 + (x2 - x1) * adj3 / ooxml_pct_denom
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx1) +
    s +
    int_to_str(y1) +
    "L" +
    int_to_str(mx1) +
    s +
    int_to_str(my) +
    "L" +
    int_to_str(mx2) +
    s +
    int_to_str(my) +
    "L" +
    int_to_str(mx2) +
    s +
    int_to_str(y2) +
    "L" +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  }
}

///|
/// Render curved connector SVG path (bezier curves).
fn render_curved_path(
  prst : String,
  avs : Array[Int],
  x1 : Int,
  y1 : Int,
  x2 : Int,
  y2 : Int,
) -> String {
  let s = " "
  if prst == "curvedConnector2" {
    // Simple S-curve
    let mx = (x1 + x2) / 2
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "C" +
    int_to_str(mx) +
    s +
    int_to_str(y1) +
    s +
    int_to_str(mx) +
    s +
    int_to_str(y2) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else if prst == "curvedConnector3" {
    // S-curve with adjustable midpoint
    let adj1 = adj_val(avs, 0)
    let mx = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "C" +
    int_to_str(mx) +
    s +
    int_to_str(y1) +
    s +
    int_to_str(mx) +
    s +
    int_to_str(y2) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else if prst == "curvedConnector4" {
    let adj1 = adj_val(avs, 0)
    let adj2 = adj_val(avs, 1)
    let mx = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    let my = y1 + (y2 - y1) * adj2 / ooxml_pct_denom
    // Two cubic bezier segments
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "C" +
    int_to_str(mx) +
    s +
    int_to_str(y1) +
    s +
    int_to_str(mx) +
    s +
    int_to_str(my) +
    s +
    int_to_str(mx) +
    s +
    int_to_str(my) +
    "C" +
    int_to_str(mx) +
    s +
    int_to_str(my) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(my) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  } else {
    // curvedConnector5
    let adj1 = adj_val(avs, 0)
    let adj2 = adj_val(avs, 1)
    let adj3 = adj_val(avs, 2)
    let mx1 = x1 + (x2 - x1) * adj1 / ooxml_pct_denom
    let my = y1 + (y2 - y1) * adj2 / ooxml_pct_denom
    let mx2 = x1 + (x2 - x1) * adj3 / ooxml_pct_denom
    "M" +
    int_to_str(x1) +
    s +
    int_to_str(y1) +
    "C" +
    int_to_str(mx1) +
    s +
    int_to_str(y1) +
    s +
    int_to_str(mx1) +
    s +
    int_to_str(my) +
    s +
    int_to_str(mx1) +
    s +
    int_to_str(my) +
    "C" +
    int_to_str(mx1) +
    s +
    int_to_str(my) +
    s +
    int_to_str(mx2) +
    s +
    int_to_str(my) +
    s +
    int_to_str(mx2) +
    s +
    int_to_str(my) +
    "C" +
    int_to_str(mx2) +
    s +
    int_to_str(my) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(y2) +
    s +
    int_to_str(x2) +
    s +
    int_to_str(y2)
  }
}

// ── data-ooxml-* attribute helpers ───────────────────────────────────────────

///|
fn shape_type_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(_) => "autoshape"
    @ooxml.ShapeKind::Picture(_) => "picture"
    @ooxml.ShapeKind::TableShape(_) => "table"
    @ooxml.ShapeKind::GroupShape(_) => "group"
    @ooxml.ShapeKind::ChartShape(_) => "chart"
    @ooxml.ShapeKind::Other => "other"
  }
}

///|
fn geom_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(g) => g.to_prst()
    _ => "rect"
  }
}

// ── Bullet auto-numbering ─────────────────────────────────────────────────────

///|
fn auto_number_text(type_ : String, idx : Int) -> String {
  let n = idx + 1
  if type_ == "arabicPeriod" {
    int_to_str(n) + "."
  } else if type_ == "arabicParenR" {
    int_to_str(n) + ")"
  } else if type_ == "alphaLcPeriod" {
    alpha_lower(n) + "."
  } else if type_ == "alphaLcParenR" {
    alpha_lower(n) + ")"
  } else if type_ == "alphaUcPeriod" {
    alpha_upper(n) + "."
  } else if type_ == "alphaUcParenR" {
    alpha_upper(n) + ")"
  } else if type_ == "romanLcPeriod" {
    roman_lower(n) + "."
  } else if type_ == "romanUcPeriod" {
    roman_upper(n) + "."
  } else {
    int_to_str(n) + "."
  }
}

///|
let alpha_count : Int = 26 // number of letters in alphabet

///|
let ascii_lower_a_minus1 : Int = 96 // 'a' - 1

///|
let ascii_upper_a_minus1 : Int = 64 // 'A' - 1

///|
let ascii_lower_a : Int = 97

///|
let ascii_lower_z : Int = 122

///|
let ascii_case_shift : Int = 32 // 'a' - 'A'

///|
fn alpha_lower(n : Int) -> String {
  if n >= 1 && n <= alpha_count {
    @ffi.ffi_char_code_to_str(ascii_lower_a_minus1 + n)
  } else {
    int_to_str(n)
  }
}

///|
fn alpha_upper(n : Int) -> String {
  if n >= 1 && n <= alpha_count {
    @ffi.ffi_char_code_to_str(ascii_upper_a_minus1 + n)
  } else {
    int_to_str(n)
  }
}

///|
fn roman_lower(n : Int) -> String {
  roman_convert(n, true)
}

///|
fn roman_upper(n : Int) -> String {
  roman_convert(n, false)
}

///|
fn roman_convert(n : Int, lower : Bool) -> String {
  if n <= 0 || n > 3999 {
    return int_to_str(n)
  }
  let vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
  let syms = [
    "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I",
  ]
  let syml = [
    "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i",
  ]
  let mut result = ""
  let mut rem = n
  let n_pairs = vals.length()
  let mut i = 0
  while i < n_pairs {
    while rem >= vals[i] {
      result = result + (if lower { syml[i] } else { syms[i] })
      rem = rem - vals[i]
    }
    i = i + 1
  }
  result
}

// ── Shape rendering ───────────────────────────────────────────────────────────

///|
fn render_shape(
  shape : @ooxml.Shape,
  rels : Array[@ooxml.Relationship],
  slide_num : Int,
  scale : Int,
  shape_idx : Int,
  theme : @ooxml.ThemeData,
  table_styles : Array[@ooxml.TableStyleDef],
) -> String {
  let t = shape.transform
  let x_p = px(t.x, scale)
  let y_p = px(t.y, scale)
  let cx_p = px(t.cx, scale)
  let mut cy_p = px(t.cy, scale)
  // Apply autoFit: expand shape height if text overflows (never shrink)
  if shape.body_props.auto_fit && shape.paragraphs.length() > 0 {
    let text_h = estimate_text_height(shape.paragraphs, scale, shape.body_props)
    if text_h > cy_p {
      cy_p = text_h
    }
  }
  let is_line = match shape.kind {
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Line) => true
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Connector(_, _)) => true
    _ => false
  }
  if cx_p <= 0 {
    return ""
  }
  if cy_p <= 0 && not(is_line) {
    return ""
  }
  let rot = rotation_attr(t, scale)

  // Build data-ooxml-* attributes for the <g> wrapper
  let g_attrs = da("shape-type", shape_type_str(shape.kind)) +
    da("geom", geom_str(shape.kind)) +
    dai("x", t.x) +
    dai("y", t.y) +
    dai("cx", t.cx) +
    dai("cy", t.cy) +
    dai("rot", t.rot) +
    da("flip-h", if t.flip_h { "1" } else { "" }) +
    da("flip-v", if t.flip_v { "1" } else { "" }) +
    da("fill", shape.fill.to_hex()) +
    da("stroke", shape.stroke.to_hex()) +
    dai("stroke-w", shape.stroke_w) +
    da("stroke-dash", shape.stroke_dash) +
    da("stroke-cap", shape.stroke_cap) +
    da("stroke-join", shape.stroke_join) +
    dai("stroke-miter-lim", shape.stroke_miter_limit) +
    da("stroke-head-type", shape.stroke_head_type) +
    da("stroke-head-w", shape.stroke_head_w) +
    da("stroke-head-len", shape.stroke_head_len) +
    da("stroke-tail-type", shape.stroke_tail_type) +
    da("stroke-tail-w", shape.stroke_tail_w) +
    da("stroke-tail-len", shape.stroke_tail_len) +
    da("stroke-cmpd", shape.stroke_cmpd) +
    da("stroke-no-fill", if shape.stroke_no_fill { "1" } else { "" }) +
    dai("shape-idx", shape_idx) +
    da("ph-type", shape.ph_type) +
    dai("ph-idx", shape.ph_idx) +
    dai("st-cxn-id", shape.st_cxn_id) +
    dai("st-cxn-idx", shape.st_cxn_idx) +
    dai("end-cxn-id", shape.end_cxn_id) +
    dai("end-cxn-idx", shape.end_cxn_idx) +
    da("sh-link-rid", shape.sh_link_rid) +
    da("sh-link-hover-rid", shape.sh_link_hover_rid) +
    // Effect data attributes for round-trip
    (if not(shape.effects.is_none()) {
      let eff = shape.effects
      let mut ea = ""
      if not(eff.outer_shadow.is_none()) {
        let os = eff.outer_shadow
        ea = ea +
          dai("eff-os-blur", os.blur_rad) +
          dai("eff-os-dist", os.dist) +
          dai("eff-os-dir", os.dir) +
          da("eff-os-clr", os.color.to_hex()) +
          dai("eff-os-sx", os.sx) +
          dai("eff-os-sy", os.sy) +
          da("eff-os-algn", os.algn) +
          da("eff-os-rot", if os.rot_with_shape { "1" } else { "0" })
      }
      if not(eff.inner_shadow.is_none()) {
        let ishdw = eff.inner_shadow
        ea = ea +
          dai("eff-is-blur", ishdw.blur_rad) +
          dai("eff-is-dist", ishdw.dist) +
          dai("eff-is-dir", ishdw.dir) +
          da("eff-is-clr", ishdw.color.to_hex())
      }
      if not(eff.glow.is_none()) {
        ea = ea +
          dai("eff-gl-rad", eff.glow.rad) +
          da("eff-gl-clr", eff.glow.color.to_hex())
      }
      if not(eff.soft_edge.is_none()) {
        ea = ea + dai("eff-se-rad", eff.soft_edge.rad)
      }
      if not(eff.reflection.is_none()) {
        let r = eff.reflection
        ea = ea +
          dai("eff-refl-blur", r.blur_rad) +
          dai("eff-refl-dist", r.dist) +
          dai("eff-refl-dir", r.dir) +
          dai("eff-refl-sta", r.st_alpha) +
          dai("eff-refl-enda", r.end_alpha) +
          dai("eff-refl-fadedir", r.fade_dir) +
          dai("eff-refl-sx", r.sx) +
          dai("eff-refl-sy", r.sy) +
          da("eff-refl-algn", r.algn) +
          da("eff-refl-rot", if r.rot_with_shape { "1" } else { "0" })
      }
      ea
    } else {
      ""
    }) +
    // 3D data attributes for round-trip
    (if not(shape.scene_3d.is_none()) {
      da("3d-cam", shape.scene_3d.camera_prst) +
      da("3d-rig", shape.scene_3d.light_rig) +
      da("3d-rig-dir", shape.scene_3d.light_dir)
    } else {
      ""
    }) +
    (if not(shape.sp_3d.is_none()) {
      let s3 = shape.sp_3d
      let mut a3 = ""
      if not(s3.bevel_t.is_none()) {
        a3 = a3 + dai("3d-bt-w", s3.bevel_t.w) + dai("3d-bt-h", s3.bevel_t.h) + da("3d-bt-prst", s3.bevel_t.prst)
      }
      if not(s3.bevel_b.is_none()) {
        a3 = a3 + dai("3d-bb-w", s3.bevel_b.w) + dai("3d-bb-h", s3.bevel_b.h) + da("3d-bb-prst", s3.bevel_b.prst)
      }
      a3 + dai("3d-ext-h", s3.extrusion_h) +
        dai("3d-cnt-w", s3.contour_w) +
        da("3d-ext-clr", s3.extrusion_clr.to_hex()) +
        da("3d-cnt-clr", s3.contour_clr.to_hex()) +
        da("3d-mat", s3.prst_material) +
        dai("3d-z", s3.z)
    } else {
      ""
    }) +
    da("anchor", shape.body_props.anchor) +
    dai("l-ins", shape.body_props.l_ins) +
    dai("t-ins", shape.body_props.t_ins) +
    dai("r-ins", shape.body_props.r_ins) +
    dai("b-ins", shape.body_props.b_ins) +
    da("auto-fit", bool_str(shape.body_props.auto_fit)) +
    dai("font-scale", shape.body_props.font_scale) +
    dai("ln-spc-reduction", shape.body_props.ln_spc_reduction) +
    da("wrap", shape.body_props.wrap) +
    dai("bp-rot", shape.body_props.rot) +
    da("bp-vert", shape.body_props.vert) +
    dai("bp-num-cols", shape.body_props.num_cols) +
    dai("bp-col-spacing", shape.body_props.col_spacing) +
    da("warp-prst", shape.body_props.warp_prst) +
    dai("warp-av1", if shape.body_props.warp_av1 >= 0 { shape.body_props.warp_av1 } else { 0 }) +
    dai("warp-av2", if shape.body_props.warp_av2 >= 0 { shape.body_props.warp_av2 } else { 0 }) +
    // Gradient fill attributes
    (if not(shape.grad_fill.is_none()) {
      da("grad-stops", grad_stops_attr(shape.grad_fill)) +
      dai("grad-angle", shape.grad_fill.angle) +
      da("grad-path", shape.grad_fill.path_type) +
      da("grad-rot", if shape.grad_fill.rot_with_shape { "1" } else { "0" }) +
      dai("grad-ftl", shape.grad_fill.fill_to_l) +
      dai("grad-ftt", shape.grad_fill.fill_to_t) +
      dai("grad-ftr", shape.grad_fill.fill_to_r) +
      dai("grad-ftb", shape.grad_fill.fill_to_b) +
      da("grad-tile-flip", shape.grad_fill.tile_flip)
    } else {
      ""
    }) +
    // Blip fill attributes
    (if not(shape.blip_fill.is_none()) {
      da("blip-rid", shape.blip_fill.rid) +
      da("blip-stretch", if shape.blip_fill.stretch { "1" } else { "0" }) +
      dai("blip-src-l", shape.blip_fill.src_l) +
      dai("blip-src-t", shape.blip_fill.src_t) +
      dai("blip-src-r", shape.blip_fill.src_r) +
      dai("blip-src-b", shape.blip_fill.src_b) +
      dai("blip-tile-tx", shape.blip_fill.tile_tx) +
      dai("blip-tile-ty", shape.blip_fill.tile_ty) +
      dai("blip-tile-sx", shape.blip_fill.tile_sx) +
      dai("blip-tile-sy", shape.blip_fill.tile_sy) +
      da("blip-tile-flip", shape.blip_fill.tile_flip) +
      da("blip-tile-algn", shape.blip_fill.tile_algn) +
      dai("blip-alpha", shape.blip_fill.alpha) +
      da("blip-svg-rid", shape.blip_fill.svg_rid) +
      dai("blip-bright", shape.blip_fill.bright) +
      dai("blip-contrast", shape.blip_fill.contrast) +
      (if not(shape.blip_fill.duotone_1.is_none()) {
        da("blip-duotone1", shape.blip_fill.duotone_1.to_hex()) +
        da("blip-duotone2", shape.blip_fill.duotone_2.to_hex())
      } else { "" }) +
      (if not(shape.blip_fill.clr_from.is_none()) {
        da("blip-clr-from", shape.blip_fill.clr_from.to_hex()) +
        da("blip-clr-to", shape.blip_fill.clr_to.to_hex())
      } else { "" })
    } else {
      ""
    }) +
    // Pattern fill attributes
    (if not(shape.patt_fill.is_none()) {
      da("patt-prst", shape.patt_fill.prst) +
      da("patt-fg", shape.patt_fill.fg_color.to_hex()) +
      da("patt-bg", shape.patt_fill.bg_color.to_hex())
    } else {
      ""
    }) +
    // Group shape / connector-specific attributes
    (match shape.kind {
      @ooxml.ShapeKind::GroupShape(grp) =>
        dai("grp-ch-off-x", grp.ch_off_x) +
        dai("grp-ch-off-y", grp.ch_off_y) +
        dai("grp-ch-ext-cx", grp.ch_ext_cx) +
        dai("grp-ch-ext-cy", grp.ch_ext_cy)
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Connector(_, avs)) =>
        if avs.length() > 0 {
          da("cxn-adj", encode_adj_values(avs))
        } else {
          ""
        }
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Other(_, avs)) =>
        if avs.length() > 0 {
          da("cxn-adj", encode_adj_values(avs))
        } else {
          ""
        }
      @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Custom(data)) =>
        da("cust-gdlst", data.gdlst) +
        da("cust-paths", data.paths) +
        dai("cust-pw", data.path_w) +
        dai("cust-ph", data.path_h) +
        da("cust-rl", data.rect_l) +
        da("cust-rt", data.rect_t) +
        da("cust-rr", data.rect_r) +
        da("cust-rb", data.rect_b) +
        da("cust-cxn", data.cxn_lst)
      _ => ""
    })

  let inner = match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => {
      let blip_tile_id = blip_clip_id("bliptile", slide_num, shape_idx)
      let fill = if not(shape.blip_fill.is_none()) &&
        not(shape.blip_fill.stretch) &&
        shape.blip_fill.tile_sx > 0 {
        // BlipFill tile: filled via <pattern> defined in defs
        " fill=\"url(#" + blip_tile_id + ")\""
      } else if not(shape.blip_fill.is_none()) {
        // BlipFill stretch: image fill rendered as separate <image> below
        a("fill", "none")
      } else if not(shape.patt_fill.is_none()) {
        " fill=\"url(#" + patt_id(slide_num, shape_idx) + ")\""
      } else if not(shape.grad_fill.is_none()) {
        " fill=\"url(#" + grad_id(slide_num, shape_idx) + ")\""
      } else {
        let fill_base = if shape.fill.is_none() {
          a("fill", "none")
        } else {
          a("fill", shape.fill.to_css())
        }
        if shape.fill.alpha < 255 && not(shape.fill.is_none()) {
          fill_base +
          " fill-opacity=\"" +
          @ooxml.alpha_to_css_str(shape.fill.alpha) +
          "\""
        } else {
          fill_base
        }
      }
      let stroke = if shape.stroke.is_none() && not(shape.stroke_no_fill) &&
        shape.stroke_grad_fill.is_none() &&
        shape.stroke_patt_fill.is_none() {
        " stroke=\"none\""
      } else if shape.stroke_no_fill {
        " stroke=\"none\""
      } else {
        let stroke_emu = if shape.stroke_w > 0 {
          shape.stroke_w
        } else {
          @ooxml.default_stroke_w
        }
        let sw_str = px_frac(stroke_emu, scale)
        let sw = px(stroke_emu, scale)
        let sw_for_dash = if sw <= 0 { 1 } else { sw }
        let stroke_color = if not(shape.stroke_grad_fill.is_none()) {
          "url(#" + stroke_grad_id(slide_num, shape_idx) + ")"
        } else if not(shape.stroke_patt_fill.is_none()) {
          "url(#" + stroke_patt_id(slide_num, shape_idx) + ")"
        } else {
          shape.stroke.to_css()
        }
        let base = a("stroke", stroke_color) + a("stroke-width", sw_str)
        let dash_attr = stroke_dasharray(shape.stroke_dash, sw_for_dash)
        let cap_attr = stroke_linecap(shape.stroke_cap)
        let join_attr = stroke_linejoin(
          shape.stroke_join,
          shape.stroke_miter_limit,
        )
        let marker_attr = stroke_markers(shape, shape_idx, slide_num)
        base + dash_attr + cap_attr + join_attr + marker_attr
      }
      let rect_svg = match geom {
        @ooxml.ShapeGeom::Ellipse =>
          "<ellipse" +
          ai("cx", x_p + cx_p / 2) +
          ai("cy", y_p + cy_p / 2) +
          ai("rx", cx_p / 2) +
          ai("ry", cy_p / 2) +
          fill +
          stroke +
          rot +
          "/>"
        @ooxml.ShapeGeom::Line =>
          render_line_geom(
            x_p,
            y_p,
            cx_p,
            cy_p,
            t.flip_h,
            t.flip_v,
            stroke,
            rot,
          )
        @ooxml.ShapeGeom::Connector(prst, avs) =>
          render_connector_path(
            prst,
            avs,
            x_p,
            y_p,
            cx_p,
            cy_p,
            t.flip_h,
            t.flip_v,
            stroke,
            rot,
          )
        @ooxml.ShapeGeom::Other(prst, avs) =>
          render_preset_geom(prst, avs, cx_p, cy_p, x_p, y_p, fill, stroke, rot)
        @ooxml.ShapeGeom::Custom(data) =>
          render_custom_geom(data, cx_p, cy_p, x_p, y_p, fill, stroke, rot)
        _ => {
          // Rect / RoundRect
          let rx = match geom {
            @ooxml.ShapeGeom::RoundRect =>
              int_to_str(cx_p / round_rect_radius_divisor)
            _ => "0"
          }
          let base = "<rect" +
            ai("x", x_p) +
            ai("y", y_p) +
            ai("width", cx_p) +
            ai("height", cy_p) +
            fill +
            stroke +
            rot
          if rx == "0" {
            base + "/>"
          } else {
            base + a("rx", rx) + a("ry", rx) + "/>"
          }
        }
      }
      // BlipFill: render image inside shape
      let blip_svg = if not(shape.blip_fill.is_none()) {
        let bf = shape.blip_fill
        let href = resolve_blip_href(rels, bf)
        if href.is_empty() {
          ""
        } else {
          let opacity = blip_opacity_attr(bf)
          let filter_id = blip_filter_id(slide_num, shape_idx)
          let (filter_defs, filter_attr) = render_blip_filter(bf, filter_id)
          if not(bf.stretch) && bf.tile_sx > 0 {
            // Tile mode: create <pattern> definition
            let tile_w = cx_p * bf.tile_sx / ooxml_pct_denom
            let tile_h = cy_p * bf.tile_sy / ooxml_pct_denom
            let tw = int_to_str(if tile_w > 0 { tile_w } else { cx_p })
            let th = int_to_str(if tile_h > 0 { tile_h } else { cy_p })
            filter_defs +
            "<defs><pattern id=\"" +
            blip_tile_id +
            "\" patternUnits=\"userSpaceOnUse\"" +
            ai("x", x_p) +
            ai("y", y_p) +
            " width=\"" +
            tw +
            "\" height=\"" +
            th +
            "\">" +
            "<image width=\"" +
            tw +
            "\" height=\"" +
            th +
            "\" preserveAspectRatio=\"none\"" +
            a("href", href) +
            opacity + filter_attr +
            "/>" +
            "</pattern></defs>"
          } else if bf.has_crop() {
            let clip_id = blip_clip_id("blipclip", slide_num, shape_idx)
            filter_defs + render_cropped_image_with_filter(bf, href, x_p, y_p, cx_p, cy_p, clip_id, opacity, rot, filter_attr)
          } else {
            filter_defs +
            "<image" +
            ai("x", x_p) +
            ai("y", y_p) +
            ai("width", cx_p) +
            ai("height", cy_p) +
            " preserveAspectRatio=\"none\"" +
            a("href", href) +
            opacity + filter_attr + rot + "/>"
          }
        }
      } else {
        ""
      }
      // Compute text rectangle for non-rectangular shapes
      let (tr_l, tr_t, tr_w, tr_h) = match geom {
        @ooxml.ShapeGeom::Other(prst, avs) =>
          compute_preset_text_rect(prst, avs, cx_p, cy_p)
        @ooxml.ShapeGeom::Custom(data) =>
          compute_custom_text_rect(data, cx_p, cy_p)
        _ => (0, 0, cx_p, cy_p)
      }
      // For autoFit, use adjusted cy for text rendering if shape was expanded
      let eff_cy = if shape.body_props.auto_fit && cy_p > px(t.cy, scale) {
        cy_p
      } else {
        tr_h
      }
      let text_t : @ooxml.ShapeTransform = {
        x: t.x + tr_l * scale,
        y: t.y + tr_t * scale,
        cx: tr_w * scale,
        cy: eff_cy * scale,
        rot: t.rot,
        flip_h: t.flip_h,
        flip_v: t.flip_v,
      }
      rect_svg +
      blip_svg +
      render_text(
        shape.paragraphs,
        text_t,
        scale,
        theme,
        shape.body_props,
        rels,
      )
    }
    @ooxml.ShapeKind::Picture(_) => {
      let bf = shape.blip_fill
      let href = resolve_blip_href(rels, bf)
      let placeholder = "<rect" +
        ai("x", x_p) +
        ai("y", y_p) +
        ai("width", cx_p) +
        ai("height", cy_p) +
        " fill=\"" +
        image_placeholder_fill +
        "\"" +
        rot +
        "/>"
      if href.is_empty() {
        placeholder
      } else {
        let opacity = blip_opacity_attr(bf)
        let filter_id = blip_filter_id(slide_num, shape_idx)
        let (filter_defs, filter_attr) = render_blip_filter(bf, filter_id)
        let img_svg = if bf.has_crop() {
          let clip_id = blip_clip_id("picclip", slide_num, shape_idx)
          filter_defs + render_cropped_image_with_filter(bf, href, x_p, y_p, cx_p, cy_p, clip_id, opacity, rot, filter_attr)
        } else {
          filter_defs +
          "<image" +
          ai("x", x_p) +
          ai("y", y_p) +
          ai("width", cx_p) +
          ai("height", cy_p) +
          " preserveAspectRatio=\"xMidYMid meet\"" +
          a("href", href) +
          opacity + filter_attr + rot + "/>"
        }
        // Render border if picture has a stroke (a:ln)
        let border = if not(shape.stroke.is_none()) && not(shape.stroke_no_fill) {
          let stroke_emu = if shape.stroke_w > 0 {
            shape.stroke_w
          } else {
            @ooxml.default_stroke_w
          }
          let sw_str = px_frac(stroke_emu, scale)
          "<rect" +
          ai("x", x_p) +
          ai("y", y_p) +
          ai("width", cx_p) +
          ai("height", cy_p) +
          " fill=\"none\"" +
          a("stroke", shape.stroke.to_css()) +
          a("stroke-width", sw_str) +
          rot +
          "/>"
        } else {
          ""
        }
        img_svg + border
      }
    }
    @ooxml.ShapeKind::TableShape(table) =>
      render_table(table, t, scale, theme, rels, table_styles)
    @ooxml.ShapeKind::ChartShape(chart) =>
      render_chart(chart, t, scale, theme)
    @ooxml.ShapeKind::GroupShape(grp) =>
      render_group(grp, t, rels, slide_num, scale, shape_idx, theme, table_styles)
    _ => ""
  }

  if inner.is_empty() {
    ""
  } else {
    let eff_attr = if not(shape.effects.is_none()) {
      let (_, fa) = render_effect_filter(
        shape.effects,
        effect_filter_id(slide_num, shape_idx),
        scale,
      )
      fa
    } else {
      ""
    }
    let has_refl = not(shape.effects.is_none()) &&
      not(shape.effects.reflection.is_none())
    let shape_id_attr = if has_refl {
      " id=\"" + refl_shape_id(slide_num, shape_idx) + "\""
    } else {
      ""
    }
    let shape_svg = "<g" + g_attrs + eff_attr + shape_id_attr + ">" + inner + "</g>"
    if has_refl {
      let t = shape.transform
      shape_svg +
        render_reflection_element(
          shape.effects.reflection,
          refl_shape_id(slide_num, shape_idx),
          refl_mask_id(slide_num, shape_idx),
          px(t.x, scale),
          px(t.y, scale),
          px(t.cx, scale),
          cy_p,
          scale,
        )
    } else {
      shape_svg
    }
  }
}

// ── Group rendering ─────────────────────────────────────────────────────────

///|
fn render_group(
  grp : @ooxml.GroupShapeData,
  t : @ooxml.ShapeTransform,
  rels : Array[@ooxml.Relationship],
  slide_num : Int,
  scale : Int,
  shape_idx : Int,
  theme : @ooxml.ThemeData,
  table_styles : Array[@ooxml.TableStyleDef],
) -> String {
  let nc = grp.children.length()
  if nc == 0 {
    return ""
  }
  // Build SVG transform: translate(off) scale(ext/chExt) translate(-chOff)
  let off_x = px(t.x, scale)
  let off_y = px(t.y, scale)
  let ch_off_x = px(grp.ch_off_x, scale)
  let ch_off_y = px(grp.ch_off_y, scale)
  let sx = format_ratio(t.cx, grp.ch_ext_cx)
  let sy = format_ratio(t.cy, grp.ch_ext_cy)
  let mut transform_str = "translate(" +
    int_to_str(off_x) +
    "," +
    int_to_str(off_y) +
    ")"
  if sx != "1" || sy != "1" {
    transform_str = transform_str + " scale(" + sx + "," + sy + ")"
  }
  if ch_off_x != 0 || ch_off_y != 0 {
    transform_str = transform_str +
      " translate(" +
      int_to_str(-ch_off_x) +
      "," +
      int_to_str(-ch_off_y) +
      ")"
  }
  // Group rotation
  if t.rot != 0 {
    let cx_p = px(t.cx / 2, scale)
    let cy_p = px(t.cy / 2, scale)
    // Rotation is applied around the group center (before the translate)
    // We prepend the rotation to the transform
    transform_str = "translate(" +
      int_to_str(off_x + cx_p) +
      "," +
      int_to_str(off_y + cy_p) +
      ") rotate(" +
      format_rotation_deg(t.rot) +
      ") translate(" +
      int_to_str(-(off_x + cx_p)) +
      "," +
      int_to_str(-(off_y + cy_p)) +
      ") " +
      transform_str
  }
  // Render child defs (gradient, pattern, arrow markers) and shapes in one pass
  let mut defs = ""
  let mut children_svg = ""
  let mut ci = 0
  while ci < nc {
    let child = grp.children[ci]
    let child_idx = shape_idx * group_child_idx_multiplier + ci
    if not(child.grad_fill.is_none()) {
      defs = defs +
        render_gradient_def(child.grad_fill, grad_id(slide_num, child_idx))
    }
    if not(child.patt_fill.is_none()) {
      defs = defs +
        render_pattern_def(child.patt_fill, patt_id(slide_num, child_idx))
    }
    if has_arrow(child.stroke_head_type) {
      defs = defs +
        render_arrow_marker(
          arrow_marker_id(slide_num, child_idx, true),
          child.stroke_head_type,
          child.stroke_head_w,
          child.stroke_head_len,
          child.stroke,
        )
    }
    if has_arrow(child.stroke_tail_type) {
      defs = defs +
        render_arrow_marker(
          arrow_marker_id(slide_num, child_idx, false),
          child.stroke_tail_type,
          child.stroke_tail_w,
          child.stroke_tail_len,
          child.stroke,
        )
    }
    children_svg = children_svg +
      render_shape(child, rels, slide_num, scale, child_idx, theme, table_styles)
    ci = ci + 1
  }
  let defs_section = if defs.is_empty() {
    ""
  } else {
    "<defs>" + defs + "</defs>"
  }
  "<g transform=\"" +
  transform_str +
  "\">" +
  defs_section +
  children_svg +
  "</g>"
}

// ── Public API ────────────────────────────────────────────────────────────────

///|
/// Render a SlideData to an SVG string with data-ooxml-* attributes.
pub fn render_slide_svg(
  slide : @ooxml.SlideData,
  rels : Array[@ooxml.Relationship],
  slide_num : Int,
  theme : @ooxml.ThemeData,
  bg_shapes : Array[@ooxml.Shape],
  table_styles : Array[@ooxml.TableStyleDef],
  bg_shapes_rels : Array[Array[@ooxml.Relationship]],
  bg_fill_rels : Array[@ooxml.Relationship],
) -> String {
  let sw = slide.slide_size.cx
  let sh = slide.slide_size.cy
  let w = svg_target_width
  let scale = if sw > 0 { sw / w } else { default_emu_per_px }
  let h = if scale > 0 { sh / scale } else { svg_target_width * 9 / 16 }
  let header = "<svg" +
    " xmlns=\"http://www.w3.org/2000/svg\"" +
    " xmlns:xlink=\"http://www.w3.org/1999/xlink\"" +
    ai("width", w) +
    ai("height", h) +
    " style=\"background:" +
    slide.background.to_css() +
    ";display:block\"" +
    // data-ooxml-* metadata on SVG root for round-trip
    dai("slide-cx", sw) +
    dai("slide-cy", sh) +
    da("bg", slide.background.to_hex()) +
    dai("scale", scale) +
    ">"
  // Collect all <defs> in a single pass per shape set
  let mut defs = ""
  // Background gradient
  if not(slide.bg_grad.is_none()) {
    defs = defs +
      render_gradient_def(slide.bg_grad, "grad-bg-s" + int_to_str(slide_num))
  }
  // Background pattern
  if not(slide.bg_patt_fill.is_none()) {
    defs = defs +
      render_pattern_def(
        slide.bg_patt_fill,
        "patt-bg-s" + int_to_str(slide_num),
      )
  }
  // Background blip fill filter
  if not(slide.bg_blip_fill.is_none()) && blip_has_effects(slide.bg_blip_fill) {
    let (fd, _) = render_blip_filter(
      slide.bg_blip_fill,
      "bfilt-bg-s" + int_to_str(slide_num),
    )
    defs = defs + fd
  }
  // Background shapes defs (master/layout non-placeholder shapes)
  let nbg = bg_shapes.length()
  let mut bdi = 0
  while bdi < nbg {
    let bs = bg_shapes[bdi]
    let bs_idx = -(bdi + 1)
    if not(bs.grad_fill.is_none()) {
      defs = defs +
        render_gradient_def(bs.grad_fill, grad_id(slide_num, bs_idx))
    }
    bdi = bdi + 1
  }
  // Slide shapes defs (gradient, pattern, arrows)
  let n = slide.shapes.length()
  let mut di = 0
  while di < n {
    let sh = slide.shapes[di]
    if not(sh.grad_fill.is_none()) {
      defs = defs + render_gradient_def(sh.grad_fill, grad_id(slide_num, di))
    }
    if not(sh.patt_fill.is_none()) {
      defs = defs + render_pattern_def(sh.patt_fill, patt_id(slide_num, di))
    }
    // Stroke gradient/pattern defs
    if not(sh.stroke_grad_fill.is_none()) {
      defs = defs +
        render_gradient_def(sh.stroke_grad_fill, stroke_grad_id(slide_num, di))
    }
    if not(sh.stroke_patt_fill.is_none()) {
      defs = defs +
        render_pattern_def(sh.stroke_patt_fill, stroke_patt_id(slide_num, di))
    }
    if has_arrow(sh.stroke_head_type) {
      defs = defs +
        render_arrow_marker(
          arrow_marker_id(slide_num, di, true),
          sh.stroke_head_type,
          sh.stroke_head_w,
          sh.stroke_head_len,
          sh.stroke,
        )
    }
    if has_arrow(sh.stroke_tail_type) {
      defs = defs +
        render_arrow_marker(
          arrow_marker_id(slide_num, di, false),
          sh.stroke_tail_type,
          sh.stroke_tail_w,
          sh.stroke_tail_len,
          sh.stroke,
        )
    }
    // Effect filter defs
    if not(sh.effects.is_none()) {
      let (eff_def, _) = render_effect_filter(
        sh.effects,
        effect_filter_id(slide_num, di),
        scale,
      )
      defs = defs + eff_def
      // Reflection mask gradient
      if not(sh.effects.reflection.is_none()) {
        defs = defs +
          render_reflection_mask_def(
            sh.effects.reflection,
            refl_mask_id(slide_num, di),
          )
      }
    }
    di = di + 1
  }
  let defs_section = if defs.is_empty() {
    ""
  } else {
    "<defs>" + defs + "</defs>"
  }

  let bg_fill = if not(slide.bg_grad.is_none()) {
    "url(#grad-bg-s" + int_to_str(slide_num) + ")"
  } else if not(slide.bg_patt_fill.is_none()) {
    "url(#patt-bg-s" + int_to_str(slide_num) + ")"
  } else {
    slide.background.to_css()
  }
  let bg_rect = "<rect" +
    ai("width", w) +
    ai("height", h) +
    a("fill", bg_fill) +
    "/>"
  // Background blip fill (image)
  let bg_blip_svg = if not(slide.bg_blip_fill.is_none()) {
    let bf = slide.bg_blip_fill
    let bg_resolve_rels = if bg_fill_rels.length() > 0 {
      bg_fill_rels
    } else {
      rels
    }
    let href = resolve_blip_href(bg_resolve_rels, bf)
    if href.is_empty() {
      ""
    } else {
      let opacity = blip_opacity_attr(bf)
      let filter_attr = if blip_has_effects(bf) {
        " filter=\"url(#bfilt-bg-s" + int_to_str(slide_num) + ")\""
      } else {
        ""
      }
      if bf.has_crop() {
        let clip_id = "bgclip-s" + int_to_str(slide_num)
        render_cropped_image_with_filter(
          bf, href, 0, 0, w, h, clip_id, opacity, "", filter_attr,
        )
      } else {
        "<image" +
        " x=\"0\" y=\"0\"" +
        ai("width", w) +
        ai("height", h) +
        " preserveAspectRatio=\"none\"" +
        a("href", href) +
        opacity +
        filter_attr +
        "/>"
      }
    }
  } else {
    ""
  }
  // Render master/layout background shapes (non-placeholder)
  let mut bg_shapes_svg = ""
  let mut bi = 0
  while bi < nbg {
    let bg_shape_rels = if bi < bg_shapes_rels.length() {
      bg_shapes_rels[bi]
    } else {
      rels
    }
    bg_shapes_svg = bg_shapes_svg +
      render_shape(bg_shapes[bi], bg_shape_rels, slide_num, scale, -(bi + 1), theme, table_styles)
    bi = bi + 1
  }
  // Render slide shapes
  let mut shapes_svg = ""
  let mut i = 0
  while i < n {
    shapes_svg = shapes_svg +
      render_shape(slide.shapes[i], rels, slide_num, scale, i, theme, table_styles)
    i = i + 1
  }
  header + defs_section + bg_rect + bg_blip_svg + bg_shapes_svg + shapes_svg + "</svg>"
}
