// OMML (Office Math Markup Language) → SVG renderer.
//
// Converts m:oMath / m:oMathPara XML into positioned SVG elements
// using a recursive layout engine. Each OMML element is measured and
// rendered as a MathBox (svg fragment, width, height, baseline).
//
// Supported elements:
//   m:r       — math run (text leaf)
//   m:f       — fraction (numerator / denominator)
//   m:sSup    — superscript
//   m:sSub    — subscript
//   m:sSubSup — sub + superscript
//   m:rad     — radical (square root)
//   m:nary    — large operator (∫ Σ ∏)
//   m:d       — delimiters (brackets)
//   m:acc     — accent (hat, tilde, bar)
//   m:m       — matrix
//   m:bar     — over/under bar
//
// Round-trip: the raw OMML XML is preserved via data-ooxml-math-xml
// attributes on a hidden <text> element alongside the visual rendering.

// ── Layout model ────────────────────────────────────────────────────────────

///|
priv struct MathBox {
  svg : String
  w : Double
  h : Double
  baseline : Double // distance from top to baseline
}

// ── Constants ───────────────────────────────────────────────────────────────

///|
let math_empty_box : MathBox = { svg: "", w: 0.0, h: 0.0, baseline: 0.0 }

///|
let math_font_family : String = "Cambria Math, Cambria, serif"

///|
let math_frac_scale : Double = 0.85

///|
let math_script_scale : Double = 0.7

///|
let math_nary_scale : Double = 1.8

///|
let math_nary_limit_scale : Double = 0.65

///|
let math_frac_bar_thick : Double = 1.0

///|
let math_frac_gap : Double = 2.0

///|
let math_sup_shift_ratio : Double = 0.4

///|
let math_sub_shift_ratio : Double = 0.25

///|
let math_rad_gap : Double = 2.0

///|
let math_matrix_gap_ratio : Double = 0.4

///|
let math_delim_pad : Double = 2.0

///|
let math_text_height_ratio : Double = 1.2 // text height = font_size × this

///|
let math_text_baseline_ratio : Double = 0.85 // baseline = font_size × this

// ── Helpers ─────────────────────────────────────────────────────────────────

///|
fn math_g(x : Double, y : Double, content : String) -> String {
  "<g transform=\"translate(" +
  dts(x) +
  "," +
  dts(y) +
  ")\">" +
  content +
  "</g>"
}

///|
fn math_text_w(text : String, ff : String, fs : Double) -> Double {
  if text.is_empty() {
    return 0.0
  }
  @ffi.ffi_measure_text(text, ff, fs)
}

///|
fn math_text_svg(
  text : String,
  x : Double,
  y : Double,
  fs : Double,
  ff : String,
  color : String,
) -> String {
  "<text" +
  " x=\"" +
  dts(x) +
  "\" y=\"" +
  dts(y) +
  "\" font-size=\"" +
  dts(fs) +
  "\" font-family=\"" +
  ff +
  ", sans-serif\" fill=\"" +
  color +
  "\">" +
  svg_escape(text) +
  "</text>"
}

///|
/// Find first child element with given tag.
fn math_child(node : @xml.XmlNode, tag : String) -> @xml.XmlNode? {
  let children = @xml.get_children(node)
  @xml.find_child(children, tag)
}

///|
/// Find all child elements with given tag.
fn math_children(node : @xml.XmlNode, tag : String) -> Array[@xml.XmlNode] {
  let children = @xml.get_children(node)
  @xml.find_children(children, tag)
}

///|
/// Get a property value from the element's properties child.
/// e.g. math_prop(node, "m:fPr", "m:type") → value of m:val on m:type inside m:fPr.
fn math_prop(node : @xml.XmlNode, pr_tag : String, prop_tag : String) -> String {
  match math_child(node, pr_tag) {
    Some(pr) =>
      match math_child(pr, prop_tag) {
        Some(prop) =>
          match @xml.get_attr(prop, "m:val") {
            Some(v) => v
            None => ""
          }
        None => ""
      }
    None => ""
  }
}

///|
/// Layout all children of a container element (m:e, m:num, m:den, etc.)
/// as a horizontal row.
fn layout_math_children(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let children = @xml.get_children(node)
  let boxes : Array[MathBox] = []
  let n = children.length()
  let mut i = 0
  while i < n {
    let child = children[i]
    let tag = @xml.get_tag(child)
    if !tag.is_empty() && tag != "m:rPr" && tag != "m:ctrlPr" {
      boxes.push(layout_math_node(child, fs, ff, color))
    }
    i = i + 1
  }
  math_hbox(boxes)
}

///|
/// Horizontally concatenate boxes, baseline-aligned.
fn math_hbox(boxes : Array[MathBox]) -> MathBox {
  let n = boxes.length()
  if n == 0 {
    return math_empty_box
  }
  if n == 1 {
    return boxes[0]
  }
  // Find max baseline (ascent) and max descent
  let mut max_bl : Double = 0.0
  let mut max_desc : Double = 0.0
  let mut i = 0
  while i < n {
    let b = boxes[i]
    if b.baseline > max_bl {
      max_bl = b.baseline
    }
    let desc = b.h - b.baseline
    if desc > max_desc {
      max_desc = desc
    }
    i = i + 1
  }
  let total_h = max_bl + max_desc
  let mut svg = ""
  let mut x : Double = 0.0
  i = 0
  while i < n {
    let b = boxes[i]
    let dy = max_bl - b.baseline
    if b.svg.length() > 0 {
      svg = svg + math_g(x, dy, b.svg)
    }
    x = x + b.w
    i = i + 1
  }
  { svg, w: x, h: total_h, baseline: max_bl }
}

// ── Recursive layout engine ─────────────────────────────────────────────────

///|
fn layout_math_node(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let tag = @xml.get_tag(node)
  if tag == "m:r" {
    return layout_math_run(node, fs, ff, color)
  }
  if tag == "m:f" {
    return layout_math_frac(node, fs, ff, color)
  }
  if tag == "m:sSup" {
    return layout_math_sup(node, fs, ff, color)
  }
  if tag == "m:sSub" {
    return layout_math_sub(node, fs, ff, color)
  }
  if tag == "m:sSubSup" {
    return layout_math_subsup(node, fs, ff, color)
  }
  if tag == "m:rad" {
    return layout_math_rad(node, fs, ff, color)
  }
  if tag == "m:nary" {
    return layout_math_nary(node, fs, ff, color)
  }
  if tag == "m:d" {
    return layout_math_delim(node, fs, ff, color)
  }
  if tag == "m:acc" {
    return layout_math_acc(node, fs, ff, color)
  }
  if tag == "m:m" {
    return layout_math_matrix(node, fs, ff, color)
  }
  if tag == "m:bar" {
    return layout_math_bar(node, fs, ff, color)
  }
  if tag == "m:oMath" ||
    tag == "m:oMathPara" ||
    tag == "m:e" ||
    tag == "m:num" ||
    tag == "m:den" ||
    tag == "m:sup" ||
    tag == "m:sub" ||
    tag == "m:deg" {
    return layout_math_children(node, fs, ff, color)
  }
  // Unknown element: try to lay out children
  layout_math_children(node, fs, ff, color)
}

// ── m:r — Math run (text leaf) ──────────────────────────────────────────────

///|
fn layout_math_run(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  // Extract text from m:t child
  let mut text = ""
  let children = @xml.get_children(node)
  let n = children.length()
  let mut i = 0
  while i < n {
    let child = children[i]
    if @xml.get_tag(child) == "m:t" {
      text = text + @xml.get_text_content(child)
    }
    i = i + 1
  }
  if text.is_empty() {
    return math_empty_box
  }
  let w = math_text_w(text, ff, fs)
  let h = fs * math_text_height_ratio
  let bl = fs * math_text_baseline_ratio
  let svg = math_text_svg(text, 0.0, bl, fs, ff, color)
  { svg, w, h, baseline: bl }
}

// ── m:f — Fraction ──────────────────────────────────────────────────────────

///|
fn layout_math_frac(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let sub_fs = fs * math_frac_scale
  let num_box = match math_child(node, "m:num") {
    Some(n) => layout_math_children(n, sub_fs, ff, color)
    None => math_empty_box
  }
  let den_box = match math_child(node, "m:den") {
    Some(d) => layout_math_children(d, sub_fs, ff, color)
    None => math_empty_box
  }
  let frac_pad : Double = 4.0 // horizontal padding around fraction bar
  let bar_w = if num_box.w > den_box.w { num_box.w } else { den_box.w }
  let total_w = bar_w + frac_pad
  // Layout: numerator centered above fraction bar, denominator below
  let num_y : Double = 0.0
  let bar_y = num_box.h + math_frac_gap
  let den_y = bar_y + math_frac_bar_thick + math_frac_gap
  let total_h = den_y + den_box.h
  // Baseline: at fraction bar (middle of surrounding text)
  let baseline = bar_y + math_frac_bar_thick / 2.0
  // Center numerator and denominator
  let num_x = (total_w - num_box.w) / 2.0
  let den_x = (total_w - den_box.w) / 2.0
  let svg = math_g(num_x, num_y, num_box.svg) +
    "<line x1=\"0\" y1=\"" +
    dts(bar_y) +
    "\" x2=\"" +
    dts(total_w) +
    "\" y2=\"" +
    dts(bar_y) +
    "\" stroke=\"" +
    color +
    "\" stroke-width=\"" +
    dts(math_frac_bar_thick) +
    "\"/>" +
    math_g(den_x, den_y, den_box.svg)
  { svg, w: total_w, h: total_h, baseline }
}

// ── m:sSup — Superscript ────────────────────────────────────────────────────

///|
fn layout_math_sup(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let base_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  let sup_fs = fs * math_script_scale
  let sup_box = match math_child(node, "m:sup") {
    Some(s) => layout_math_children(s, sup_fs, ff, color)
    None => math_empty_box
  }
  // Superscript shifted up relative to base baseline
  let sup_shift = base_box.baseline * math_sup_shift_ratio
  let sup_y = base_box.baseline - sup_shift - sup_box.h + sup_box.baseline * 0.5
  let sup_y_clamped = if sup_y < 0.0 { 0.0 } else { sup_y }
  let top_extra = if sup_y < 0.0 { -sup_y } else { 0.0 }
  let total_w = base_box.w + sup_box.w
  let total_h = if base_box.h + top_extra >
    sup_y_clamped + sup_box.h + top_extra {
    base_box.h + top_extra
  } else {
    sup_y_clamped + sup_box.h + top_extra
  }
  let svg = math_g(0.0, top_extra, base_box.svg) +
    math_g(base_box.w, sup_y_clamped, sup_box.svg)
  { svg, w: total_w, h: total_h, baseline: base_box.baseline + top_extra }
}

// ── m:sSub — Subscript ──────────────────────────────────────────────────────

///|
fn layout_math_sub(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let base_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  let sub_fs = fs * math_script_scale
  let sub_box = match math_child(node, "m:sub") {
    Some(s) => layout_math_children(s, sub_fs, ff, color)
    None => math_empty_box
  }
  let sub_y = base_box.baseline +
    base_box.baseline * math_sub_shift_ratio -
    sub_box.baseline * 0.5
  let total_w = base_box.w + sub_box.w
  let bottom = if sub_y + sub_box.h > base_box.h {
    sub_y + sub_box.h
  } else {
    base_box.h
  }
  let svg = base_box.svg + math_g(base_box.w, sub_y, sub_box.svg)
  { svg, w: total_w, h: bottom, baseline: base_box.baseline }
}

// ── m:sSubSup — Subscript + Superscript ─────────────────────────────────────

///|
fn layout_math_subsup(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let base_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  let script_fs = fs * math_script_scale
  let sup_box = match math_child(node, "m:sup") {
    Some(s) => layout_math_children(s, script_fs, ff, color)
    None => math_empty_box
  }
  let sub_box = match math_child(node, "m:sub") {
    Some(s) => layout_math_children(s, script_fs, ff, color)
    None => math_empty_box
  }
  // Superscript above baseline, subscript below
  let sup_shift = base_box.baseline * math_sup_shift_ratio
  let sup_y = base_box.baseline - sup_shift - sup_box.h + sup_box.baseline * 0.5
  let sup_y_clamped = if sup_y < 0.0 { 0.0 } else { sup_y }
  let top_extra = if sup_y < 0.0 { -sup_y } else { 0.0 }
  let sub_y = base_box.baseline +
    base_box.baseline * math_sub_shift_ratio -
    sub_box.baseline * 0.5 +
    top_extra
  let script_w = if sup_box.w > sub_box.w { sup_box.w } else { sub_box.w }
  let total_w = base_box.w + script_w
  let bottom_candidates = [
    base_box.h + top_extra,
    sub_y + sub_box.h,
    sup_y_clamped + sup_box.h,
  ]
  let mut total_h : Double = 0.0
  let mut ci = 0
  while ci < bottom_candidates.length() {
    if bottom_candidates[ci] > total_h {
      total_h = bottom_candidates[ci]
    }
    ci = ci + 1
  }
  let svg = math_g(0.0, top_extra, base_box.svg) +
    math_g(base_box.w, sup_y_clamped, sup_box.svg) +
    math_g(base_box.w, sub_y, sub_box.svg)
  { svg, w: total_w, h: total_h, baseline: base_box.baseline + top_extra }
}

// ── m:rad — Radical (square root) ───────────────────────────────────────────

///|
fn layout_math_rad(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let content_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  // Check if degree should be hidden
  let deg_hide = math_prop(node, "m:radPr", "m:degHide")
  let deg_box = if deg_hide != "1" {
    match math_child(node, "m:deg") {
      Some(d) => {
        let db = layout_math_children(d, fs * 0.6, ff, color)
        if db.w > 0.0 {
          db
        } else {
          math_empty_box
        }
      }
      None => math_empty_box
    }
  } else {
    math_empty_box
  }
  let gap = math_rad_gap
  let vinculum_thick : Double = 1.0
  // Radical symbol dimensions
  let rad_tail_w = fs * 0.15 // short tail
  let rad_body_w = fs * 0.35 // V-shape width
  let rad_w = rad_tail_w + rad_body_w
  let deg_offset_x = if deg_box.w > rad_tail_w {
    deg_box.w - rad_tail_w
  } else {
    0.0
  }
  let total_w = deg_offset_x + rad_w + content_box.w + 2.0
  let vinculum_y = gap
  let content_y = vinculum_y + vinculum_thick
  let total_h = content_y + content_box.h
  // Radical sign SVG path: tail → V bottom → top → vinculum start
  let rx = deg_offset_x // radical start x
  let tail_y = total_h * 0.55 // tail starts at middle-ish
  let v_bottom_y = total_h - 1.0 // V bottom near bottom
  let v_top_y = vinculum_y // V top at vinculum
  let rad_path = "M " +
    dts(rx) +
    " " +
    dts(tail_y) +
    " L " +
    dts(rx + rad_tail_w) +
    " " +
    dts(tail_y + fs * 0.1) +
    " L " +
    dts(rx + rad_tail_w + rad_body_w * 0.4) +
    " " +
    dts(v_bottom_y) +
    " L " +
    dts(rx + rad_w) +
    " " +
    dts(v_top_y) +
    " L " +
    dts(total_w) +
    " " +
    dts(v_top_y)
  let svg_path = "<path d=\"" +
    rad_path +
    "\" fill=\"none\" stroke=\"" +
    color +
    "\" stroke-width=\"" +
    dts(vinculum_thick) +
    "\"/>"
  // Degree (if present) positioned at top-left
  let deg_svg = if deg_box.w > 0.0 { math_g(0.0, 0.0, deg_box.svg) } else { "" }
  let content_x = deg_offset_x + rad_w + 1.0
  let svg = deg_svg + svg_path + math_g(content_x, content_y, content_box.svg)
  { svg, w: total_w, h: total_h, baseline: content_y + content_box.baseline }
}

// ── m:nary — Large operator (integral, sum, product) ────────────────────────

///|
fn layout_math_nary(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  // Get operator character (default: integral ∫)
  let chr_val = math_prop(node, "m:naryPr", "m:chr")
  let op_char = if chr_val.is_empty() { "\u{222B}" } else { chr_val } // ∫
  let lim_loc = math_prop(node, "m:naryPr", "m:limLoc") // "subSup" or "undOvr"
  let op_fs = fs * math_nary_scale
  let limit_fs = fs * math_nary_limit_scale
  // Layout operator
  let op_w = math_text_w(op_char, ff, op_fs)
  let op_h = op_fs * math_text_height_ratio
  let op_bl = op_fs * math_text_baseline_ratio
  // Layout limits
  let sub_box = match math_child(node, "m:sub") {
    Some(s) => layout_math_children(s, limit_fs, ff, color)
    None => math_empty_box
  }
  let sup_box = match math_child(node, "m:sup") {
    Some(s) => layout_math_children(s, limit_fs, ff, color)
    None => math_empty_box
  }
  // Layout content
  let content_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  if lim_loc == "undOvr" {
    // Limits above and below operator
    let col_w_candidates = [op_w, sub_box.w, sup_box.w]
    let mut col_w : Double = 0.0
    let mut ci = 0
    while ci < col_w_candidates.length() {
      if col_w_candidates[ci] > col_w {
        col_w = col_w_candidates[ci]
      }
      ci = ci + 1
    }
    let sup_y : Double = 0.0
    let op_y = sup_box.h + 1.0
    let sub_y = op_y + op_h + 1.0
    let total_h = sub_y + sub_box.h
    let baseline = op_y + op_bl
    let op_svg = math_text_svg(
      op_char,
      (col_w - op_w) / 2.0,
      op_y + op_bl,
      op_fs,
      ff,
      color,
    )
    let sup_svg = if sup_box.w > 0.0 {
      math_g((col_w - sup_box.w) / 2.0, sup_y, sup_box.svg)
    } else {
      ""
    }
    let sub_svg = if sub_box.w > 0.0 {
      math_g((col_w - sub_box.w) / 2.0, sub_y, sub_box.svg)
    } else {
      ""
    }
    let content_x = col_w + 2.0
    let content_y = baseline - content_box.baseline
    let total_w = col_w + 2.0 + content_box.w
    let svg = sup_svg +
      op_svg +
      sub_svg +
      math_g(content_x, content_y, content_box.svg)
    { svg, w: total_w, h: total_h, baseline }
  } else {
    // Sub/sup style (limits as sub/superscript of operator)
    let op_svg = math_text_svg(op_char, 0.0, op_bl, op_fs, ff, color)
    let script_x = op_w
    let sup_y : Double = 0.0
    let sub_y = op_h - sub_box.h
    let script_w = if sup_box.w > sub_box.w { sup_box.w } else { sub_box.w }
    let content_x = op_w + script_w + 2.0
    let baseline = op_bl
    let content_y = baseline - content_box.baseline
    let total_w = content_x + content_box.w
    let mut total_h = op_h
    if content_y + content_box.h > total_h {
      total_h = content_y + content_box.h
    }
    let sup_svg = if sup_box.w > 0.0 {
      math_g(script_x, sup_y, sup_box.svg)
    } else {
      ""
    }
    let sub_svg = if sub_box.w > 0.0 {
      math_g(script_x, sub_y, sub_box.svg)
    } else {
      ""
    }
    let svg = op_svg +
      sup_svg +
      sub_svg +
      math_g(
        content_x,
        if content_y > 0.0 {
          content_y
        } else {
          0.0
        },
        content_box.svg,
      )
    { svg, w: total_w, h: total_h, baseline }
  }
}

// ── m:d — Delimiters (brackets) ─────────────────────────────────────────────

///|
fn layout_math_delim(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  // Get delimiter characters
  let beg_chr_val = match math_child(node, "m:dPr") {
    Some(pr) =>
      match math_child(pr, "m:begChr") {
        Some(bc) =>
          match @xml.get_attr(bc, "m:val") {
            Some(v) => v
            None => "("
          }
        None => "("
      }
    None => "("
  }
  let end_chr_val = match math_child(node, "m:dPr") {
    Some(pr) =>
      match math_child(pr, "m:endChr") {
        Some(ec) =>
          match @xml.get_attr(ec, "m:val") {
            Some(v) => v
            None => ")"
          }
        None => ")"
      }
    None => ")"
  }
  // Layout content (may have multiple m:e children separated by separator)
  let e_nodes = math_children(node, "m:e")
  let content_boxes : Array[MathBox] = []
  let ne = e_nodes.length()
  let mut ei = 0
  while ei < ne {
    if ei > 0 {
      // Add separator (comma by default)
      let sep_w = math_text_w(",", ff, fs)
      let sep_h = fs * math_text_height_ratio
      let sep_bl = fs * math_text_baseline_ratio
      let sep_svg = math_text_svg(",", 0.0, sep_bl, fs, ff, color)
      content_boxes.push({ svg: sep_svg, w: sep_w, h: sep_h, baseline: sep_bl })
    }
    content_boxes.push(layout_math_children(e_nodes[ei], fs, ff, color))
    ei = ei + 1
  }
  let content = math_hbox(content_boxes)
  // Scale delimiters to content height
  let delim_fs = if content.h > fs * math_text_height_ratio {
    content.h * 0.9
  } else {
    fs
  }
  let beg_w = math_text_w(beg_chr_val, ff, delim_fs) + math_delim_pad
  let end_w = math_text_w(end_chr_val, ff, delim_fs) + math_delim_pad
  let total_w = beg_w + content.w + end_w
  let total_h = content.h
  let baseline = content.baseline
  let beg_svg = math_text_svg(beg_chr_val, 0.0, baseline, delim_fs, ff, color)
  let end_svg = math_text_svg(
    end_chr_val,
    beg_w + content.w,
    baseline,
    delim_fs,
    ff,
    color,
  )
  let svg = beg_svg + math_g(beg_w, 0.0, content.svg) + end_svg
  { svg, w: total_w, h: total_h, baseline }
}

// ── m:acc — Accent ──────────────────────────────────────────────────────────

///|
fn layout_math_acc(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let content_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  // Get accent character (default: combining circumflex accent)
  let acc_chr = match math_child(node, "m:accPr") {
    Some(pr) =>
      match math_child(pr, "m:chr") {
        Some(c) =>
          match @xml.get_attr(c, "m:val") {
            Some(v) => v
            None => "\u{0302}" // combining circumflex
          }
        None => "\u{0302}"
      }
    None => "\u{0302}"
  }
  // Map combining characters to non-combining display characters
  // Combining chars (U+0300 range) render poorly standalone in SVG <text>;
  // use full-width non-combining equivalents instead.
  let display_acc = if acc_chr == "\u{0302}" || acc_chr == "\u{005E}" {
    "\u{02C6}" // modifier letter circumflex accent
  } else if acc_chr == "\u{0303}" || acc_chr == "\u{007E}" {
    "\u{02DC}" // small tilde
  } else if acc_chr == "\u{0304}" || acc_chr == "\u{00AF}" {
    "\u{00AF}" // macron
  } else if acc_chr == "\u{0307}" {
    "\u{02D9}" // dot above
  } else if acc_chr == "\u{0308}" {
    "\u{00A8}" // diaeresis
  } else if acc_chr == "\u{030C}" {
    "\u{02C7}" // caron
  } else if acc_chr == "\u{0301}" {
    "\u{00B4}" // acute accent
  } else if acc_chr == "\u{0300}" {
    "\u{0060}" // grave accent
  } else {
    acc_chr
  }
  // Draw accent as SVG line/path for common accents that don't render well as text
  let use_line = display_acc == "\u{00AF}" // macron → draw as line
  if use_line {
    // Draw overline as SVG <line>
    let line_thick : Double = 1.2
    let line_gap : Double = 1.0
    let content_y = line_thick + line_gap
    let total_h = content_y + content_box.h
    let baseline = content_y + content_box.baseline
    let svg = "<line x1=\"0\" y1=\"" +
      dts(line_thick / 2.0) +
      "\" x2=\"" +
      dts(content_box.w) +
      "\" y2=\"" +
      dts(line_thick / 2.0) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"" +
      dts(line_thick) +
      "\"/>" +
      math_g(0.0, content_y, content_box.svg)
    return { svg, w: content_box.w, h: total_h, baseline }
  }
  // Draw common accents as SVG paths for precise positioning.
  // SVG <text> y is the baseline, causing glyphs to float above — paths are more reliable.
  let is_circumflex = display_acc == "\u{02C6}" || display_acc == "\u{005E}"
  let is_tilde = display_acc == "\u{02DC}" || display_acc == "\u{007E}"
  let is_dot = display_acc == "\u{02D9}"
  let is_diaeresis = display_acc == "\u{00A8}"
  let acc_thick : Double = 1.2
  let acc_gap2 : Double = 1.0
  if is_circumflex {
    // Draw ^ as SVG path: /\ shape
    let cw = content_box.w
    let peak_h = fs * 0.18
    let half_w = cw * 0.25
    let cx = cw / 2.0
    let content_y = peak_h + acc_gap2
    let total_h = content_y + content_box.h
    let baseline = content_y + content_box.baseline
    let svg = "<path d=\"M " +
      dts(cx - half_w) +
      " " +
      dts(peak_h) +
      " L " +
      dts(cx) +
      " 0" +
      " L " +
      dts(cx + half_w) +
      " " +
      dts(peak_h) +
      "\" fill=\"none\" stroke=\"" +
      color +
      "\" stroke-width=\"" +
      dts(acc_thick) +
      "\" stroke-linejoin=\"miter\"/>" +
      math_g(0.0, content_y, content_box.svg)
    return { svg, w: cw, h: total_h, baseline }
  }
  if is_tilde {
    // Draw ~ as SVG path: sinusoidal curve
    let cw = content_box.w
    let wave_h = fs * 0.1
    let cx = cw / 2.0
    let content_y = wave_h * 2.0 + acc_gap2
    let total_h = content_y + content_box.h
    let baseline = content_y + content_box.baseline
    let svg = "<path d=\"M " +
      dts(cx - cw * 0.25) +
      " " +
      dts(wave_h * 1.5) +
      " Q " +
      dts(cx - cw * 0.1) +
      " " +
      dts(wave_h * 0.3) +
      " " +
      dts(cx) +
      " " +
      dts(wave_h) +
      " Q " +
      dts(cx + cw * 0.1) +
      " " +
      dts(wave_h * 1.7) +
      " " +
      dts(cx + cw * 0.25) +
      " " +
      dts(wave_h * 0.5) +
      "\" fill=\"none\" stroke=\"" +
      color +
      "\" stroke-width=\"" +
      dts(acc_thick) +
      "\"/>" +
      math_g(0.0, content_y, content_box.svg)
    return { svg, w: cw, h: total_h, baseline }
  }
  if is_dot {
    // Draw dot as SVG circle
    let cw = content_box.w
    let dot_r = fs * 0.04
    let content_y = dot_r * 2.0 + acc_gap2
    let total_h = content_y + content_box.h
    let baseline = content_y + content_box.baseline
    let svg = "<circle cx=\"" +
      dts(cw / 2.0) +
      "\" cy=\"" +
      dts(dot_r) +
      "\" r=\"" +
      dts(dot_r) +
      "\" fill=\"" +
      color +
      "\"/>" +
      math_g(0.0, content_y, content_box.svg)
    return { svg, w: cw, h: total_h, baseline }
  }
  if is_diaeresis {
    // Draw diaeresis as two dots
    let cw = content_box.w
    let dot_r = fs * 0.035
    let dot_spacing = cw * 0.2
    let content_y = dot_r * 2.0 + acc_gap2
    let total_h = content_y + content_box.h
    let baseline = content_y + content_box.baseline
    let svg = "<circle cx=\"" +
      dts(cw / 2.0 - dot_spacing) +
      "\" cy=\"" +
      dts(dot_r) +
      "\" r=\"" +
      dts(dot_r) +
      "\" fill=\"" +
      color +
      "\"/>" +
      "<circle cx=\"" +
      dts(cw / 2.0 + dot_spacing) +
      "\" cy=\"" +
      dts(dot_r) +
      "\" r=\"" +
      dts(dot_r) +
      "\" fill=\"" +
      color +
      "\"/>" +
      math_g(0.0, content_y, content_box.svg)
    return { svg, w: cw, h: total_h, baseline }
  }
  // Fallback: use text-based accent with negative offset to bring glyph closer
  let acc_fs = fs * 0.7
  let acc_w = math_text_w(display_acc, ff, acc_fs)
  let acc_region = acc_fs * 0.15
  let total_h = acc_region + content_box.h
  let baseline = acc_region + content_box.baseline
  let acc_x = (content_box.w - acc_w) / 2.0
  let acc_svg = math_text_svg(display_acc, acc_x, acc_region, acc_fs, ff, color)
  let svg = acc_svg + math_g(0.0, acc_region, content_box.svg)
  { svg, w: content_box.w, h: total_h, baseline }
}

// ── m:m — Matrix ────────────────────────────────────────────────────────────

///|
fn layout_math_matrix(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let rows = math_children(node, "m:mr")
  let nrows = rows.length()
  if nrows == 0 {
    return math_empty_box
  }
  // Measure all cells
  let cell_boxes : Array[Array[MathBox]] = []
  let mut max_cols = 0
  let mut ri = 0
  while ri < nrows {
    let cells = math_children(rows[ri], "m:e")
    let row : Array[MathBox] = []
    let nc = cells.length()
    if nc > max_cols {
      max_cols = nc
    }
    let mut ci = 0
    while ci < nc {
      row.push(layout_math_children(cells[ci], fs, ff, color))
      ci = ci + 1
    }
    cell_boxes.push(row)
    ri = ri + 1
  }
  if max_cols == 0 {
    return math_empty_box
  }
  // Compute column widths and row heights
  let col_widths : Array[Double] = Array::make(max_cols, 0.0)
  let row_heights : Array[Double] = Array::make(nrows, 0.0)
  let row_baselines : Array[Double] = Array::make(nrows, 0.0)
  ri = 0
  while ri < nrows {
    let row = cell_boxes[ri]
    let nc = row.length()
    let mut ci = 0
    while ci < nc {
      let box_ = row[ci]
      if box_.w > col_widths[ci] {
        col_widths[ci] = box_.w
      }
      if box_.h > row_heights[ri] {
        row_heights[ri] = box_.h
      }
      if box_.baseline > row_baselines[ri] {
        row_baselines[ri] = box_.baseline
      }
      ci = ci + 1
    }
    ri = ri + 1
  }
  let gap = fs * math_matrix_gap_ratio
  // Compute total size
  let mut total_w : Double = 0.0
  let mut ci = 0
  while ci < max_cols {
    total_w = total_w + col_widths[ci]
    if ci < max_cols - 1 {
      total_w = total_w + gap
    }
    ci = ci + 1
  }
  let mut total_h : Double = 0.0
  ri = 0
  while ri < nrows {
    total_h = total_h + row_heights[ri]
    if ri < nrows - 1 {
      total_h = total_h + gap
    }
    ri = ri + 1
  }
  // Render cells
  let mut svg = ""
  let mut cur_y : Double = 0.0
  ri = 0
  while ri < nrows {
    let row = cell_boxes[ri]
    let nc = row.length()
    let mut cur_x : Double = 0.0
    let mut ci2 = 0
    while ci2 < nc {
      let box_ = row[ci2]
      // Center horizontally in column, align baseline in row
      let cx = cur_x + (col_widths[ci2] - box_.w) / 2.0
      let cy = cur_y + (row_baselines[ri] - box_.baseline)
      if box_.svg.length() > 0 {
        svg = svg + math_g(cx, cy, box_.svg)
      }
      cur_x = cur_x + col_widths[ci2] + gap
      ci2 = ci2 + 1
    }
    cur_y = cur_y + row_heights[ri] + gap
    ri = ri + 1
  }
  let baseline = total_h / 2.0 + fs * 0.3 // center baseline
  { svg, w: total_w, h: total_h, baseline }
}

// ── m:bar — Over/under bar ──────────────────────────────────────────────────

///|
fn layout_math_bar(
  node : @xml.XmlNode,
  fs : Double,
  ff : String,
  color : String,
) -> MathBox {
  let content_box = match math_child(node, "m:e") {
    Some(e) => layout_math_children(e, fs, ff, color)
    None => math_empty_box
  }
  let pos = math_prop(node, "m:barPr", "m:pos")
  let bar_thick : Double = 1.5
  let bar_gap : Double = 2.0
  if pos == "bot" {
    // Underbar
    let bar_y = content_box.h + bar_gap
    let total_h = bar_y + bar_thick
    let svg = content_box.svg +
      "<line x1=\"0\" y1=\"" +
      dts(bar_y) +
      "\" x2=\"" +
      dts(content_box.w) +
      "\" y2=\"" +
      dts(bar_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"" +
      dts(bar_thick) +
      "\"/>"
    { svg, w: content_box.w, h: total_h, baseline: content_box.baseline }
  } else {
    // Overbar (default)
    let bar_y = bar_thick
    let content_y = bar_y + bar_gap
    let total_h = content_y + content_box.h
    let svg = "<line x1=\"0\" y1=\"" +
      dts(bar_y) +
      "\" x2=\"" +
      dts(content_box.w) +
      "\" y2=\"" +
      dts(bar_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"" +
      dts(bar_thick) +
      "\"/>" +
      math_g(0.0, content_y, content_box.svg)
    {
      svg,
      w: content_box.w,
      h: total_h,
      baseline: content_y + content_box.baseline,
    }
  }
}

// ── Entry points ────────────────────────────────────────────────────────────

///|
/// Check if all paragraphs contain only math runs.
fn is_all_math(paragraphs : Array[@ooxml.TextParagraph]) -> Bool {
  let np = paragraphs.length()
  if np == 0 {
    return false
  }
  let mut has_math = false
  let mut i = 0
  while i < np {
    let para = paragraphs[i]
    let nr = para.runs.length()
    let mut j = 0
    while j < nr {
      let run = para.runs[j]
      if run.math_xml.length() > 0 {
        has_math = true
      } else if !run.text.is_empty() {
        // Non-empty text run without math → mixed content
        // Allow whitespace-only runs
        let mut all_ws = true
        let mut k = 0
        let tlen = run.text.length()
        while k < tlen {
          match run.text.get_char(k) {
            Some(c) =>
              if c != ' ' && c != '\t' && c != '\n' && c != '\r' {
                all_ws = false
              }
            None => ()
          }
          k = k + 1
        }
        if !all_ws {
          return false
        }
      }
      j = j + 1
    }
    i = i + 1
  }
  has_math
}

///|
/// Render math-only text box as SVG.
fn render_math_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  _theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
  _rels : Array[@ooxml.Relationship],
  _slide_num : Int,
  _shape_idx : Int,
) -> String {
  // Collect all math XML from runs
  let np = paragraphs.length()
  let mut all_math_xml = ""
  let mut first_run_color : @ooxml.Color = { r: -1, g: 0, b: 0, alpha: 255 }
  let mut first_run_fs = 0
  let mut first_run_font = ""
  let mut found_first = false
  let mut i = 0
  while i < np {
    let para = paragraphs[i]
    let nr = para.runs.length()
    let mut j = 0
    while j < nr {
      let run = para.runs[j]
      if run.math_xml.length() > 0 {
        all_math_xml = run.math_xml
        if !found_first {
          first_run_color = run.color
          first_run_fs = run.font_size
          first_run_font = run.font_face
          found_first = true
        }
      }
      j = j + 1
    }
    i = i + 1
  }
  if all_math_xml.is_empty() {
    return ""
  }
  // Determine font size in pixels
  let fs_emu = if first_run_fs > 0 {
    hundredth_pt_to_emu(first_run_fs)
  } else {
    default_font_size_emu()
  }
  let fs_px = px(fs_emu, scale).to_double()
  // Determine color — use same fallback as regular text rendering
  let color_str = if first_run_color.r >= 0 {
    first_run_color.to_css()
  } else {
    default_text_fill
  }
  // Font family
  let ff = if !first_run_font.is_empty() {
    first_run_font + ", " + math_font_family
  } else {
    math_font_family
  }
  // Parse the math XML
  let nodes = @xml.parse_xml(all_math_xml)
  if nodes.length() == 0 {
    return ""
  }
  // Layout the math
  let boxes : Array[MathBox] = []
  let nn = nodes.length()
  let mut ni = 0
  while ni < nn {
    boxes.push(layout_math_node(nodes[ni], fs_px, ff, color_str))
    ni = ni + 1
  }
  let math_box = math_hbox(boxes)
  if math_box.w <= 0.0 {
    return ""
  }
  // Position within shape
  let l_ins = if body_props.l_ins >= 0 { body_props.l_ins } else { 91440 }
  let t_ins = if body_props.t_ins >= 0 { body_props.t_ins } else { 45720 }
  let r_ins = if body_props.r_ins >= 0 { body_props.r_ins } else { 91440 }
  let b_ins = if body_props.b_ins >= 0 { body_props.b_ins } else { 45720 }
  let shape_x = px(t.x, scale).to_double()
  let shape_y = px(t.y, scale).to_double()
  let shape_w = px(t.cx, scale).to_double()
  let shape_h = px(t.cy, scale).to_double()
  let avail_w = shape_w -
    px(l_ins, scale).to_double() -
    px(r_ins, scale).to_double()
  let avail_h = shape_h -
    px(t_ins, scale).to_double() -
    px(b_ins, scale).to_double()
  // Center horizontally within available area
  let math_x = shape_x +
    px(l_ins, scale).to_double() +
    (avail_w - math_box.w) / 2.0
  // Vertical alignment
  let math_y = if body_props.anchor == "ctr" {
    shape_y + px(t_ins, scale).to_double() + (avail_h - math_box.h) / 2.0
  } else if body_props.anchor == "b" {
    shape_y + shape_h - px(b_ins, scale).to_double() - math_box.h
  } else {
    // top (default)
    shape_y + px(t_ins, scale).to_double()
  }
  // Build SVG: visual math rendering + hidden text for round-trip
  let visual_svg = "<g" +
    da("math-xml", @xml.xml_escape(all_math_xml)) +
    ">" +
    math_g(math_x, math_y, math_box.svg) +
    "</g>"
  // Hidden text element preserving round-trip data attributes
  let hidden_svg = "<text" +
    " x=\"" +
    dts(shape_x) +
    "\" y=\"" +
    dts(shape_y) +
    "\" font-size=\"0\" visibility=\"hidden\">" +
    "<tspan" +
    da("para-idx", "0") +
    da("para-align", "c") +
    "><tspan" +
    da("run-idx", "0") +
    da("font-size", int_to_str(first_run_fs)) +
    da("color", first_run_color.to_hex()) +
    da("run-font", first_run_font) +
    da("math-xml", @xml.xml_escape(all_math_xml)) +
    "></tspan></tspan></text>"
  visual_svg + hidden_svg
}
