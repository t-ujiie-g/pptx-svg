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

// ── Integer → String ──────────────────────────────────────────────────────────

fn digit_str(d : Int) -> String {
  match d {
    0 => "0"; 1 => "1"; 2 => "2"; 3 => "3"; 4 => "4"
    5 => "5"; 6 => "6"; 7 => "7"; 8 => "8"; _ => "9"
  }
}

fn int_to_str(n : Int) -> String {
  if n < 0 {
    "-" + int_to_str(-n)
  } else if n < 10 {
    digit_str(n)
  } else {
    int_to_str(n / 10) + digit_str(n % 10)
  }
}

// ── SVG attribute helpers ─────────────────────────────────────────────────────

fn a(name : String, value : String) -> String {
  " " + name + "=\"" + value + "\""
}

fn ai(name : String, value : Int) -> String {
  a(name, int_to_str(value))
}

fn bool_str(b : Bool) -> String {
  if b { "1" } else { "0" }
}

// Escape text for SVG content.
fn svg_escape(s : String) -> String {
  let len = s.length()
  let mut needs = false
  let mut k = 0
  while k < len {
    let c = s.get_char(k).unwrap()
    if c == '<' || c == '>' || c == '&' || c == '"' {
      needs = true
      break
    }
    k = k + 1
  }
  if not(needs) { return s }
  let mut result = ""
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '<' {
      result = result + "&lt;"
    } else if c == '>' {
      result = result + "&gt;"
    } else if c == '&' {
      result = result + "&amp;"
    } else if c == '"' {
      result = result + "&quot;"
    } else {
      result = result + @ffi.ffi_char_code_to_str(Char::to_int(c))
    }
    i = i + 1
  }
  result
}

// ── EMU → pixel conversion ────────────────────────────────────────────────────

fn px(emu : Int, scale : Int) -> Int {
  if scale <= 0 { 0 } else { emu / scale }
}

// ── Font size ─────────────────────────────────────────────────────────────────

fn hundredth_pt_to_emu(sz : Int) -> Int {
  sz * 127
}

fn default_font_size_emu() -> Int {
  228600 // 18 pt
}

// ── Rotation transform ────────────────────────────────────────────────────────

fn rotation_attr(t : @ooxml.ShapeTransform, scale : Int) -> String {
  if t.rot == 0 { return "" }
  let deg_x100 = t.rot / 600
  let deg_int  = deg_x100 / 100
  let deg_frac = deg_x100 % 100
  let frac_str = if deg_frac < 10 {
    "0" + int_to_str(deg_frac)
  } else {
    int_to_str(deg_frac)
  }
  let cx_p = px(t.x + t.cx / 2, scale)
  let cy_p = px(t.y + t.cy / 2, scale)
  " transform=\"rotate(" +
  int_to_str(deg_int) + "." + frac_str + "," +
  int_to_str(cx_p) + "," + int_to_str(cy_p) + ")\""
}

// ── Image helpers ─────────────────────────────────────────────────────────────

fn resolve_image_path(target : String) -> String {
  if target.length() > 3 &&
    target.get_char(0).unwrap() == '.' &&
    target.get_char(1).unwrap() == '.' &&
    target.get_char(2).unwrap() == '/' {
    "ppt/" + suffix(target, 3)
  } else {
    target
  }
}

fn suffix(s : String, from : Int) -> String {
  let len = s.length()
  let mut result = ""
  let mut i = from
  while i < len {
    result = result + @ffi.ffi_char_code_to_str(Char::to_int(s.get_char(i).unwrap()))
    i = i + 1
  }
  result
}

fn mime_from_ext(path : String) -> String {
  let len = path.length()
  let mut dot = len - 1
  while dot >= 0 {
    if path.get_char(dot).unwrap() == '.' { break }
    dot = dot - 1
  }
  if dot < 0 { return "image/png" }
  let ext = suffix(path, dot + 1)
  if ext == "png" { "image/png" }
  else if ext == "jpg" || ext == "jpeg" { "image/jpeg" }
  else if ext == "gif" { "image/gif" }
  else if ext == "svg" || ext == "svgz" { "image/svg+xml" }
  else { "image/png" }
}

// ── data-ooxml-* attribute helpers ───────────────────────────────────────────

fn shape_type_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(_) => "autoshape"
    @ooxml.ShapeKind::Picture(_) => "picture"
    @ooxml.ShapeKind::Table => "table"
    @ooxml.ShapeKind::GroupShape => "group"
    @ooxml.ShapeKind::Other => "other"
  }
}

fn geom_str(kind : @ooxml.ShapeKind) -> String {
  match kind {
    @ooxml.ShapeKind::AutoShape(g) => g.to_prst()
    _ => "rect"
  }
}

// ── Text rendering ────────────────────────────────────────────────────────────

fn render_text(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int
) -> String {
  if paragraphs.length() == 0 { return "" }
  let pad_x   = px(t.cx / 20, scale)
  let pad_y   = px(t.cy / 20, scale)
  let text_x  = px(t.x, scale) + pad_x
  let text_w  = px(t.cx, scale) - pad_x * 2
  let def_sz  = px(default_font_size_emu(), scale) // 24 px for standard slides
  let line_h  = def_sz * 12 / 10

  let mut out = "<text" +
    ai("x", text_x) +
    a("xml:space", "preserve") +
    a("font-family", "Arial, sans-serif") +
    ai("font-size", def_sz) +
    a("fill", "black") +
    ">"

  let mut cur_y = px(t.y, scale) + pad_y + def_sz
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    if nr == 0 {
      cur_y = cur_y + line_h
      pi = pi + 1
    } else {
      let (anchor_x, anchor_attr) = if para.align == "ctr" || para.align == "c" {
        (text_x + text_w / 2, " text-anchor=\"middle\"")
      } else if para.align == "r" {
        (text_x + text_w, " text-anchor=\"end\"")
      } else {
        (text_x, " text-anchor=\"start\"")
      }
      // Paragraph-level tspan with data-ooxml-* attributes
      let para_align = if para.align.is_empty() { "l" } else { para.align }
      out = out + "<tspan" +
        ai("x", anchor_x) +
        ai("y", cur_y) +
        anchor_attr +
        a("data-ooxml-para-idx", int_to_str(pi)) +
        a("data-ooxml-para-align", para_align) +
        ">"
      let mut ri = 0
      while ri < nr {
        let run = para.runs[ri]
        let mut span = "<tspan"
        // Visual attributes
        if run.bold   { span = span + " font-weight=\"bold\""  }
        if run.italic { span = span + " font-style=\"italic\"" }
        if run.font_size > 0 {
          span = span + ai("font-size", px(hundredth_pt_to_emu(run.font_size), scale))
        }
        if not(run.color.is_none()) {
          span = span + a("fill", run.color.to_css())
        }
        // data-ooxml-* attributes for round-trip
        span = span +
          a("data-ooxml-run-idx", int_to_str(ri)) +
          a("data-ooxml-bold", bool_str(run.bold)) +
          a("data-ooxml-font-size", int_to_str(run.font_size)) +
          a("data-ooxml-color", run.color.to_hex())
        out = out + span + ">" + svg_escape(run.text) + "</tspan>"
        ri = ri + 1
      }
      out = out + "</tspan>"
      cur_y = cur_y + line_h
      pi = pi + 1
    }
  }
  out + "</text>"
}

// ── Shape rendering ───────────────────────────────────────────────────────────

fn render_shape(
  shape : @ooxml.Shape,
  rels : Array[@ooxml.Relationship],
  _slide_num : Int,
  scale : Int,
  shape_idx : Int
) -> String {
  let t    = shape.transform
  let x_p  = px(t.x,  scale)
  let y_p  = px(t.y,  scale)
  let cx_p = px(t.cx, scale)
  let cy_p = px(t.cy, scale)
  if cx_p <= 0 || cy_p <= 0 { return "" }
  let rot = rotation_attr(t, scale)

  // Build data-ooxml-* attributes for the <g> wrapper
  let g_attrs =
    a("data-ooxml-shape-type", shape_type_str(shape.kind)) +
    a("data-ooxml-geom", geom_str(shape.kind)) +
    ai("data-ooxml-x", t.x) +
    ai("data-ooxml-y", t.y) +
    ai("data-ooxml-cx", t.cx) +
    ai("data-ooxml-cy", t.cy) +
    ai("data-ooxml-rot", t.rot) +
    a("data-ooxml-fill", shape.fill.to_hex()) +
    a("data-ooxml-stroke", shape.stroke.to_hex()) +
    ai("data-ooxml-stroke-w", shape.stroke_w) +
    ai("data-ooxml-shape-idx", shape_idx)

  let inner = match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => {
      let fill   = a("fill", shape.fill.to_css())
      let stroke = if shape.stroke.is_none() {
        " stroke=\"none\""
      } else {
        a("stroke", shape.stroke.to_css()) +
        ai("stroke-width",
          px(if shape.stroke_w > 0 { shape.stroke_w } else { 12700 }, scale))
      }
      let rect_svg = match geom {
        @ooxml.ShapeGeom::Ellipse => {
          "<ellipse" +
          ai("cx", x_p + cx_p / 2) + ai("cy", y_p + cy_p / 2) +
          ai("rx", cx_p / 2) + ai("ry", cy_p / 2) +
          fill + stroke + rot + "/>"
        }
        _ => {
          let rx = match geom {
            @ooxml.ShapeGeom::RoundRect => int_to_str(cx_p / 20)
            _ => "0"
          }
          let base = "<rect" +
            ai("x", x_p) + ai("y", y_p) +
            ai("width", cx_p) + ai("height", cy_p) +
            fill + stroke + rot
          if rx == "0" {
            base + "/>"
          } else {
            base + a("rx", rx) + a("ry", rx) + "/>"
          }
        }
      }
      rect_svg + render_text(shape.paragraphs, t, scale)
    }
    @ooxml.ShapeKind::Picture(rid) => {
      let target = @ooxml.find_rel_target(rels, rid)
      let placeholder = "<rect" +
        ai("x", x_p) + ai("y", y_p) +
        ai("width", cx_p) + ai("height", cy_p) +
        " fill=\"#dddddd\"" + rot + "/>"
      if target.is_empty() {
        placeholder
      } else {
        let img_path = resolve_image_path(target)
        let b64 = @ffi.ffi_get_file_base64(img_path)
        if b64.is_empty() {
          placeholder
        } else {
          let mime = mime_from_ext(img_path)
          "<image" +
          ai("x", x_p) + ai("y", y_p) +
          ai("width", cx_p) + ai("height", cy_p) +
          " preserveAspectRatio=\"xMidYMid meet\"" +
          a("href", "data:" + mime + ";base64," + b64) +
          rot + "/>"
        }
      }
    }
    _ => ""
  }

  if inner.is_empty() {
    ""
  } else {
    "<g" + g_attrs + ">" + inner + "</g>"
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

///| Render a SlideData to an SVG string with data-ooxml-* attributes.
pub fn render_slide_svg(
  slide : @ooxml.SlideData,
  rels : Array[@ooxml.Relationship],
  slide_num : Int
) -> String {
  let sw = slide.slide_size.cx
  let sh = slide.slide_size.cy
  let w = 960
  let scale = if sw > 0 { sw / w } else { 9525 }
  let h = if scale > 0 { sh / scale } else { 540 }
  let header = "<svg" +
    " xmlns=\"http://www.w3.org/2000/svg\"" +
    " xmlns:xlink=\"http://www.w3.org/1999/xlink\"" +
    ai("width", w) + ai("height", h) +
    " style=\"background:white;display:block\"" +
    // data-ooxml-* metadata on SVG root for round-trip
    ai("data-ooxml-slide-cx", sw) +
    ai("data-ooxml-slide-cy", sh) +
    a("data-ooxml-bg", slide.background.to_hex()) +
    ai("data-ooxml-scale", scale) +
    ">"
  let bg = "<rect" +
    ai("width", w) + ai("height", h) +
    a("fill", slide.background.to_css()) +
    "/>"
  let mut shapes_svg = ""
  let n = slide.shapes.length()
  let mut i = 0
  while i < n {
    shapes_svg = shapes_svg + render_shape(slide.shapes[i], rels, slide_num, scale, i)
    i = i + 1
  }
  header + bg + shapes_svg + "</svg>"
}
