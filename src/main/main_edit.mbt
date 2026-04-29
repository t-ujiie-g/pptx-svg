// pptx-svg: Shape and text editing API
//
// This file contains all shape-level editing exports:
//   - render_shape_svg, update_shape_transform, update_shape_text, update_shape_fill
//   - delete_shape, add_shape, add_shape_text, duplicate_shape
//   - update_shape_gradient_fill, update_shape_stroke
//   - Text editing (E2.5): add/delete paragraph/run, style/font/color/align/decoration

// ── Helpers ─────────────────────────────────────────────────────────────────

///|
/// Resolve a composite shape index (used by group children: parent*1000+child)
/// to the containing array and the local index within it.
/// Returns (shapes_array, local_index) or None if out of range.
fn resolve_shape_location(
  slide : @ooxml.SlideData,
  composite_idx : Int,
) -> (Array[@ooxml.Shape], Int)? {
  if composite_idx < 1000 {
    if composite_idx >= 0 && composite_idx < slide.shapes.length() {
      Some((slide.shapes, composite_idx))
    } else {
      None
    }
  } else {
    let parent_composite = composite_idx / 1000
    let child_local = composite_idx % 1000
    match resolve_shape_location(slide, parent_composite) {
      Some((parent_arr, parent_local)) =>
        match parent_arr[parent_local].kind {
          GroupShape(grp) =>
            if child_local >= 0 && child_local < grp.children.length() {
              Some((grp.children, child_local))
            } else {
              None
            }
          _ => None
        }
      None => None
    }
  }
}

///|
/// Validate slide index and resolve shape. Returns error string or calls `f` with (arr, local_idx).
fn with_shape(
  slide_idx : Int,
  shape_idx : Int,
  f : (Array[@ooxml.Shape], Int) -> String,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide = g_slides.val[slide_idx]
  match resolve_shape_location(slide, shape_idx) {
    None => "ERROR:shape index out of range"
    Some((arr, local_idx)) => f(arr, local_idx)
  }
}

///|
/// Validate slide/shape/paragraph and resolve the text run.
/// Returns error string or calls `f` with (shape, paragraph, run).
fn with_run(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  f : (@ooxml.Shape, @ooxml.TextParagraph, @ooxml.TextRun) -> String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    if para_idx < 0 || para_idx >= shape.paragraphs.length() {
      return "ERROR:paragraph index out of range"
    }
    let para = shape.paragraphs[para_idx]
    if run_idx < 0 || run_idx >= para.runs.length() {
      return "ERROR:run index out of range"
    }
    f(shape, para, para.runs[run_idx])
  })
}

///|
/// Create a default TextRun with the given text and optional settings.
fn make_default_run(
  text : String,
  font_size? : Int = 0,
  color? : @ooxml.Color = @ooxml.Color::none(),
) -> @ooxml.TextRun {
  {
    text,
    bold: false,
    bold_explicit: false,
    italic: false,
    font_size,
    color,
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
    math_xml: "",
  }
}

///|
/// Create a default TextParagraph with a single run.
fn make_default_paragraph(
  run : @ooxml.TextRun,
  align? : String = "",
) -> @ooxml.TextParagraph {
  {
    runs: [run],
    align,
    level: 0,
    spc_before: 0,
    spc_after: 0,
    line_spacing: 0,
    mar_l: -1,
    indent: 0,
    bullet: "",
    bullet_auto: "",
    bullet_none: false,
    bullet_font: "",
    bullet_size: 0,
    bullet_color: @ooxml.Color::none(),
    bullet_img_rid: "",
    tab_stops: [],
    rtl: false,
    epr_font_size: 0,
    epr_color: @ooxml.Color::none(),
    epr_font_face: "",
    epr_ea_font: "",
  }
}

///|
/// Remove an element at `idx` from an array in-place.
fn[T] array_remove_at(arr : Array[T], idx : Int) -> Unit {
  let new_arr : Array[T] = []
  let len = arr.length()
  let mut i = 0
  while i < len {
    if i != idx {
      new_arr.push(arr[i])
    }
    i = i + 1
  }
  while arr.length() > 0 {
    let _ = arr.pop()
  }
  let mut j = 0
  while j < new_arr.length() {
    arr.push(new_arr[j])
    j = j + 1
  }
}

///|
fn parse_gradient_stops(s : String) -> Array[@ooxml.GradientStop] {
  let stops : Array[@ooxml.GradientStop] = []
  if s.is_empty() {
    return stops
  }
  let parts = split_string(s, ';')
  let np = parts.length()
  let mut i = 0
  while i < np {
    let part = parts[i]
    if !part.is_empty() {
      let vals = split_string(part, ',')
      if vals.length() >= 4 {
        let pos = @xml.parse_int(vals[0])
        let r = @xml.parse_int(vals[1])
        let g = @xml.parse_int(vals[2])
        let b = @xml.parse_int(vals[3])
        stops.push({ pos, color: ({ r, g, b, alpha: 255 } : @ooxml.Color) })
      }
    }
    i = i + 1
  }
  stops
}

///|
fn split_string(s : String, sep : Char) -> Array[String] {
  let result : Array[String] = []
  let len = s.length()
  let mut start = 0
  let mut i = 0
  while i < len {
    if s.get_char(i).unwrap() == sep {
      result.push(substring(s, start, i))
      start = i + 1
    }
    i = i + 1
  }
  if start <= len {
    result.push(substring(s, start, len))
  }
  result
}

// ── Shape rendering and transform ──────────────────────────────────────────

///|
/// Render a single shape as SVG (0-indexed slide and shape).
/// Returns the shape SVG with `<defs>` prepended, or "ERROR:<message>".
pub fn render_shape_svg(slide_idx : Int, shape_idx : Int) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let scale = get_scale()
    let slide_num = slide_idx + 1
    let rels_path = "ppt/slides/_rels/slide" +
      int_to_str(slide_num) +
      ".xml.rels"
    let rels_xml = @ffi.ffi_get_file(rels_path)
    let rels = if rels_xml.is_empty() {
      []
    } else {
      @ooxml.parse_rels(rels_xml)
    }
    let defs = @renderer.render_shape_defs(shape, slide_num, shape_idx, scale)
    let defs_section = if defs.is_empty() {
      ""
    } else {
      "<defs>" + defs + "</defs>"
    }
    let shape_svg = @renderer.render_shape(
      shape,
      rels,
      slide_num,
      scale,
      shape_idx,
      g_theme.val,
      g_table_styles.val,
    )
    defs_section + shape_svg
  })
}

///|
/// Update a shape's transform (position, size, rotation) and return re-rendered SVG.
/// All values are in EMU.
pub fn update_shape_transform(
  slide_idx : Int,
  shape_idx : Int,
  x : Int,
  y : Int,
  cx : Int,
  cy : Int,
  rot : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    arr[local_idx] = {
      ..shape,
      transform: {
        x,
        y,
        cx,
        cy,
        rot,
        flip_h: shape.transform.flip_h,
        flip_v: shape.transform.flip_v,
      },
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a text run's content and return the shape's re-rendered SVG.
pub fn update_shape_text(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  text : String,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    para.runs[run_idx] = { ..run, text, }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a shape's solid fill color (RGB 0-255) and return re-rendered SVG.
pub fn update_shape_fill(
  slide_idx : Int,
  shape_idx : Int,
  r : Int,
  g : Int,
  b : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    arr[local_idx] = {
      ..arr[local_idx],
      fill: ({ r, g, b, alpha: 255 } : @ooxml.Color),
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

// ── Shape CRUD ─────────────────────────────────────────────────────────────

///|
/// Delete a shape from a slide. Returns "OK" or "ERROR:<message>".
pub fn delete_shape(slide_idx : Int, shape_idx : Int) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    array_remove_at(arr, local_idx)
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Add a basic shape to a slide. Returns "OK:<shape_index>" or "ERROR:<message>".
/// `geom_type`: "rect", "ellipse", "roundRect", "line", etc.
/// Position/size in EMU. fill_r/g/b: -1 = no fill.
pub fn add_shape(
  slide_idx : Int,
  geom_type : String,
  x : Int,
  y : Int,
  cx : Int,
  cy : Int,
  fill_r : Int,
  fill_g : Int,
  fill_b : Int,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide = g_slides.val[slide_idx]
  let transform : @ooxml.ShapeTransform = {
    x,
    y,
    cx,
    cy,
    rot: 0,
    flip_h: false,
    flip_v: false,
  }
  let geom = if geom_type == "ellipse" {
    @ooxml.ShapeGeom::Ellipse
  } else if geom_type == "roundRect" {
    @ooxml.ShapeGeom::RoundRect(@ooxml.round_rect_default_adj)
  } else if geom_type == "line" {
    @ooxml.ShapeGeom::Line
  } else {
    @ooxml.ShapeGeom::Rect
  }
  let kind = @ooxml.ShapeKind::AutoShape(geom)
  let fill = if fill_r >= 0 && fill_g >= 0 && fill_b >= 0 {
    ({ r: fill_r, g: fill_g, b: fill_b, alpha: 255 } : @ooxml.Color)
  } else {
    @ooxml.Color::none()
  }
  let stroke = if geom_type == "line" {
    ({ r: 0, g: 0, b: 0, alpha: 255 } : @ooxml.Color)
  } else {
    @ooxml.Color::none()
  }
  let stroke_w = if geom_type == "line" { 12700 } else { 0 }
  let shape = {
    ..@ooxml.Shape::default_with(kind, transform),
    fill,
    stroke,
    stroke_w,
  }
  slide.shapes.push(shape)
  let new_idx = slide.shapes.length() - 1
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(new_idx)
}

///|
/// Add a Picture shape to a slide. The image binary must already be in the ZIP
/// and referenced via the given relationship ID in the slide's .rels file.
/// Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn add_picture_shape(
  slide_idx : Int,
  rid : String,
  x : Int,
  y : Int,
  cx : Int,
  cy : Int,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide = g_slides.val[slide_idx]
  let transform : @ooxml.ShapeTransform = {
    x,
    y,
    cx,
    cy,
    rot: 0,
    flip_h: false,
    flip_v: false,
  }
  let kind = @ooxml.ShapeKind::Picture(rid)
  let blip_fill : @ooxml.BlipFill = {
    rid,
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
    alpha: 100000,
    svg_rid: "",
    bright: 0,
    contrast: 0,
    duotone_1: @ooxml.Color::none(),
    duotone_2: @ooxml.Color::none(),
    clr_from: @ooxml.Color::none(),
    clr_to: @ooxml.Color::none(),
  }
  let shape = @ooxml.Shape::default_with_blip(kind, transform, blip_fill)
  slide.shapes.push(shape)
  let new_idx = slide.shapes.length() - 1
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(new_idx)
}

///|
/// Replace the image reference of a Picture shape.
/// Updates the blip_fill rid and ShapeKind::Picture(rid).
/// Returns "OK" or "ERROR:<message>".
pub fn replace_picture_rid(
  slide_idx : Int,
  shape_idx : Int,
  new_rid : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    match shape.kind {
      Picture(_) => {
        arr[local_idx] = {
          ..shape,
          kind: @ooxml.ShapeKind::Picture(new_rid),
          blip_fill: { ..shape.blip_fill, rid: new_rid },
        }
        g_modified.val[slide_idx] = true
        "OK"
      }
      _ => "ERROR:shape is not a picture"
    }
  })
}

///|
/// Add a text paragraph with a single run to a shape.
/// Returns "OK:<paraIndex>" or "ERROR:<message>".
pub fn add_shape_text(
  slide_idx : Int,
  shape_idx : Int,
  text : String,
  font_size : Int,
  color_r : Int,
  color_g : Int,
  color_b : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let color : @ooxml.Color = if color_r >= 0 && color_g >= 0 && color_b >= 0 {
      { r: color_r, g: color_g, b: color_b, alpha: 255 }
    } else {
      @ooxml.Color::none()
    }
    let run = make_default_run(text, font_size~, color~)
    let para = make_default_paragraph(run)
    shape.paragraphs.push(para)
    let para_idx = shape.paragraphs.length() - 1
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(para_idx)
  })
}

///|
/// Duplicate a shape on a slide, offset by (dx_emu, dy_emu).
/// Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn duplicate_shape(
  slide_idx : Int,
  shape_idx : Int,
  dx_emu : Int,
  dy_emu : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let src = arr[local_idx]
    let dup = {
      ..src,
      transform: {
        x: src.transform.x + dx_emu,
        y: src.transform.y + dy_emu,
        cx: src.transform.cx,
        cy: src.transform.cy,
        rot: src.transform.rot,
        flip_h: src.transform.flip_h,
        flip_v: src.transform.flip_v,
      },
    }
    arr.push(dup)
    let new_idx = if shape_idx < 1000 {
      g_slides.val[slide_idx].shapes.length() - 1
    } else {
      let parent_composite = shape_idx / 1000
      parent_composite * 1000 + (arr.length() - 1)
    }
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(new_idx)
  })
}

// ── Fill and stroke ────────────────────────────────────────────────────────

///|
/// Update a shape's gradient fill. Returns re-rendered SVG or "ERROR:<message>".
/// `angle`: gradient angle in 60000ths of a degree (e.g. 5400000 = 90deg)
/// `stops_data`: "pos,r,g,b;pos,r,g,b;..." where pos is 0-100000
pub fn update_shape_gradient_fill(
  slide_idx : Int,
  shape_idx : Int,
  angle : Int,
  stops_data : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let stops = parse_gradient_stops(stops_data)
    if stops.is_empty() {
      return "ERROR:invalid gradient stops"
    }
    let grad : @ooxml.GradientFill = {
      stops,
      angle,
      path_type: "",
      rot_with_shape: true,
      fill_to_l: -1,
      fill_to_t: -1,
      fill_to_r: -1,
      fill_to_b: -1,
      tile_flip: "",
    }
    arr[local_idx] = {
      ..arr[local_idx],
      fill: @ooxml.Color::none(),
      grad_fill: grad,
      blip_fill: @ooxml.BlipFill::none(),
      patt_fill: @ooxml.PatternFill::none(),
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a shape's stroke color, width, and dash. Returns re-rendered SVG.
/// r/g/b: -1 = remove stroke. width_emu: 0 = default. dash: "" = solid.
pub fn update_shape_stroke(
  slide_idx : Int,
  shape_idx : Int,
  r : Int,
  g : Int,
  b : Int,
  width_emu : Int,
  dash : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let stroke_color = if r >= 0 && g >= 0 && b >= 0 {
      ({ r, g, b, alpha: 255 } : @ooxml.Color)
    } else {
      @ooxml.Color::none()
    }
    arr[local_idx] = {
      ..arr[local_idx],
      stroke: stroke_color,
      stroke_w: width_emu,
      stroke_dash: dash,
      stroke_no_fill: r < 0,
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

// ── Text editing API (E2.5) ─────────────────────────────────────────────────

///|
/// Add a new paragraph to a shape. Creates a paragraph with a single run.
/// Returns "OK:<paraIndex>" or "ERROR:<message>".
pub fn add_paragraph(
  slide_idx : Int,
  shape_idx : Int,
  text : String,
  align : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let run = make_default_run(text)
    let para = make_default_paragraph(run, align~)
    shape.paragraphs.push(para)
    let para_idx = shape.paragraphs.length() - 1
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(para_idx)
  })
}

///|
/// Delete a paragraph from a shape. Returns "OK" or "ERROR:<message>".
pub fn delete_paragraph(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    if para_idx < 0 || para_idx >= shape.paragraphs.length() {
      return "ERROR:paragraph index out of range"
    }
    array_remove_at(shape.paragraphs, para_idx)
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Add a new text run to a paragraph. Returns "OK:<runIndex>" or "ERROR:<message>".
pub fn add_run(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  text : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    if para_idx < 0 || para_idx >= shape.paragraphs.length() {
      return "ERROR:paragraph index out of range"
    }
    let para = shape.paragraphs[para_idx]
    let run = make_default_run(text)
    para.runs.push(run)
    let run_idx = para.runs.length() - 1
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(run_idx)
  })
}

///|
/// Delete a text run from a paragraph. Returns "OK" or "ERROR:<message>".
pub fn delete_run(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, _run) {
    array_remove_at(para.runs, run_idx)
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Update a text run's bold/italic style.
/// bold/italic: 1 = set true, 0 = set false, -1 = no change.
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_text_run_style(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  bold : Int,
  italic : Int,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    para.runs[run_idx] = {
      ..run,
      bold: if bold >= 0 {
        bold != 0
      } else {
        run.bold
      },
      bold_explicit: if bold >= 0 {
        true
      } else {
        run.bold_explicit
      },
      italic: if italic >= 0 {
        italic != 0
      } else {
        run.italic
      },
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a text run's font size (hundredths of a point, e.g. 1800 = 18pt). 0 = inherit.
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_text_run_font_size(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  font_size : Int,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    para.runs[run_idx] = { ..run, font_size, }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a text run's color (RGB 0-255). Pass r = -1 to clear (inherit).
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_text_run_color(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  r : Int,
  g_ : Int,
  b : Int,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    let color : @ooxml.Color = if r >= 0 && g_ >= 0 && b >= 0 {
      { r, g: g_, b, alpha: 255 }
    } else {
      @ooxml.Color::none()
    }
    para.runs[run_idx] = { ..run, color, }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a text run's font family. Empty string = no change.
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_text_run_font(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  font_face : String,
  ea_font : String,
  cs_font : String,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    para.runs[run_idx] = {
      ..run,
      font_face: if font_face != "" {
        font_face
      } else {
        run.font_face
      },
      ea_font: if ea_font != "" {
        ea_font
      } else {
        run.ea_font
      },
      cs_font: if cs_font != "" {
        cs_font
      } else {
        run.cs_font
      },
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a paragraph's alignment.
/// align: "l" (left), "ctr" (center), "r" (right), "just" (justify), "" (inherit).
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_paragraph_align(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  align : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    if para_idx < 0 || para_idx >= shape.paragraphs.length() {
      return "ERROR:paragraph index out of range"
    }
    shape.paragraphs[para_idx] = { ..shape.paragraphs[para_idx], align, }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Update a text run's decoration (underline, strikethrough, baseline shift).
/// underline: "sng", "dbl", "" (no change), "none" (remove).
/// strike: "sngStrike", "dblStrike", "" (no change), "none" (remove).
/// baseline: 30000 = superscript, -25000 = subscript, 0 = normal, -1 = no change.
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_text_run_decoration(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  underline : String,
  strike : String,
  baseline : Int,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    let new_underline = if underline == "" {
      run.underline
    } else if underline == "none" {
      ""
    } else {
      underline
    }
    let new_strike = if strike == "" {
      run.strike
    } else if strike == "none" {
      ""
    } else {
      strike
    }
    para.runs[run_idx] = {
      ..run,
      underline: new_underline,
      strike: new_strike,
      baseline: if baseline == -1 {
        run.baseline
      } else {
        baseline
      },
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}
