///|
/// Composite shape-index radix: a group child is `parent * group_index_base + local`.
let group_index_base : Int32 = 1000

///|
/// The array holding composite shape index `idx`, and its local index there.
fn resolve_shape_location(
  slide : @ooxml.SlideData,
  composite_idx : Int32,
) -> (Array[@ooxml.Shape], Int32)? {
  if composite_idx < group_index_base {
    if composite_idx >= 0 && composite_idx < slide.shapes.length() {
      Some((slide.shapes, composite_idx))
    } else {
      None
    }
  } else {
    let parent_composite = composite_idx / group_index_base
    let child_local = composite_idx % group_index_base
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
/// Parse a slide into `g_slides` if it hasn't been yet. Until its first render a slide
/// is an empty placeholder, so an edit would silently find no shapes.
fn ensure_slide_parsed(slide_idx : Int32) -> Unit {
  let n = g_slide_count.val
  if n == 0 || slide_idx < 0 || slide_idx >= n {
    return
  }
  if g_modified.val[slide_idx] || g_parsed.val[slide_idx] {
    return
  }
  // Rendering parses and caches the slide; the SVG itself is not needed.
  let _ = render_slide_svg(slide_idx)
}

///|
/// Call `f` with the shape's containing array and local index, or return an error.
fn with_shape(
  slide_idx : Int32,
  shape_idx : Int32,
  f : (Array[@ooxml.Shape], Int32) -> String,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  ensure_slide_parsed(slide_idx)
  let slide = g_slides.val[slide_idx]
  match resolve_shape_location(slide, shape_idx) {
    None => "ERROR:shape index out of range"
    Some((arr, local_idx)) => f(arr, local_idx)
  }
}

///|
/// Call `f` with the addressed shape, paragraph and run, or return an error.
fn with_run(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
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
  font_size~ : Int32 = 0,
  color~ : @ooxml.Color = @ooxml.Color::none(),
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
    underline_color: @ooxml.Color::none(),
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
  align~ : String = "",
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
/// `shape` with p:cNvPr ids above every id already on `slide`, for a shape being added.
fn with_new_nv_ids(
  slide : @ooxml.SlideData,
  shape : @ooxml.Shape,
) -> @ooxml.Shape {
  @ooxml.with_fresh_nv_ids(shape, { val: @ooxml.max_nv_id(slide.shapes), })
}

///|
fn parse_gradient_stops(s : String) -> Array[@ooxml.GradientStop] {
  let stops : Array[@ooxml.GradientStop] = []
  if s.is_empty() {
    return stops
  }
  let parts = split_string(s, ';')
  for part in parts {
    if !part.is_empty() {
      let vals = split_string(part, ',')
      if vals.length() >= 4 {
        let pos = @xml.parse_int(vals[0])
        let r = @xml.parse_int(vals[1])
        let g = @xml.parse_int(vals[2])
        let b = @xml.parse_int(vals[3])
        stops.push({ pos, color: ({ r, g, b, alpha: 255, } : @ooxml.Color), })
      }
    }
  }
  stops
}

///|
fn split_string(s : String, sep : Char) -> Array[String] {
  let result : Array[String] = []
  let len = s.length()
  let mut start = 0
  for i in 0..<len {
    if s.get_char(i).unwrap() == sep {
      result.push(substring(s, start, i))
      start = i + 1
    }
  }
  if start <= len {
    result.push(substring(s, start, len))
  }
  result
}

///|
/// Render a single shape as SVG (0-indexed slide and shape).
/// Returns the shape SVG with `<defs>` prepended, or "ERROR:<message>".
pub fn render_shape_svg(slide_idx : Int32, shape_idx : Int32) -> String {
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
/// Set a shape's position, size and rotation (EMU); returns the re-rendered shape SVG.
pub fn update_shape_transform(
  slide_idx : Int32,
  shape_idx : Int32,
  x : Int32,
  y : Int32,
  cx : Int32,
  cy : Int32,
  rot : Int32,
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
/// Set several transforms from `"shapeIdx,x,y,cx,cy,rot;..."` (EMU, rot in 1/60000°). All
/// indices are validated before anything changes. Returns "OK:<count>" or "ERROR:...".
pub fn update_shapes_transform(
  slide_idx : Int32,
  items_data : String,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  ensure_slide_parsed(slide_idx)
  let slide = g_slides.val[slide_idx]
  let targets : Array[(Array[@ooxml.Shape], Int32, @ooxml.ShapeTransform)] = []
  let parts = split_string(items_data, ';')
  let np = parts.length()
  let mut i = 0
  while i < np {
    let part = parts[i]
    i = i + 1
    if part.is_empty() {
      continue
    }
    let vals = split_string(part, ',')
    if vals.length() < 6 {
      return "ERROR:invalid transform item: " + part
    }
    let s_idx = @xml.parse_int(vals[0])
    match resolve_shape_location(slide, s_idx) {
      None => return "ERROR:shape index out of range: " + vals[0]
      Some((arr, local_idx)) => {
        let cur = arr[local_idx]
        let t : @ooxml.ShapeTransform = {
          x: @xml.parse_int(vals[1]),
          y: @xml.parse_int(vals[2]),
          cx: @xml.parse_int(vals[3]),
          cy: @xml.parse_int(vals[4]),
          rot: @xml.parse_int(vals[5]),
          flip_h: cur.transform.flip_h,
          flip_v: cur.transform.flip_v,
        }
        targets.push((arr, local_idx, t))
      }
    }
  }
  if targets.is_empty() {
    return "ERROR:no transform items"
  }
  // Transform edits never change indices, so applying after validation is safe.
  let nt = targets.length()
  for j in 0..<nt {
    let (arr, local_idx, t) = targets[j]
    arr[local_idx] = { ..arr[local_idx], transform: t, }
  }
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(nt)
}

///|
/// Update a text run's content and return the shape's re-rendered SVG.
pub fn update_shape_text(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  r : Int32,
  g : Int32,
  b : Int32,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    arr[local_idx] = {
      ..arr[local_idx],
      fill: ({ r, g, b, alpha: 255, } : @ooxml.Color),
    }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Delete a shape from a slide. Returns "OK" or "ERROR:<message>".
pub fn delete_shape(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    arr.remove(local_idx) |> ignore
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Add a preset shape (`"rect"`, `"ellipse"`, …) at an EMU position; fill -1 = none.
/// Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn add_shape(
  slide_idx : Int32,
  geom_type : String,
  x : Int32,
  y : Int32,
  cx : Int32,
  cy : Int32,
  fill_r : Int32,
  fill_g : Int32,
  fill_b : Int32,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  ensure_slide_parsed(slide_idx)
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
    ({ r: fill_r, g: fill_g, b: fill_b, alpha: 255, } : @ooxml.Color)
  } else {
    @ooxml.Color::none()
  }
  let stroke = if geom_type == "line" {
    ({ r: 0, g: 0, b: 0, alpha: 255, } : @ooxml.Color)
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
  slide.shapes.push(with_new_nv_ids(slide, shape))
  let new_idx = slide.shapes.length() - 1
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(new_idx)
}

///|
/// Add a picture whose image the slide's .rels already references as `rid`.
/// Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn add_picture_shape(
  slide_idx : Int32,
  rid : String,
  x : Int32,
  y : Int32,
  cx : Int32,
  cy : Int32,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  ensure_slide_parsed(slide_idx)
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
  slide.shapes.push(with_new_nv_ids(slide, shape))
  let new_idx = slide.shapes.length() - 1
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(new_idx)
}

///|
/// Point a picture at another image rId. Returns "OK" or "ERROR:<message>".
pub fn replace_picture_rid(
  slide_idx : Int32,
  shape_idx : Int32,
  new_rid : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    match shape.kind {
      Picture(_) => {
        arr[local_idx] = {
          ..shape,
          kind: @ooxml.ShapeKind::Picture(new_rid),
          blip_fill: { ..shape.blip_fill, rid: new_rid, },
        }
        g_modified.val[slide_idx] = true
        "OK"
      }
      _ => "ERROR:shape is not a picture"
    }
  })
}

///|
/// Add a one-run paragraph to a shape. Returns "OK:<paraIndex>" or "ERROR:<message>".
pub fn add_shape_text(
  slide_idx : Int32,
  shape_idx : Int32,
  text : String,
  font_size : Int32,
  color_r : Int32,
  color_g : Int32,
  color_b : Int32,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let color : @ooxml.Color = if color_r >= 0 && color_g >= 0 && color_b >= 0 {
      { r: color_r, g: color_g, b: color_b, alpha: 255, }
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
/// Duplicate a shape, offset by (dx_emu, dy_emu). Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn duplicate_shape(
  slide_idx : Int32,
  shape_idx : Int32,
  dx_emu : Int32,
  dy_emu : Int32,
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
    arr.push(with_new_nv_ids(g_slides.val[slide_idx], dup))
    let new_idx = if shape_idx < group_index_base {
      g_slides.val[slide_idx].shapes.length() - 1
    } else {
      let parent_composite = shape_idx / group_index_base
      parent_composite * group_index_base + (arr.length() - 1)
    }
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(new_idx)
  })
}

///|
/// Set a linear gradient: `angle` in 60000ths of a degree, `stops_data` `"pos,r,g,b;..."`
/// (pos 0-100000). Returns the re-rendered SVG or "ERROR:<message>".
pub fn update_shape_gradient_fill(
  slide_idx : Int32,
  shape_idx : Int32,
  angle : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  r : Int32,
  g : Int32,
  b : Int32,
  width_emu : Int32,
  dash : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let stroke_color = if r >= 0 && g >= 0 && b >= 0 {
      ({ r, g, b, alpha: 255, } : @ooxml.Color)
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

///|
/// Add a one-run paragraph. Returns "OK:<paraIndex>" or "ERROR:<message>".
pub fn add_paragraph(
  slide_idx : Int32,
  shape_idx : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    if para_idx < 0 || para_idx >= shape.paragraphs.length() {
      return "ERROR:paragraph index out of range"
    }
    shape.paragraphs.remove(para_idx) |> ignore
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Add a new text run to a paragraph. Returns "OK:<runIndex>" or "ERROR:<message>".
pub fn add_run(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, _run) {
    para.runs.remove(run_idx) |> ignore
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Set bold / italic (1 = on, 0 = off, -1 = unchanged). Returns the shape SVG or "ERROR:...".
pub fn update_text_run_style(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
  bold : Int32,
  italic : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
  font_size : Int32,
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
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
  r : Int32,
  g_ : Int32,
  b : Int32,
) -> String {
  with_run(slide_idx, shape_idx, para_idx, run_idx, fn(_shape, para, run) {
    let color : @ooxml.Color = if r >= 0 && g_ >= 0 && b >= 0 {
      { r, g: g_, b, alpha: 255, }
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
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
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
/// Set alignment: "l", "ctr", "r", "just", or "" (inherit). Returns the shape SVG or "ERROR:...".
pub fn update_paragraph_align(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
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
/// Set underline ("sng"/"dbl") and strike ("sngStrike"/"dblStrike") — "" unchanged, "none"
/// removes — and baseline (30000 super, -25000 sub, 0 normal, -1 unchanged).
/// Returns the shape SVG or "ERROR:...".
pub fn update_text_run_decoration(
  slide_idx : Int32,
  shape_idx : Int32,
  para_idx : Int32,
  run_idx : Int32,
  underline : String,
  strike : String,
  baseline : Int32,
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

///|
/// Restore a slide from OOXML captured by `get_slide_ooxml` (Undo/Redo). Runs the full
/// `resolve_slide_data` pipeline, which is idempotent on serialized slides.
/// Returns "OK" or "ERROR:<message>".
pub fn restore_slide_ooxml(slide_idx : Int32, xml : String) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  if xml.is_empty() {
    return "ERROR:empty xml"
  }
  let slide_num = slide_idx + 1
  let ctx = slide_context(slide_idx)
  let (slide_data, bg_rels) = resolve_slide_data(
    slide_num,
    xml,
    ctx.rels,
    ctx.layout_opt,
    ctx.master_opt,
    ctx.layout_path,
    ctx.master_path,
    ctx.base_theme,
  )
  g_slides.val[slide_idx] = slide_data
  g_modified.val[slide_idx] = true
  g_parsed.val[slide_idx] = true
  g_bg_rels.val[slide_idx] = bg_rels
  "OK"
}

// z-order is array order (last = front-most), within the shape's containing array.

///|
/// Composite index for a sibling at `local_idx` within the same container as `shape_idx`.
fn sibling_composite_idx(shape_idx : Int32, local_idx : Int32) -> Int32 {
  let base = if shape_idx < group_index_base {
    0
  } else {
    shape_idx / group_index_base * group_index_base
  }
  base + local_idx
}

///|
/// Move a shape to the front. Returns "OK:<new_composite_idx>" or "ERROR:<message>".
pub fn bring_to_front(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let sh = arr[local_idx]
    arr.remove(local_idx) |> ignore
    arr.push(sh)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(sibling_composite_idx(shape_idx, arr.length() - 1))
  })
}

///|
/// Move a shape to the back. Returns "OK:<new_composite_idx>" or "ERROR:<message>".
pub fn send_to_back(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let sh = arr[local_idx]
    arr.remove(local_idx) |> ignore
    arr.insert(0, sh)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(sibling_composite_idx(shape_idx, 0))
  })
}

///|
/// Move a shape one step forward (no-op at the front). Returns "OK:<new_composite_idx>".
pub fn bring_forward(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let new_local = if local_idx < arr.length() - 1 {
      let sh = arr[local_idx]
      arr[local_idx] = arr[local_idx + 1]
      arr[local_idx + 1] = sh
      g_modified.val[slide_idx] = true
      local_idx + 1
    } else {
      local_idx
    }
    "OK:" + int_to_str(sibling_composite_idx(shape_idx, new_local))
  })
}

///|
/// Move a shape one step backward (no-op at the back). Returns "OK:<new_composite_idx>".
pub fn send_backward(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let new_local = if local_idx > 0 {
      let sh = arr[local_idx]
      arr[local_idx] = arr[local_idx - 1]
      arr[local_idx - 1] = sh
      g_modified.val[slide_idx] = true
      local_idx - 1
    } else {
      local_idx
    }
    "OK:" + int_to_str(sibling_composite_idx(shape_idx, new_local))
  })
}

///|
/// Serialize a single shape to its OOXML fragment (for copy). Returns the XML, or
/// "ERROR:<message>" (e.g. charts, whose chart part a paste cannot carry over).
pub fn get_shape_ooxml(slide_idx : Int32, shape_idx : Int32) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    // A paste cannot carry a chart's chart part and embedded workbook.
    if arr[local_idx].kind is @ooxml.ShapeKind::ChartShape(_) {
      return "ERROR:charts are not copyable"
    }
    let xml = @serializer.serialize_shape(arr[local_idx])
    if xml.is_empty() {
      "ERROR:shape not serializable"
    } else {
      xml
    }
  })
}

///|
/// Append a shape OOXML fragment, offset by (dx, dy) EMU, for paste; its rIds must
/// already point at this slide's rels. Returns "OK:<shape_index>" or "ERROR:<message>".
pub fn add_shape_from_ooxml(
  slide_idx : Int32,
  shape_xml : String,
  dx_emu : Int32,
  dy_emu : Int32,
) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  ensure_slide_parsed(slide_idx)
  let shapes = @ooxml.parse_shapes_xml(shape_xml, g_theme.val)
  if shapes.is_empty() {
    return "ERROR:no shape parsed from spec"
  }
  let src = shapes[0]
  let pasted = {
    ..src,
    transform: {
      ..src.transform,
      x: src.transform.x + dx_emu,
      y: src.transform.y + dy_emu,
    },
  }
  let slide = g_slides.val[slide_idx]
  slide.shapes.push(with_new_nv_ids(slide, pasted))
  g_modified.val[slide_idx] = true
  "OK:" + int_to_str(slide.shapes.length() - 1)
}
