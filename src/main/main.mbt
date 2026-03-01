// pptx-render: Round-trip PPTX ↔ SVG engine
//
// Architecture:
//   JavaScript host:
//     - Parses the ZIP archive (reads local file headers)
//     - Decompresses entries using DecompressionStream API
//     - Exposes decompressed content via FFI (pptx_ffi.get_file, etc.)
//
//   MoonBit (this module):
//     - Receives OOXML XML strings via FFI
//     - Parses XML to build document model (SlideData)
//     - Generates SVG output with data-ooxml-* attributes
//     - Parses SVG back to SlideData (round-trip)
//     - Serializes SlideData to OOXML slide XML

// ── Global state ─────────────────────────────────────────────────────────────
//
// All mutable state is stored in Ref cells. Populated by initialize_pptx().
//
// Slide data:
//   g_slide_count    — total number of slides
//   g_slide_size     — slide dimensions from presentation.xml
//   g_slides         — parsed SlideData cache (0-indexed), updated by round-trip
//   g_modified       — per-slide dirty flag for export
//   g_theme          — parsed theme1.xml color/font scheme
//
// Inheritance chain (slide → layout → master → theme):
//   g_slide_layout_paths — slide idx → layout file path
//   g_layout_paths       — unique layout paths (parallel with g_layouts)
//   g_layouts            — parsed SlideLayoutData + path pairs
//   g_master_paths       — unique master paths (parallel with g_masters)
//   g_masters            — parsed SlideMasterData + path pairs
//   g_layout_to_master   — (layout_path, master_path) mapping

let g_slide_count : Ref[Int] = { val: 0 }
let g_slide_size : Ref[@ooxml.SlideSize] = { val: @ooxml.SlideSize::default() }
let g_slides : Ref[Array[@ooxml.SlideData]] = { val: [] }
let g_modified : Ref[Array[Bool]] = { val: [] }
let g_theme : Ref[@ooxml.ThemeData] = { val: @ooxml.ThemeData::default() }

let g_slide_layout_paths : Ref[Array[String]] = { val: [] }
let g_layout_paths : Ref[Array[String]] = { val: [] }
let g_layouts : Ref[Array[(@ooxml.SlideLayoutData, String)]] = { val: [] }
let g_master_paths : Ref[Array[String]] = { val: [] }
let g_masters : Ref[Array[(@ooxml.SlideMasterData, String)]] = { val: [] }
let g_layout_to_master : Ref[Array[(String, String)]] = { val: [] }

// ── Path resolution helpers ──────────────────────────────────────────────────

///| Resolve a relative target path against a base directory.
///  e.g. resolve_rel_path("ppt/slides", "../slideLayouts/slideLayout1.xml")
///    → "ppt/slideLayouts/slideLayout1.xml"
fn resolve_rel_path(base_dir : String, target : String) -> String {
  if starts_with(target, "/") {
    return suffix(target, 1)
  }
  // Handle ../ prefixes
  let mut dir = base_dir
  let mut tgt = target
  while starts_with(tgt, "../") {
    // Remove trailing "/" from dir if present
    let dlen = dir.length()
    let d2 = if dlen > 0 && dir.get_char(dlen - 1).unwrap() == '/' {
      substring(dir, 0, dlen - 1)
    } else {
      dir
    }
    // Go up one directory
    let slash = last_index_of(d2, '/')
    if slash >= 0 {
      dir = substring(d2, 0, slash)
    } else {
      dir = ""
    }
    tgt = suffix(tgt, 3)
  }
  if dir.is_empty() {
    tgt
  } else {
    dir + "/" + tgt
  }
}

fn last_index_of(s : String, c : Char) -> Int {
  let len = s.length()
  let mut i = len - 1
  while i >= 0 {
    if s.get_char(i).unwrap() == c { return i }
    i = i - 1
  }
  -1
}

fn substring(s : String, from : Int, to : Int) -> String { @xml.str_substring(s, from, to) }

///| Get the directory portion of a path.
fn dir_of(path : String) -> String {
  let slash = last_index_of(path, '/')
  if slash >= 0 { substring(path, 0, slash) } else { "" }
}

///| Find relationship by type suffix in a rels array.
fn find_rel_by_type_suffix(rels : Array[@ooxml.Relationship], type_suffix : String) -> String {
  let n = rels.length()
  let mut i = 0
  while i < n {
    if ends_with(rels[i].type_, type_suffix) {
      return rels[i].target
    }
    i = i + 1
  }
  ""
}

///| Check if an array contains a string.
fn array_contains(arr : Array[String], val : String) -> Bool {
  let n = arr.length()
  let mut i = 0
  while i < n {
    if arr[i] == val { return true }
    i = i + 1
  }
  false
}

///| Find index of a string in an array, or -1.
fn array_index_of(arr : Array[String], val : String) -> Int {
  let n = arr.length()
  let mut i = 0
  while i < n {
    if arr[i] == val { return i }
    i = i + 1
  }
  -1
}

///| Get the .rels path for a given content path.
fn rels_path_for(path : String) -> String {
  let dir = dir_of(path)
  let filename = suffix(path, dir.length() + 1)
  dir + "/_rels/" + filename + ".rels"
}

///| Look up the master for a given layout path.
fn find_master_for_layout(layout_path : String) -> String {
  let pairs = g_layout_to_master.val
  let n = pairs.length()
  let mut i = 0
  while i < n {
    let (lp, mp) = pairs[i]
    if lp == layout_path { return mp }
    i = i + 1
  }
  // Fallback: first master
  if g_master_paths.val.length() > 0 {
    g_master_paths.val[0]
  } else {
    ""
  }
}

///| Look up parsed layout data by path.
fn get_layout(layout_path : String) -> @ooxml.SlideLayoutData? {
  let idx = array_index_of(g_layout_paths.val, layout_path)
  if idx >= 0 {
    let (data, _) = g_layouts.val[idx]
    Some(data)
  } else {
    None
  }
}

///| Look up parsed master data by path.
fn get_master(master_path : String) -> @ooxml.SlideMasterData? {
  let idx = array_index_of(g_master_paths.val, master_path)
  if idx >= 0 {
    let (data, _) = g_masters.val[idx]
    Some(data)
  } else {
    None
  }
}

///| Initialize the PPTX processor.
///
/// Reads presentation.xml via FFI and counts the number of slides.
/// Returns "OK:<count>" on success, "ERROR:<message>" on failure.
pub fn initialize_pptx() -> String {
  @ffi.ffi_log("initialize_pptx: reading ppt/presentation.xml")
  let prs_xml = @ffi.ffi_get_file("ppt/presentation.xml")
  if prs_xml.is_empty() {
    return "ERROR:ppt/presentation.xml not found"
  }
  @ffi.ffi_log("presentation.xml: " + int_to_str(prs_xml.length()) + " chars")
  let count = count_slide_ids(prs_xml)
  g_slide_count.val = count
  g_slide_size.val = @ooxml.parse_slide_size(prs_xml)
  // Load theme from presentation rels
  let prs_rels_xml = @ffi.ffi_get_file("ppt/_rels/presentation.xml.rels")
  if not(prs_rels_xml.is_empty()) {
    let prs_rels = @ooxml.parse_rels(prs_rels_xml)
    let theme_path = find_theme_path(prs_rels)
    if not(theme_path.is_empty()) {
      let theme_xml = @ffi.ffi_get_file(theme_path)
      if not(theme_xml.is_empty()) {
        g_theme.val = @ooxml.parse_theme(theme_xml)
        @ffi.ffi_log("Theme loaded from " + theme_path)
      }
    }
  }
  // ── Resolve slide → layout → master relationships ──
  let slide_layout_paths : Array[String] = []
  let layout_paths : Array[String] = []
  let layout_to_master : Array[(String, String)] = []
  let master_paths : Array[String] = []
  // For each slide, find its layout from the slide's .rels
  let mut si = 0
  while si < count {
    let slide_num = si + 1
    let slide_rels_path = "ppt/slides/_rels/slide" + int_to_str(slide_num) + ".xml.rels"
    let slide_rels_xml = @ffi.ffi_get_file(slide_rels_path)
    let layout_path = if slide_rels_xml.is_empty() {
      ""
    } else {
      let srels = @ooxml.parse_rels(slide_rels_xml)
      let layout_target = find_rel_by_type_suffix(srels, "/slideLayout")
      if layout_target.is_empty() {
        ""
      } else {
        resolve_rel_path("ppt/slides", layout_target)
      }
    }
    slide_layout_paths.push(layout_path)
    if not(layout_path.is_empty()) && not(array_contains(layout_paths, layout_path)) {
      layout_paths.push(layout_path)
    }
    si = si + 1
  }
  // For each unique layout, find its master from the layout's .rels
  let nlay = layout_paths.length()
  let mut li = 0
  while li < nlay {
    let lpath = layout_paths[li]
    let lrels_path = rels_path_for(lpath)
    let lrels_xml = @ffi.ffi_get_file(lrels_path)
    let mpath = if lrels_xml.is_empty() {
      ""
    } else {
      let lrels = @ooxml.parse_rels(lrels_xml)
      let master_target = find_rel_by_type_suffix(lrels, "/slideMaster")
      if master_target.is_empty() {
        ""
      } else {
        resolve_rel_path(dir_of(lpath), master_target)
      }
    }
    layout_to_master.push((lpath, mpath))
    if not(mpath.is_empty()) && not(array_contains(master_paths, mpath)) {
      master_paths.push(mpath)
    }
    li = li + 1
  }
  // Parse all unique masters
  let masters : Array[(@ooxml.SlideMasterData, String)] = []
  let nmas = master_paths.length()
  let mut mi = 0
  while mi < nmas {
    let mp = master_paths[mi]
    let mxml = @ffi.ffi_get_file(mp)
    if not(mxml.is_empty()) {
      let mdata = @ooxml.parse_slide_master(mxml, g_slide_size.val, g_theme.val)
      masters.push((mdata, mp))
      @ffi.ffi_log("Parsed master: " + mp)
    } else {
      let empty_master : @ooxml.SlideMasterData = {
        background: @ooxml.Color::none(),
        shapes: [],
        text_styles: @ooxml.MasterTextStyles::empty(),
      }
      masters.push((empty_master, mp))
    }
    mi = mi + 1
  }
  // Parse all unique layouts
  let layouts : Array[(@ooxml.SlideLayoutData, String)] = []
  let nlay2 = layout_paths.length()
  let mut li2 = 0
  while li2 < nlay2 {
    let lp = layout_paths[li2]
    let lxml = @ffi.ffi_get_file(lp)
    if not(lxml.is_empty()) {
      let ldata = @ooxml.parse_slide_layout(lxml, g_slide_size.val, g_theme.val)
      layouts.push((ldata, lp))
      @ffi.ffi_log("Parsed layout: " + lp)
    } else {
      let empty_layout : @ooxml.SlideLayoutData = {
        background: @ooxml.Color::none(),
        shapes: [],
      }
      layouts.push((empty_layout, lp))
    }
    li2 = li2 + 1
  }
  // Store in globals
  g_slide_layout_paths.val = slide_layout_paths
  g_layout_paths.val = layout_paths
  g_layouts.val = layouts
  g_master_paths.val = master_paths
  g_masters.val = masters
  g_layout_to_master.val = layout_to_master
  // Initialize slide cache
  let slides : Array[@ooxml.SlideData] = []
  let modified : Array[Bool] = []
  let mut i = 0
  while i < count {
    slides.push(
      {
        slide_size: g_slide_size.val,
        background: @ooxml.Color::none(),
        shapes: [],
      },
    )
    modified.push(false)
    i = i + 1
  }
  g_slides.val = slides
  g_modified.val = modified
  @ffi.ffi_log("Slide count: " + int_to_str(count) + ", layouts: " + int_to_str(nlay) + ", masters: " + int_to_str(nmas))
  "OK:" + int_to_str(count)
}

///| Get the number of slides in the loaded presentation.
pub fn get_slide_count() -> Int {
  g_slide_count.val
}

///| Get the raw XML of a slide (0-indexed).
pub fn get_slide_xml_raw(slide_idx : Int) -> String {
  let path = "ppt/slides/slide" + int_to_str(slide_idx + 1) + ".xml"
  @ffi.ffi_log("get_slide_xml_raw: " + path)
  let xml = @ffi.ffi_get_file(path)
  if xml.is_empty() {
    "ERROR:not found: " + path
  } else {
    xml
  }
}

///| Get a newline-separated list of all PPTX entry paths.
pub fn get_entry_list() -> String {
  @ffi.ffi_get_entry_list()
}

///| Render slide `slide_idx` (0-indexed) as an SVG string.
///
/// Returns the SVG markup on success, or "ERROR:<message>" on failure.
pub fn render_slide_svg(slide_idx : Int) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide_num = slide_idx + 1
  // Load slide XML
  let slide_path = "ppt/slides/slide" + int_to_str(slide_num) + ".xml"
  let slide_xml = @ffi.ffi_get_file(slide_path)
  if slide_xml.is_empty() {
    return "ERROR:slide XML not found: " + slide_path
  }
  // Load slide relationships (.rels)
  let rels_path =
    "ppt/slides/_rels/slide" + int_to_str(slide_num) + ".xml.rels"
  let rels_xml = @ffi.ffi_get_file(rels_path)
  let rels = if rels_xml.is_empty() {
    []
  } else {
    @ooxml.parse_rels(rels_xml)
  }
  // Parse slide and cache it
  let slide_data = @ooxml.parse_slide(slide_xml, g_slide_size.val, g_theme.val)
  // ── Resolve layout/master for this slide ──
  let layout_path = if slide_idx < g_slide_layout_paths.val.length() {
    g_slide_layout_paths.val[slide_idx]
  } else {
    ""
  }
  let layout_opt = if not(layout_path.is_empty()) { get_layout(layout_path) } else { None }
  let master_path = if not(layout_path.is_empty()) {
    find_master_for_layout(layout_path)
  } else {
    ""
  }
  let master_opt = if not(master_path.is_empty()) { get_master(master_path) } else { None }
  // ── Placeholder inheritance: resolve transforms from layout/master ──
  resolve_placeholder_inheritance(slide_data.shapes, layout_opt, master_opt)
  // ── Background inheritance: slide → layout → master → white ──
  let mut bg = slide_data.background
  if bg.is_none() {
    match layout_opt {
      Some(layout) => {
        if not(layout.background.is_none()) {
          bg = layout.background
        }
      }
      None => ()
    }
    if bg.is_none() {
      match master_opt {
        Some(master) => {
          if not(master.background.is_none()) {
            bg = master.background
          }
        }
        None => ()
      }
    }
    if bg.is_none() {
      bg = @ooxml.Color::white()
    }
  }
  // ── Apply text style defaults from master ──
  let ns = slide_data.shapes.length()
  let mut shape_i = 0
  while shape_i < ns {
    let shape = slide_data.shapes[shape_i]
    match master_opt {
      Some(master) => apply_text_defaults(shape, master.text_styles)
      None => ()
    }
    shape_i = shape_i + 1
  }
  let resolved_slide : @ooxml.SlideData = {
    slide_size: slide_data.slide_size,
    background: bg,
    shapes: slide_data.shapes,
  }
  g_slides.val[slide_idx] = resolved_slide
  // ── Collect master/layout background shapes (non-placeholder) ──
  let bg_shapes : Array[@ooxml.Shape] = []
  match master_opt {
    Some(master) => {
      let ms = master.shapes.length()
      let mut mi = 0
      while mi < ms {
        let s = master.shapes[mi]
        if s.ph_type.is_empty() {
          bg_shapes.push(s)
        }
        mi = mi + 1
      }
    }
    None => ()
  }
  match layout_opt {
    Some(layout) => {
      let ls = layout.shapes.length()
      let mut li = 0
      while li < ls {
        let s = layout.shapes[li]
        if s.ph_type.is_empty() {
          bg_shapes.push(s)
        }
        li = li + 1
      }
    }
    None => ()
  }
  // Render to SVG
  @renderer.render_slide_svg(resolved_slide, rels, slide_num, g_theme.val, bg_shapes)
}

///| Update a slide's internal data from an edited SVG string.
///
/// Parses the SVG's data-ooxml-* attributes back into SlideData.
/// Returns "OK" on success, "ERROR:<message>" on failure.
pub fn update_slide_from_svg(slide_idx : Int, svg_string : String) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide_data = @svg_parser.parse_svg_to_slide(svg_string)
  g_slides.val[slide_idx] = slide_data
  g_modified.val[slide_idx] = true
  "OK"
}

///| Get the OOXML slide XML for a slide (0-indexed).
///
/// If the slide has been modified via update_slide_from_svg(),
/// returns the serialized modified data. Otherwise returns the
/// original slide XML from the PPTX.
pub fn get_slide_ooxml(slide_idx : Int) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  if g_modified.val[slide_idx] {
    @serializer.serialize_slide(g_slides.val[slide_idx])
  } else {
    // Return original XML
    let slide_num = slide_idx + 1
    let path = "ppt/slides/slide" + int_to_str(slide_num) + ".xml"
    let xml = @ffi.ffi_get_file(path)
    if xml.is_empty() {
      "ERROR:slide XML not found: " + path
    } else {
      xml
    }
  }
}

///| Get tab-separated list of modified entries for ZIP rebuild.
///
/// Format: "path\tcontent\n" for each modified slide.
/// JS uses this to replace entries in the ZIP archive.
pub fn get_modified_entries() -> String {
  let n = g_slide_count.val
  let mut result = ""
  let mut i = 0
  while i < n {
    if g_modified.val[i] {
      let path = "ppt/slides/slide" + int_to_str(i + 1) + ".xml"
      let content = @serializer.serialize_slide(g_slides.val[i])
      result = result + path + "\t" + content + "\n"
    }
    i = i + 1
  }
  result
}

// ── Placeholder inheritance ──────────────────────────────────────────────────

///| Find a matching placeholder shape in a shapes array by ph_type and ph_idx.
fn find_placeholder(shapes : Array[@ooxml.Shape], ph_type : String, ph_idx : Int) -> @ooxml.Shape? {
  let n = shapes.length()
  let mut i = 0
  // First try matching by ph_idx (if set)
  if ph_idx >= 0 {
    while i < n {
      if shapes[i].ph_idx == ph_idx {
        return Some(shapes[i])
      }
      i = i + 1
    }
  }
  // Fallback: match by ph_type
  i = 0
  while i < n {
    if shapes[i].ph_type == ph_type {
      return Some(shapes[i])
    }
    i = i + 1
  }
  None
}

///| Resolve placeholder transforms and fills from layout/master for a slide's shapes.
///  Modifies shapes in-place by replacing zero-size transforms with inherited ones.
fn resolve_placeholder_inheritance(
  shapes : Array[@ooxml.Shape],
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?
) -> Unit {
  let ns = shapes.length()
  let mut i = 0
  while i < ns {
    let shape = shapes[i]
    if shape.ph_type.is_empty() {
      i = i + 1
      continue
    }
    let t = shape.transform
    let needs_transform = t.cx == 0 && t.cy == 0
    if not(needs_transform) {
      i = i + 1
      continue
    }
    // Try layout first, then master
    let mut resolved = false
    match layout_opt {
      Some(layout) => {
        match find_placeholder(layout.shapes, shape.ph_type, shape.ph_idx) {
          Some(layout_ph) => {
            if layout_ph.transform.cx > 0 && layout_ph.transform.cy > 0 {
              shapes[i] = {
                kind: shape.kind,
                transform: layout_ph.transform,
                fill: shape.fill,
                stroke: shape.stroke,
                stroke_w: shape.stroke_w,
                paragraphs: shape.paragraphs,
                body_props: if shape.body_props.anchor.is_empty() && not(layout_ph.body_props.anchor.is_empty()) {
                  layout_ph.body_props
                } else {
                  shape.body_props
                },
                ph_type: shape.ph_type,
                ph_idx: shape.ph_idx,
              }
              resolved = true
            }
          }
          None => ()
        }
      }
      None => ()
    }
    if not(resolved) {
      match master_opt {
        Some(master) => {
          match find_placeholder(master.shapes, shape.ph_type, shape.ph_idx) {
            Some(master_ph) => {
              if master_ph.transform.cx > 0 && master_ph.transform.cy > 0 {
                shapes[i] = {
                  kind: shape.kind,
                  transform: master_ph.transform,
                  fill: shape.fill,
                  stroke: shape.stroke,
                  stroke_w: shape.stroke_w,
                  paragraphs: shape.paragraphs,
                  body_props: if shape.body_props.anchor.is_empty() && not(master_ph.body_props.anchor.is_empty()) {
                    master_ph.body_props
                  } else {
                    shape.body_props
                  },
                  ph_type: shape.ph_type,
                  ph_idx: shape.ph_idx,
                }
              }
            }
            None => ()
          }
        }
        None => ()
      }
    }
    i = i + 1
  }
}

// ── Text style defaults application ─────────────────────────────────────────

///| Apply master text style defaults to a shape's paragraphs based on ph_type.
fn apply_text_defaults(shape : @ooxml.Shape, text_styles : @ooxml.MasterTextStyles) -> Unit {
  let ph = shape.ph_type
  if ph.is_empty() { return }
  let style = if ph == "title" || ph == "ctrTitle" {
    text_styles.title_style
  } else if ph == "body" || ph == "subTitle" || ph == "obj" {
    text_styles.body_style
  } else {
    text_styles.other_style
  }
  let np = shape.paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = shape.paragraphs[pi]
    let lvl = para.level
    let lvl_idx = if lvl >= 0 && lvl < 9 { lvl } else { 0 }
    let defaults = style.levels[lvl_idx]
    // Apply defaults to each run
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      // Apply defaults if any field is unset
      let need_font_size = run.font_size == 0 && defaults.font_size > 0
      let need_color = run.color.is_none() && not(defaults.color.is_none())
      let need_font_face = run.font_face.is_empty() && not(defaults.font_face.is_empty())
      let need_ea_font = run.ea_font.is_empty() && not(defaults.ea_font.is_empty())
      if need_font_size || need_color || need_font_face || need_ea_font {
        let new_run : @ooxml.TextRun = {
          text: run.text,
          bold: run.bold,
          italic: run.italic,
          font_size: if need_font_size { defaults.font_size } else { run.font_size },
          color: if need_color { defaults.color } else { run.color },
          font_face: if need_font_face { defaults.font_face } else { run.font_face },
          ea_font: if need_ea_font { defaults.ea_font } else { run.ea_font },
          underline: run.underline,
          strike: run.strike,
          baseline: run.baseline,
          char_spacing: run.char_spacing,
        }
        para.runs[ri] = new_run
      }
      ri = ri + 1
    }
    pi = pi + 1
  }
}

// ── Theme path resolution ────────────────────────────────────────────────────

///| Find the theme file path from presentation.xml.rels.
fn find_theme_path(rels : Array[@ooxml.Relationship]) -> String {
  let n = rels.length()
  let mut i = 0
  while i < n {
    let rel = rels[i]
    // Theme relationship type ends with "/theme"
    if ends_with(rel.type_, "/theme") {
      // Target is relative to ppt/ directory
      if starts_with(rel.target, "/") {
        return suffix(rel.target, 1)
      } else {
        return "ppt/" + rel.target
      }
    }
    i = i + 1
  }
  // Fallback: try the standard location
  ""
}

fn ends_with(s : String, pat : String) -> Bool { @xml.str_ends_with(s, pat) }

fn starts_with(s : String, pat : String) -> Bool { @xml.str_starts_with(s, pat) }

fn suffix(s : String, from : Int) -> String { @xml.str_suffix(s, from) }

// ── Integer → String conversion ───────────────────────────────────────────────

fn int_to_str(n : Int) -> String { @xml.int_to_str(n) }

// ── Slide counting ────────────────────────────────────────────────────────────

fn count_slide_ids(xml : String) -> Int {
  count_occurrences(xml, "<p:sldId ") + count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") + count_occurrences(xml, "<p:sldId/>")
}

fn count_occurrences(text : String, pattern : String) -> Int {
  let plen = pattern.length()
  let tlen = text.length()
  if plen == 0 || tlen < plen {
    return 0
  }
  let mut count = 0
  let mut i = 0
  while i <= tlen - plen {
    if match_at(text, pattern, i) {
      count = count + 1
      i = i + plen
    } else {
      i = i + 1
    }
  }
  count
}

fn match_at(text : String, pattern : String, pos : Int) -> Bool {
  let plen = pattern.length()
  for j = 0; j < plen; j = j + 1 {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}
