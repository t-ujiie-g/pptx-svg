// pptx-svg: Round-trip PPTX ↔ SVG engine
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

///|
let g_slide_count : Ref[Int] = { val: 0 }

///|
let g_slide_size : Ref[@ooxml.SlideSize] = { val: @ooxml.SlideSize::default() }

///|
let g_slides : Ref[Array[@ooxml.SlideData]] = { val: [] }

///|
let g_modified : Ref[Array[Bool]] = { val: [] }

///|
let g_theme : Ref[@ooxml.ThemeData] = { val: @ooxml.ThemeData::default() }

///|
let g_slide_layout_paths : Ref[Array[String]] = { val: [] }

///|
let g_layout_paths : Ref[Array[String]] = { val: [] }

///|
let g_layouts : Ref[Array[(@ooxml.SlideLayoutData, String)]] = { val: [] }

///|
let g_master_paths : Ref[Array[String]] = { val: [] }

///|
let g_masters : Ref[Array[(@ooxml.SlideMasterData, String)]] = { val: [] }

///|
let g_layout_to_master : Ref[Array[(String, String)]] = { val: [] }

///|
let g_layout_rels : Ref[Array[Array[@ooxml.Relationship]]] = { val: [] }

///|
let g_master_rels : Ref[Array[Array[@ooxml.Relationship]]] = { val: [] }

///|
let g_master_themes : Ref[Array[@ooxml.ThemeData]] = { val: [] }

///|
let g_table_styles : Ref[Array[@ooxml.TableStyleDef]] = { val: [] }

// ── Package-local string/int helpers ─────────────────────────────────────────

///|
fn ends_with(s : String, pat : String) -> Bool {
  @xml.str_ends_with(s, pat)
}

///|
fn starts_with(s : String, pat : String) -> Bool {
  @xml.str_starts_with(s, pat)
}

///|
fn suffix(s : String, from : Int) -> String {
  @xml.str_suffix(s, from)
}

///|
fn int_to_str(n : Int) -> String {
  @xml.int_to_str(n)
}

// ── Slide counting ──────────────────────────────────────────────────────────

///|
fn count_slide_ids(xml : String) -> Int {
  count_occurrences(xml, "<p:sldId ") +
  count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") +
  count_occurrences(xml, "<p:sldId/>")
}

///|
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

///|
fn match_at(text : String, pattern : String, pos : Int) -> Bool {
  let plen = pattern.length()
  for j = 0; j < plen; j = j + 1 {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}

// ── Theme path resolution ────────────────────────────────────────────────────

///|
/// Find the theme file path from presentation.xml.rels.
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
  ""
}

///|
/// Find and parse the theme for a slide master from its relationships.
/// Falls back to g_theme if no theme relationship is found.
fn load_master_theme(
  master_path : String,
  mrels : Array[@ooxml.Relationship],
) -> @ooxml.ThemeData {
  // Find theme rel in master rels
  let n = mrels.length()
  let mut i = 0
  while i < n {
    let rel = mrels[i]
    if ends_with(rel.type_, "/theme") {
      // Resolve target relative to master's directory
      let slash = last_index_of(master_path, '/')
      let base_dir = if slash >= 0 {
        substring(master_path, 0, slash)
      } else {
        ""
      }
      let theme_path = resolve_rel_path(base_dir, rel.target)
      let theme_xml = @ffi.ffi_get_file(theme_path)
      if not(theme_xml.is_empty()) {
        @ffi.ffi_log("Loaded per-master theme: " + theme_path)
        return @ooxml.parse_theme(theme_xml)
      }
    }
    i = i + 1
  }
  g_theme.val
}

///|
/// Get the theme for a given master path from g_master_themes.
fn get_master_theme(master_path : String) -> @ooxml.ThemeData {
  let idx = array_index_of(g_master_paths.val, master_path)
  if idx >= 0 && idx < g_master_themes.val.length() {
    g_master_themes.val[idx]
  } else {
    g_theme.val
  }
}

// ── Path resolution helpers ──────────────────────────────────────────────────

///|
/// Resolve a relative target path against a base directory.
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

///|
/// Walk shapes to find ChartShape placeholders (with rid stored in chart_xml)
/// and resolve them by loading the actual chart XML file.
fn resolve_chart_shapes(
  shapes : Array[@ooxml.Shape],
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData,
) -> Unit {
  let n = shapes.length()
  let mut i = 0
  while i < n {
    match shapes[i].kind {
      @ooxml.ShapeKind::ChartShape(cd) =>
        // chart_xml field temporarily holds the relationship ID
        if not(cd.chart_xml.is_empty()) && cd.groups.is_empty() {
          let rid = cd.chart_xml
          let target = @ooxml.find_rel_target(rels, rid)
          if not(target.is_empty()) {
            let chart_path = resolve_rel_path("ppt/slides", target)
            let chart_xml = @ffi.ffi_get_file(chart_path)
            if not(chart_xml.is_empty()) {
              let parsed = @ooxml.parse_chart_xml(chart_xml, theme)
              // Replace the shape in-place
              shapes[i] = { ..shapes[i], kind: @ooxml.ShapeKind::ChartShape(parsed) }
            }
          }
        }
      @ooxml.ShapeKind::GroupShape(grp) =>
        // Recurse into group children
        resolve_chart_shapes(grp.children, rels, theme)
      _ => ()
    }
    i = i + 1
  }
}

///|
fn last_index_of(s : String, c : Char) -> Int {
  let len = s.length()
  let mut i = len - 1
  while i >= 0 {
    if s.get_char(i).unwrap() == c {
      return i
    }
    i = i - 1
  }
  -1
}

///|
fn substring(s : String, from : Int, to : Int) -> String {
  @xml.str_substring(s, from, to)
}

///|
/// Get the directory portion of a path.
fn dir_of(path : String) -> String {
  let slash = last_index_of(path, '/')
  if slash >= 0 {
    substring(path, 0, slash)
  } else {
    ""
  }
}

///|
/// Find relationship by type suffix in a rels array.
fn find_rel_by_type_suffix(
  rels : Array[@ooxml.Relationship],
  type_suffix : String,
) -> String {
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

///|
/// Check if an array contains a string.
fn array_contains(arr : Array[String], val : String) -> Bool {
  let n = arr.length()
  let mut i = 0
  while i < n {
    if arr[i] == val {
      return true
    }
    i = i + 1
  }
  false
}

///|
/// Find index of a string in an array, or -1.
fn array_index_of(arr : Array[String], val : String) -> Int {
  let n = arr.length()
  let mut i = 0
  while i < n {
    if arr[i] == val {
      return i
    }
    i = i + 1
  }
  -1
}

///|
/// Get the .rels path for a given content path.
fn rels_path_for(path : String) -> String {
  let dir = dir_of(path)
  let filename = suffix(path, dir.length() + 1)
  dir + "/_rels/" + filename + ".rels"
}

///|
/// Look up the master for a given layout path.
fn find_master_for_layout(layout_path : String) -> String {
  let pairs = g_layout_to_master.val
  let n = pairs.length()
  let mut i = 0
  while i < n {
    let (lp, mp) = pairs[i]
    if lp == layout_path {
      return mp
    }
    i = i + 1
  }
  // Fallback: first master
  if g_master_paths.val.length() > 0 {
    g_master_paths.val[0]
  } else {
    ""
  }
}

///|
/// Look up parsed layout data by path.
fn get_layout(layout_path : String) -> @ooxml.SlideLayoutData? {
  let idx = array_index_of(g_layout_paths.val, layout_path)
  if idx >= 0 {
    let (data, _) = g_layouts.val[idx]
    Some(data)
  } else {
    None
  }
}

///|
/// Look up parsed layout rels by path.
fn get_layout_rels(layout_path : String) -> Array[@ooxml.Relationship] {
  let idx = array_index_of(g_layout_paths.val, layout_path)
  if idx >= 0 && idx < g_layout_rels.val.length() {
    g_layout_rels.val[idx]
  } else {
    []
  }
}

///|
/// Look up parsed master data by path.
fn get_master(master_path : String) -> @ooxml.SlideMasterData? {
  let idx = array_index_of(g_master_paths.val, master_path)
  if idx >= 0 {
    let (data, _) = g_masters.val[idx]
    Some(data)
  } else {
    None
  }
}

///|
/// Look up parsed master rels by path.
fn get_master_rels(master_path : String) -> Array[@ooxml.Relationship] {
  let idx = array_index_of(g_master_paths.val, master_path)
  if idx >= 0 && idx < g_master_rels.val.length() {
    g_master_rels.val[idx]
  } else {
    []
  }
}

///|
/// Find the owning master path for a given layout path.
fn find_master_path_for_layout(
  layout_to_master : Array[(String, String)],
  layout_path : String,
) -> String {
  let n = layout_to_master.length()
  let mut i = 0
  while i < n {
    if layout_to_master[i].0 == layout_path {
      return layout_to_master[i].1
    }
    i = i + 1
  }
  ""
}

///|
/// Find the owning master's ColorMap for a given layout path.
fn find_master_clr_map(
  layout_to_master : Array[(String, String)],
  masters : Array[(@ooxml.SlideMasterData, String)],
  layout_path : String,
) -> @ooxml.ColorMap {
  let master_path = find_master_path_for_layout(layout_to_master, layout_path)
  let mlen = masters.length()
  let mut mi = 0
  while mi < mlen {
    if masters[mi].1 == master_path {
      return masters[mi].0.clr_map
    }
    mi = mi + 1
  }
  @ooxml.ColorMap::default()
}

///|
/// Find the owning master's theme for a given layout path.
fn find_master_theme_for_layout(
  layout_to_master : Array[(String, String)],
  master_themes : Array[@ooxml.ThemeData],
  master_paths : Array[String],
  layout_path : String,
) -> @ooxml.ThemeData {
  let master_path = find_master_path_for_layout(layout_to_master, layout_path)
  let n = master_paths.length()
  let mut i = 0
  while i < n {
    if master_paths[i] == master_path {
      return master_themes[i]
    }
    i = i + 1
  }
  g_theme.val
}

///|
/// Initialize the PPTX processor.
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
  // ── Parse table styles from ppt/tableStyles.xml ──
  let tbl_styles_xml = @ffi.ffi_get_file("ppt/tableStyles.xml")
  if not(tbl_styles_xml.is_empty()) {
    g_table_styles.val = @ooxml.parse_table_styles(tbl_styles_xml, g_theme.val)
    if g_table_styles.val.length() > 0 {
      @ffi.ffi_log(
        "Table styles loaded: " + int_to_str(g_table_styles.val.length()),
      )
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
    let slide_rels_path = "ppt/slides/_rels/slide" +
      int_to_str(slide_num) +
      ".xml.rels"
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
    if not(layout_path.is_empty()) &&
      not(array_contains(layout_paths, layout_path)) {
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
  let master_rels_arr : Array[Array[@ooxml.Relationship]] = []
  let master_themes : Array[@ooxml.ThemeData] = []
  let nmas = master_paths.length()
  let mut mi = 0
  while mi < nmas {
    let mp = master_paths[mi]
    let mxml = @ffi.ffi_get_file(mp)
    // Parse master rels
    let mrels_path = rels_path_for(mp)
    let mrels_xml = @ffi.ffi_get_file(mrels_path)
    let mrels = if mrels_xml.is_empty() {
      ([] : Array[@ooxml.Relationship])
    } else {
      @ooxml.parse_rels(mrels_xml)
    }
    // Load per-master theme (falls back to g_theme)
    let m_theme = load_master_theme(mp, mrels)
    master_themes.push(m_theme)
    if not(mxml.is_empty()) {
      let mdata = @ooxml.parse_slide_master(mxml, m_theme)
      masters.push((mdata, mp))
      master_rels_arr.push(mrels)
      @ffi.ffi_log("Parsed master: " + mp)
    } else {
      let empty_master : @ooxml.SlideMasterData = {
        background: @ooxml.Color::none(),
        bg_grad: @ooxml.GradientFill::none(),
        bg_blip_fill: @ooxml.BlipFill::none(),
        bg_patt_fill: @ooxml.PatternFill::none(),
        shapes: [],
        text_styles: @ooxml.MasterTextStyles::empty(),
        clr_map: @ooxml.ColorMap::default(),
      }
      masters.push((empty_master, mp))
      master_rels_arr.push(mrels)
    }
    mi = mi + 1
  }
  // Parse all unique layouts
  let layouts : Array[(@ooxml.SlideLayoutData, String)] = []
  let layout_rels_arr : Array[Array[@ooxml.Relationship]] = []
  let nlay2 = layout_paths.length()
  let mut li2 = 0
  while li2 < nlay2 {
    let lp = layout_paths[li2]
    let lxml = @ffi.ffi_get_file(lp)
    // Parse layout rels (needed for bg_shapes image resolution)
    let layout_rels_path = rels_path_for(lp)
    let layout_rels_xml = @ffi.ffi_get_file(layout_rels_path)
    let cur_layout_rels = if layout_rels_xml.is_empty() {
      ([] : Array[@ooxml.Relationship])
    } else {
      @ooxml.parse_rels(layout_rels_xml)
    }
    if not(lxml.is_empty()) {
      // Compute effective theme: layout's own clrMapOvr > master's clrMap
      let layout_effective_theme = {
        // Use the master's own theme (not g_theme) for color map application
        let base_theme = find_master_theme_for_layout(
          layout_to_master, master_themes, master_paths, lp,
        )
        let master_clr_map = find_master_clr_map(layout_to_master, masters, lp)
        let lxml_nodes = @xml.parse_xml(lxml)
        let layout_root_ch = match @xml.find_child(lxml_nodes, "p:sldLayout") {
          Some(n) => @xml.get_children(n)
          None => []
        }
        match @ooxml.extract_clr_map_ovr_from_nodes(layout_root_ch) {
          Some(ovr) => @ooxml.apply_color_map(base_theme, ovr)
          None => @ooxml.apply_color_map(base_theme, master_clr_map)
        }
      }
      let ldata = @ooxml.parse_slide_layout(lxml, layout_effective_theme)
      layouts.push((ldata, lp))
      layout_rels_arr.push(cur_layout_rels)
      @ffi.ffi_log("Parsed layout: " + lp)
    } else {
      let empty_layout : @ooxml.SlideLayoutData = {
        background: @ooxml.Color::none(),
        bg_grad: @ooxml.GradientFill::none(),
        bg_blip_fill: @ooxml.BlipFill::none(),
        bg_patt_fill: @ooxml.PatternFill::none(),
        shapes: [],
        clr_map_ovr: None,
      }
      layouts.push((empty_layout, lp))
      layout_rels_arr.push(cur_layout_rels)
    }
    li2 = li2 + 1
  }
  // Store in globals
  g_slide_layout_paths.val = slide_layout_paths
  g_layout_paths.val = layout_paths
  g_layouts.val = layouts
  g_layout_rels.val = layout_rels_arr
  g_master_paths.val = master_paths
  g_masters.val = masters
  g_master_rels.val = master_rels_arr
  g_master_themes.val = master_themes
  g_layout_to_master.val = layout_to_master
  // Initialize slide cache
  let slides : Array[@ooxml.SlideData] = []
  let modified : Array[Bool] = []
  let mut i = 0
  while i < count {
    slides.push({
      slide_size: g_slide_size.val,
      background: @ooxml.Color::none(),
      bg_grad: @ooxml.GradientFill::none(),
      bg_blip_fill: @ooxml.BlipFill::none(),
      bg_patt_fill: @ooxml.PatternFill::none(),
      shapes: [],
    })
    modified.push(false)
    i = i + 1
  }
  g_slides.val = slides
  g_modified.val = modified
  @ffi.ffi_log(
    "Slide count: " +
    int_to_str(count) +
    ", layouts: " +
    int_to_str(nlay) +
    ", masters: " +
    int_to_str(nmas),
  )
  "OK:" + int_to_str(count)
}

///|
/// Get the number of slides in the loaded presentation.
pub fn get_slide_count() -> Int {
  g_slide_count.val
}

///|
/// Get the raw XML of a slide (0-indexed).
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

///|
/// Get a newline-separated list of all PPTX entry paths.
pub fn get_entry_list() -> String {
  @ffi.ffi_get_entry_list()
}

///|
/// Render slide `slide_idx` (0-indexed) as an SVG string.
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
  // Load slide relationships (.rels) — needed for both fresh and cached paths
  let rels_path = "ppt/slides/_rels/slide" + int_to_str(slide_num) + ".xml.rels"
  let rels_xml = @ffi.ffi_get_file(rels_path)
  let rels = if rels_xml.is_empty() { [] } else { @ooxml.parse_rels(rels_xml) }
  // ── Resolve layout/master for this slide ──
  let layout_path = if slide_idx < g_slide_layout_paths.val.length() {
    g_slide_layout_paths.val[slide_idx]
  } else {
    ""
  }
  let layout_opt = if not(layout_path.is_empty()) {
    get_layout(layout_path)
  } else {
    None
  }
  let master_path = if not(layout_path.is_empty()) {
    find_master_for_layout(layout_path)
  } else {
    ""
  }
  let master_opt = if not(master_path.is_empty()) {
    get_master(master_path)
  } else {
    None
  }
  // ── Compute effective theme with color map overrides ──
  // Use the master's own theme (not g_theme) for correct per-master colors
  let base_theme = if not(master_path.is_empty()) {
    get_master_theme(master_path)
  } else {
    g_theme.val
  }
  // ── If slide is modified, skip XML re-parse and use cached SlideData ──
  let resolved_slide = if g_modified.val[slide_idx] {
    g_slides.val[slide_idx]
  } else {
    // Load slide XML
    let slide_path = "ppt/slides/slide" + int_to_str(slide_num) + ".xml"
    let slide_xml = @ffi.ffi_get_file(slide_path)
    if slide_xml.is_empty() {
      return "ERROR:slide XML not found: " + slide_path
    }
    let effective_theme = {
      // Start with master's clr_map applied to master's theme
      let mut eff = match master_opt {
        Some(master) => @ooxml.apply_color_map(base_theme, master.clr_map)
        None => base_theme
      }
      // Apply layout clrMapOvr if present
      match layout_opt {
        Some(layout) =>
          match layout.clr_map_ovr {
            Some(ovr) => eff = @ooxml.apply_color_map(base_theme, ovr)
            None => ()
          }
        None => ()
      }
      // Apply slide's own clrMapOvr if present
      let slide_nodes = @xml.parse_xml(slide_xml)
      let slide_root_ch = match @xml.find_child(slide_nodes, "p:sld") {
        Some(n) => @xml.get_children(n)
        None => []
      }
      match @ooxml.extract_clr_map_ovr_from_nodes(slide_root_ch) {
        Some(ovr) => eff = @ooxml.apply_color_map(base_theme, ovr)
        None => ()
      }
      eff
    }
    // Parse slide and cache it
    let slide_data = @ooxml.parse_slide(
      slide_xml,
      g_slide_size.val,
      effective_theme,
    )
    // ── Resolve chart shapes: load chart XML and parse ChartML ──
    resolve_chart_shapes(slide_data.shapes, rels, effective_theme)
    // ── Placeholder inheritance: resolve transforms from layout/master ──
    resolve_placeholder_inheritance(slide_data.shapes, layout_opt, master_opt)
    // ── Placeholder auto-content: inject slide number/date/footer text ──
    inject_placeholder_auto_content(
      slide_data.shapes,
      slide_num,
      layout_opt,
      master_opt,
    )
    // ── Background inheritance: slide → layout → master → white ──
    let mut bg = slide_data.background
    let mut bg_grad = slide_data.bg_grad
    let mut bg_blip_fill = slide_data.bg_blip_fill
    let mut bg_patt_fill = slide_data.bg_patt_fill
    let mut bg_fill_rels : Array[@ooxml.Relationship] = rels
    // Check if slide has any explicit background
    let slide_has_bg = not(bg.is_none()) || not(bg_grad.is_none()) ||
      not(bg_blip_fill.is_none()) ||
      not(bg_patt_fill.is_none())
    if not(slide_has_bg) {
      // Try layout
      let mut inherited = false
      match layout_opt {
        Some(layout) => {
          let layout_has_bg = not(layout.background.is_none()) ||
            not(layout.bg_grad.is_none()) ||
            not(layout.bg_blip_fill.is_none()) ||
            not(layout.bg_patt_fill.is_none())
          if layout_has_bg {
            bg = layout.background
            bg_grad = layout.bg_grad
            bg_blip_fill = layout.bg_blip_fill
            bg_patt_fill = layout.bg_patt_fill
            bg_fill_rels = if not(layout_path.is_empty()) {
              get_layout_rels(layout_path)
            } else {
              rels
            }
            inherited = true
          }
        }
        None => ()
      }
      // Try master
      if not(inherited) {
        match master_opt {
          Some(master) => {
            let master_has_bg = not(master.background.is_none()) ||
              not(master.bg_grad.is_none()) ||
              not(master.bg_blip_fill.is_none()) ||
              not(master.bg_patt_fill.is_none())
            if master_has_bg {
              bg = master.background
              bg_grad = master.bg_grad
              bg_blip_fill = master.bg_blip_fill
              bg_patt_fill = master.bg_patt_fill
              bg_fill_rels = if not(master_path.is_empty()) {
                get_master_rels(master_path)
              } else {
                rels
              }
              inherited = true
            }
          }
          None => ()
        }
      }
      if not(inherited) && bg.is_none() {
        bg = @ooxml.Color::white()
      }
    }
    // ── Apply text style defaults: layout first (more specific), then master ──
    let ns = slide_data.shapes.length()
    let mut shape_i = 0
    while shape_i < ns {
      let shape = slide_data.shapes[shape_i]
      // Apply the shape's own inline lstStyle (p:txBody/a:lstStyle) first
      // This applies to all shapes (both placeholders and free-standing text boxes)
      apply_inline_lst_style(shape)
      // Layout placeholder defaults first (more specific)
      apply_layout_text_defaults(shape, layout_opt)
      // Master placeholder's lstStyle as second fallback
      apply_master_ph_text_defaults(shape, master_opt)
      // Master text styles (p:txStyles) as final fallback
      match master_opt {
        Some(master) => apply_text_defaults(shape, master.text_styles)
        None => ()
      }
      shape_i = shape_i + 1
    }
    let s : @ooxml.SlideData = {
      slide_size: slide_data.slide_size,
      background: bg,
      bg_grad,
      bg_blip_fill,
      bg_patt_fill,
      shapes: slide_data.shapes,
    }
    g_slides.val[slide_idx] = s
    ignore(bg_fill_rels)
    s
  }
  // ── Resolve effective theme for rendering (needed for both paths) ──
  let effective_theme = {
    let mut eff = match master_opt {
      Some(master) => @ooxml.apply_color_map(base_theme, master.clr_map)
      None => base_theme
    }
    match layout_opt {
      Some(layout) =>
        match layout.clr_map_ovr {
          Some(ovr) => eff = @ooxml.apply_color_map(base_theme, ovr)
          None => ()
        }
      None => ()
    }
    eff
  }
  // ── Resolve background fill rels for rendering ──
  let bg_fill_rels : Array[@ooxml.Relationship] = {
    let slide_has_bg = not(resolved_slide.background.is_none()) ||
      not(resolved_slide.bg_grad.is_none()) ||
      not(resolved_slide.bg_blip_fill.is_none()) ||
      not(resolved_slide.bg_patt_fill.is_none())
    if slide_has_bg {
      rels
    } else {
      // Check layout then master for bg fill rels
      let mut r = rels
      let mut found = false
      match layout_opt {
        Some(layout) => {
          if not(layout.background.is_none()) || not(layout.bg_grad.is_none()) ||
            not(layout.bg_blip_fill.is_none()) || not(layout.bg_patt_fill.is_none()) {
            r = if not(layout_path.is_empty()) { get_layout_rels(layout_path) } else { rels }
            found = true
          }
        }
        None => ()
      }
      if not(found) {
        match master_opt {
          Some(master) => {
            if not(master.background.is_none()) || not(master.bg_grad.is_none()) ||
              not(master.bg_blip_fill.is_none()) || not(master.bg_patt_fill.is_none()) {
              r = if not(master_path.is_empty()) { get_master_rels(master_path) } else { rels }
            }
          }
          None => ()
        }
      }
      r
    }
  }
  // ── Collect master/layout background shapes (non-placeholder) ──
  let bg_shapes : Array[@ooxml.Shape] = []
  let bg_shapes_rels : Array[Array[@ooxml.Relationship]] = []
  let m_rels = if not(master_path.is_empty()) {
    get_master_rels(master_path)
  } else {
    ([] : Array[@ooxml.Relationship])
  }
  match master_opt {
    Some(master) => {
      let ms = master.shapes.length()
      let mut mi = 0
      while mi < ms {
        let s = master.shapes[mi]
        if s.ph_type.is_empty() {
          bg_shapes.push(s)
          bg_shapes_rels.push(m_rels)
        }
        mi = mi + 1
      }
    }
    None => ()
  }
  let l_rels = if not(layout_path.is_empty()) {
    get_layout_rels(layout_path)
  } else {
    ([] : Array[@ooxml.Relationship])
  }
  match layout_opt {
    Some(layout) => {
      let ls = layout.shapes.length()
      let mut li = 0
      while li < ls {
        let s = layout.shapes[li]
        if s.ph_type.is_empty() {
          bg_shapes.push(s)
          bg_shapes_rels.push(l_rels)
        }
        li = li + 1
      }
    }
    None => ()
  }
  // Render to SVG
  @renderer.render_slide_svg(
    resolved_slide, rels, slide_num, effective_theme, bg_shapes, g_table_styles.val,
    bg_shapes_rels, bg_fill_rels,
  )
}

///|
/// Update a slide's internal data from an edited SVG string.
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

///|
/// Get the OOXML slide XML for a slide (0-indexed).
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

///|
/// Get tab-separated list of modified entries for ZIP rebuild.
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

// ── Shape-level editing API ──────────────────────────────────────────────────

///|
/// Compute the EMU-to-pixel scale factor for the current presentation.
fn get_scale() -> Int {
  let sw = g_slide_size.val.cx
  let w = 960 // svg_target_width
  if sw > 0 { sw / w } else { 9525 }
}

///|
/// Resolve a composite shape index (used by group children: parent*1000+child)
/// to the containing array and the local index within it.
/// Returns (shapes_array, local_index) or None if out of range.
fn resolve_shape_location(
  slide : @ooxml.SlideData,
  composite_idx : Int,
) -> (Array[@ooxml.Shape], Int)? {
  if composite_idx < 1000 {
    // Direct top-level shape
    if composite_idx >= 0 && composite_idx < slide.shapes.length() {
      Some((slide.shapes, composite_idx))
    } else {
      None
    }
  } else {
    // Group child: parent = composite_idx / 1000, child = composite_idx % 1000
    let parent_composite = composite_idx / 1000
    let child_local = composite_idx % 1000
    // Recursively resolve the parent
    match resolve_shape_location(slide, parent_composite) {
      Some((parent_arr, parent_local)) => {
        let parent_shape = parent_arr[parent_local]
        match parent_shape.kind {
          GroupShape(grp) =>
            if child_local >= 0 && child_local < grp.children.length() {
              Some((grp.children, child_local))
            } else {
              None
            }
          _ => None
        }
      }
      None => None
    }
  }
}

///|
/// Render a single shape as SVG (0-indexed slide and shape).
///
/// Returns the shape SVG (`<g>...</g>`) with its own `<defs>` prepended.
/// Returns "ERROR:<message>" on failure.
pub fn render_shape_svg(slide_idx : Int, shape_idx : Int) -> String {
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
    Some((arr, local_idx)) => {
      let shape = arr[local_idx]
      let scale = get_scale()
      let slide_num = slide_idx + 1
      // Load slide rels for image resolution
      let rels_path = "ppt/slides/_rels/slide" + int_to_str(slide_num) + ".xml.rels"
      let rels_xml = @ffi.ffi_get_file(rels_path)
      let rels = if rels_xml.is_empty() { [] } else { @ooxml.parse_rels(rels_xml) }
      // Generate defs for this shape
      let defs = @renderer.render_shape_defs(shape, slide_num, shape_idx, scale)
      let defs_section = if defs.is_empty() { "" } else { "<defs>" + defs + "</defs>" }
      // Render the shape
      let shape_svg = @renderer.render_shape(
        shape, rels, slide_num, scale, shape_idx, g_theme.val, g_table_styles.val,
      )
      defs_section + shape_svg
    }
  }
}

///|
/// Update a shape's transform (position, size, rotation) and return its re-rendered SVG.
///
/// All values are in EMU. Updates the cached SlideData in-place and marks the slide
/// as modified. Returns the shape SVG with defs, or "ERROR:<message>" on failure.
pub fn update_shape_transform(
  slide_idx : Int,
  shape_idx : Int,
  x : Int,
  y : Int,
  cx : Int,
  cy : Int,
  rot : Int,
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
    Some((arr, local_idx)) => {
      let shape = arr[local_idx]
      let new_transform : @ooxml.ShapeTransform = {
        x,
        y,
        cx,
        cy,
        rot,
        flip_h: shape.transform.flip_h,
        flip_v: shape.transform.flip_v,
      }
      arr[local_idx] = { ..shape, transform: new_transform }
      g_modified.val[slide_idx] = true
      render_shape_svg(slide_idx, shape_idx)
    }
  }
}

///|
/// Update a text run's content and return the shape's re-rendered SVG.
///
/// Updates the text of the specified paragraph/run in-place. Returns the shape SVG
/// with defs, or "ERROR:<message>" on failure.
pub fn update_shape_text(
  slide_idx : Int,
  shape_idx : Int,
  para_idx : Int,
  run_idx : Int,
  text : String,
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
    Some((arr, local_idx)) => {
      let shape = arr[local_idx]
      if para_idx < 0 || para_idx >= shape.paragraphs.length() {
        return "ERROR:paragraph index out of range"
      }
      let para = shape.paragraphs[para_idx]
      if run_idx < 0 || run_idx >= para.runs.length() {
        return "ERROR:run index out of range"
      }
      para.runs[run_idx] = { ..para.runs[run_idx], text, }
      g_modified.val[slide_idx] = true
      render_shape_svg(slide_idx, shape_idx)
    }
  }
}

///|
/// Update a shape's solid fill color and return its re-rendered SVG.
///
/// Sets the shape fill to the given RGB color (0-255 each). Returns the shape SVG
/// with defs, or "ERROR:<message>" on failure.
pub fn update_shape_fill(
  slide_idx : Int,
  shape_idx : Int,
  r : Int,
  g : Int,
  b : Int,
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
    Some((arr, local_idx)) => {
      let new_color : @ooxml.Color = { r, g: g, b, alpha: 255 }
      arr[local_idx] = { ..arr[local_idx], fill: new_color }
      g_modified.val[slide_idx] = true
      render_shape_svg(slide_idx, shape_idx)
    }
  }
}
