// Document state, filled by initialize_pptx. Per-slide arrays are 0-indexed;
// the *_paths arrays run parallel to the data arrays after them.

///|
let g_slide_count : Ref[Int32] = { val: 0, }

///|
let g_slide_size : Ref[@ooxml.SlideSize] = { val: @ooxml.SlideSize::default(), }

///|
let g_slides : Ref[Array[@ooxml.SlideData]] = { val: [], }

///|
let g_modified : Ref[Array[Bool]] = { val: [], }

///|
/// Whether the slide's XML has been parsed into g_slides (see ensure_slide_parsed).
let g_parsed : Ref[Array[Bool]] = { val: [], }

///|
/// Rels of the part each slide's background came from; empty = the slide's own.
let g_bg_rels : Ref[Array[Array[@ooxml.Relationship]]] = { val: [], }

///|
let g_theme : Ref[@ooxml.ThemeData] = { val: @ooxml.ThemeData::default(), }

///|
let g_slide_layout_paths : Ref[Array[String]] = { val: [], }

///|
let g_layout_paths : Ref[Array[String]] = { val: [], }

///|
let g_layouts : Ref[Array[(@ooxml.SlideLayoutData, String)]] = { val: [], }

///|
let g_master_paths : Ref[Array[String]] = { val: [], }

///|
let g_masters : Ref[Array[(@ooxml.SlideMasterData, String)]] = { val: [], }

///|
let g_layout_to_master : Ref[Array[(String, String)]] = { val: [], }

///|
let g_layout_rels : Ref[Array[Array[@ooxml.Relationship]]] = { val: [], }

///|
let g_master_rels : Ref[Array[Array[@ooxml.Relationship]]] = { val: [], }

///|
let g_master_themes : Ref[Array[@ooxml.ThemeData]] = { val: [], }

///|
let g_table_styles : Ref[Array[@ooxml.TableStyleDef]] = { val: [], }

///|
fn int_to_str(n : Int32) -> String {
  @xml.int_to_str(n)
}

///|
fn count_slide_ids(xml : String) -> Int32 {
  count_occurrences(xml, "<p:sldId ") +
  count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") +
  count_occurrences(xml, "<p:sldId/>")
}

///|
fn count_occurrences(text : String, pattern : String) -> Int32 {
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
fn match_at(text : String, pattern : String, pos : Int32) -> Bool {
  let plen = pattern.length()
  for j in 0..<plen {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}

///|
/// Find the theme file path from presentation.xml.rels.
fn find_theme_path(rels : Array[@ooxml.Relationship]) -> String {
  for rel in rels {
    if @xml.str_ends_with(rel.type_, "/theme") {
      if @xml.str_starts_with(rel.target, "/") {
        return @xml.str_suffix(rel.target, 1)
      } else {
        return "ppt/" + rel.target
      }
    }
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
  for rel in mrels {
    if @xml.str_ends_with(rel.type_, "/theme") {
      let slash = last_index_of(master_path, '/')
      let base_dir = if slash >= 0 {
        substring(master_path, 0, slash)
      } else {
        ""
      }
      let theme_path = resolve_rel_path(base_dir, rel.target)
      let theme_xml = @ffi.ffi_get_file(theme_path)
      if !theme_xml.is_empty() {
        @ffi.ffi_log("Loaded per-master theme: " + theme_path)
        return @ooxml.parse_theme(theme_xml)
      }
    }
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

///|
/// Resolve a relative target against a base directory
/// (`"ppt/slides"` + `"../slideLayouts/x.xml"` → `"ppt/slideLayouts/x.xml"`).
fn resolve_rel_path(base_dir : String, target : String) -> String {
  if @xml.str_starts_with(target, "/") {
    return @xml.str_suffix(target, 1)
  }
  let mut dir = base_dir
  let mut tgt = target
  while @xml.str_starts_with(tgt, "../") {
    let dlen = dir.length()
    let d2 = if dlen > 0 && dir.get_char(dlen - 1).unwrap() == '/' {
      substring(dir, 0, dlen - 1)
    } else {
      dir
    }
    let slash = last_index_of(d2, '/')
    if slash >= 0 {
      dir = substring(d2, 0, slash)
    } else {
      dir = ""
    }
    tgt = @xml.str_suffix(tgt, 3)
  }
  if dir.is_empty() {
    tgt
  } else {
    dir + "/" + tgt
  }
}

///|
/// Load and parse the chart part of every unresolved ChartShape, recursing into groups.
fn resolve_chart_shapes(
  shapes : Array[@ooxml.Shape],
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData,
) -> Unit {
  let n = shapes.length()
  for i in 0..<n {
    match shapes[i].kind {
      @ooxml.ShapeKind::ChartShape(cd) =>
        // Until resolved, chart_xml holds the chart's rId ("cx:"-prefixed for ChartEx).
        if !cd.chart_xml.is_empty() && cd.groups.is_empty() {
          let raw_rid = cd.chart_xml
          let is_chartex = @xml.str_starts_with(raw_rid, "cx:")
          let rid = if is_chartex {
            @xml.str_substring(raw_rid, 3, raw_rid.length())
          } else {
            raw_rid
          }
          let target = @ooxml.find_rel_target(rels, rid)
          if !target.is_empty() {
            let chart_path = resolve_rel_path("ppt/slides", target)
            let chart_xml = @ffi.ffi_get_file(chart_path)
            if !chart_xml.is_empty() {
              let parsed = if is_chartex {
                @ooxml.parse_chartex_xml(chart_xml, theme)
              } else {
                @ooxml.parse_chart_xml(chart_xml, theme)
              }
              shapes[i] = {
                ..shapes[i],
                kind: @ooxml.ShapeKind::ChartShape(parsed),
              }
            }
          }
        }
      @ooxml.ShapeKind::GroupShape(grp) =>
        resolve_chart_shapes(grp.children, rels, theme)
      _ => ()
    }
  }
}

///|
/// Read an attribute off the first `<*:relIds>` descendant of a parsed
/// SmartArt graphicFrame (e.g. `r:dm`, `r:cs`). Returns "" when absent.
fn relids_attr(gf_nodes : Array[@xml.XmlNode], attr : String) -> String {
  match @xml.find_descendant(gf_nodes, "relIds") {
    Some(ri) =>
      match @xml.get_attr(ri, attr) {
        Some(v) => v
        None => ""
      }
    None => ""
  }
}

///|
/// The shapes of a SmartArt frame's cached drawing (`ppt/drawings/drawingN.xml`), reached
/// via `dgm:relIds/@r:dm` → data part → `dsp:dataModelExt/@relId`; `[]` if any link is missing.
fn resolve_diagram_children(
  ole_xml : String,
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData,
) -> Array[@ooxml.Shape] {
  let gf_nodes = @xml.parse_xml(ole_xml)
  let dm_rid = relids_attr(gf_nodes, "r:dm")
  if dm_rid.is_empty() {
    return []
  }
  let data_target = @ooxml.find_rel_target(rels, dm_rid)
  if data_target.is_empty() {
    return []
  }
  let data_path = resolve_rel_path("ppt/slides", data_target)
  let data_xml = @ffi.ffi_get_file(data_path)
  if data_xml.is_empty() {
    return []
  }
  // dsp:dataModelExt/@relId → drawing rel in the data part's own rels.
  let draw_rid = match
    @xml.find_descendant(@xml.parse_xml(data_xml), "dataModelExt") {
    Some(ext) =>
      match @xml.get_attr(ext, "relId") {
        Some(v) => v
        None => ""
      }
    None => ""
  }
  if draw_rid.is_empty() {
    return []
  }
  let data_rels = @ooxml.parse_rels(@ffi.ffi_get_file(rels_path_for(data_path)))
  let draw_target = @ooxml.find_rel_target(data_rels, draw_rid)
  if draw_target.is_empty() {
    return []
  }
  let draw_path = resolve_rel_path(dir_of(data_path), draw_target)
  let children = @ooxml.parse_diagram_drawing(
    @ffi.ffi_get_file(draw_path),
    theme,
  )
  if children.is_empty() {
    return []
  }
  // SmartArt cached drawings omit run colors; derive the text color from the
  // diagram colors part (dgm:relIds/@r:cs) and fill it into uncolored runs.
  let cs_rid = relids_attr(gf_nodes, "r:cs")
  if !cs_rid.is_empty() {
    let cs_target = @ooxml.find_rel_target(rels, cs_rid)
    if !cs_target.is_empty() {
      let cs_path = resolve_rel_path("ppt/slides", cs_target)
      let txt_color = @ooxml.parse_diagram_text_color(
        @ffi.ffi_get_file(cs_path),
        theme,
      )
      @ooxml.apply_diagram_text_color(children, txt_color)
    }
  }
  children
}

///|
/// Fill every unresolved SmartArt group (no children, frame in `ole_xml`) from its cached drawing.
fn resolve_diagram_shapes(
  shapes : Array[@ooxml.Shape],
  rels : Array[@ooxml.Relationship],
  theme : @ooxml.ThemeData,
) -> Unit {
  let n = shapes.length()
  for i in 0..<n {
    match shapes[i].kind {
      @ooxml.ShapeKind::GroupShape(grp) =>
        if grp.children.is_empty() && !shapes[i].ole_xml.is_empty() {
          let children = resolve_diagram_children(
            shapes[i].ole_xml,
            rels,
            theme,
          )
          if !children.is_empty() {
            shapes[i] = {
              ..shapes[i],
              kind: @ooxml.ShapeKind::GroupShape({ ..grp, children, }),
            }
          }
        } else {
          resolve_diagram_shapes(grp.children, rels, theme)
        }
      _ => ()
    }
  }
}

///|
fn last_index_of(s : String, c : Char) -> Int32 {
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
fn substring(s : String, from : Int32, to : Int32) -> String {
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
  for rel in rels {
    if @xml.str_ends_with(rel.type_, type_suffix) {
      return rel.target
    }
  }
  ""
}

///|
/// Check if an array contains a string.
fn array_contains(arr : Array[String], val : String) -> Bool {
  for s in arr {
    if s == val {
      return true
    }
  }
  false
}

///|
/// Find index of a string in an array, or -1.
fn array_index_of(arr : Array[String], val : String) -> Int32 {
  let n = arr.length()
  for i in 0..<n {
    if arr[i] == val {
      return i
    }
  }
  -1
}

///|
/// Get the .rels path for a given content path.
fn rels_path_for(path : String) -> String {
  let dir = dir_of(path)
  let filename = @xml.str_suffix(path, dir.length() + 1)
  dir + "/_rels/" + filename + ".rels"
}

///|
/// Look up the master for a given layout path.
fn find_master_for_layout(layout_path : String) -> String {
  let pairs = g_layout_to_master.val
  for pair in pairs {
    let (lp, mp) = pair
    if lp == layout_path {
      return mp
    }
  }
  if !g_master_paths.val.is_empty() {
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
  for pair in layout_to_master {
    if pair.0 == layout_path {
      return pair.1
    }
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
  for master in masters {
    if master.1 == master_path {
      return master.0.clr_map
    }
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
  for i in 0..<n {
    if master_paths[i] == master_path {
      return master_themes[i]
    }
  }
  g_theme.val
}

///|
/// Load the presentation through the FFI. Returns "OK:<slide count>" or "ERROR:<message>".
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
  let prs_rels_xml = @ffi.ffi_get_file("ppt/_rels/presentation.xml.rels")
  if !prs_rels_xml.is_empty() {
    let prs_rels = @ooxml.parse_rels(prs_rels_xml)
    let theme_path = find_theme_path(prs_rels)
    if !theme_path.is_empty() {
      let theme_xml = @ffi.ffi_get_file(theme_path)
      if !theme_xml.is_empty() {
        g_theme.val = @ooxml.parse_theme(theme_xml)
        @ffi.ffi_log("Theme loaded from " + theme_path)
      }
    }
  }
  let tbl_styles_xml = @ffi.ffi_get_file("ppt/tableStyles.xml")
  if !tbl_styles_xml.is_empty() {
    g_table_styles.val = @ooxml.parse_table_styles(tbl_styles_xml, g_theme.val)
    if !g_table_styles.val.is_empty() {
      @ffi.ffi_log(
        "Table styles loaded: " + int_to_str(g_table_styles.val.length()),
      )
    }
  }
  let slide_layout_paths : Array[String] = []
  let layout_paths : Array[String] = []
  let layout_to_master : Array[(String, String)] = []
  let master_paths : Array[String] = []
  for si = 0; si < count; si = si + 1 {
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
    if !layout_path.is_empty() && !array_contains(layout_paths, layout_path) {
      layout_paths.push(layout_path)
    }
  }
  let nlay = layout_paths.length()
  for li in 0..<nlay {
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
    if !mpath.is_empty() && !array_contains(master_paths, mpath) {
      master_paths.push(mpath)
    }
  }
  let masters : Array[(@ooxml.SlideMasterData, String)] = []
  let master_rels_arr : Array[Array[@ooxml.Relationship]] = []
  let master_themes : Array[@ooxml.ThemeData] = []
  let nmas = master_paths.length()
  for mi in 0..<nmas {
    let mp = master_paths[mi]
    let mxml = @ffi.ffi_get_file(mp)
    let mrels_path = rels_path_for(mp)
    let mrels_xml = @ffi.ffi_get_file(mrels_path)
    let mrels = if mrels_xml.is_empty() {
      ([] : Array[@ooxml.Relationship])
    } else {
      @ooxml.parse_rels(mrels_xml)
    }
    let m_theme = load_master_theme(mp, mrels)
    master_themes.push(m_theme)
    if !mxml.is_empty() {
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
  }
  let layouts : Array[(@ooxml.SlideLayoutData, String)] = []
  let layout_rels_arr : Array[Array[@ooxml.Relationship]] = []
  for lp in layout_paths {
    let lxml = @ffi.ffi_get_file(lp)
    // Layout rels resolve the images of the layout's non-placeholder shapes.
    let layout_rels_path = rels_path_for(lp)
    let layout_rels_xml = @ffi.ffi_get_file(layout_rels_path)
    let cur_layout_rels = if layout_rels_xml.is_empty() {
      ([] : Array[@ooxml.Relationship])
    } else {
      @ooxml.parse_rels(layout_rels_xml)
    }
    if !lxml.is_empty() {
      let layout_effective_theme = {
        // Each master has its own theme; g_theme is only the presentation default.
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
  }
  g_slide_layout_paths.val = slide_layout_paths
  g_layout_paths.val = layout_paths
  g_layouts.val = layouts
  g_layout_rels.val = layout_rels_arr
  g_master_paths.val = master_paths
  g_masters.val = masters
  g_master_rels.val = master_rels_arr
  g_master_themes.val = master_themes
  g_layout_to_master.val = layout_to_master
  let slides : Array[@ooxml.SlideData] = []
  let modified : Array[Bool] = []
  let parsed : Array[Bool] = []
  let bg_rels : Array[Array[@ooxml.Relationship]] = []
  for i = 0; i < count; i = i + 1 {
    slides.push({
      slide_size: g_slide_size.val,
      background: @ooxml.Color::none(),
      bg_grad: @ooxml.GradientFill::none(),
      bg_blip_fill: @ooxml.BlipFill::none(),
      bg_patt_fill: @ooxml.PatternFill::none(),
      shapes: [],
      transition_xml: "",
      timing_xml: "",
      hidden: false,
      bg_inherited: false,
    })
    modified.push(false)
    parsed.push(false)
    bg_rels.push([])
  }
  g_slides.val = slides
  g_modified.val = modified
  g_parsed.val = parsed
  g_bg_rels.val = bg_rels
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
pub fn get_slide_count() -> Int32 {
  g_slide_count.val
}

///|
/// Check if a slide is hidden (0-indexed). Returns 1 if hidden, 0 otherwise.
pub fn is_slide_hidden(slide_idx : Int32) -> Int32 {
  if slide_idx < 0 || slide_idx >= g_slide_count.val {
    return 0
  }
  if g_modified.val[slide_idx] {
    return if g_slides.val[slide_idx].hidden { 1 } else { 0 }
  }
  let slide_num = slide_idx + 1
  let slide_path = "ppt/slides/slide" + int_to_str(slide_num) + ".xml"
  let slide_xml = @ffi.ffi_get_file(slide_path)
  if slide_xml.is_empty() {
    return 0
  }
  let nodes = @xml.parse_xml(slide_xml)
  match @xml.find_child(nodes, "p:sld") {
    Some(sld) =>
      match @xml.get_attr(sld, "show") {
        Some("0") => 1
        _ => 0
      }
    None => 0
  }
}

///|
/// Get the raw XML of a slide (0-indexed).
pub fn get_slide_xml_raw(slide_idx : Int32) -> String {
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
/// A slide's own .rels plus the layout / master / theme it inherits from.
priv struct SlideContext {
  rels : Array[@ooxml.Relationship]
  layout_path : String
  layout_opt : @ooxml.SlideLayoutData?
  master_path : String
  master_opt : @ooxml.SlideMasterData?
  base_theme : @ooxml.ThemeData // the master's own theme (g_theme without one)
}

///|
/// Resolve the inheritance context of slide `slide_idx` (0-indexed, in range).
fn slide_context(slide_idx : Int32) -> SlideContext {
  let rels_path = "ppt/slides/_rels/slide" +
    int_to_str(slide_idx + 1) +
    ".xml.rels"
  let rels_xml = @ffi.ffi_get_file(rels_path)
  let rels = if rels_xml.is_empty() { [] } else { @ooxml.parse_rels(rels_xml) }
  let layout_path = if slide_idx < g_slide_layout_paths.val.length() {
    g_slide_layout_paths.val[slide_idx]
  } else {
    ""
  }
  let layout_opt = if !layout_path.is_empty() {
    get_layout(layout_path)
  } else {
    None
  }
  let master_path = if !layout_path.is_empty() {
    find_master_for_layout(layout_path)
  } else {
    ""
  }
  let master_opt = if !master_path.is_empty() {
    get_master(master_path)
  } else {
    None
  }
  let base_theme = if !master_path.is_empty() {
    get_master_theme(master_path)
  } else {
    g_theme.val
  }
  { rels, layout_path, layout_opt, master_path, master_opt, base_theme, }
}

///|
/// `base_theme` under the layout's clrMapOvr, else the master's clrMap. A slide's own
/// clrMapOvr is applied by `resolve_slide_data`.
fn layout_effective_theme(
  base_theme : @ooxml.ThemeData,
  master_opt : @ooxml.SlideMasterData?,
  layout_opt : @ooxml.SlideLayoutData?,
) -> @ooxml.ThemeData {
  if layout_opt is Some(layout) && layout.clr_map_ovr is Some(ovr) {
    return @ooxml.apply_color_map(base_theme, ovr)
  }
  match master_opt {
    Some(master) => @ooxml.apply_color_map(base_theme, master.clr_map)
    None => base_theme
  }
}

///|
/// True when any of the four background fill kinds is set.
fn has_background(
  color : @ooxml.Color,
  grad : @ooxml.GradientFill,
  blip : @ooxml.BlipFill,
  patt : @ooxml.PatternFill,
) -> Bool {
  !color.is_none() || !grad.is_none() || !blip.is_none() || !patt.is_none()
}

///|
/// Parse slide XML into a fully resolved `SlideData` (theme, charts, SmartArt, placeholder
/// and background inheritance, text defaults). Idempotent on its own serialized output,
/// which Undo/Redo (`restore_slide_ooxml`) relies on.
fn resolve_slide_data(
  slide_num : Int32,
  slide_xml : String,
  rels : Array[@ooxml.Relationship],
  layout_opt : @ooxml.SlideLayoutData?,
  master_opt : @ooxml.SlideMasterData?,
  layout_path : String,
  master_path : String,
  base_theme : @ooxml.ThemeData,
) -> (@ooxml.SlideData, Array[@ooxml.Relationship]) {
  let effective_theme = {
    let mut eff = layout_effective_theme(base_theme, master_opt, layout_opt)
    let slide_nodes = @xml.parse_xml(slide_xml)
    let slide_root_ch = match @xml.find_child(slide_nodes, "p:sld") {
      Some(n) => @xml.get_children(n)
      None => []
    }
    if @ooxml.extract_clr_map_ovr_from_nodes(slide_root_ch) is Some(ovr) {
      eff = @ooxml.apply_color_map(base_theme, ovr)
    }
    eff
  }
  let slide_data = @ooxml.parse_slide(
    slide_xml,
    g_slide_size.val,
    effective_theme,
  )
  resolve_chart_shapes(slide_data.shapes, rels, effective_theme)
  resolve_diagram_shapes(slide_data.shapes, rels, effective_theme)
  resolve_placeholder_inheritance(slide_data.shapes, layout_opt, master_opt)
  inject_placeholder_auto_content(
    slide_data.shapes,
    slide_num,
    layout_opt,
    master_opt,
  )
  // `bg_rels` follows the part the background came from: an inherited
  // blipFill rId belongs to the layout's / master's .rels, not the slide's.
  let mut bg = slide_data.background
  let mut bg_grad = slide_data.bg_grad
  let mut bg_blip_fill = slide_data.bg_blip_fill
  let mut bg_patt_fill = slide_data.bg_patt_fill
  let mut bg_rels = rels
  let own_bg = has_background(bg, bg_grad, bg_blip_fill, bg_patt_fill)
  if !own_bg {
    if layout_opt is Some(layout) &&
      has_background(
        layout.background,
        layout.bg_grad,
        layout.bg_blip_fill,
        layout.bg_patt_fill,
      ) {
      bg = layout.background
      bg_grad = layout.bg_grad
      bg_blip_fill = layout.bg_blip_fill
      bg_patt_fill = layout.bg_patt_fill
      bg_rels = get_layout_rels(layout_path)
    } else if master_opt is Some(master) &&
      has_background(
        master.background,
        master.bg_grad,
        master.bg_blip_fill,
        master.bg_patt_fill,
      ) {
      bg = master.background
      bg_grad = master.bg_grad
      bg_blip_fill = master.bg_blip_fill
      bg_patt_fill = master.bg_patt_fill
      bg_rels = get_master_rels(master_path)
    } else {
      bg = @ooxml.Color::white()
    }
  }
  for shape in slide_data.shapes {
    // Text defaults, most specific first: the shape's own lstStyle, the layout placeholder,
    // the master placeholder, then the master's p:txStyles.
    apply_inline_lst_style(shape)
    apply_layout_text_defaults(shape, layout_opt)
    apply_master_ph_text_defaults(shape, master_opt)
    if master_opt is Some(master) {
      apply_text_defaults(shape, master.text_styles)
    }
    // Last resorts for runs still uncolored: endParaRPr, then the p:style fontRef color.
    apply_epr_fallbacks(shape)
    apply_font_ref_fallback(shape)
  }
  let resolved : @ooxml.SlideData = {
    slide_size: slide_data.slide_size,
    background: bg,
    bg_grad,
    bg_blip_fill,
    bg_patt_fill,
    shapes: slide_data.shapes,
    transition_xml: slide_data.transition_xml,
    timing_xml: slide_data.timing_xml,
    hidden: slide_data.hidden,
    bg_inherited: !own_bg,
  }
  (resolved, bg_rels)
}

///|
/// Render slide `slide_idx` (0-indexed) as SVG markup, or "ERROR:<message>".
pub fn render_slide_svg(slide_idx : Int32) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide_num = slide_idx + 1
  let ctx = slide_context(slide_idx)
  let rels = ctx.rels
  let layout_path = ctx.layout_path
  let layout_opt = ctx.layout_opt
  let master_path = ctx.master_path
  let master_opt = ctx.master_opt
  let resolved_slide = if g_modified.val[slide_idx] {
    g_slides.val[slide_idx]
  } else {
    let slide_path = "ppt/slides/slide" + int_to_str(slide_num) + ".xml"
    let slide_xml = @ffi.ffi_get_file(slide_path)
    if slide_xml.is_empty() {
      return "ERROR:slide XML not found: " + slide_path
    }
    let (s, bg_rels) = resolve_slide_data(
      slide_num,
      slide_xml,
      rels,
      layout_opt,
      master_opt,
      layout_path,
      master_path,
      ctx.base_theme,
    )
    g_slides.val[slide_idx] = s
    g_parsed.val[slide_idx] = true
    g_bg_rels.val[slide_idx] = bg_rels
    s
  }
  let effective_theme = layout_effective_theme(
    ctx.base_theme,
    master_opt,
    layout_opt,
  )
  let bg_shapes : Array[@ooxml.Shape] = []
  let bg_shapes_rels : Array[Array[@ooxml.Relationship]] = []
  let m_rels = if !master_path.is_empty() {
    get_master_rels(master_path)
  } else {
    ([] : Array[@ooxml.Relationship])
  }
  if master_opt is Some(master) {
    for s in master.shapes {
      if s.ph_type.is_empty() {
        bg_shapes.push(s)
        bg_shapes_rels.push(m_rels)
      }
    }
  }
  let l_rels = if !layout_path.is_empty() {
    get_layout_rels(layout_path)
  } else {
    ([] : Array[@ooxml.Relationship])
  }
  if layout_opt is Some(layout) {
    for s in layout.shapes {
      if s.ph_type.is_empty() {
        bg_shapes.push(s)
        bg_shapes_rels.push(l_rels)
      }
    }
  }
  @renderer.render_slide_svg(
    resolved_slide,
    rels,
    slide_num,
    effective_theme,
    bg_shapes,
    g_table_styles.val,
    bg_shapes_rels,
    g_bg_rels.val[slide_idx],
  )
}

///|
/// Replace a slide's data with an edited SVG, read back from its `data-ooxml-*`.
/// Returns "OK" or "ERROR:<message>".
pub fn update_slide_from_svg(slide_idx : Int32, svg_string : String) -> String {
  let n = g_slide_count.val
  if n == 0 {
    return "ERROR:not initialized"
  }
  if slide_idx < 0 || slide_idx >= n {
    return "ERROR:slide index out of range"
  }
  let slide_data = @svg_parser.parse_svg_to_slide(svg_string)
  // The SVG carries no chart data: re-read each chart from its preserved frame.
  let ctx = slide_context(slide_idx)
  let shapes = slide_data.shapes
  for i in 0..<shapes.length() {
    let s = shapes[i]
    if s.kind is @ooxml.ShapeKind::ChartShape(_) && !s.ole_xml.is_empty() {
      if @ooxml.parse_shapes_xml(s.ole_xml, ctx.base_theme) is [frame, ..] &&
        frame.kind is @ooxml.ShapeKind::ChartShape(_) {
        shapes[i] = { ..s, kind: frame.kind, }
      }
    }
  }
  resolve_chart_shapes(
    shapes,
    ctx.rels,
    layout_effective_theme(ctx.base_theme, ctx.master_opt, ctx.layout_opt),
  )
  g_slides.val[slide_idx] = slide_data
  g_modified.val[slide_idx] = true
  "OK"
}

///|
/// A slide's OOXML: serialized from g_slides if modified, otherwise the original part.
pub fn get_slide_ooxml(slide_idx : Int32) -> String {
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
/// The modified slides as `"path\tcontent\n"` records, for the host to rebuild the ZIP.
pub fn get_modified_entries() -> String {
  let n = g_slide_count.val
  let mut result = ""
  for i in 0..<n {
    if g_modified.val[i] {
      let path = "ppt/slides/slide" + int_to_str(i + 1) + ".xml"
      let content = @serializer.serialize_slide(g_slides.val[i])
      result = result + path + "\t" + content + "\n"
    }
  }
  result
}

///|
/// Compute the EMU-to-pixel scale factor for the current presentation.
fn get_scale() -> Int32 {
  let sw = g_slide_size.val.cx
  if sw > 0 {
    sw / @renderer.svg_target_width
  } else {
    @renderer.default_emu_per_px
  }
}
