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

// Global state
let g_slide_count : Ref[Int] = { val: 0 }

// Cache parsed slide size from presentation.xml
let g_slide_size : Ref[@ooxml.SlideSize] = {
  val: @ooxml.SlideSize::default(),
}

// Slide cache: parsed SlideData for each slide (0-indexed)
// Updated by update_slide_from_svg() for round-trip editing
let g_slides : Ref[Array[@ooxml.SlideData]] = { val: [] }

// Track which slides have been modified
let g_modified : Ref[Array[Bool]] = { val: [] }

// Cached theme data parsed from theme1.xml
let g_theme : Ref[@ooxml.ThemeData] = { val: @ooxml.ThemeData::default() }

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
  // Initialize slide cache
  let slides : Array[@ooxml.SlideData] = []
  let modified : Array[Bool] = []
  let mut i = 0
  while i < count {
    slides.push(
      {
        slide_size: g_slide_size.val,
        background: @ooxml.Color::white(),
        shapes: [],
      },
    )
    modified.push(false)
    i = i + 1
  }
  g_slides.val = slides
  g_modified.val = modified
  @ffi.ffi_log("Slide count: " + int_to_str(count))
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
  g_slides.val[slide_idx] = slide_data
  // Render to SVG
  @renderer.render_slide_svg(slide_data, rels, slide_num, g_theme.val)
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

fn ends_with(s : String, pat : String) -> Bool {
  let slen = s.length()
  let plen = pat.length()
  if slen < plen { return false }
  let start = slen - plen
  for j = 0; j < plen; j = j + 1 {
    if s.get_char(start + j).unwrap() != pat.get_char(j).unwrap() {
      return false
    }
  }
  true
}

fn starts_with(s : String, pat : String) -> Bool {
  let slen = s.length()
  let plen = pat.length()
  if slen < plen { return false }
  for j = 0; j < plen; j = j + 1 {
    if s.get_char(j).unwrap() != pat.get_char(j).unwrap() {
      return false
    }
  }
  true
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

// ── Integer → String conversion ───────────────────────────────────────────────

fn int_digit_str(d : Int) -> String {
  match d {
    0 => "0"
    1 => "1"
    2 => "2"
    3 => "3"
    4 => "4"
    5 => "5"
    6 => "6"
    7 => "7"
    8 => "8"
    _ => "9"
  }
}

fn int_to_str(n : Int) -> String {
  if n < 0 {
    "-" + int_to_str(-n)
  } else if n < 10 {
    int_digit_str(n)
  } else {
    int_to_str(n / 10) + int_digit_str(n % 10)
  }
}

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
