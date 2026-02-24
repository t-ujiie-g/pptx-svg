// pptx-render: Phase 1 — Foundation
//
// Architecture:
//   JavaScript host:
//     - Parses the ZIP archive (reads local file headers)
//     - Decompresses entries using DecompressionStream API
//     - Exposes decompressed content via FFI (pptx_ffi.get_file, etc.)
//
//   MoonBit (this module):
//     - Receives OOXML XML strings via FFI
//     - Parses XML to build document model
//     - Generates SVG output
//
// Phase 1 scope:
//   - Prove the end-to-end Wasm FFI works
//   - Read presentation.xml and count slides
//   - Return slide XML on request

// Global state
let g_slide_count : Ref[Int] = { val: 0 }

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
  // Avoid string interpolation of integers (uses fromCharCodeArray internally).
  // Use explicit int_to_str() which only relies on concat + string literals.
  @ffi.ffi_log("presentation.xml: " + int_to_str(prs_xml.length()) + " chars")
  let count = count_slide_ids(prs_xml)
  g_slide_count.val = count
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

// ── Integer → String conversion ───────────────────────────────────────────────
//
// MoonBit's built-in integer interpolation uses `wasm:js-string.fromCharCodeArray`,
// which requires browser support for the WebAssembly JS String Builtins proposal.
// These helpers use only `wasm:js-string.concat` and string literals, which are
// far more widely supported and can be polyfilled when builtins are unavailable.

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

// Count slide IDs by finding <p:sldId or <sldId occurrences in presentation.xml.
// PPTX spec: each slide is referenced by a <p:sldId> element inside <p:sldIdLst>.
fn count_slide_ids(xml : String) -> Int {
  count_occurrences(xml, "<p:sldId ") + count_occurrences(xml, "<p:sldId\t") +
  count_occurrences(xml, "<p:sldId\n") + count_occurrences(xml, "<p:sldId/>")
}

// Count non-overlapping occurrences of `pattern` in `text`.
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

// Check if `pattern` matches `text` starting at position `pos`.
fn match_at(text : String, pattern : String, pos : Int) -> Bool {
  let plen = pattern.length()
  for j = 0; j < plen; j = j + 1 {
    if text.get_char(pos + j).unwrap() != pattern.get_char(j).unwrap() {
      return false
    }
  }
  true
}
