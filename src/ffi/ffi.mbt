// FFI declarations for pptx-svg Wasm module
// These are host function imports provided by the JavaScript host layer.
//
// Architecture:
//   The JavaScript host handles ZIP parsing and DEFLATE decompression
//   (using browser's native DecompressionStream API), then exposes
//   decompressed PPTX entries to MoonBit via these FFI functions.

///|
/// Get the decompressed content of a PPTX ZIP entry as a UTF-8 string.
/// path: e.g. "ppt/slides/slide1.xml", "ppt/presentation.xml"
/// Returns the file content string, or an empty string if not found.
pub fn ffi_get_file(path : String) -> String = "pptx_ffi" "get_file"

///|
/// Get a newline-separated list of all entry paths in the PPTX archive.
pub fn ffi_get_entry_list() -> String = "pptx_ffi" "get_entry_list"

///|
/// Get the raw bytes of a PPTX entry as a base64-encoded string.
/// Used for binary files (images, fonts, etc.)
pub fn ffi_get_file_base64(path : String) -> String = "pptx_ffi" "get_file_base64"

///|
/// Convert a Unicode code point (int) to a single-character string.
/// Implemented via String.fromCharCode in the JavaScript host.
/// Used to build strings without StringBuilder (avoids fromCharCodeArray).
pub fn ffi_char_code_to_str(code : Int) -> String = "pptx_ffi" "char_code_to_str"

///|
/// Log an informational message to the browser console.
pub fn ffi_log(msg : String) -> Unit = "pptx_ffi" "log"

///|
/// Log a warning message to the browser console.
pub fn ffi_warn(msg : String) -> Unit = "pptx_ffi" "warn"

///|
/// Log an error message to the browser console.
pub fn ffi_error(msg : String) -> Unit = "pptx_ffi" "error"

///|
/// Math: sine (radians). Used by preset geometry guide formula evaluator.
pub fn ffi_sin(x : Double) -> Double = "pptx_ffi" "math_sin"

///|
/// Math: cosine (radians).
pub fn ffi_cos(x : Double) -> Double = "pptx_ffi" "math_cos"

///|
/// Math: atan2(y, x) → radians.
pub fn ffi_atan2(y : Double, x : Double) -> Double = "pptx_ffi" "math_atan2"

///|
/// Math: square root.
pub fn ffi_sqrt(x : Double) -> Double = "pptx_ffi" "math_sqrt"

///|
/// Measure the rendered width of text in pixels.
/// Used for text layout in SVG rendering (calls canvas.measureText).
pub fn ffi_measure_text(
  text : String,
  font_face : String,
  font_size_pt : Double,
) -> Double = "pptx_ffi" "measure_text"

///|
/// Get font fallback list for a given font name.
/// Returns comma-separated fallback font names, or empty string if none.
pub fn ffi_get_font_fallback(font : String) -> String = "pptx_ffi" "get_font_fallback"

///|
/// Convert an EMF file to an SVG data URI.
/// path: ZIP entry path (e.g. "ppt/media/image1.emf")
/// Returns a data:image/svg+xml URI, or empty string if conversion fails.
pub fn ffi_convert_emf(path : String) -> String = "pptx_ffi" "convert_emf"
