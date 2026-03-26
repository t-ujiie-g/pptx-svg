// FFI stubs for MoonBit unit tests (JS target).
// Provides the `pptx_ffi` global that MoonBit FFI imports resolve to
// when compiled to JavaScript.
//
// Usage: NODE_OPTIONS="--require ./test_fixtures/ffi_stub.js" moon test --target js

globalThis.pptx_ffi = {
  char_code_to_str: (code) => String.fromCharCode(code),
  get_file: () => "",
  get_entry_list: () => "",
  get_file_base64: () => "",
  log: (msg) => console.log(msg),
  warn: (msg) => console.warn(msg),
  error: (msg) => console.error(msg),
  math_sin: Math.sin,
  math_cos: Math.cos,
  math_atan2: Math.atan2,
  math_sqrt: Math.sqrt,
  measure_text: () => 0.0,
  get_font_fallback: () => "",
  convert_emf: () => "",
};
