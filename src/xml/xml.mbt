// Simple XML parser for OOXML/PPTX documents.
//
// NOTE: This module intentionally avoids StringBuilder entirely.
// StringBuilder::to_string() calls wasm:js-string "fromCharCodeArray" which
// cannot be polyfilled in JS for Tier-2/3 fallback browsers.
// Instead, all strings are built via the `+` concat operator
// (→ wasm:js-string "concat") which IS polyfillable.
//
// Char→String conversion uses ffi_char_code_to_str (→ String.fromCharCode).

// ── Data types ────────────────────────────────────────────────────────────────

///|
pub struct XmlAttr {
  name : String
  value : String
}

///|
pub enum XmlNode {
  Element(String, Array[XmlAttr], Array[XmlNode])
  Text(String)
}

// ── Char → String without fromCharCodeArray ───────────────────────────────────

///|
fn char_to_str(c : Char) -> String {
  @ffi.ffi_char_code_to_str(Char::to_int(c))
}

// ── Parser state ──────────────────────────────────────────────────────────────

///|
priv struct Parser {
  src : String
  mut pos : Int
}

///|
fn Parser::new(src : String) -> Parser {
  { src, pos: 0 }
}

///|
fn Parser::at_end(self : Parser) -> Bool {
  self.pos >= self.src.length()
}

///|
fn Parser::peek_char(self : Parser) -> Char {
  match self.src.get_char(self.pos) {
    Some(c) => c
    None => '\u{0}'
  }
}

///|
fn Parser::advance(self : Parser) -> Unit {
  self.pos = self.pos + 1
}

///|
fn Parser::starts_with(self : Parser, s : String) -> Bool {
  let slen = s.length()
  if self.pos + slen > self.src.length() {
    return false
  }
  let mut i = 0
  while i < slen {
    if self.src.get_char(self.pos + i).unwrap() != s.get_char(i).unwrap() {
      return false
    }
    i = i + 1
  }
  true
}

///|
fn Parser::skip_str(self : Parser, s : String) -> Unit {
  self.pos = self.pos + s.length()
}

///|
fn Parser::skip_whitespace(self : Parser) -> Unit {
  while not(self.at_end()) {
    match self.peek_char() {
      ' ' | '\t' | '\n' | '\r' => self.advance()
      _ => break
    }
  }
}

// ── String helpers (no StringBuilder) ────────────────────────────────────────

///|
fn collect_chars(s : String, start : Int, end_exclusive : Int) -> String {
  let lim = if end_exclusive < s.length() { end_exclusive } else { s.length() }
  let mut result = ""
  let mut i = start
  while i < lim {
    result = result + char_to_str(s.get_char(i).unwrap())
    i = i + 1
  }
  result
}

///|
fn Parser::read_until_char(self : Parser, stop : Char) -> String {
  let mut result = ""
  while not(self.at_end()) {
    let c = self.peek_char()
    if c == stop {
      break
    }
    result = result + char_to_str(c)
    self.advance()
  }
  result
}

///|
fn Parser::read_until_str(self : Parser, stop : String) -> String {
  let mut result = ""
  while not(self.at_end()) {
    if self.starts_with(stop) {
      break
    }
    result = result + char_to_str(self.peek_char())
    self.advance()
  }
  result
}

// ── Entity decoding ───────────────────────────────────────────────────────────

///|
fn decode_entities(s : String) -> String {
  let len = s.length()
  // Fast path: no '&'
  let mut has_amp = false
  let mut k = 0
  while k < len {
    if s.get_char(k).unwrap() == '&' {
      has_amp = true
      break
    }
    k = k + 1
  }
  if not(has_amp) {
    return s
  }
  let mut result = ""
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '&' {
      let mut j = i + 1
      while j < len && s.get_char(j).unwrap() != ';' {
        j = j + 1
      }
      let entity = collect_chars(s, i + 1, j)
      if entity == "amp" {
        result = result + "&"
      } else if entity == "lt" {
        result = result + "<"
      } else if entity == "gt" {
        result = result + ">"
      } else if entity == "quot" {
        result = result + "\""
      } else if entity == "apos" {
        result = result + "'"
      } else {
        result = result + "&" + entity + ";"
      }
      i = j + 1
    } else {
      result = result + char_to_str(c)
      i = i + 1
    }
  }
  result
}

// ── Name and attribute parsing ────────────────────────────────────────────────

///|
fn Parser::parse_name(self : Parser) -> String {
  let mut result = ""
  while not(self.at_end()) {
    let c = self.peek_char()
    match c {
      ' ' | '\t' | '\n' | '\r' | '>' | '/' | '=' => break
      _ => {
        result = result + char_to_str(c)
        self.advance()
      }
    }
  }
  result
}

///|
fn Parser::parse_attr_value(self : Parser) -> String {
  let quote = self.peek_char()
  self.advance()
  let raw = self.read_until_char(quote)
  if not(self.at_end()) {
    self.advance() // skip closing quote
  }
  decode_entities(raw)
}

///|
fn Parser::parse_attrs(self : Parser) -> Array[XmlAttr] {
  let attrs : Array[XmlAttr] = []
  let mut running = true
  while running {
    self.skip_whitespace()
    if self.at_end() {
      running = false
    } else {
      let c = self.peek_char()
      if c == '>' || c == '/' {
        running = false
      } else {
        let name = self.parse_name()
        if name.is_empty() {
          self.advance() // skip unexpected char
        } else {
          self.skip_whitespace()
          if not(self.at_end()) && self.peek_char() == '=' {
            self.advance()
            self.skip_whitespace()
            if not(self.at_end()) {
              let q = self.peek_char()
              if q == '"' || q == '\'' {
                let value = self.parse_attr_value()
                attrs.push({ name, value })
              } else {
                let value = self.parse_name()
                attrs.push({ name, value })
              }
            }
          } else {
            attrs.push({ name, value: "" })
          }
        }
      }
    }
  }
  attrs
}

// ── Text helpers ──────────────────────────────────────────────────────────────

///|
fn is_all_whitespace(s : String) -> Bool {
  let len = s.length()
  let mut i = 0
  while i < len {
    match s.get_char(i).unwrap() {
      ' ' | '\t' | '\n' | '\r' => ()
      _ => return false
    }
    i = i + 1
  }
  true
}

// ── Core parsing ──────────────────────────────────────────────────────────────

///|
fn Parser::parse_children(self : Parser) -> Array[XmlNode] {
  let children : Array[XmlNode] = []
  let mut running = true
  while running {
    if self.at_end() {
      running = false
    } else {
      let c = self.peek_char()
      if c == '<' {
        self.advance()
        if self.at_end() {
          running = false
        } else {
          let nc = self.peek_char()
          if nc == '/' {
            // Closing tag: skip content to '>' (result unused, advances position)
            self.advance()
            let _ = self.read_until_char('>')
            if not(self.at_end()) {
              self.advance()
            }
            running = false
          } else if nc == '!' {
            self.advance()
            if self.starts_with("--") {
              // HTML comment: skip to end marker
              self.skip_str("--")
              let _ = self.read_until_str("-->")
              if not(self.at_end()) {
                self.skip_str("-->")
              }
            } else if self.starts_with("[CDATA[") {
              self.skip_str("[CDATA[")
              let text = self.read_until_str("]]>")
              if not(self.at_end()) {
                self.skip_str("]]>")
              }
              if not(text.is_empty()) {
                children.push(XmlNode::Text(text))
              }
            } else {
              // DOCTYPE or other declaration: skip to '>'
              let _ = self.read_until_char('>')
              if not(self.at_end()) {
                self.advance()
              }
            }
          } else if nc == '?' {
            // Processing instruction: skip to end marker
            self.advance()
            let _ = self.read_until_str("?>")
            if not(self.at_end()) {
              self.skip_str("?>")
            }
          } else {
            let tag = self.parse_name()
            if not(tag.is_empty()) {
              let attrs = self.parse_attrs()
              self.skip_whitespace()
              if not(self.at_end()) {
                let ec = self.peek_char()
                if ec == '/' {
                  self.advance()
                  if not(self.at_end()) {
                    self.advance()
                  }
                  children.push(XmlNode::Element(tag, attrs, []))
                } else if ec == '>' {
                  self.advance()
                  let ch = self.parse_children()
                  children.push(XmlNode::Element(tag, attrs, ch))
                }
              }
            } else {
              self.advance()
            }
          }
        }
      } else {
        let text = self.read_until_char('<')
        if not(is_all_whitespace(text)) {
          children.push(XmlNode::Text(decode_entities(text)))
        }
      }
    }
  }
  children
}

// ── Public API ────────────────────────────────────────────────────────────────

///|
/// Parse an XML string into a list of top-level XmlNode values.
pub fn parse_xml(src : String) -> Array[XmlNode] {
  Parser::new(src).parse_children()
}

///|
/// Get the children of an Element node (empty for Text nodes).
pub fn get_children(node : XmlNode) -> Array[XmlNode] {
  match node {
    XmlNode::Element(_, _, children) => children
    XmlNode::Text(_) => []
  }
}

///|
/// Get the tag name of an Element node, or "" for Text nodes.
pub fn get_tag(node : XmlNode) -> String {
  match node {
    XmlNode::Element(t, _, _) => t
    XmlNode::Text(_) => ""
  }
}

///|
/// Get an attribute value by name, or None if absent.
pub fn get_attr(node : XmlNode, name : String) -> String? {
  match node {
    XmlNode::Element(_, attrs, _) => {
      let n = attrs.length()
      let mut i = 0
      while i < n {
        if attrs[i].name == name {
          return Some(attrs[i].value)
        }
        i = i + 1
      }
      None
    }
    XmlNode::Text(_) => None
  }
}

///|
/// Find the first direct child Element with the given tag.
pub fn find_child(nodes : Array[XmlNode], tag : String) -> XmlNode? {
  let n = nodes.length()
  let mut i = 0
  while i < n {
    match nodes[i] {
      XmlNode::Element(t, _, _) => if t == tag { return Some(nodes[i]) }
      _ => ()
    }
    i = i + 1
  }
  None
}

///|
/// Find all direct child Elements with the given tag.
pub fn find_children(nodes : Array[XmlNode], tag : String) -> Array[XmlNode] {
  let result : Array[XmlNode] = []
  let n = nodes.length()
  let mut i = 0
  while i < n {
    match nodes[i] {
      XmlNode::Element(t, _, _) => if t == tag { result.push(nodes[i]) }
      _ => ()
    }
    i = i + 1
  }
  result
}

///|
/// Get all Text content within a node (concatenates Text children).
pub fn get_text_content(node : XmlNode) -> String {
  match node {
    XmlNode::Text(t) => t
    XmlNode::Element(_, _, children) => {
      let mut result = ""
      let n = children.length()
      let mut i = 0
      while i < n {
        match children[i] {
          XmlNode::Text(t) => result = result + t
          _ => ()
        }
        i = i + 1
      }
      result
    }
  }
}

///|
/// Navigate a tag-name path from a node list, returning the first match.
pub fn find_path(nodes : Array[XmlNode], path : Array[String]) -> XmlNode? {
  let plen = path.length()
  if plen == 0 {
    return None
  }
  match find_child(nodes, path[0]) {
    None => None
    Some(node) =>
      if plen == 1 {
        Some(node)
      } else {
        let rest : Array[String] = []
        let mut i = 1
        while i < plen {
          rest.push(path[i])
          i = i + 1
        }
        find_path(get_children(node), rest)
      }
  }
}

///|
/// Parse an integer from a string, returning 0 on failure.
pub fn parse_int(s : String) -> Int {
  let len = s.length()
  if len == 0 {
    return 0
  }
  let mut i = 0
  let mut neg = false
  if s.get_char(0).unwrap() == '-' {
    neg = true
    i = 1
  }
  let mut result = 0
  while i < len {
    let d = Char::to_int(s.get_char(i).unwrap()) - 48
    if d < 0 || d > 9 {
      break
    }
    result = result * 10 + d
    i = i + 1
  }
  if neg {
    -result
  } else {
    result
  }
}

///|
/// Convert a single digit (0-9) to its string representation.
fn digit_str(d : Int) -> String {
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

///|
/// Convert an integer to its decimal string representation.
pub fn int_to_str(n : Int) -> String {
  if n < 0 {
    "-" + int_to_str(-n)
  } else if n < 10 {
    digit_str(n)
  } else {
    int_to_str(n / 10) + digit_str(n % 10)
  }
}

// ── String helpers ──────────────────────────────────────────────────────────

///|
/// Extract a substring from index `from` to the end.
pub fn str_suffix(s : String, from : Int) -> String {
  let len = s.length()
  let mut result = ""
  let mut i = from
  while i < len {
    result = result +
      @ffi.ffi_char_code_to_str(Char::to_int(s.get_char(i).unwrap()))
    i = i + 1
  }
  result
}

///|
/// Check if string `s` starts with `pat`.
pub fn str_starts_with(s : String, pat : String) -> Bool {
  let slen = s.length()
  let plen = pat.length()
  if slen < plen {
    return false
  }
  for j = 0; j < plen; j = j + 1 {
    if s.get_char(j).unwrap() != pat.get_char(j).unwrap() {
      return false
    }
  }
  true
}

///|
/// Check if string `s` ends with `pat`.
pub fn str_ends_with(s : String, pat : String) -> Bool {
  let slen = s.length()
  let plen = pat.length()
  if slen < plen {
    return false
  }
  let start = slen - plen
  for j = 0; j < plen; j = j + 1 {
    if s.get_char(start + j).unwrap() != pat.get_char(j).unwrap() {
      return false
    }
  }
  true
}

///|
/// Extract substring from `from` (inclusive) to `to` (exclusive).
pub fn str_substring(s : String, from : Int, to : Int) -> String {
  let mut result = ""
  let mut i = from
  while i < to {
    result = result +
      @ffi.ffi_char_code_to_str(Char::to_int(s.get_char(i).unwrap()))
    i = i + 1
  }
  result
}

///|
/// Escape XML special characters: < > & " '
pub fn xml_escape(s : String) -> String {
  let len = s.length()
  let mut needs = false
  let mut k = 0
  while k < len {
    let c = s.get_char(k).unwrap()
    if c == '<' || c == '>' || c == '&' || c == '"' || c == '\'' {
      needs = true
      break
    }
    k = k + 1
  }
  if not(needs) {
    return s
  }
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
    } else if c == '\'' {
      result = result + "&apos;"
    } else {
      result = result + @ffi.ffi_char_code_to_str(Char::to_int(c))
    }
    i = i + 1
  }
  result
}

///|
/// Split a string by spaces into an array of tokens.
pub fn split_by_space(s : String) -> Array[String] {
  let tokens : Array[String] = []
  let len = s.length()
  let mut start = 0
  let mut i = 0
  while i < len {
    if s.get_char(i).unwrap() == ' ' {
      if i > start {
        tokens.push(str_substring(s, start, i))
      }
      start = i + 1
    }
    i = i + 1
  }
  if start < len {
    tokens.push(str_substring(s, start, len))
  }
  tokens
}
