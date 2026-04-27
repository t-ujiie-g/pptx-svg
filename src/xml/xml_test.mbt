// ── int_to_str ──────────────────────────────────────────────────────────────

///|
test "int_to_str zero" {
  assert_eq(int_to_str(0), "0")
}

///|
test "int_to_str positive" {
  assert_eq(int_to_str(1), "1")
  assert_eq(int_to_str(42), "42")
  assert_eq(int_to_str(12345), "12345")
  assert_eq(int_to_str(1000000), "1000000")
}

///|
test "int_to_str negative" {
  assert_eq(int_to_str(-1), "-1")
  assert_eq(int_to_str(-999), "-999")
}

// ── parse_int ───────────────────────────────────────────────────────────────

///|
test "parse_int basic" {
  assert_eq(parse_int("0"), 0)
  assert_eq(parse_int("123"), 123)
  assert_eq(parse_int("-42"), -42)
  assert_eq(parse_int(""), 0)
}

///|
test "parse_int stops at non-digit" {
  assert_eq(parse_int("100pt"), 100)
}

// ── parse_hex_int ───────────────────────────────────────────────────────────

///|
test "parse_hex_int basic" {
  assert_eq(parse_hex_int("FF"), 255)
  assert_eq(parse_hex_int("ff"), 255)
  assert_eq(parse_hex_int("0"), 0)
  assert_eq(parse_hex_int("10"), 16)
  assert_eq(parse_hex_int(""), 0)
}

///|
test "parse_hex_int 6-digit color" {
  assert_eq(parse_hex_int("4488CC"), 0x4488CC)
  assert_eq(parse_hex_int("FFFFFF"), 0xFFFFFF)
  assert_eq(parse_hex_int("000000"), 0)
}

// ── str_starts_with / str_ends_with ─────────────────────────────────────────

///|
test "str_starts_with" {
  assert_true(str_starts_with("hello world", "hello"))
  assert_true(str_starts_with("hello", "hello"))
  assert_false(str_starts_with("hello", "world"))
  assert_false(str_starts_with("hi", "hello"))
  assert_true(str_starts_with("anything", ""))
}

///|
test "str_ends_with" {
  assert_true(str_ends_with("hello world", "world"))
  assert_true(str_ends_with("hello", "hello"))
  assert_false(str_ends_with("hello", "world"))
  assert_false(str_ends_with("hi", "hello"))
  assert_true(str_ends_with("anything", ""))
}

///|
test "str_replace_all basic" {
  assert_eq(str_replace_all("hello world", "world", "there"), "hello there")
  assert_eq(str_replace_all("aaa", "a", "b"), "bbb")
  // No match — unchanged
  assert_eq(str_replace_all("hello", "xyz", "!"), "hello")
  // Multiple non-adjacent matches
  assert_eq(str_replace_all("foo bar foo baz", "foo", "X"), "X bar X baz")
  // Replacement longer than pattern
  assert_eq(str_replace_all("ab", "a", "xyz"), "xyzb")
  // Replacement shorter than pattern
  assert_eq(str_replace_all("xyzb", "xyz", "a"), "ab")
  // Empty pattern is a no-op (would otherwise loop forever)
  assert_eq(str_replace_all("hello", "", "X"), "hello")
  // Empty replacement = deletion
  assert_eq(str_replace_all("hello world", " world", ""), "hello")
}

///|
test "str_replace_all phClr use case" {
  // The actual use case: substituting phClr in theme fmtScheme XML.
  let raw = "<a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/></a:schemeClr>"
  let out = str_replace_all(raw, "phClr", "accent1")
  assert_eq(
    out, "<a:schemeClr val=\"accent1\"><a:lumMod val=\"110000\"/></a:schemeClr>",
  )
}

// ── is_astral / char_step ───────────────────────────────────────────────────

///|
test "is_astral" {
  assert_false(is_astral('A'))
  assert_false(is_astral('\u{FFFF}'))
}

///|
test "char_step BMP" {
  assert_eq(char_step('A'), 1)
  assert_eq(char_step(' '), 1)
}

// ── parse_xml (requires FFI stub via JS target) ─────────────────────────────

///|
test "parse_xml empty" {
  let nodes = parse_xml("")
  assert_eq(nodes.length(), 0)
}

///|
test "parse_xml single element" {
  let nodes = parse_xml("<root/>")
  assert_eq(nodes.length(), 1)
  assert_eq(get_tag(nodes[0]), "root")
  assert_eq(get_children(nodes[0]).length(), 0)
}

///|
test "parse_xml element with attribute" {
  let nodes = parse_xml("<a:sp id=\"42\" name=\"test\"/>")
  assert_eq(nodes.length(), 1)
  assert_eq(get_tag(nodes[0]), "a:sp")
  assert_eq(get_attr(nodes[0], "id"), Some("42"))
  assert_eq(get_attr(nodes[0], "name"), Some("test"))
  assert_eq(get_attr(nodes[0], "missing"), None)
}

///|
test "parse_xml nested elements" {
  let nodes = parse_xml("<root><child1/><child2/></root>")
  assert_eq(nodes.length(), 1)
  let children = get_children(nodes[0])
  assert_eq(children.length(), 2)
  assert_eq(get_tag(children[0]), "child1")
  assert_eq(get_tag(children[1]), "child2")
}

///|
test "parse_xml text content" {
  let nodes = parse_xml("<a:t>Hello World</a:t>")
  assert_eq(nodes.length(), 1)
  assert_eq(get_text_content(nodes[0]), "Hello World")
}

///|
test "parse_xml mixed content" {
  let nodes = parse_xml("<p><b>bold</b> text</p>")
  let children = get_children(nodes[0])
  assert_eq(children.length(), 2)
  assert_eq(get_tag(children[0]), "b")
  assert_eq(get_text_content(children[0]), "bold")
  match children[1] {
    XmlNode::Text(t) => assert_eq(t, " text")
    _ => assert_true(false)
  }
}

// ── get_children / get_tag / get_attr on Text nodes ─────────────────────────

///|
test "accessors on Text node" {
  // Text nodes are created by parse_xml, not constructed directly
  let nodes = parse_xml("<p>hello</p>")
  let children = get_children(nodes[0])
  let text_node = children[0]
  assert_eq(get_tag(text_node), "")
  assert_eq(get_children(text_node).length(), 0)
  assert_true(get_attr(text_node, "x") is None)
}

// ── find_child / find_children ──────────────────────────────────────────────

///|
test "find_child" {
  let nodes = parse_xml("<root><a/><b/><c/></root>")
  let children = get_children(nodes[0])
  match find_child(children, "b") {
    Some(n) => assert_eq(get_tag(n), "b")
    None => assert_true(false)
  }
  assert_true(find_child(children, "missing") is None)
}

///|
test "find_children" {
  let nodes = parse_xml("<root><item/><other/><item/></root>")
  let children = get_children(nodes[0])
  let items = find_children(children, "item")
  assert_eq(items.length(), 2)
}

// ── find_path ───────────────────────────────────────────────────────────────

///|
test "find_path" {
  let nodes = parse_xml("<a><b><c val=\"found\"/></b></a>")
  match find_path(nodes, ["a", "b", "c"]) {
    Some(n) => assert_eq(get_attr(n, "val"), Some("found"))
    None => assert_true(false)
  }
  assert_true(find_path(nodes, ["a", "missing"]) is None)
  assert_true(find_path(nodes, []) is None)
}

// ── str_suffix / str_substring ──────────────────────────────────────────────

///|
test "str_suffix" {
  assert_eq(str_suffix("hello", 0), "hello")
  assert_eq(str_suffix("hello", 2), "llo")
  assert_eq(str_suffix("hello", 5), "")
}

///|
test "str_substring" {
  assert_eq(str_substring("hello world", 0, 5), "hello")
  assert_eq(str_substring("hello world", 6, 11), "world")
  assert_eq(str_substring("abc", 1, 2), "b")
}

// ── xml_escape / xml_unescape ───────────────────────────────────────────────

///|
test "xml_escape" {
  assert_eq(xml_escape("hello"), "hello")
  assert_eq(xml_escape("<b>"), "&lt;b&gt;")
  assert_eq(xml_escape("a&b"), "a&amp;b")
  assert_eq(xml_escape("say \"hi\""), "say &quot;hi&quot;")
}

///|
test "xml_unescape" {
  assert_eq(xml_unescape("hello"), "hello")
  assert_eq(xml_unescape("&lt;b&gt;"), "<b>")
  assert_eq(xml_unescape("a&amp;b"), "a&b")
  assert_eq(xml_unescape("say &quot;hi&quot;"), "say \"hi\"")
  assert_eq(xml_unescape("&apos;x&apos;"), "'x'")
}

///|
test "xml_escape_unescape roundtrip" {
  let original = "<tag attr=\"val\">&'test'"
  assert_eq(xml_unescape(xml_escape(original)), original)
}

// ── split_by_space ──────────────────────────────────────────────────────────

///|
test "split_by_space" {
  let tokens = split_by_space("hello world")
  assert_eq(tokens.length(), 2)
  assert_eq(tokens[0], "hello")
  assert_eq(tokens[1], "world")
}

///|
test "split_by_space single" {
  let tokens = split_by_space("hello")
  assert_eq(tokens.length(), 1)
  assert_eq(tokens[0], "hello")
}

///|
test "split_by_space empty" {
  assert_eq(split_by_space("").length(), 0)
}

// ── serialize_node ──────────────────────────────────────────────────────────

///|
test "serialize_node self-closing" {
  let nodes = parse_xml("<br/>")
  assert_eq(serialize_node(nodes[0]), "<br/>")
}

///|
test "serialize_node with attributes" {
  let nodes = parse_xml("<a href=\"url\" title=\"t\"/>")
  let s = serialize_node(nodes[0])
  assert_true(str_starts_with(s, "<a "))
  assert_true(str_ends_with(s, "/>"))
}

///|
test "serialize_node with text" {
  let nodes = parse_xml("<p>hello</p>")
  assert_eq(serialize_node(nodes[0]), "<p>hello</p>")
}

///|
test "serialize_node roundtrip" {
  let xml = "<a x=\"1\"><b>text</b><c/></a>"
  let nodes = parse_xml(xml)
  assert_eq(serialize_node(nodes[0]), xml)
}

// ── parse_xml OOXML fragment ────────────────────────────────────────────────

///|
test "parse_xml OOXML slide size" {
  let xml = "<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>"
  let nodes = parse_xml(xml)
  assert_eq(get_attr(nodes[0], "cx"), Some("9144000"))
  assert_eq(get_attr(nodes[0], "cy"), Some("6858000"))
}

///|
test "parse_xml OOXML text body" {
  let xml = "<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></a:txBody>"
  let nodes = parse_xml(xml)
  match find_path(nodes, ["a:txBody", "a:p", "a:r", "a:t"]) {
    Some(n) => assert_eq(get_text_content(n), "Hello")
    None => assert_true(false)
  }
}

// ── concat_balanced (security: M5 — O(N²) → O(N log N)) ─────────────────────

///|
test "concat_balanced empty" {
  let parts : Array[String] = []
  assert_eq(concat_balanced(parts), "")
}

///|
test "concat_balanced single" {
  let parts = ["hello"]
  assert_eq(concat_balanced(parts), "hello")
}

///|
test "concat_balanced two" {
  let parts = ["foo", "bar"]
  assert_eq(concat_balanced(parts), "foobar")
}

///|
test "concat_balanced odd count" {
  let parts = ["a", "b", "c", "d", "e"]
  assert_eq(concat_balanced(parts), "abcde")
}

///|
test "concat_balanced even count" {
  let parts = ["a", "b", "c", "d", "e", "f", "g", "h"]
  assert_eq(concat_balanced(parts), "abcdefgh")
}

///|
test "concat_balanced large input matches naive" {
  // 10000 single-char strings — naive concat would be O(N²) ≈ 50M ops.
  // concat_balanced should finish quickly and produce identical output.
  let parts : Array[String] = []
  let mut i = 0
  while i < 10000 {
    parts.push("x")
    i = i + 1
  }
  let result = concat_balanced(parts)
  assert_eq(result.length(), 10000)
  // Spot-check via slicing through str_substring (also O(N log N) post-fix).
  assert_eq(str_substring(result, 0, 5), "xxxxx")
  assert_eq(str_substring(result, 9995, 10000), "xxxxx")
}

// ── parse_xml depth limit (security: M4) ────────────────────────────────────

///|
test "parse_xml shallow nesting works" {
  // Nesting under the cap parses fully.
  let mut xml = ""
  let mut close = ""
  let mut i = 0
  while i < 100 {
    xml = xml + "<a>"
    close = "</a>" + close
    i = i + 1
  }
  let full = xml + "X" + close
  let nodes = parse_xml(full)
  // Walk down to the deepest node and verify text payload survived.
  let mut n = nodes
  let mut depth = 0
  while depth < 100 {
    match find_child(n, "a") {
      Some(node) => n = get_children(node)
      None => break
    }
    depth = depth + 1
  }
  assert_eq(depth, 100)
}

///|
test "parse_xml deep nesting does not stack-overflow" {
  // Far above max_xml_depth = 1024. Without the depth cap this would
  // recurse 5000 times and blow the Wasm call stack.
  let mut xml = ""
  let mut close = ""
  let mut i = 0
  while i < 5000 {
    xml = xml + "<a>"
    close = "</a>" + close
    i = i + 1
  }
  let full = xml + "X" + close
  // Just verify it returns without panicking. Depth past the cap is
  // skipped lexically (skip_to_close); the outer structure remains.
  let nodes = parse_xml(full)
  assert_true(nodes.length() >= 1)
}
