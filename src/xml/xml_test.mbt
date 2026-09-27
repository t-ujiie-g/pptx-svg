///|
using @test {assert_eq}

// ── int_to_str ──────────────────────────────────────────────────────────────

///|
test "int_to_str zero" {
  assert_eq(@xml.int_to_str(0), "0")
}

///|
test "int_to_str positive" {
  assert_eq(@xml.int_to_str(1), "1")
  assert_eq(@xml.int_to_str(42), "42")
  assert_eq(@xml.int_to_str(12345), "12345")
  assert_eq(@xml.int_to_str(1000000), "1000000")
}

///|
test "int_to_str negative" {
  assert_eq(@xml.int_to_str(-1), "-1")
  assert_eq(@xml.int_to_str(-999), "-999")
}

// ── parse_int ───────────────────────────────────────────────────────────────

///|
test "parse_int basic" {
  assert_eq(@xml.parse_int("0"), 0)
  assert_eq(@xml.parse_int("123"), 123)
  assert_eq(@xml.parse_int("-42"), -42)
  assert_eq(@xml.parse_int(""), 0)
}

///|
test "parse_int stops at non-digit" {
  assert_eq(@xml.parse_int("100pt"), 100)
}

// ── parse_hex_int ───────────────────────────────────────────────────────────

///|
test "parse_hex_int basic" {
  assert_eq(@xml.parse_hex_int("FF"), 255)
  assert_eq(@xml.parse_hex_int("ff"), 255)
  assert_eq(@xml.parse_hex_int("0"), 0)
  assert_eq(@xml.parse_hex_int("10"), 16)
  assert_eq(@xml.parse_hex_int(""), 0)
}

///|
test "parse_hex_int 6-digit color" {
  assert_eq(@xml.parse_hex_int("4488CC"), 0x4488CC)
  assert_eq(@xml.parse_hex_int("FFFFFF"), 0xFFFFFF)
  assert_eq(@xml.parse_hex_int("000000"), 0)
}

// ── str_starts_with / str_ends_with ─────────────────────────────────────────

///|
test "str_starts_with" {
  assert_true(@xml.str_starts_with("hello world", "hello"))
  assert_true(@xml.str_starts_with("hello", "hello"))
  assert_false(@xml.str_starts_with("hello", "world"))
  assert_false(@xml.str_starts_with("hi", "hello"))
  assert_true(@xml.str_starts_with("anything", ""))
}

///|
test "str_ends_with" {
  assert_true(@xml.str_ends_with("hello world", "world"))
  assert_true(@xml.str_ends_with("hello", "hello"))
  assert_false(@xml.str_ends_with("hello", "world"))
  assert_false(@xml.str_ends_with("hi", "hello"))
  assert_true(@xml.str_ends_with("anything", ""))
}

///|
test "str_replace_all basic" {
  assert_eq(
    @xml.str_replace_all("hello world", "world", "there"),
    "hello there",
  )
  assert_eq(@xml.str_replace_all("aaa", "a", "b"), "bbb")
  // No match — unchanged
  assert_eq(@xml.str_replace_all("hello", "xyz", "!"), "hello")
  // Multiple non-adjacent matches
  assert_eq(@xml.str_replace_all("foo bar foo baz", "foo", "X"), "X bar X baz")
  // Replacement longer than pattern
  assert_eq(@xml.str_replace_all("ab", "a", "xyz"), "xyzb")
  // Replacement shorter than pattern
  assert_eq(@xml.str_replace_all("xyzb", "xyz", "a"), "ab")
  // Empty pattern is a no-op (would otherwise loop forever)
  assert_eq(@xml.str_replace_all("hello", "", "X"), "hello")
  // Empty replacement = deletion
  assert_eq(@xml.str_replace_all("hello world", " world", ""), "hello")
}

///|
test "str_replace_all phClr use case" {
  // The actual use case: substituting phClr in theme fmtScheme XML.
  let raw = "<a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/></a:schemeClr>"
  let out = @xml.str_replace_all(raw, "phClr", "accent1")
  assert_eq(
    out, "<a:schemeClr val=\"accent1\"><a:lumMod val=\"110000\"/></a:schemeClr>",
  )
}

// ── is_astral / char_step ───────────────────────────────────────────────────

///|
test "is_astral" {
  assert_false(@xml.is_astral('A'))
  assert_false(@xml.is_astral('\u{FFFF}'))
}

///|
test "char_step BMP" {
  assert_eq(@xml.char_step('A'), 1)
  assert_eq(@xml.char_step(' '), 1)
}

// ── parse_xml (requires FFI stub via JS target) ─────────────────────────────

///|
test "parse_xml empty" {
  let nodes = @xml.parse_xml("")
  assert_eq(nodes.length(), 0)
}

///|
test "parse_xml single element" {
  let nodes = @xml.parse_xml("<root/>")
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_tag(nodes[0]), "root")
  assert_eq(@xml.get_children(nodes[0]).length(), 0)
}

///|
test "parse_xml element with attribute" {
  let nodes = @xml.parse_xml("<a:sp id=\"42\" name=\"test\"/>")
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_tag(nodes[0]), "a:sp")
  assert_eq(@xml.get_attr(nodes[0], "id"), Some("42"))
  assert_eq(@xml.get_attr(nodes[0], "name"), Some("test"))
  assert_eq(@xml.get_attr(nodes[0], "missing"), None)
}

///|
test "parse_xml nested elements" {
  let nodes = @xml.parse_xml("<root><child1/><child2/></root>")
  assert_eq(nodes.length(), 1)
  let children = @xml.get_children(nodes[0])
  assert_eq(children.length(), 2)
  assert_eq(@xml.get_tag(children[0]), "child1")
  assert_eq(@xml.get_tag(children[1]), "child2")
}

///|
test "parse_xml text content" {
  let nodes = @xml.parse_xml("<a:t>Hello World</a:t>")
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_text_content(nodes[0]), "Hello World")
}

///|
test "parse_xml mixed content" {
  let nodes = @xml.parse_xml("<p><b>bold</b> text</p>")
  let children = @xml.get_children(nodes[0])
  assert_eq(children.length(), 2)
  assert_eq(@xml.get_tag(children[0]), "b")
  assert_eq(@xml.get_text_content(children[0]), "bold")
  match children[1] {
    @xml.XmlNode::Text(t) => assert_eq(t, " text")
    _ => assert_true(false)
  }
}

// ── get_children / get_tag / get_attr on Text nodes ─────────────────────────

///|
test "accessors on Text node" {
  // Text nodes are created by parse_xml, not constructed directly
  let nodes = @xml.parse_xml("<p>hello</p>")
  let children = @xml.get_children(nodes[0])
  let text_node = children[0]
  assert_eq(@xml.get_tag(text_node), "")
  assert_eq(@xml.get_children(text_node).length(), 0)
  assert_true(@xml.get_attr(text_node, "x") is None)
}

// ── find_child / find_children ──────────────────────────────────────────────

///|
test "find_child" {
  let nodes = @xml.parse_xml("<root><a/><b/><c/></root>")
  let children = @xml.get_children(nodes[0])
  match @xml.find_child(children, "b") {
    Some(n) => assert_eq(@xml.get_tag(n), "b")
    None => assert_true(false)
  }
  assert_true(@xml.find_child(children, "missing") is None)
}

///|
test "find_children" {
  let nodes = @xml.parse_xml("<root><item/><other/><item/></root>")
  let children = @xml.get_children(nodes[0])
  let items = @xml.find_children(children, "item")
  assert_eq(items.length(), 2)
}

// ── find_path ───────────────────────────────────────────────────────────────

///|
test "find_path" {
  let nodes = @xml.parse_xml("<a><b><c val=\"found\"/></b></a>")
  match @xml.find_path(nodes, ["a", "b", "c"]) {
    Some(n) => assert_eq(@xml.get_attr(n, "val"), Some("found"))
    None => assert_true(false)
  }
  assert_true(@xml.find_path(nodes, ["a", "missing"]) is None)
  assert_true(@xml.find_path(nodes, []) is None)
}

// ── str_suffix / str_substring ──────────────────────────────────────────────

///|
test "str_suffix" {
  assert_eq(@xml.str_suffix("hello", 0), "hello")
  assert_eq(@xml.str_suffix("hello", 2), "llo")
  assert_eq(@xml.str_suffix("hello", 5), "")
}

///|
test "str_substring" {
  assert_eq(@xml.str_substring("hello world", 0, 5), "hello")
  assert_eq(@xml.str_substring("hello world", 6, 11), "world")
  assert_eq(@xml.str_substring("abc", 1, 2), "b")
}

// ── xml_escape / xml_unescape ───────────────────────────────────────────────

///|
test "xml_escape" {
  assert_eq(@xml.xml_escape("hello"), "hello")
  assert_eq(@xml.xml_escape("<b>"), "&lt;b&gt;")
  assert_eq(@xml.xml_escape("a&b"), "a&amp;b")
  assert_eq(@xml.xml_escape("say \"hi\""), "say &quot;hi&quot;")
}

///|
test "xml_unescape" {
  assert_eq(@xml.xml_unescape("hello"), "hello")
  assert_eq(@xml.xml_unescape("&lt;b&gt;"), "<b>")
  assert_eq(@xml.xml_unescape("a&amp;b"), "a&b")
  assert_eq(@xml.xml_unescape("say &quot;hi&quot;"), "say \"hi\"")
  assert_eq(@xml.xml_unescape("&apos;x&apos;"), "'x'")
}

///|
test "xml_escape_unescape roundtrip" {
  let original = "<tag attr=\"val\">&'test'"
  assert_eq(@xml.xml_unescape(@xml.xml_escape(original)), original)
}

// ── split_by_space ──────────────────────────────────────────────────────────

///|
test "split_by_space" {
  let tokens = @xml.split_by_space("hello world")
  assert_eq(tokens.length(), 2)
  assert_eq(tokens[0], "hello")
  assert_eq(tokens[1], "world")
}

///|
test "split_by_space single" {
  let tokens = @xml.split_by_space("hello")
  assert_eq(tokens.length(), 1)
  assert_eq(tokens[0], "hello")
}

///|
test "split_by_space empty" {
  assert_eq(@xml.split_by_space("").length(), 0)
}

// ── serialize_node ──────────────────────────────────────────────────────────

///|
test "serialize_node self-closing" {
  let nodes = @xml.parse_xml("<br/>")
  assert_eq(@xml.serialize_node(nodes[0]), "<br/>")
}

///|
test "serialize_node with attributes" {
  let nodes = @xml.parse_xml("<a href=\"url\" title=\"t\"/>")
  let s = @xml.serialize_node(nodes[0])
  assert_true(@xml.str_starts_with(s, "<a "))
  assert_true(@xml.str_ends_with(s, "/>"))
}

///|
test "serialize_node with text" {
  let nodes = @xml.parse_xml("<p>hello</p>")
  assert_eq(@xml.serialize_node(nodes[0]), "<p>hello</p>")
}

///|
test "serialize_node roundtrip" {
  let xml = "<a x=\"1\"><b>text</b><c/></a>"
  let nodes = @xml.parse_xml(xml)
  assert_eq(@xml.serialize_node(nodes[0]), xml)
}

// ── parse_xml OOXML fragment ────────────────────────────────────────────────

///|
test "parse_xml OOXML slide size" {
  let xml = "<p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>"
  let nodes = @xml.parse_xml(xml)
  assert_eq(@xml.get_attr(nodes[0], "cx"), Some("9144000"))
  assert_eq(@xml.get_attr(nodes[0], "cy"), Some("6858000"))
}

///|
test "parse_xml OOXML text body" {
  let xml = "<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></a:txBody>"
  let nodes = @xml.parse_xml(xml)
  match @xml.find_path(nodes, ["a:txBody", "a:p", "a:r", "a:t"]) {
    Some(n) => assert_eq(@xml.get_text_content(n), "Hello")
    None => assert_true(false)
  }
}

// ── concat_balanced (security: M5 — O(N²) → O(N log N)) ─────────────────────

///|
test "concat_balanced empty" {
  let parts : Array[String] = []
  assert_eq(@xml.concat_balanced(parts), "")
}

///|
test "concat_balanced single" {
  let parts = ["hello"]
  assert_eq(@xml.concat_balanced(parts), "hello")
}

///|
test "concat_balanced two" {
  let parts = ["foo", "bar"]
  assert_eq(@xml.concat_balanced(parts), "foobar")
}

///|
test "concat_balanced odd count" {
  let parts = ["a", "b", "c", "d", "e"]
  assert_eq(@xml.concat_balanced(parts), "abcde")
}

///|
test "concat_balanced even count" {
  let parts = ["a", "b", "c", "d", "e", "f", "g", "h"]
  assert_eq(@xml.concat_balanced(parts), "abcdefgh")
}

///|
test "concat_balanced large input matches naive" {
  // 10000 single-char strings — naive concat would be O(N²) ≈ 50M ops.
  // concat_balanced should finish quickly and produce identical output.
  let parts : Array[String] = []
  for _ in 0..<10000 {
    parts.push("x")
  }
  let result = @xml.concat_balanced(parts)
  assert_eq(result.length(), 10000)
  // Spot-check via slicing through str_substring (also O(N log N) post-fix).
  assert_eq(@xml.str_substring(result, 0, 5), "xxxxx")
  assert_eq(@xml.str_substring(result, 9995, 10000), "xxxxx")
}

// ── parse_xml depth limit (security: M4) ────────────────────────────────────

///|
test "parse_xml shallow nesting works" {
  // Nesting under the cap parses fully.
  let mut xml = ""
  let mut close = ""
  for _ in 0..<100 {
    xml = xml + "<a>"
    close = "</a>" + close
  }
  let full = xml + "X" + close
  let nodes = @xml.parse_xml(full)
  // Walk down to the deepest node and verify text payload survived.
  let mut n = nodes
  let mut depth = 0
  while depth < 100 {
    match @xml.find_child(n, "a") {
      Some(node) => n = @xml.get_children(node)
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
  for _ in 0..<5000 {
    xml = xml + "<a>"
    close = "</a>" + close
  }
  let full = xml + "X" + close
  // Just verify it returns without panicking. Depth past the cap is
  // skipped lexically (skip_to_close); the outer structure remains.
  let nodes = @xml.parse_xml(full)
  assert_true(nodes.length() >= 1)
}

///|
/// `find_descendant` matches a local name at any depth regardless of namespace
/// prefix (used for diagram `relIds`/`dataModelExt` whose prefix varies).
test "find_descendant matches local name regardless of prefix" {
  let nodes = @xml.parse_xml(
    "<root><a:wrap><ext0:relIds r:dm=\"rId2\" r:cs=\"rId5\"/></a:wrap></root>",
  )
  match @xml.find_descendant(nodes, "relIds") {
    Some(ri) => assert_eq(@xml.get_attr(ri, "r:dm"), Some("rId2"))
    None => fail("relIds not found")
  }
  // Exact (unprefixed) tag also matches
  let plain = @xml.parse_xml("<root><dataModelExt relId=\"rIdX\"/></root>")
  match @xml.find_descendant(plain, "dataModelExt") {
    Some(ext) => assert_eq(@xml.get_attr(ext, "relId"), Some("rIdX"))
    None => fail("dataModelExt not found")
  }
  // No match returns None
  assert_true(
    match @xml.find_descendant(nodes, "missing") {
      None => true
      Some(_) => false
    },
  )
}
