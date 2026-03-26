// ── serialize_slide minimal ──────────────────────────────────────────────────

test "serialize_slide empty slide" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  // Should start with XML declaration
  assert_true(@xml.str_starts_with(xml, "<?xml version"))
  // Should contain p:sld root element
  assert_true(@xml.str_starts_with(xml, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><p:sld"))
  // Should contain shape tree
  let nodes = @xml.parse_xml(xml)
  assert_eq(nodes.length(), 1)
  assert_eq(@xml.get_tag(nodes[0]), "p:sld")
}

test "serialize_slide hidden slide" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: true,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  assert_eq(@xml.get_attr(nodes[0], "show"), Some("0"))
}

test "serialize_slide with colored background" {
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.parse_hex_color("FF0000"),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  // p:bg is inside p:cSld
  match @xml.find_path(nodes, ["p:sld", "p:cSld", "p:bg"]) {
    Some(_) => ()
    None => assert_true(false)
  }
}

test "serialize_slide with single shape" {
  let shape = @ooxml.Shape::default_with(
    @ooxml.ShapeKind::AutoShape(@ooxml.ShapeGeom::Rect),
    { x: 100000, y: 200000, cx: 500000, cy: 300000, rot: 0, flip_h: false, flip_v: false },
  )
  let slide : @ooxml.SlideData = {
    slide_size: @ooxml.SlideSize::default(),
    background: @ooxml.Color::white(),
    bg_grad: @ooxml.GradientFill::none(),
    bg_blip_fill: @ooxml.BlipFill::none(),
    bg_patt_fill: @ooxml.PatternFill::none(),
    shapes: [shape],
    transition_xml: "",
    timing_xml: "",
    hidden: false,
  }
  let xml = serialize_slide(slide)
  let nodes = @xml.parse_xml(xml)
  let csld = @xml.get_children(nodes[0])
  match @xml.find_child(csld, "p:cSld") {
    Some(csld_node) => {
      let sp_tree = @xml.find_child(@xml.get_children(csld_node), "p:spTree")
      assert_false(sp_tree is None)
    }
    None => assert_true(false)
  }
}
