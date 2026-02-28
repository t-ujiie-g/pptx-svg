// OOXML slide XML serializer — SlideData → OOXML XML
//
// Generates valid pptx slide XML from the internal SlideData model.
// Used for the round-trip pipeline: PPTX → SlideData → SVG → SlideData → PPTX.
//
// NOTE: No StringBuilder — all strings built via `+` (wasm:js-string "concat").

// ── Integer → String ──────────────────────────────────────────────────────────

fn digit_str(d : Int) -> String {
  match d {
    0 => "0"; 1 => "1"; 2 => "2"; 3 => "3"; 4 => "4"
    5 => "5"; 6 => "6"; 7 => "7"; 8 => "8"; _ => "9"
  }
}

fn int_to_str(n : Int) -> String {
  if n < 0 {
    "-" + int_to_str(-n)
  } else if n < 10 {
    digit_str(n)
  } else {
    int_to_str(n / 10) + digit_str(n % 10)
  }
}

// ── XML helpers ──────────────────────────────────────────────────────────────

fn xml_escape(s : String) -> String {
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
  if not(needs) { return s }
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

// ── Color serialization ─────────────────────────────────────────────────────

fn serialize_color(color : @ooxml.Color) -> String {
  if color.is_none() {
    "<a:noFill/>"
  } else {
    "<a:solidFill><a:srgbClr val=\"" + color.to_hex() + "\"/></a:solidFill>"
  }
}

// ── Text body serialization ──────────────────────────────────────────────────

///| Serialize paragraphs as OOXML text body content (shared by shapes and table cells).
///  `wrapper` is the outer tag: "p:txBody" for shapes, "a:txBody" for table cells.
fn serialize_text_body_inner(
  paragraphs : Array[@ooxml.TextParagraph],
  wrapper : String
) -> String {
  let mut out = "<" + wrapper + ">"
  out = out + "<a:bodyPr/><a:lstStyle/>"
  let np = paragraphs.length()
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    out = out + "<a:p>"
    // Paragraph properties
    let has_align = not(para.align.is_empty())
    let has_level = para.level > 0
    if has_align || has_level {
      let mut ppr = "<a:pPr"
      if has_align {
        ppr = ppr + " algn=\"" + para.align + "\""
      }
      if has_level {
        ppr = ppr + " lvl=\"" + int_to_str(para.level) + "\""
      }
      out = out + ppr + "/>"
    }
    // Runs
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let run = para.runs[ri]
      out = out + "<a:r>"
      out = out + serialize_run_props(run)
      out = out + "<a:t>" + xml_escape(run.text) + "</a:t>"
      out = out + "</a:r>"
      ri = ri + 1
    }
    // If no runs, emit an empty run to keep the paragraph
    if nr == 0 {
      out = out + "<a:endParaRPr lang=\"en-US\"/>"
    }
    out = out + "</a:p>"
    pi = pi + 1
  }
  out + "</" + wrapper + ">"
}

///| Serialize a:rPr element for a text run.
fn serialize_run_props(run : @ooxml.TextRun) -> String {
  let mut rpr = "<a:rPr lang=\"en-US\" dirty=\"0\""
  if run.bold {
    rpr = rpr + " b=\"1\""
  }
  if run.italic {
    rpr = rpr + " i=\"1\""
  }
  if run.font_size > 0 {
    rpr = rpr + " sz=\"" + int_to_str(run.font_size) + "\""
  }
  let has_children = not(run.color.is_none()) || not(run.font_face.is_empty())
  if has_children {
    rpr = rpr + ">"
    if not(run.color.is_none()) {
      rpr = rpr + serialize_color(run.color)
    }
    if not(run.font_face.is_empty()) {
      rpr = rpr + "<a:latin typeface=\"" + xml_escape(run.font_face) + "\"/>"
    }
    rpr + "</a:rPr>"
  } else {
    rpr + "/>"
  }
}

fn serialize_text_body(paragraphs : Array[@ooxml.TextParagraph]) -> String {
  if paragraphs.length() == 0 {
    return ""
  }
  serialize_text_body_inner(paragraphs, "p:txBody")
}

// ── Shape serialization ──────────────────────────────────────────────────────

fn serialize_transform(t : @ooxml.ShapeTransform) -> String {
  let mut xfrm = "<a:xfrm"
  if t.rot != 0 {
    xfrm = xfrm + " rot=\"" + int_to_str(t.rot) + "\""
  }
  if t.flip_h {
    xfrm = xfrm + " flipH=\"1\""
  }
  if t.flip_v {
    xfrm = xfrm + " flipV=\"1\""
  }
  xfrm = xfrm + ">"
  xfrm = xfrm + "<a:off x=\"" + int_to_str(t.x) + "\" y=\"" + int_to_str(t.y) + "\"/>"
  xfrm = xfrm + "<a:ext cx=\"" + int_to_str(t.cx) + "\" cy=\"" + int_to_str(t.cy) + "\"/>"
  xfrm + "</a:xfrm>"
}

fn serialize_autoshape(shape : @ooxml.Shape, geom : @ooxml.ShapeGeom) -> String {
  let mut out = "<p:sp>"
  // Non-visual properties (minimal)
  out = out + "<p:nvSpPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvSpPr/>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvSpPr>"
  // Shape properties
  out = out + "<p:spPr>"
  out = out + serialize_transform(shape.transform)
  out = out + "<a:prstGeom prst=\"" + geom.to_prst() + "\"><a:avLst/></a:prstGeom>"
  // Fill
  if shape.fill.is_none() {
    out = out + "<a:noFill/>"
  } else {
    out = out + serialize_color(shape.fill)
  }
  // Stroke
  if not(shape.stroke.is_none()) {
    let w = if shape.stroke_w > 0 { shape.stroke_w } else { 12700 }
    out = out + "<a:ln w=\"" + int_to_str(w) + "\">"
    out = out + serialize_color(shape.stroke)
    out = out + "</a:ln>"
  }
  out = out + "</p:spPr>"
  // Text body
  out = out + serialize_text_body(shape.paragraphs)
  out + "</p:sp>"
}

fn serialize_picture(shape : @ooxml.Shape, rid : String) -> String {
  let mut out = "<p:pic>"
  // Non-visual properties
  out = out + "<p:nvPicPr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvPicPr>"
  // Blip fill
  out = out + "<p:blipFill>"
  out = out + "<a:blip r:embed=\"" + rid + "\"/>"
  out = out + "<a:stretch><a:fillRect/></a:stretch>"
  out = out + "</p:blipFill>"
  // Shape properties
  out = out + "<p:spPr>"
  out = out + serialize_transform(shape.transform)
  out = out + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
  out = out + "</p:spPr>"
  out + "</p:pic>"
}

fn serialize_table_cell(cell : @ooxml.TableCell) -> String {
  let mut out = "<a:tc>"
  // Text body
  if cell.paragraphs.length() > 0 {
    out = out + serialize_table_text_body(cell.paragraphs)
  } else {
    out = out + "<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\"/></a:p></a:txBody>"
  }
  // Cell properties
  out = out + "<a:tcPr>"
  if not(cell.fill.is_none()) {
    out = out + serialize_color(cell.fill)
  }
  out = out + "</a:tcPr>"
  out + "</a:tc>"
}

fn serialize_table_text_body(paragraphs : Array[@ooxml.TextParagraph]) -> String {
  serialize_text_body_inner(paragraphs, "a:txBody")
}

fn serialize_table(shape : @ooxml.Shape, table : @ooxml.TableData) -> String {
  let mut out = "<p:graphicFrame>"
  // Non-visual properties
  out = out + "<p:nvGraphicFramePr>"
  out = out + "<p:cNvPr id=\"0\" name=\"\"/>"
  out = out + "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGraphicFramePr>"
  // Transform (directly under graphicFrame)
  out = out + serialize_transform(shape.transform)
  // Table content
  out = out + "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">"
  out = out + "<a:tbl>"
  out = out + "<a:tblPr firstRow=\"1\" bandRow=\"1\"/>"
  // Grid columns
  out = out + "<a:tblGrid>"
  let ncols = table.col_widths.length()
  let mut ci = 0
  while ci < ncols {
    out = out + "<a:gridCol w=\"" + int_to_str(table.col_widths[ci]) + "\"/>"
    ci = ci + 1
  }
  out = out + "</a:tblGrid>"
  // Rows
  let nrows = table.rows.length()
  let mut ri = 0
  while ri < nrows {
    let row = table.rows[ri]
    out = out + "<a:tr h=\"" + int_to_str(row.height) + "\">"
    let ncells = row.cells.length()
    let mut cj = 0
    while cj < ncells {
      out = out + serialize_table_cell(row.cells[cj])
      cj = cj + 1
    }
    out = out + "</a:tr>"
    ri = ri + 1
  }
  out = out + "</a:tbl>"
  out = out + "</a:graphicData></a:graphic>"
  out + "</p:graphicFrame>"
}

fn serialize_shape(shape : @ooxml.Shape) -> String {
  match shape.kind {
    @ooxml.ShapeKind::AutoShape(geom) => serialize_autoshape(shape, geom)
    @ooxml.ShapeKind::Picture(rid) => serialize_picture(shape, rid)
    @ooxml.ShapeKind::TableShape(table) => serialize_table(shape, table)
    _ => ""
  }
}

// ── Public API ────────────────────────────────────────────────────────────────

///| Serialize a SlideData into a complete OOXML slide XML string.
pub fn serialize_slide(slide : @ooxml.SlideData) -> String {
  let mut out = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  out = out + "<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
  out = out + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
  out = out + " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
  out = out + "<p:cSld>"
  // Background
  if not(slide.background.is_none()) && not(
    slide.background.r == 255 && slide.background.g == 255 && slide.background.b == 255
  ) {
    out = out + "<p:bg><p:bgPr>"
    out = out + serialize_color(slide.background)
    out = out + "<a:effectLst/></p:bgPr></p:bg>"
  }
  // Shape tree
  out = out + "<p:spTree>"
  out = out + "<p:nvGrpSpPr>"
  out = out + "<p:cNvPr id=\"1\" name=\"\"/>"
  out = out + "<p:cNvGrpSpPr/>"
  out = out + "<p:nvPr/>"
  out = out + "</p:nvGrpSpPr>"
  out = out + "<p:grpSpPr>"
  out = out + "<a:xfrm>"
  out = out + "<a:off x=\"0\" y=\"0\"/>"
  out = out + "<a:ext cx=\"0\" cy=\"0\"/>"
  out = out + "<a:chOff x=\"0\" y=\"0\"/>"
  out = out + "<a:chExt cx=\"0\" cy=\"0\"/>"
  out = out + "</a:xfrm>"
  out = out + "</p:grpSpPr>"
  // Shapes
  let n = slide.shapes.length()
  let mut i = 0
  while i < n {
    out = out + serialize_shape(slide.shapes[i])
    i = i + 1
  }
  out = out + "</p:spTree>"
  out = out + "</p:cSld>"
  out = out + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
  out + "</p:sld>"
}
