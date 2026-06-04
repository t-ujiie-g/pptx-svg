// pptx-svg: text geometry for inline editing (E6.2)
//
// `build_text_layout` reproduces `render_text`'s line/run/char positioning as a
// structured geometry (EMU) so the consumer can draw a caret / selection over the
// rendered SVG and hit-test clicks. It shares `solve_text_autofit`, `wrap_paragraph`,
// and `measure_line_width` with the renderer, so line and run counts match the
// rendered output by construction.
//
// v1 scope: horizontal LTR text (left/center/right/justify). Vertical text, warp,
// OMML math, and multi-column bodies return only the bounding box with no lines
// (the consumer falls back to coarse positioning for those). Bullets are accounted
// for approximately (left-anchored lines only).

///|
pub(all) struct GlyphBox {
  x : Int // EMU, absolute
  w : Int // EMU
}

///|
pub(all) struct RunBox {
  para_idx : Int
  run_idx : Int
  x : Int // EMU, absolute left edge
  w : Int // EMU
  run_char_start : Int // char offset within the run where this box begins
  para_char_start : Int // char offset within the paragraph where this box begins
  chars : Array[GlyphBox]
}

///|
pub(all) struct LineBox {
  para_idx : Int
  y : Int // EMU, top edge
  h : Int // EMU, line height
  runs : Array[RunBox]
}

///|
pub(all) struct TextLayout {
  box_x : Int
  box_y : Int
  box_cx : Int
  box_cy : Int
  lines : Array[LineBox]
}

///|
/// Approximate bullet glyph width (px) for a paragraph's first line.
fn layout_bullet_width_px(
  para : @ooxml.TextParagraph,
  p_font_sz : Int,
  theme : @ooxml.ThemeData,
) -> Double {
  if para.bullet_none || para.bullet.is_empty() {
    // Auto-number / image bullets: skip width estimation in v1.
    return 0.0
  }
  let ff = if !para.bullet_font.is_empty() {
    para.bullet_font
  } else {
    theme.minor_font
  }
  @ffi.ffi_measure_text(para.bullet + " ", ff, p_font_sz.to_double())
}

///|
/// Build text geometry (EMU) for a shape's paragraphs. See file header for scope.
pub fn build_text_layout(
  paragraphs : Array[@ooxml.TextParagraph],
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  body_props : @ooxml.BodyProps,
) -> TextLayout {
  let lines : Array[LineBox] = []
  let layout : TextLayout = {
    box_x: t.x,
    box_y: t.y,
    box_cx: t.cx,
    box_cy: t.cy,
    lines,
  }
  let np = paragraphs.length()
  // Unsupported modes → bounding box only (no per-line geometry).
  let wp = body_props.warp_prst
  let is_warp = !wp.is_empty() && wp != "textNoShape" && wp != "textPlain"
  let is_vert = body_props.vert == "vert" ||
    body_props.vert == "eaVert" ||
    body_props.vert == "vert270" ||
    body_props.vert == "wordArtVert"
  if np == 0 || is_warp || is_vert || is_all_math(paragraphs) {
    return layout
  }
  let scale_d = scale.to_double()

  // Insets / text rectangle (mirror render_text)
  let l_ins = if body_props.l_ins >= 0 {
    body_props.l_ins
  } else {
    default_l_inset
  }
  let t_ins = if body_props.t_ins >= 0 {
    body_props.t_ins
  } else {
    default_t_inset
  }
  let r_ins = if body_props.r_ins >= 0 {
    body_props.r_ins
  } else {
    default_l_inset
  }
  let b_ins = if body_props.b_ins >= 0 {
    body_props.b_ins
  } else {
    default_t_inset
  }
  let pad_l = px(l_ins, scale)
  let pad_t = px(t_ins, scale)
  let pad_r = px(r_ins, scale)
  let pad_b = px(b_ins, scale)
  let text_x = px(t.x, scale) + pad_l
  let text_w = px(t.cx, scale) - pad_l - pad_r
  let def_sz = px(default_font_size_emu, scale)
  let def_line_h = def_sz * line_height_numer / line_height_denom
  let do_wrap = body_props.wrap != "none"
  let col_w = text_w // single column only in v1
  let shape_h = px(t.cy, scale)
  let avail_h = shape_h - pad_t - pad_b

  // Solve autofit (shared with render_text → identical line/run counts).
  let (
    fs,
    lr,
    total_text_h,
    last_para_line_h,
    last_para_font_sz,
    last_para_spc_after,
  ) = solve_text_autofit(
    paragraphs,
    scale,
    theme,
    def_sz,
    def_line_h,
    body_props.font_scale,
    body_props.ln_spc_reduction,
    do_wrap,
    col_w,
    avail_h,
  )

  // Vertical anchor offset (mirror render_text)
  let trailing_space = if last_para_line_h > last_para_font_sz {
    last_para_line_h - last_para_font_sz + last_para_spc_after
  } else {
    last_para_spc_after
  }
  let text_h_for_anchor = total_text_h - trailing_space
  let anchor_offset = if body_props.anchor == "ctr" {
    (avail_h - text_h_for_anchor) / 2
  } else if body_props.anchor == "b" {
    avail_h - text_h_for_anchor
  } else {
    0
  }
  let anchor_offset_clamped = if body_props.anchor == "ctr" {
    if anchor_offset >= 0 - pad_t {
      anchor_offset
    } else {
      0 - pad_t
    }
  } else if anchor_offset > 0 {
    anchor_offset
  } else {
    0
  }
  let base_y = px(t.y, scale) + pad_t + anchor_offset_clamped
  let mut cur_y = base_y

  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    let nr = para.runs.length()
    let p_line_h = apply_autofit_scale(
      para_line_height(para, scale, def_line_h),
      fs,
      lr,
    )
    let p_font_sz = apply_font_scale(para_font_size(para, scale, def_sz), fs)
    if para.spc_before > 0 && pi > 0 {
      cur_y = cur_y + px(para.spc_before, scale)
    }
    let (bl_extra_above, bl_extra_below) = para_baseline_shift_extra(
      para, scale, def_sz,
    )
    cur_y = cur_y + bl_extra_above
    cur_y = cur_y + p_font_sz
    if nr == 0 {
      let top = cur_y - p_font_sz
      lines.push({ para_idx: pi, y: top * scale, h: p_line_h * scale, runs: [] })
      cur_y = cur_y + (p_line_h - p_font_sz)
      pi = pi + 1
      continue
    }
    // Indent + alignment (mirror render_text)
    let mar_l_emu = if para.mar_l >= 0 {
      para.mar_l
    } else {
      para.level * level_indent_emu
    }
    let indent = px(mar_l_emu, scale)
    let first_line_indent_emu = mar_l_emu + para.indent
    let first_line_indent = if first_line_indent_emu > 0 {
      px(first_line_indent_emu, scale)
    } else {
      0
    }
    let effective_align = if para.rtl && para.align.is_empty() {
      "r"
    } else {
      para.align
    }
    let bullet_w = layout_bullet_width_px(para, p_font_sz, theme)

    // Per-run paragraph offset prefix + consumed tracker
    let run_para_off : Array[Int] = []
    let mut acc_off = 0
    let mut ro = 0
    while ro < nr {
      run_para_off.push(acc_off)
      acc_off = acc_off + para.runs[ro].text.length()
      ro = ro + 1
    }
    let consumed : Array[Int] = Array::make(nr, 0)

    let wrap_w = if do_wrap { text_w - indent } else { no_wrap_width_px }
    let vlines = wrap_paragraph(para, wrap_w, scale, def_sz, fs, theme)
    let nvl = vlines.length()
    let mut vli = 0
    while vli < nvl {
      if vli > 0 {
        cur_y = cur_y + p_line_h
      }
      let baseline = cur_y
      let top = baseline - p_font_sz
      let vline = vlines[vli]
      let (line_w, _) = measure_line_width(
        vline, para, scale, def_sz, fs, theme,
      )
      // Left edge of the line content (px, Double)
      let ind = if vli == 0 { first_line_indent } else { indent }
      let bullet_off = if vli == 0 { bullet_w } else { 0.0 }
      let line_start : Double = if effective_align == "ctr" ||
        effective_align == "c" {
        (text_x + ind).to_double() +
        (text_w - ind).to_double() / 2.0 -
        line_w / 2.0
      } else if effective_align == "r" {
        (text_x + text_w).to_double() - line_w
      } else {
        // left / justify / default
        (text_x + ind).to_double() + bullet_off
      }
      // Walk segments left→right, merging consecutive same-run segments.
      let run_boxes : Array[RunBox] = []
      let mut cur_run_idx = -1
      let mut cur_box_x = 0.0
      let mut cur_box_w = 0.0
      let mut cur_chars : Array[GlyphBox] = []
      let mut cur_run_char_start = 0
      let mut x = line_start
      let nf = vline.length()
      let mut fi = 0
      fn flush_run() -> Unit {
        if cur_run_idx >= 0 {
          run_boxes.push({
            para_idx: pi,
            run_idx: cur_run_idx,
            x: (cur_box_x * scale_d).to_int(),
            w: (cur_box_w * scale_d).to_int(),
            run_char_start: cur_run_char_start,
            para_char_start: run_para_off[cur_run_idx] + cur_run_char_start,
            chars: cur_chars,
          })
        }
      }
      while fi < nf {
        let (run_idx, frag_text) = vline[fi]
        let run = para.runs[run_idx]
        let ff = run_font_face(run, theme)
        let fs_px = run_fs_px(run, scale, def_sz, fs)
        let text = if run.cap == "all" {
          to_upper_ascii(frag_text)
        } else {
          frag_text
        }
        let seg_run_char_start = consumed[run_idx]
        // Start a new run box on run change
        if run_idx != cur_run_idx {
          flush_run()
          cur_run_idx = run_idx
          cur_box_x = x
          cur_box_w = 0.0
          cur_chars = []
          cur_run_char_start = seg_run_char_start
        }
        // Per-char advance via cumulative prefix measurement
        let tlen = text.length()
        let mut ci = 0
        let mut prev_w = 0.0
        while ci < tlen {
          let step = @xml.char_step(text.get_char(ci).unwrap())
          let prefix = @xml.str_substring(text, 0, ci + step)
          let cum_w = @ffi.ffi_measure_text(prefix, ff, fs_px)
          let ch_w = cum_w - prev_w
          let ch_x = x + prev_w
          cur_chars.push({
            x: (ch_x * scale_d).to_int(),
            w: (ch_w * scale_d).to_int(),
          })
          prev_w = cum_w
          ci = ci + step
        }
        let seg_w = if tlen > 0 {
          @ffi.ffi_measure_text(text, ff, fs_px)
        } else {
          0.0
        }
        x = x + seg_w
        cur_box_w = cur_box_w + seg_w
        consumed[run_idx] = consumed[run_idx] + frag_text.length()
        fi = fi + 1
      }
      flush_run()
      lines.push({
        para_idx: pi,
        y: top * scale,
        h: p_line_h * scale,
        runs: run_boxes,
      })
      vli = vli + 1
    }
    cur_y = cur_y + (p_line_h - p_font_sz) + bl_extra_below
    if para.spc_after > 0 {
      cur_y = cur_y + px(para.spc_after, scale)
    }
    pi = pi + 1
  }
  layout
}

///|
/// Serialize a `TextLayout` to JSON (all numeric → no escaping needed).
pub fn text_layout_to_json(tl : TextLayout) -> String {
  let mut s = "{\"box\":{\"x\":" +
    int_to_str(tl.box_x) +
    ",\"y\":" +
    int_to_str(tl.box_y) +
    ",\"cx\":" +
    int_to_str(tl.box_cx) +
    ",\"cy\":" +
    int_to_str(tl.box_cy) +
    "},\"lines\":["
  let nl = tl.lines.length()
  let mut li = 0
  while li < nl {
    let line = tl.lines[li]
    if li > 0 {
      s = s + ","
    }
    s = s +
      "{\"paraIdx\":" +
      int_to_str(line.para_idx) +
      ",\"y\":" +
      int_to_str(line.y) +
      ",\"h\":" +
      int_to_str(line.h) +
      ",\"runs\":["
    let nr = line.runs.length()
    let mut ri = 0
    while ri < nr {
      let rb = line.runs[ri]
      if ri > 0 {
        s = s + ","
      }
      s = s +
        "{\"paraIdx\":" +
        int_to_str(rb.para_idx) +
        ",\"runIdx\":" +
        int_to_str(rb.run_idx) +
        ",\"x\":" +
        int_to_str(rb.x) +
        ",\"w\":" +
        int_to_str(rb.w) +
        ",\"runCharStart\":" +
        int_to_str(rb.run_char_start) +
        ",\"paraCharStart\":" +
        int_to_str(rb.para_char_start) +
        ",\"chars\":["
      let nc = rb.chars.length()
      let mut cj = 0
      while cj < nc {
        let g = rb.chars[cj]
        if cj > 0 {
          s = s + ","
        }
        s = s + "{\"x\":" + int_to_str(g.x) + ",\"w\":" + int_to_str(g.w) + "}"
        cj = cj + 1
      }
      s = s + "]}"
      ri = ri + 1
    }
    s = s + "]}"
    li = li + 1
  }
  s + "]}"
}

///|
/// Hit-test a point (EMU) against a layout → caret insertion position.
/// Returns `(para_idx, run_idx, char_in_run, para_offset)`.
/// Falls back to the nearest line / line end when the point is outside content.
pub fn hit_test_layout(
  tl : TextLayout,
  x_emu : Int,
  y_emu : Int,
) -> (Int, Int, Int, Int) {
  let nl = tl.lines.length()
  if nl == 0 {
    return (0, 0, 0, 0)
  }
  // Pick the line whose vertical band contains y, else the nearest by center.
  let mut line_idx = -1
  let mut li = 0
  while li < nl {
    let l = tl.lines[li]
    if y_emu >= l.y && y_emu < l.y + l.h {
      line_idx = li
      break
    }
    li = li + 1
  }
  if line_idx < 0 {
    if y_emu < tl.lines[0].y {
      line_idx = 0
    } else {
      line_idx = nl - 1
    }
  }
  let line = tl.lines[line_idx]
  let nr = line.runs.length()
  if nr == 0 {
    return (line.para_idx, 0, 0, 0)
  }
  // Before the line start → caret at the first run's start.
  let first = line.runs[0]
  if x_emu <= first.x {
    return (
      line.para_idx,
      first.run_idx,
      first.run_char_start,
      first.para_char_start,
    )
  }
  let mut ri = 0
  while ri < nr {
    let rb = line.runs[ri]
    if x_emu <= rb.x + rb.w {
      // Within this run: find the char by midpoint.
      let nc = rb.chars.length()
      let mut ci = 0
      while ci < nc {
        let g = rb.chars[ci]
        if x_emu < g.x + g.w / 2 {
          return (
            line.para_idx,
            rb.run_idx,
            rb.run_char_start + ci,
            rb.para_char_start + ci,
          )
        }
        ci = ci + 1
      }
      // Past the last char of this run box.
      return (
        line.para_idx,
        rb.run_idx,
        rb.run_char_start + nc,
        rb.para_char_start + nc,
      )
    }
    ri = ri + 1
  }
  // Past the line end → caret after the last run box.
  let last = line.runs[nr - 1]
  let lc = last.chars.length()
  (
    line.para_idx,
    last.run_idx,
    last.run_char_start + lc,
    last.para_char_start + lc,
  )
}
