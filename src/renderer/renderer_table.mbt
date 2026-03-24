// SVG renderer — Table rendering
//
// Same package as renderer.mbt — all constants and helpers are shared.

// ── Table constants ─────────────────────────────────────────────────────────

///|
let default_cell_mar_lr : Int = 91440 // OOXML default L/R margin: 0.1 inch

///|
let default_cell_mar_tb : Int = 45720 // OOXML default T/B margin: 0.05 inch

///|
let auto_contrast_lum_threshold : Int = 128000 // luminance threshold (×1000) for auto-contrast text color

// ── Table cell text rendering ───────────────────────────────────────────────

///|
fn render_table_cell_text(
  paragraphs : Array[@ooxml.TextParagraph],
  cell_x : Int,
  cell_y : Int,
  cell_w : Int,
  cell_h : Int,
  pad_l : Int,
  pad_r : Int,
  pad_t : Int,
  pad_b : Int,
  anchor : String,
  scale : Int,
  theme : @ooxml.ThemeData,
  rels : Array[@ooxml.Relationship],
  style_cell : @ooxml.TableStyleCell,
  cell_fill : @ooxml.Color,
) -> String {
  if paragraphs.length() == 0 {
    return ""
  }
  // Use smaller default font size for table cells (10pt vs 18pt for shapes)
  let def_sz = px(default_table_font_size_emu, scale)
  let available_h = cell_h - pad_t - pad_b
  if available_h <= 0 {
    return ""
  }
  let content_w = cell_w - pad_l - pad_r
  let np = paragraphs.length()
  // Pre-compute per-paragraph: line height, wrapped lines count, bullet indent
  let para_line_hs : Array[Int] = []
  let para_wrap_counts : Array[Int] = []
  let para_bul_mars : Array[Int] = []
  let mut pi = 0
  while pi < np {
    let para = paragraphs[pi]
    // Line height from max run font size
    let mut max_run_sz = 0
    let nr = para.runs.length()
    let mut ri = 0
    while ri < nr {
      let rsz = para.runs[ri].font_size
      if rsz > 0 {
        let rsp = px(rsz * @ooxml.emu_per_hundredth_pt, scale)
        if rsp > max_run_sz {
          max_run_sz = rsp
        }
      }
      ri = ri + 1
    }
    let font_sz = if max_run_sz > 0 { max_run_sz } else { def_sz }
    let lh = if para.line_spacing > 0 {
      // Absolute spacing in EMU
      px(para.line_spacing, scale)
    } else if para.line_spacing < 0 {
      // Percentage spacing (-100000 = 100% = single spacing)
      let pct = -para.line_spacing
      font_sz * pct / @ooxml.pct_100
    } else {
      font_sz * line_height_numer / line_height_denom
    }
    para_line_hs.push(lh)
    // Bullet indent
    let bul_mar = if para.mar_l >= 0 { px(para.mar_l, scale) } else { 0 }
    para_bul_mars.push(bul_mar)
    // Wrapped line count
    let wrap_w = content_w - bul_mar
    if nr == 0 {
      para_wrap_counts.push(1)
    } else if wrap_w > 0 {
      let vlines = wrap_paragraph(para, wrap_w, scale, def_sz, -1, theme)
      para_wrap_counts.push(vlines.length())
    } else {
      para_wrap_counts.push(1)
    }
    pi = pi + 1
  }
  // Total text height (all wrapped lines)
  let mut text_height = 0
  let mut thi = 0
  while thi < np {
    text_height = text_height + para_line_hs[thi] * para_wrap_counts[thi]
    // Space before/after
    if paragraphs[thi].spc_before > 0 && thi > 0 {
      text_height = text_height + px(paragraphs[thi].spc_before, scale)
    }
    if paragraphs[thi].spc_after > 0 {
      text_height = text_height + px(paragraphs[thi].spc_after, scale)
    }
    thi = thi + 1
  }
  // Vertical anchor offset (clamp to pad_t so text doesn't overflow above the cell)
  let y_raw = if anchor == "ctr" {
    pad_t + (available_h - text_height) / 2
  } else if anchor == "b" {
    pad_t + available_h - text_height
  } else {
    pad_t
  }
  let y_offset = if y_raw < pad_t { pad_t } else { y_raw }
  // Default text color: style > auto-contrast > black
  let style_fill = if not(style_cell.font_color.is_none()) {
    style_cell.font_color.to_css()
  } else if not(cell_fill.is_none()) {
    let lum = cell_fill.r * 299 + cell_fill.g * 587 + cell_fill.b * 114
    if lum < auto_contrast_lum_threshold { "#FFFFFF" } else { default_text_fill }
  } else {
    default_text_fill
  }
  let font_face = theme.minor_font + ", sans-serif"
  let mut out = "<text" +
    a("xml:space", "preserve") +
    a("font-family", font_face) +
    ai("font-size", def_sz) +
    a("fill", style_fill) +
    (if style_cell.bold { " font-weight=\"bold\"" } else { "" }) +
    ">"
  let mut cur_y = cell_y + y_offset
  let mut p_idx = 0
  while p_idx < np {
    let para = paragraphs[p_idx]
    let lh = para_line_hs[p_idx]
    let nr = para.runs.length()
    let bul_mar = para_bul_mars[p_idx]
    // Space before
    if para.spc_before > 0 && p_idx > 0 {
      cur_y = cur_y + px(para.spc_before, scale)
    }
    if nr == 0 && para.bullet.is_empty() {
      cur_y = cur_y + lh
      p_idx = p_idx + 1
      continue
    }
    // Paragraph alignment
    let algn = para.align
    let bul_indent = if para.indent != 0 { px(para.indent, scale) } else { 0 }
    // Text start x after bullet indent
    let text_x = cell_x + pad_l + bul_mar
    // Wrap text into visual lines
    let wrap_w = content_w - bul_mar
    let wrapped = if wrap_w > 0 {
      wrap_paragraph(para, wrap_w, scale, def_sz, -1, theme)
    } else {
      // Fallback: all segments on one line
      let segs : Array[(Int, String)] = []
      let mut ri = 0
      while ri < nr {
        segs.push((ri, para.runs[ri].text))
        ri = ri + 1
      }
      [segs]
    }
    let nlines = wrapped.length()
    let mut li = 0
    while li < nlines {
      cur_y = cur_y + lh
      let line = wrapped[li]
      // Compute x position based on alignment
      let (line_x, text_anchor) = if algn == "ctr" {
        (cell_x + cell_w / 2, "middle")
      } else if algn == "r" {
        (cell_x + cell_w - pad_r, "end")
      } else {
        (text_x, "start")
      }
      // Render bullet on first line only
      if li == 0 && not(para.bullet.is_empty()) && algn != "ctr" && algn != "r" {
        let bul_x = cell_x + pad_l + bul_mar + bul_indent
        out = out + "<tspan" + ai("x", bul_x) + ai("y", cur_y) + ">" +
          para.bullet + " </tspan>"
      }
      // Line tspan
      out = out + "<tspan" + ai("x", line_x) + ai("y", cur_y)
      if text_anchor != "start" {
        out = out + a("text-anchor", text_anchor)
      }
      out = out + ">"
      // Render segments
      let nseg = line.length()
      let mut si = 0
      while si < nseg {
        let (run_idx, seg_text) = line[si]
        if run_idx < nr {
          let run = para.runs[run_idx]
          let display = if run.cap == "all" {
            to_upper_ascii(seg_text)
          } else {
            seg_text
          }
          // Use render_run_tspan with segment text instead of full run text
          out = out + render_run_tspan(
            run, run_idx, display, scale, -1, rels, theme,
          )
        }
        si = si + 1
      }
      out = out + "</tspan>"
      li = li + 1
    }
    // Space after
    if para.spc_after > 0 {
      cur_y = cur_y + px(para.spc_after, scale)
    }
    p_idx = p_idx + 1
  }
  out + "</text>"
}

// ── Table border helpers ────────────────────────────────────────────────────

///|
fn render_table_border_line(
  x1 : Int,
  y1 : Int,
  x2 : Int,
  y2 : Int,
  w : Int,
  color : @ooxml.Color,
  scale : Int,
) -> String {
  // w == 0 → default 1px border, w == -1 → no border
  if w == -1 {
    return ""
  }
  let sw = if w > 0 { let v = px(w, scale); if v < 1 { 1 } else { v } } else { 1 }
  let sc = if color.is_none() { table_border_color } else { color.to_css() }
  "<line" + ai("x1", x1) + ai("y1", y1) + ai("x2", x2) + ai("y2", y2) +
    a("stroke", sc) + ai("stroke-width", sw) + "/>"
}

///|
fn eff_bdr_w(cell_w : Int, style_w : Int) -> Int {
  if cell_w != 0 { cell_w } else { style_w }
}

///|
fn eff_bdr_c(cell_c : @ooxml.Color, style_c : @ooxml.Color) -> @ooxml.Color {
  if not(cell_c.is_none()) { cell_c } else { style_c }
}

// ── Table style resolution ──────────────────────────────────────────────────

///|
fn find_table_style(
  style_id : String,
  table_styles : Array[@ooxml.TableStyleDef],
) -> @ooxml.TableStyleDef {
  if style_id.is_empty() {
    return @ooxml.TableStyleDef::none()
  }
  let n = table_styles.length()
  let mut i = 0
  while i < n {
    if table_styles[i].id == style_id {
      return table_styles[i]
    }
    i = i + 1
  }
  @ooxml.TableStyleDef::none()
}

///|
fn resolve_table_style_cell(
  style : @ooxml.TableStyleDef,
  table : @ooxml.TableData,
  row_idx : Int,
  col_idx : Int,
  nrows : Int,
  ncols : Int,
) -> @ooxml.TableStyleCell {
  // Priority: firstRow/lastRow/firstCol/lastCol > band > wholeTbl
  // Check special row/column positions first
  if table.first_row && row_idx == 0 {
    if not(style.first_row.is_none()) {
      return style.first_row
    }
  }
  if table.last_row && row_idx == nrows - 1 {
    if not(style.last_row.is_none()) {
      return style.last_row
    }
  }
  if table.first_col && col_idx == 0 {
    if not(style.first_col.is_none()) {
      return style.first_col
    }
  }
  if table.last_col && col_idx == ncols - 1 {
    if not(style.last_col.is_none()) {
      return style.last_col
    }
  }
  // Band rows (skip header row if firstRow is enabled)
  if table.band_row {
    let effective_row = if table.first_row { row_idx - 1 } else { row_idx }
    if effective_row >= 0 {
      if effective_row % 2 == 0 {
        if not(style.band1_h.is_none()) {
          return style.band1_h
        }
      } else {
        if not(style.band2_h.is_none()) {
          return style.band2_h
        }
      }
    }
  }
  // Band columns (skip first col if firstCol is enabled)
  if table.band_col {
    let effective_col = if table.first_col { col_idx - 1 } else { col_idx }
    if effective_col >= 0 {
      if effective_col % 2 == 0 {
        if not(style.band1_v.is_none()) {
          return style.band1_v
        }
      } else {
        if not(style.band2_v.is_none()) {
          return style.band2_v
        }
      }
    }
  }
  // Fallback to wholeTbl
  style.whole_tbl
}

// ── Table rendering ─────────────────────────────────────────────────────────

///|
fn render_table(
  table : @ooxml.TableData,
  t : @ooxml.ShapeTransform,
  scale : Int,
  theme : @ooxml.ThemeData,
  rels : Array[@ooxml.Relationship],
  table_styles : Array[@ooxml.TableStyleDef],
) -> String {
  let base_x = px(t.x, scale)
  let base_y = px(t.y, scale)
  let mut out = ""
  // Precompute column x-offsets in EMU
  let ncols = table.col_widths.length()
  let col_x_emu : Array[Int] = []
  let mut cx_accum = 0
  let mut ci = 0
  while ci < ncols {
    col_x_emu.push(cx_accum)
    cx_accum = cx_accum + table.col_widths[ci]
    ci = ci + 1
  }
  // Precompute row heights in EMU (auto-size rows with h=0)
  let nrows = table.rows.length()
  let row_heights_emu : Array[Int] = []
  // Minimum row height: default font size + top/bottom margins
  let min_row_h_emu = default_table_font_size_emu * line_height_numer /
    line_height_denom +
    default_cell_mar_tb * 2
  let mut rhi = 0
  while rhi < nrows {
    let h = table.rows[rhi].height
    if h > 0 {
      row_heights_emu.push(h)
    } else {
      // Auto-size: compute max text height across all cells in this row
      let row = table.rows[rhi]
      let mut max_cell_h_emu = min_row_h_emu
      let nc = row.cells.length()
      let mut ci2 = 0
      let mut col2 = 0
      while ci2 < nc {
        let cell = row.cells[ci2]
        if not(cell.h_merge) && not(cell.v_merge) {
          let mar_t_emu = if cell.mar_t >= 0 { cell.mar_t } else { default_cell_mar_tb }
          let mar_b_emu = if cell.mar_b >= 0 { cell.mar_b } else { default_cell_mar_tb }
          let mar_l_emu = if cell.mar_l >= 0 { cell.mar_l } else { default_cell_mar_lr }
          let mar_r_emu = if cell.mar_r >= 0 { cell.mar_r } else { default_cell_mar_lr }
          // Compute cell width from grid spans
          let gs = if cell.grid_span > 1 { cell.grid_span } else { 1 }
          let mut cell_w_emu = 0
          let mut gsi = 0
          while gsi < gs && col2 + gsi < ncols {
            cell_w_emu = cell_w_emu + table.col_widths[col2 + gsi]
            gsi = gsi + 1
          }
          let content_w_emu = cell_w_emu - mar_l_emu - mar_r_emu
          // Estimate text height from paragraphs
          let np = cell.paragraphs.length()
          let mut text_h_emu = 0
          let mut pi = 0
          while pi < np {
            let para = cell.paragraphs[pi]
            // Find max font size in this paragraph
            let mut max_fs_emu = 0
            let nr = para.runs.length()
            let mut rri = 0
            while rri < nr {
              let rsz = para.runs[rri].font_size
              if rsz > 0 {
                let fs_emu = rsz * @ooxml.emu_per_hundredth_pt
                if fs_emu > max_fs_emu {
                  max_fs_emu = fs_emu
                }
              }
              rri = rri + 1
            }
            if max_fs_emu == 0 {
              max_fs_emu = default_table_font_size_emu
            }
            let lh_emu = max_fs_emu * line_height_numer / line_height_denom
            // Estimate wrapped line count
            let bul_mar_emu = if para.mar_l >= 0 { para.mar_l } else { 0 }
            let wrap_w_emu = content_w_emu - bul_mar_emu
            let mut wrap_lines = 1
            if nr > 0 && wrap_w_emu > 0 {
              // Estimate total text width and compute line count
              let wrap_w_px = px(wrap_w_emu, scale)
              let vlines = wrap_paragraph(para, wrap_w_px, scale, px(default_table_font_size_emu, scale), -1, theme)
              wrap_lines = vlines.length()
              if wrap_lines < 1 {
                wrap_lines = 1
              }
            }
            if para.spc_before > 0 && pi > 0 {
              text_h_emu = text_h_emu + para.spc_before
            }
            text_h_emu = text_h_emu + lh_emu * wrap_lines
            if para.spc_after > 0 {
              text_h_emu = text_h_emu + para.spc_after
            }
            pi = pi + 1
          }
          if np == 0 {
            text_h_emu = default_table_font_size_emu * line_height_numer / line_height_denom
          }
          let cell_h_emu = text_h_emu + mar_t_emu + mar_b_emu
          if cell_h_emu > max_cell_h_emu {
            max_cell_h_emu = cell_h_emu
          }
        }
        let gs3 = if row.cells[ci2].grid_span > 1 { row.cells[ci2].grid_span } else { 1 }
        col2 = col2 + gs3
        ci2 = ci2 + 1
      }
      row_heights_emu.push(max_cell_h_emu)
    }
    rhi = rhi + 1
  }
  // Precompute row y-offsets in EMU
  let row_y_emu : Array[Int] = []
  let mut ry_accum = 0
  let mut ri2 = 0
  while ri2 < nrows {
    row_y_emu.push(ry_accum)
    ry_accum = ry_accum + row_heights_emu[ri2]
    ri2 = ri2 + 1
  }
  // Look up table style definition
  let style = find_table_style(table.style_id, table_styles)
  // Render cells (vMerge=true cells are skipped)
  let mut ri = 0
  while ri < nrows {
    let row = table.rows[ri]
    let ry = base_y + px(row_y_emu[ri], scale)
    let ncells = row.cells.length()
    let mut cj = 0
    let mut col_idx = 0 // actual grid column index
    while cj < ncells && col_idx < ncols {
      let cell = row.cells[cj]
      // Skip merged continuation cells
      if cell.v_merge {
        col_idx = col_idx + cell.grid_span
        cj = cj + 1
        continue
      }
      if cell.h_merge {
        // hMerge cells don't advance col_idx — parent's gridSpan already covers them
        cj = cj + 1
        continue
      }
      // Compute merged cell width (sum of grid_span columns)
      let span = cell.grid_span
      let mut merged_w_emu = 0
      let mut si = 0
      while si < span && (col_idx + si) < ncols {
        merged_w_emu = merged_w_emu + table.col_widths[col_idx + si]
        si = si + 1
      }
      // Compute merged cell height (sum of row_span rows)
      let rspan = cell.row_span
      let mut merged_h_emu = 0
      let mut sri = 0
      while sri < rspan && (ri + sri) < nrows {
        merged_h_emu = merged_h_emu + row_heights_emu[ri + sri]
        sri = sri + 1
      }
      let cx_px = base_x + px(col_x_emu[col_idx], scale)
      let cw = px(merged_w_emu, scale)
      let ch = px(merged_h_emu, scale)
      // Resolve table style for this cell position
      let style_cell = resolve_table_style_cell(
        style, table, ri, col_idx, nrows, ncols,
      )
      // Cell background: explicit fill > style fill > none
      let cell_fill = if not(cell.fill.is_none()) {
        cell.fill.to_css()
      } else if not(style_cell.fill.is_none()) {
        style_cell.fill.to_css()
      } else {
        "none"
      }
      out = out +
        "<rect" +
        ai("x", cx_px) +
        ai("y", ry) +
        ai("width", cw) +
        ai("height", ch) +
        a("fill", cell_fill) +
        " stroke=\"none\"/>"
      // Cell borders: explicit > style > default (top, bottom, left, right)
      out = out + render_table_border_line(cx_px, ry, cx_px + cw, ry, eff_bdr_w(cell.bdr_t_w, style_cell.bdr_t_w), eff_bdr_c(cell.bdr_t_color, style_cell.bdr_t_color), scale)
      out = out + render_table_border_line(cx_px, ry + ch, cx_px + cw, ry + ch, eff_bdr_w(cell.bdr_b_w, style_cell.bdr_b_w), eff_bdr_c(cell.bdr_b_color, style_cell.bdr_b_color), scale)
      out = out + render_table_border_line(cx_px, ry, cx_px, ry + ch, eff_bdr_w(cell.bdr_l_w, style_cell.bdr_l_w), eff_bdr_c(cell.bdr_l_color, style_cell.bdr_l_color), scale)
      out = out + render_table_border_line(cx_px + cw, ry, cx_px + cw, ry + ch, eff_bdr_w(cell.bdr_r_w, style_cell.bdr_r_w), eff_bdr_c(cell.bdr_r_color, style_cell.bdr_r_color), scale)
      // Diagonal borders
      if cell.bdr_tl_br_w > 0 {
        out = out + render_table_border_line(cx_px, ry, cx_px + cw, ry + ch, cell.bdr_tl_br_w, cell.bdr_tl_br_color, scale)
      }
      if cell.bdr_bl_tr_w > 0 {
        out = out + render_table_border_line(cx_px, ry + ch, cx_px + cw, ry, cell.bdr_bl_tr_w, cell.bdr_bl_tr_color, scale)
      }
      // Compute cell margins in pixels
      let pad_l = px(if cell.mar_l >= 0 { cell.mar_l } else { default_cell_mar_lr }, scale)
      let pad_t = px(if cell.mar_t >= 0 { cell.mar_t } else { default_cell_mar_tb }, scale)
      let pad_b = px(if cell.mar_b >= 0 { cell.mar_b } else { default_cell_mar_tb }, scale)
      // Cell text
      let pad_r = px(if cell.mar_r >= 0 { cell.mar_r } else { default_cell_mar_lr }, scale)
      // Resolve effective cell fill for auto-contrast text color
      let eff_fill = if not(cell.fill.is_none()) {
        cell.fill
      } else if not(style_cell.fill.is_none()) {
        style_cell.fill
      } else {
        @ooxml.Color::none()
      }
      out = out +
        render_table_cell_text(
          cell.paragraphs,
          cx_px,
          ry,
          cw,
          ch,
          pad_l,
          pad_r,
          pad_t,
          pad_b,
          cell.anchor,
          scale,
          theme,
          rels,
          style_cell,
          eff_fill,
        )
      col_idx = col_idx + span
      cj = cj + 1
    }
    ri = ri + 1
  }
  out
}
