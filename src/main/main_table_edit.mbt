// pptx-svg: Table editing API (E6.6)
//
// Cell text + row/column add/delete. Cells/rows/col_widths are mutable arrays in
// the cached TableData, so edits happen in place. NOTE: merged cells
// (gridSpan/rowSpan/hMerge/vMerge) are not span-adjusted on row/column changes —
// editing the structure of a table with merged cells may corrupt it (v1 limit).

///|
/// Default width (EMU) for a new column when no width is given and there is no
/// neighbour to copy from. 914400 EMU = 1 inch.
let default_table_col_width_emu : Int = 914400

///|
/// Resolve a shape and require it to be a table. Calls `f` with its TableData,
/// or returns "ERROR:shape is not a table".
fn with_table(
  slide_idx : Int,
  shape_idx : Int,
  f : (@ooxml.TableData) -> String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    match arr[local_idx].kind {
      @ooxml.ShapeKind::TableShape(t) => f(t)
      _ => "ERROR:shape is not a table"
    }
  })
}

///|
/// Insertion index for an "insert after `after`" operation into an array of
/// length `len`: `after = -1` → front (0); otherwise `after + 1`, clamped to `len`.
fn clamp_insert_index(after : Int, len : Int) -> Int {
  if after < 0 {
    0
  } else if after + 1 > len {
    len
  } else {
    after + 1
  }
}

///|
/// Set a table cell's text (plain text, replacing its contents). The new run
/// inherits the cell's existing first-run formatting when present.
/// Returns re-rendered shape SVG or "ERROR:<message>".
pub fn update_table_cell_text(
  slide_idx : Int,
  shape_idx : Int,
  row : Int,
  col : Int,
  text : String,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    if row < 0 || row >= t.rows.length() {
      return "ERROR:row index out of range"
    }
    let cells = t.rows[row].cells
    if col < 0 || col >= cells.length() {
      return "ERROR:column index out of range"
    }
    let cell = cells[col]
    let run = if cell.paragraphs.length() > 0 &&
      cell.paragraphs[0].runs.length() > 0 {
      { ..cell.paragraphs[0].runs[0], text, }
    } else {
      make_default_run(text)
    }
    cells[col] = { ..cell, paragraphs: [make_default_paragraph(run)] }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Insert a new (empty) row. `after_row = -1` inserts at the top; otherwise after
/// the given row (clamped). Returns "OK:<new_row_idx>" or "ERROR:<message>".
pub fn add_table_row(
  slide_idx : Int,
  shape_idx : Int,
  after_row : Int,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    let ncols = t.col_widths.length()
    let cells : Array[@ooxml.TableCell] = []
    let mut c = 0
    while c < ncols {
      cells.push(@ooxml.TableCell::empty())
      c = c + 1
    }
    let height = if t.rows.length() > 0 {
      let ref_idx = if after_row >= 0 && after_row < t.rows.length() {
        after_row
      } else {
        0
      }
      t.rows[ref_idx].height
    } else {
      0
    }
    let new_row : @ooxml.TableRow = { height, cells }
    let insert_at = clamp_insert_index(after_row, t.rows.length())
    array_insert_at(t.rows, insert_at, new_row)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a row. At least one row must remain. Returns "OK" or "ERROR:<message>".
pub fn delete_table_row(slide_idx : Int, shape_idx : Int, row : Int) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    if t.rows.length() <= 1 {
      return "ERROR:cannot delete the last row"
    }
    if row < 0 || row >= t.rows.length() {
      return "ERROR:row index out of range"
    }
    array_remove_at(t.rows, row)
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Insert a new (empty) column. `after_col = -1` inserts at the left; otherwise
/// after the given column (clamped). `width_emu <= 0` copies a neighbour's width
/// (or 1 inch). Returns "OK:<new_col_idx>" or "ERROR:<message>".
pub fn add_table_column(
  slide_idx : Int,
  shape_idx : Int,
  after_col : Int,
  width_emu : Int,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    let ncols = t.col_widths.length()
    let insert_at = clamp_insert_index(after_col, ncols)
    let width = if width_emu > 0 {
      width_emu
    } else if ncols > 0 {
      let ref_idx = if after_col >= 0 && after_col < ncols {
        after_col
      } else {
        0
      }
      t.col_widths[ref_idx]
    } else {
      default_table_col_width_emu
    }
    array_insert_at(t.col_widths, insert_at, width)
    let mut r = 0
    while r < t.rows.length() {
      let cells = t.rows[r].cells
      // Clamp per row (defensive: rows should have one cell per grid column).
      let at = if insert_at > cells.length() {
        cells.length()
      } else {
        insert_at
      }
      array_insert_at(cells, at, @ooxml.TableCell::empty())
      r = r + 1
    }
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a column. At least one column must remain. Returns "OK" or "ERROR:...".
pub fn delete_table_column(
  slide_idx : Int,
  shape_idx : Int,
  col : Int,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    if t.col_widths.length() <= 1 {
      return "ERROR:cannot delete the last column"
    }
    if col < 0 || col >= t.col_widths.length() {
      return "ERROR:column index out of range"
    }
    array_remove_at(t.col_widths, col)
    let mut r = 0
    while r < t.rows.length() {
      let cells = t.rows[r].cells
      if col < cells.length() {
        array_remove_at(cells, col)
      }
      r = r + 1
    }
    g_modified.val[slide_idx] = true
    "OK"
  })
}
