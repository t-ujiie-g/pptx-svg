// Merged cells (gridSpan/rowSpan/hMerge/vMerge) are not span-adjusted by row/column
// edits, so changing the structure of a table with merged cells can corrupt it.

///|
/// Width of a new column with no neighbour to copy (1 inch, EMU).
let default_table_col_width_emu : Int32 = 914400

///|
/// Call `f` with the shape's TableData, or return "ERROR:shape is not a table".
fn with_table(
  slide_idx : Int32,
  shape_idx : Int32,
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
/// Index for inserting after `after` (-1 = front), clamped to `len`.
fn clamp_insert_index(after : Int32, len : Int32) -> Int32 {
  if after < 0 {
    0
  } else if after + 1 > len {
    len
  } else {
    after + 1
  }
}

///|
/// Set a cell's text, keeping its first run's formatting. Returns the shape SVG or "ERROR:...".
pub fn update_table_cell_text(
  slide_idx : Int32,
  shape_idx : Int32,
  row : Int32,
  col : Int32,
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
    let run = if !cell.paragraphs.is_empty() &&
      !cell.paragraphs[0].runs.is_empty() {
      { ..cell.paragraphs[0].runs[0], text, }
    } else {
      make_default_run(text)
    }
    cells[col] = { ..cell, paragraphs: [make_default_paragraph(run)], }
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}

///|
/// Insert an empty row after `after_row` (-1 = top). Returns "OK:<new_row_idx>" or "ERROR:...".
pub fn add_table_row(
  slide_idx : Int32,
  shape_idx : Int32,
  after_row : Int32,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    let ncols = t.col_widths.length()
    let cells : Array[@ooxml.TableCell] = []
    for _ in 0..<ncols {
      cells.push(@ooxml.TableCell::empty())
    }
    let height = if !t.rows.is_empty() {
      let ref_idx = if after_row >= 0 && after_row < t.rows.length() {
        after_row
      } else {
        0
      }
      t.rows[ref_idx].height
    } else {
      0
    }
    let new_row : @ooxml.TableRow = { height, cells, }
    let insert_at = clamp_insert_index(after_row, t.rows.length())
    array_insert_at(t.rows, insert_at, new_row)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a row. At least one row must remain. Returns "OK" or "ERROR:<message>".
pub fn delete_table_row(
  slide_idx : Int32,
  shape_idx : Int32,
  row : Int32,
) -> String {
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
/// Insert an empty column after `after_col` (-1 = left); `width_emu <= 0` copies a
/// neighbour's width. Returns "OK:<new_col_idx>" or "ERROR:<message>".
pub fn add_table_column(
  slide_idx : Int32,
  shape_idx : Int32,
  after_col : Int32,
  width_emu : Int32,
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
    for row in t.rows {
      let cells = row.cells
      // Clamp per row: malformed tables can have short rows.
      let at = if insert_at > cells.length() {
        cells.length()
      } else {
        insert_at
      }
      array_insert_at(cells, at, @ooxml.TableCell::empty())
    }
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a column. At least one column must remain. Returns "OK" or "ERROR:...".
pub fn delete_table_column(
  slide_idx : Int32,
  shape_idx : Int32,
  col : Int32,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
    if t.col_widths.length() <= 1 {
      return "ERROR:cannot delete the last column"
    }
    if col < 0 || col >= t.col_widths.length() {
      return "ERROR:column index out of range"
    }
    array_remove_at(t.col_widths, col)
    for row in t.rows {
      let cells = row.cells
      if col < cells.length() {
        array_remove_at(cells, col)
      }
    }
    g_modified.val[slide_idx] = true
    "OK"
  })
}
