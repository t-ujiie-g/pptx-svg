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
/// Insert an empty row after `after_row` (-1 = top), widening merged cells it falls
/// inside. Returns "OK:<new_row_idx>" or "ERROR:...".
pub fn add_table_row(
  slide_idx : Int32,
  shape_idx : Int32,
  after_row : Int32,
) -> String {
  with_table(slide_idx, shape_idx, fn(t) {
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
    let insert_at = clamp_insert_index(after_row, t.rows.length())
    t.insert_row(insert_at, height)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a row (one must remain); merged cells through it shrink, and one anchored
/// on it moves down. Returns "OK" or "ERROR:<message>".
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
    t.delete_row(row)
    g_modified.val[slide_idx] = true
    "OK"
  })
}

///|
/// Insert an empty column after `after_col` (-1 = left), widening merged cells it falls
/// inside; `width_emu <= 0` copies a neighbour's width. Returns "OK:<new_col_idx>".
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
    t.insert_column(insert_at, width)
    g_modified.val[slide_idx] = true
    "OK:" + int_to_str(insert_at)
  })
}

///|
/// Delete a column (one must remain); merged cells through it shrink, and one anchored
/// in it moves right. Returns "OK" or "ERROR:...".
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
    t.delete_column(col)
    g_modified.val[slide_idx] = true
    "OK"
  })
}
