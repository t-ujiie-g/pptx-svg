///|
/// A merged range of a table: its anchor cell and the rows × columns it spans.
priv struct MergeBlock {
  mut row : Int32
  mut col : Int32
  mut rows : Int32
  mut cols : Int32
}

///|
/// The table's merged ranges, read from the anchors' gridSpan/rowSpan. The covered
/// cells' hMerge/vMerge flags are not trusted: some generators omit hMerge.
fn TableData::merge_blocks(self : TableData) -> Array[MergeBlock] {
  let nrows = self.rows.length()
  let covered = self.rows.map(fn(r) { Array::make(r.cells.length(), false) })
  let blocks : Array[MergeBlock] = []
  for r in 0..<nrows {
    let cells = self.rows[r].cells
    for c in 0..<cells.length() {
      if covered[r][c] {
        continue
      }
      let rows = if cells[c].row_span > 1 { cells[c].row_span } else { 1 }
      let cols = if cells[c].grid_span > 1 { cells[c].grid_span } else { 1 }
      if rows == 1 && cols == 1 {
        continue
      }
      blocks.push({ row: r, col: c, rows, cols, })
      for rr in r..<(r + rows) {
        for cc in c..<(c + cols) {
          if rr < nrows && cc < covered[rr].length() {
            covered[rr][cc] = true
          }
        }
      }
    }
  }
  blocks
}

///|
/// Rewrite every cell's spans and merge flags from `blocks`, as PowerPoint writes them:
/// the first row of a block carries rowSpan, its first column gridSpan.
fn TableData::apply_merge_blocks(
  self : TableData,
  blocks : Array[MergeBlock],
) -> Unit {
  for row in self.rows {
    for c in 0..<row.cells.length() {
      row.cells[c] = {
        ..row.cells[c],
        grid_span: 1,
        row_span: 1,
        h_merge: false,
        v_merge: false,
      }
    }
  }
  for b in blocks {
    if b.rows <= 1 && b.cols <= 1 {
      continue
    }
    for r in b.row..<(b.row + b.rows) {
      if r >= self.rows.length() {
        break
      }
      let cells = self.rows[r].cells
      for c in b.col..<(b.col + b.cols) {
        if c >= cells.length() {
          break
        }
        cells[c] = {
          ..cells[c],
          grid_span: if c == b.col {
            b.cols
          } else {
            1
          },
          row_span: if r == b.row {
            b.rows
          } else {
            1
          },
          h_merge: c > b.col,
          v_merge: r > b.row,
        }
      }
    }
  }
}

///|
/// Insert an empty row at index `at` (0..=rows), widening every merged range it
/// falls inside.
pub fn TableData::insert_row(
  self : TableData,
  at : Int32,
  height : Int32,
) -> Unit {
  let blocks = self.merge_blocks()
  for b in blocks {
    if b.row < at && at < b.row + b.rows {
      b.rows = b.rows + 1
    } else if b.row >= at {
      b.row = b.row + 1
    }
  }
  let cells = Array::makei(self.col_widths.length(), fn(_) {
    TableCell::empty()
  })
  self.rows.insert(at, { height, cells, })
  self.apply_merge_blocks(blocks)
}

///|
/// Delete row `at`, shrinking the merged ranges through it. A range anchored on the
/// row moves its anchor (and the anchor's content) to the next row.
pub fn TableData::delete_row(self : TableData, at : Int32) -> Unit {
  let blocks = self.merge_blocks()
  for b in blocks {
    if b.row == at {
      if b.rows > 1 {
        let next = self.rows[at + 1].cells
        if b.col < next.length() {
          next[b.col] = self.rows[at].cells[b.col]
        }
      }
      b.rows = b.rows - 1
    } else if b.row < at && at < b.row + b.rows {
      b.rows = b.rows - 1
    } else if b.row > at {
      b.row = b.row - 1
    }
  }
  self.rows.remove(at) |> ignore
  self.apply_merge_blocks(blocks.filter(fn(b) { b.rows > 0 }))
}

///|
/// Insert an empty column of `width` EMU at index `at` (0..=cols), widening every
/// merged range it falls inside.
pub fn TableData::insert_column(
  self : TableData,
  at : Int32,
  width : Int32,
) -> Unit {
  let blocks = self.merge_blocks()
  for b in blocks {
    if b.col < at && at < b.col + b.cols {
      b.cols = b.cols + 1
    } else if b.col >= at {
      b.col = b.col + 1
    }
  }
  self.col_widths.insert(at, width)
  for row in self.rows {
    // Malformed tables can have short rows.
    row.cells.insert(
      if at > row.cells.length() {
        row.cells.length()
      } else {
        at
      },
      TableCell::empty(),
    )
  }
  self.apply_merge_blocks(blocks)
}

///|
/// Delete column `at`, shrinking the merged ranges through it. A range anchored in the
/// column moves its anchor (and the anchor's content) to the next column.
pub fn TableData::delete_column(self : TableData, at : Int32) -> Unit {
  let blocks = self.merge_blocks()
  for b in blocks {
    if b.col == at {
      if b.cols > 1 {
        let cells = self.rows[b.row].cells
        if at + 1 < cells.length() {
          cells[at + 1] = cells[at]
        }
      }
      b.cols = b.cols - 1
    } else if b.col < at && at < b.col + b.cols {
      b.cols = b.cols - 1
    } else if b.col > at {
      b.col = b.col - 1
    }
  }
  self.col_widths.remove(at) |> ignore
  for row in self.rows {
    if at < row.cells.length() {
      row.cells.remove(at) |> ignore
    }
  }
  self.apply_merge_blocks(blocks.filter(fn(b) { b.cols > 0 }))
}
