// pptx-svg: Inline text editing API (E6.2)
//
// Wasm exports backing a PowerPoint/Google-Slides–style direct-typing experience
// (double-click to edit, IME input) via a contentEditable overlay:
//   - get_text_layout   — text geometry (EMU) for caret / selection rendering
//   - hit_test_text     — click point (EMU) → caret insertion position
//   - replace_text_range — range replacement with boundary-run format preservation
//
// These read/mutate the `g_slides` cache through `with_shape` (which calls
// `ensure_slide_parsed`), so they work regardless of render order.

///|
/// Total character length of a paragraph's concatenated run text.
fn para_text_len(para : @ooxml.TextParagraph) -> Int {
  let mut total = 0
  let nr = para.runs.length()
  let mut i = 0
  while i < nr {
    total = total + para.runs[i].text.length()
    i = i + 1
  }
  total
}

///|
/// Split a paragraph's runs at a paragraph-level char offset, preserving each
/// run's formatting and splitting the boundary run. Empty pieces are skipped.
/// Returns (left_runs, right_runs).
fn split_para_runs(
  para : @ooxml.TextParagraph,
  char_off : Int,
) -> (Array[@ooxml.TextRun], Array[@ooxml.TextRun]) {
  let left : Array[@ooxml.TextRun] = []
  let right : Array[@ooxml.TextRun] = []
  let nr = para.runs.length()
  let mut acc = 0
  let mut i = 0
  while i < nr {
    let run = para.runs[i]
    let len = run.text.length()
    let run_start = acc
    let run_end = acc + len
    if char_off <= run_start {
      if len > 0 {
        right.push(run)
      }
    } else if char_off >= run_end {
      if len > 0 {
        left.push(run)
      }
    } else {
      let cut = char_off - run_start
      let left_text = substring(run.text, 0, cut)
      let right_text = substring(run.text, cut, len)
      if !left_text.is_empty() {
        left.push({ ..run, text: left_text })
      }
      if !right_text.is_empty() {
        right.push({ ..run, text: right_text })
      }
    }
    acc = acc + len
    i = i + 1
  }
  (left, right)
}

///|
/// Replace paragraphs `[start..end]` (inclusive) of an array in-place with `items`.
fn[T] array_splice(
  arr : Array[T],
  start : Int,
  end_inclusive : Int,
  items : Array[T],
) -> Unit {
  let new_arr : Array[T] = []
  let len = arr.length()
  let mut i = 0
  while i < len {
    if i == start {
      let mut k = 0
      while k < items.length() {
        new_arr.push(items[k])
        k = k + 1
      }
    }
    if i < start || i > end_inclusive {
      new_arr.push(arr[i])
    }
    i = i + 1
  }
  // start may equal len (append case): handled by the i == start check only when
  // start < len; for start == len, append items at the end.
  if start >= len {
    let mut k = 0
    while k < items.length() {
      new_arr.push(items[k])
      k = k + 1
    }
  }
  while arr.length() > 0 {
    let _ = arr.pop()
  }
  let mut j = 0
  while j < new_arr.length() {
    arr.push(new_arr[j])
    j = j + 1
  }
}

///|
/// Get the text geometry of a shape's text body as JSON. Coordinates in EMU.
/// See `@renderer.build_text_layout` for the schema and scope.
/// Returns JSON or "ERROR:<message>".
pub fn get_text_layout(slide_idx : Int, shape_idx : Int) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let scale = get_scale()
    let tl = @renderer.build_text_layout(
      shape.paragraphs,
      shape.transform,
      scale,
      g_theme.val,
      shape.body_props,
    )
    @renderer.text_layout_to_json(tl)
  })
}

///|
/// Hit-test a point (EMU) against a shape's text → caret insertion position.
/// Returns JSON `{paraIdx, runIdx, charOffset, paraOffset}` or "ERROR:<message>".
/// `charOffset` is the offset within the run; `paraOffset` is within the paragraph
/// (use the latter with `replace_text_range`).
pub fn hit_test_text(
  slide_idx : Int,
  shape_idx : Int,
  x_emu : Int,
  y_emu : Int,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let scale = get_scale()
    let tl = @renderer.build_text_layout(
      shape.paragraphs,
      shape.transform,
      scale,
      g_theme.val,
      shape.body_props,
    )
    let (p, r, c, po) = @renderer.hit_test_layout(tl, x_emu, y_emu)
    "{\"paraIdx\":" +
    int_to_str(p) +
    ",\"runIdx\":" +
    int_to_str(r) +
    ",\"charOffset\":" +
    int_to_str(c) +
    ",\"paraOffset\":" +
    int_to_str(po) +
    "}"
  })
}

///|
/// Replace a text range with `new_text`, preserving boundary run formatting.
/// `start_char`/`end_char` are paragraph-level char offsets (see `hit_test_text`
/// `paraOffset`). `\n` in `new_text` splits into multiple paragraphs. A collapsed
/// range (start == end) inserts; an empty `new_text` deletes. Returns re-rendered
/// shape SVG or "ERROR:<message>".
pub fn replace_text_range(
  slide_idx : Int,
  shape_idx : Int,
  start_para : Int,
  start_char : Int,
  end_para : Int,
  end_char : Int,
  new_text : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let nparas = shape.paragraphs.length()
    // Normalize so (sp, sc) precedes (ep, ec).
    let swap = start_para > end_para ||
      (start_para == end_para && start_char > end_char)
    let sp = if swap { end_para } else { start_para }
    let ep = if swap { start_para } else { end_para }
    let sc_in = if swap { end_char } else { start_char }
    let ec_in = if swap { start_char } else { end_char }
    if sp < 0 || sp >= nparas || ep < 0 || ep >= nparas {
      return "ERROR:paragraph index out of range"
    }
    // Clamp char offsets to paragraph text lengths.
    let sp_len = para_text_len(shape.paragraphs[sp])
    let ep_len = para_text_len(shape.paragraphs[ep])
    let sc = if sc_in < 0 { 0 } else if sc_in > sp_len { sp_len } else { sc_in }
    let ec = if ec_in < 0 { 0 } else if ec_in > ep_len { ep_len } else { ec_in }
    let (prefix, _) = split_para_runs(shape.paragraphs[sp], sc)
    let (_, suffix) = split_para_runs(shape.paragraphs[ep], ec)
    // Insertion format template: the run at the caret, else first run, else default.
    let base_para = shape.paragraphs[sp]
    let template : @ooxml.TextRun? = if prefix.length() > 0 {
      Some(prefix[prefix.length() - 1])
    } else if base_para.runs.length() > 0 {
      Some(base_para.runs[0])
    } else {
      None
    }
    fn mk_run(seg : String) -> @ooxml.TextRun {
      match template {
        Some(tpl) => { ..tpl, text: seg }
        None => make_default_run(seg)
      }
    }

    fn append_all(
      dst : Array[@ooxml.TextRun],
      src : Array[@ooxml.TextRun],
    ) -> Unit {
      let mut k = 0
      while k < src.length() {
        dst.push(src[k])
        k = k + 1
      }
    }

    let segs = split_string(new_text, '\n')
    let new_paras : Array[@ooxml.TextParagraph] = []
    if segs.length() <= 1 {
      let seg0 = if segs.length() == 1 { segs[0] } else { "" }
      let runs : Array[@ooxml.TextRun] = []
      append_all(runs, prefix)
      if !seg0.is_empty() {
        runs.push(mk_run(seg0))
      }
      append_all(runs, suffix)
      new_paras.push({ ..base_para, runs, })
    } else {
      let last = segs.length() - 1
      // First paragraph: prefix + seg0
      let first_runs : Array[@ooxml.TextRun] = []
      append_all(first_runs, prefix)
      if !segs[0].is_empty() {
        first_runs.push(mk_run(segs[0]))
      }
      new_paras.push({ ..base_para, runs: first_runs })
      // Middle paragraphs
      let mut k = 1
      while k < last {
        let mid_runs : Array[@ooxml.TextRun] = []
        if !segs[k].is_empty() {
          mid_runs.push(mk_run(segs[k]))
        }
        new_paras.push({ ..base_para, runs: mid_runs })
        k = k + 1
      }
      // Last paragraph: segLast + suffix
      let last_runs : Array[@ooxml.TextRun] = []
      if !segs[last].is_empty() {
        last_runs.push(mk_run(segs[last]))
      }
      append_all(last_runs, suffix)
      new_paras.push({ ..base_para, runs: last_runs })
    }
    array_splice(shape.paragraphs, sp, ep, new_paras)
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}
