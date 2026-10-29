///|
/// Total character length of a paragraph's concatenated run text.
fn para_text_len(para : @ooxml.TextParagraph) -> Int32 {
  let mut total = 0
  for run in para.runs {
    total = total + run.text.length()
  }
  total
}

///|
/// Split a paragraph's runs at a paragraph-level offset, keeping each run's formatting.
/// Returns (left_runs, right_runs); empty pieces are dropped.
fn split_para_runs(
  para : @ooxml.TextParagraph,
  char_off : Int32,
) -> (Array[@ooxml.TextRun], Array[@ooxml.TextRun]) {
  let left : Array[@ooxml.TextRun] = []
  let right : Array[@ooxml.TextRun] = []
  let mut acc = 0
  for run in para.runs {
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
        left.push({ ..run, text: left_text, })
      }
      if !right_text.is_empty() {
        right.push({ ..run, text: right_text, })
      }
    }
    acc = acc + len
  }
  (left, right)
}

///|
/// Replace paragraphs `[start..end]` (inclusive) of an array in-place with `items`.
fn[T] array_splice(
  arr : Array[T],
  start : Int32,
  end_inclusive : Int32,
  items : Array[T],
) -> Unit {
  let new_arr : Array[T] = []
  let len = arr.length()
  for i in 0..<len {
    if i == start {
      for item in items {
        new_arr.push(item)
      }
    }
    if i < start || i > end_inclusive {
      new_arr.push(arr[i])
    }
  }
  // start == len appends.
  if start >= len {
    for item in items {
      new_arr.push(item)
    }
  }
  while !arr.is_empty() {
    let _ = arr.pop()
  }
  for item in new_arr {
    arr.push(item)
  }
}

///|
/// A shape's text geometry as JSON (EMU; schema in `@renderer.build_text_layout`),
/// or "ERROR:<message>".
pub fn get_text_layout(slide_idx : Int32, shape_idx : Int32) -> String {
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
/// The caret position at an EMU point: `{paraIdx, runIdx, charOffset, paraOffset}`
/// (offsets within the run / paragraph), or "ERROR:<message>".
pub fn hit_test_text(
  slide_idx : Int32,
  shape_idx : Int32,
  x_emu : Int32,
  y_emu : Int32,
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
/// Replace a paragraph-offset range with `new_text`, keeping the boundary runs' formatting:
/// start == end inserts, "" deletes, `\n` splits paragraphs. Returns the shape SVG.
pub fn replace_text_range(
  slide_idx : Int32,
  shape_idx : Int32,
  start_para : Int32,
  start_char : Int32,
  end_para : Int32,
  end_char : Int32,
  new_text : String,
) -> String {
  with_shape(slide_idx, shape_idx, fn(arr, local_idx) {
    let shape = arr[local_idx]
    let nparas = shape.paragraphs.length()
    let swap = start_para > end_para ||
      (start_para == end_para && start_char > end_char)
    let sp = if swap { end_para } else { start_para }
    let ep = if swap { start_para } else { end_para }
    let sc_in = if swap { end_char } else { start_char }
    let ec_in = if swap { start_char } else { end_char }
    if sp < 0 || sp >= nparas || ep < 0 || ep >= nparas {
      return "ERROR:paragraph index out of range"
    }
    let sp_len = para_text_len(shape.paragraphs[sp])
    let ep_len = para_text_len(shape.paragraphs[ep])
    let sc = if sc_in < 0 { 0 } else if sc_in > sp_len { sp_len } else { sc_in }
    let ec = if ec_in < 0 { 0 } else if ec_in > ep_len { ep_len } else { ec_in }
    let (prefix, _) = split_para_runs(shape.paragraphs[sp], sc)
    let (_, suffix) = split_para_runs(shape.paragraphs[ep], ec)
    // Insertion format template: the run at the caret, else first run, else default.
    let base_para = shape.paragraphs[sp]
    let template : @ooxml.TextRun? = if !prefix.is_empty() {
      Some(prefix[prefix.length() - 1])
    } else if !base_para.runs.is_empty() {
      Some(base_para.runs[0])
    } else {
      None
    }
    fn mk_run(seg : String) -> @ooxml.TextRun {
      match template {
        Some(tpl) => { ..tpl, text: seg, }
        None => make_default_run(seg)
      }
    }

    fn append_all(
      dst : Array[@ooxml.TextRun],
      src : Array[@ooxml.TextRun],
    ) -> Unit {
      for run in src {
        dst.push(run)
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
      let first_runs : Array[@ooxml.TextRun] = []
      append_all(first_runs, prefix)
      if !segs[0].is_empty() {
        first_runs.push(mk_run(segs[0]))
      }
      new_paras.push({ ..base_para, runs: first_runs, })
      for k in 1..<last {
        let mid_runs : Array[@ooxml.TextRun] = []
        if !segs[k].is_empty() {
          mid_runs.push(mk_run(segs[k]))
        }
        new_paras.push({ ..base_para, runs: mid_runs, })
      }
      let last_runs : Array[@ooxml.TextRun] = []
      if !segs[last].is_empty() {
        last_runs.push(mk_run(segs[last]))
      }
      append_all(last_runs, suffix)
      new_paras.push({ ..base_para, runs: last_runs, })
    }
    array_splice(shape.paragraphs, sp, ep, new_paras)
    g_modified.val[slide_idx] = true
    render_shape_svg(slide_idx, shape_idx)
  })
}
