// SVG renderer — Chart rendering
//
// Renders ChartData (parsed from ChartML / ECMA-376 Part 1 Chapter 21) as SVG.
// Supports: bar/column, line, pie, donut, scatter, area, radar charts.
//
// Layout strategy:
//   - Title at top (if present)
//   - Legend at right (if enabled)
//   - Plot area fills remaining space, with axis label margins
//   - Pie/donut: no axis margins (centered in available space)
//   - Horizontal bar: category labels on left, value labels at bottom
//
// Trigonometry for pie/donut uses integer lookup table (cos_table)
// with 1° granularity × 1000 scale — sufficient for SVG path rendering.
//
// Same package as renderer.mbt — all constants and helpers are shared.

// ── Chart layout constants ──────────────────────────────────────────────────

///|
let chart_title_h : Int = 24 // title area height in px

///|
let chart_legend_w : Int = 80 // legend area width in px

///|
let chart_pad : Int = 10 // padding around chart

///|
let chart_axis_label_h : Int = 18 // space for bottom axis labels

///|
let chart_axis_label_w : Int = 40 // space for left axis labels

///|
let chart_hbar_extra_pad : Int = 20 // extra left padding for horizontal bar category labels

///|
let chart_axis_ticks : Int = 5 // number of tick divisions on value axis

///|
let chart_title_font_size : Int = 14

///|
let chart_title_y_offset : Int = 6 // title text baseline offset from title_h

///|
let chart_line_marker_r : Int = 3 // circle marker radius for line charts

///|
let chart_scatter_marker_r : Int = 4 // circle marker radius for scatter charts

///|
let chart_pie_pad : Int = 4 // radius padding for pie/donut charts

// ── Chart style constants ─────────────────────────────────────────────────

///|
let chart_bg_color : String = "white"

///|
let chart_border_color : String = "#cccccc"

///|
let chart_title_color : String = "#333333"

///|
let chart_axis_label_color : String = "#666666"

///|
let chart_gridline_color : String = "#e0e0e0"

///|
let chart_slice_stroke : String = "white"

// Default series colors (Office theme accent palette approximation)
///|
let chart_colors : Array[String] = [
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
  "#264478", "#9B5700", "#636363", "#997300", "#255E91", "#43682B",
]

// ── Main chart renderer ─────────────────────────────────────────────────────

///|
/// Detect if any chart group is a pie or donut chart (no axes needed).
fn is_pie_or_donut_chart(groups : Array[@ooxml.ChartGroup]) -> Bool {
  if groups.is_empty() {
    return false
  }
  match groups[0].chart_type {
    @ooxml.ChartKind::PieChart => true
    @ooxml.ChartKind::DoughnutChart => true
    _ => false
  }
}

///|
/// Detect if the primary chart group is a horizontal bar chart.
fn is_horizontal_bar(groups : Array[@ooxml.ChartGroup]) -> Bool {
  if groups.is_empty() {
    return false
  }
  match groups[0].chart_type {
    @ooxml.ChartKind::BarChart => groups[0].bar_dir == "bar"
    _ => false
  }
}

///|
/// Render a chart as SVG within the shape's transform bounds.
fn render_chart(
  chart : @ooxml.ChartData,
  t : @ooxml.ShapeTransform,
  scale : Int,
  _theme : @ooxml.ThemeData,
) -> String {
  let ox = px(t.x, scale)
  let oy = px(t.y, scale)
  let w = px(t.cx, scale)
  let h = px(t.cy, scale)
  if w <= 0 || h <= 0 {
    return ""
  }
  // Wrap in translate group for positioning
  let mut svg = "<g transform=\"translate(" + int_to_str(ox) + "," +
    int_to_str(oy) + ")\">"
  // Background rect
  svg = svg + "<rect width=\"" + int_to_str(w) + "\" height=\"" +
    int_to_str(h) + "\" fill=\"" + chart_bg_color + "\" stroke=\"" +
    chart_border_color + "\" stroke-width=\"1\"/>"
  // Title
  let title_h = if chart.title.is_empty() { 0 } else { chart_title_h }
  if not(chart.title.is_empty()) {
    svg = svg + "<text x=\"" + int_to_str(w / 2) + "\" y=\"" +
      int_to_str(title_h - chart_title_y_offset) +
      "\" text-anchor=\"middle\" font-size=\"" +
      int_to_str(chart_title_font_size) +
      "\" font-weight=\"bold\" fill=\"" + chart_title_color + "\">" +
      svg_escape(chart.title) + "</text>"
  }
  // Legend
  let legend_w = if chart.legend.show { chart_legend_w } else { 0 }
  // Plot area bounds — adapt based on chart type
  let no_axes = is_pie_or_donut_chart(chart.groups)
  let h_bar = is_horizontal_bar(chart.groups)
  // For horizontal bar: category labels on left (wider), value labels at bottom
  // For pie/donut: no axis labels at all
  let left_pad = if no_axes {
    chart_pad
  } else if h_bar {
    chart_pad + chart_axis_label_w + chart_hbar_extra_pad
  } else {
    chart_pad + chart_axis_label_w
  }
  let bottom_pad = if no_axes {
    chart_pad
  } else {
    chart_pad + chart_axis_label_h
  }
  let plot_x = left_pad
  let plot_y = chart_pad + title_h
  let plot_w = w - plot_x - chart_pad - legend_w
  let plot_h = h - plot_y - bottom_pad
  if plot_w <= 0 || plot_h <= 0 {
    return svg
  }
  // Render chart groups
  let ng = chart.groups.length()
  let mut gi = 0
  while gi < ng {
    let group = chart.groups[gi]
    svg = svg + render_chart_group(group, plot_x, plot_y, plot_w, plot_h)
    gi = gi + 1
  }
  // Render axes
  svg = svg + render_axes(chart.axes, plot_x, plot_y, plot_w, plot_h, chart.groups)
  // Render legend
  if chart.legend.show {
    svg = svg + render_legend(chart.groups, w - legend_w + 4, plot_y)
  }
  svg + "</g>"
}

// ── Chart group rendering ───────────────────────────────────────────────────

///|
fn render_chart_group(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  match group.chart_type {
    @ooxml.ChartKind::BarChart => render_bar_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::LineChart => render_line_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::PieChart => render_pie_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::DoughnutChart =>
      render_donut_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::ScatterChart =>
      render_scatter_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::AreaChart => render_area_chart(group, px_, py, pw, ph)
    @ooxml.ChartKind::RadarChart => render_line_chart(group, px_, py, pw, ph)
  }
}

// ── Bar / Column chart ──────────────────────────────────────────────────────

///|
fn render_bar_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  let is_horizontal = group.bar_dir == "bar"
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  // Get data range
  let (cats, max_val) = get_series_range(group.series)
  let nc = cats.length()
  if nc == 0 || max_val == 0 {
    return ""
  }
  let gap_pct = group.gap_width // percentage, e.g. 150 = 150% gap
  // Calculate bar layout
  let mut svg = ""
  // For clustered: each category has ns bars + gap
  // total units per category = ns * 100 + gap_pct
  let units_per_cat = ns * 100 + gap_pct
  let cat_size = if is_horizontal { ph } else { pw }
  let val_size = if is_horizontal { pw } else { ph }
  let mut si = 0
  while si < ns {
    let ser = group.series[si]
    let color = series_color(ser, si)
    let values = get_num_values(ser.val)
    let nv = values.length()
    let mut ci = 0
    while ci < nc && ci < nv {
      let val = values[ci]
      if val > 0 {
        let bar_len = val * val_size / max_val
        let cat_start = ci * cat_size / nc
        let bar_off = cat_start + (gap_pct / 2 + si * 100) * cat_size /
          nc /
          units_per_cat
        let bar_w = 100 * cat_size / nc / units_per_cat
        if is_horizontal {
          svg = svg + "<rect x=\"" + int_to_str(px_) + "\" y=\"" +
            int_to_str(py + bar_off) + "\" width=\"" + int_to_str(bar_len) +
            "\" height=\"" + int_to_str(if bar_w > 1 { bar_w } else { 1 }) +
            "\" fill=\"" + color + "\"/>"
        } else {
          svg = svg + "<rect x=\"" + int_to_str(px_ + bar_off) + "\" y=\"" +
            int_to_str(py + ph - bar_len) + "\" width=\"" +
            int_to_str(if bar_w > 1 { bar_w } else { 1 }) + "\" height=\"" +
            int_to_str(bar_len) + "\" fill=\"" + color + "\"/>"
        }
      }
      ci = ci + 1
    }
    si = si + 1
  }
  svg
}

// ── Line chart ──────────────────────────────────────────────────────────────

///|
fn render_line_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  let (cats, max_val) = get_series_range(group.series)
  let nc = cats.length()
  if nc == 0 || max_val == 0 {
    return ""
  }
  let mut svg = ""
  let mut si = 0
  while si < ns {
    let ser = group.series[si]
    let color = series_color(ser, si)
    let values = get_num_values(ser.val)
    let nv = values.length()
    // Build polyline points
    let mut points = ""
    let mut ci = 0
    while ci < nc && ci < nv {
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y = py + ph - values[ci] * ph / max_val
      if not(points.is_empty()) {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
      ci = ci + 1
    }
    if not(points.is_empty()) {
      svg = svg + "<polyline points=\"" + points +
        "\" fill=\"none\" stroke=\"" + color +
        "\" stroke-width=\"2\"/>"
      // Draw markers
      ci = 0
      while ci < nc && ci < nv {
        let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
        let y = py + ph - values[ci] * ph / max_val
        svg = svg + "<circle cx=\"" + int_to_str(x) + "\" cy=\"" +
          int_to_str(y) + "\" r=\"" + int_to_str(chart_line_marker_r) +
          "\" fill=\"" + color + "\"/>"
        ci = ci + 1
      }
    }
    si = si + 1
  }
  svg
}

// ── Pie chart ───────────────────────────────────────────────────────────────

///|
fn render_pie_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let ser = group.series[0]
  let values = get_num_values(ser.val)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  let cx = px_ + pw / 2
  let cy = py + ph / 2
  let r = (if pw < ph { pw } else { ph }) / 2 - chart_pie_pad
  if r <= 0 {
    return ""
  }
  render_pie_slices(values, nv, cx, cy, r, 0)
}

///|
fn render_donut_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let ser = group.series[0]
  let values = get_num_values(ser.val)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  let cx = px_ + pw / 2
  let cy = py + ph / 2
  let r = (if pw < ph { pw } else { ph }) / 2 - chart_pie_pad
  let inner_r = r * group.hole_size / 100
  if r <= 0 {
    return ""
  }
  render_pie_slices(values, nv, cx, cy, r, inner_r)
}

///|
/// Render pie/donut slices as SVG arc paths.
/// Angles use hundredths-of-degrees (e.g. -9000 = -90° = top of circle).
/// inner_r > 0 produces donut slices; inner_r == 0 produces pie slices.
fn render_pie_slices(
  values : Array[Int],
  nv : Int,
  cx : Int,
  cy : Int,
  r : Int,
  inner_r : Int,
) -> String {
  // Sum values
  let mut total = 0
  let mut i = 0
  while i < nv {
    total = total + values[i]
    i = i + 1
  }
  if total == 0 {
    return ""
  }
  let mut svg = ""
  let mut start_angle = -90 * 100 // in hundredths of degrees, start from top
  i = 0
  while i < nv {
    let sweep = values[i] * 36000 / total // hundredths of degrees
    if sweep > 0 {
      let color = chart_colors[i % chart_colors.length()]
      let end_angle = start_angle + sweep
      let large_arc = if sweep > 18000 { "1" } else { "0" }
      // Outer arc
      let sx = cx + r * cos1000(start_angle) / 1000
      let sy = cy + r * sin1000(start_angle) / 1000
      let ex = cx + r * cos1000(end_angle) / 1000
      let ey = cy + r * sin1000(end_angle) / 1000
      if inner_r > 0 {
        // Donut slice
        let isx = cx + inner_r * cos1000(start_angle) / 1000
        let isy = cy + inner_r * sin1000(start_angle) / 1000
        let iex = cx + inner_r * cos1000(end_angle) / 1000
        let iey = cy + inner_r * sin1000(end_angle) / 1000
        svg = svg + "<path d=\"M" + int_to_str(sx) + " " + int_to_str(sy) +
          " A" + int_to_str(r) + " " + int_to_str(r) + " 0 " + large_arc +
          " 1 " + int_to_str(ex) + " " + int_to_str(ey) +
          " L" + int_to_str(iex) + " " + int_to_str(iey) +
          " A" + int_to_str(inner_r) + " " + int_to_str(inner_r) +
          " 0 " + large_arc + " 0 " + int_to_str(isx) + " " +
          int_to_str(isy) + " Z\" fill=\"" + color + "\" stroke=\"" + chart_slice_stroke + "\" stroke-width=\"1\"/>"
      } else {
        // Pie slice
        svg = svg + "<path d=\"M" + int_to_str(cx) + " " + int_to_str(cy) +
          " L" + int_to_str(sx) + " " + int_to_str(sy) +
          " A" + int_to_str(r) + " " + int_to_str(r) + " 0 " + large_arc +
          " 1 " + int_to_str(ex) + " " + int_to_str(ey) +
          " Z\" fill=\"" + color + "\" stroke=\"" + chart_slice_stroke + "\" stroke-width=\"1\"/>"
      }
      start_angle = end_angle
    }
    i = i + 1
  }
  svg
}

// ── Scatter chart ───────────────────────────────────────────────────────────

///|
fn render_scatter_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  // Find X/Y range across all series
  let mut x_max = 1
  let mut y_max = 1
  let mut si = 0
  while si < ns {
    let xv = get_num_values(group.series[si].x_val)
    let yv = get_num_values(group.series[si].y_val)
    let nx = xv.length()
    let ny = yv.length()
    let mut j = 0
    while j < nx {
      if xv[j] > x_max { x_max = xv[j] }
      j = j + 1
    }
    j = 0
    while j < ny {
      if yv[j] > y_max { y_max = yv[j] }
      j = j + 1
    }
    si = si + 1
  }
  let mut svg = ""
  si = 0
  while si < ns {
    let ser = group.series[si]
    let color = series_color(ser, si)
    let xv = get_num_values(ser.x_val)
    let yv = get_num_values(ser.y_val)
    let n = if xv.length() < yv.length() { xv.length() } else { yv.length() }
    let mut ci = 0
    while ci < n {
      let x = px_ + xv[ci] * pw / x_max
      let y = py + ph - yv[ci] * ph / y_max
      svg = svg + "<circle cx=\"" + int_to_str(x) + "\" cy=\"" +
        int_to_str(y) + "\" r=\"" + int_to_str(chart_scatter_marker_r) +
        "\" fill=\"" + color + "\"/>"
      ci = ci + 1
    }
    si = si + 1
  }
  svg
}

// ── Area chart ──────────────────────────────────────────────────────────────

///|
fn render_area_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
) -> String {
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  let (_cats, max_val) = get_series_range(group.series)
  let nc = get_max_points(group.series)
  if nc == 0 || max_val == 0 {
    return ""
  }
  let mut svg = ""
  let mut si = 0
  while si < ns {
    let ser = group.series[si]
    let color = series_color(ser, si)
    let values = get_num_values(ser.val)
    let nv = values.length()
    // Build polygon points: top line + baseline
    let mut points = ""
    let mut ci = 0
    while ci < nc && ci < nv {
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y = py + ph - values[ci] * ph / max_val
      if not(points.is_empty()) {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
      ci = ci + 1
    }
    // Close to baseline
    let actual = if nv < nc { nv } else { nc }
    if actual > 0 {
      let last_x = px_ + ((actual - 1) * 2 + 1) * pw / (nc * 2)
      let first_x = px_ + pw / (nc * 2)
      points = points + " " + int_to_str(last_x) + "," +
        int_to_str(py + ph) + " " + int_to_str(first_x) + "," +
        int_to_str(py + ph)
      svg = svg + "<polygon points=\"" + points + "\" fill=\"" + color +
        "\" fill-opacity=\"0.5\" stroke=\"" + color +
        "\" stroke-width=\"1\"/>"
    }
    si = si + 1
  }
  svg
}

// ── Axes rendering ──────────────────────────────────────────────────────────

///|
fn render_axes(
  axes : Array[@ooxml.ChartAxis],
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  groups : Array[@ooxml.ChartGroup],
) -> String {
  // If no axes, this might be a pie/donut chart — skip
  if axes.is_empty() {
    return ""
  }
  let h_bar = is_horizontal_bar(groups)
  let mut svg = ""
  // Plot area border
  svg = svg + "<rect x=\"" + int_to_str(px_) + "\" y=\"" + int_to_str(py) +
    "\" width=\"" + int_to_str(pw) + "\" height=\"" + int_to_str(ph) +
    "\" fill=\"none\" stroke=\"" + chart_border_color +
    "\" stroke-width=\"0.5\"/>"
  // Category labels
  let cats = if not(groups.is_empty()) && not(groups[0].series.is_empty()) {
    get_cat_labels(groups[0].series[0].cat)
  } else {
    []
  }
  let nc = cats.length()
  if nc > 0 {
    if h_bar {
      // Horizontal bar: category labels on left (Y axis)
      let mut ci = 0
      while ci < nc {
        let y = py + (ci * 2 + 1) * ph / (nc * 2)
        svg = svg + "<text x=\"" + int_to_str(px_ - 4) + "\" y=\"" +
          int_to_str(y + 4) +
          "\" text-anchor=\"end\" font-size=\"10\" fill=\"" +
          chart_axis_label_color + "\">" +
          svg_escape(cats[ci]) + "</text>"
        ci = ci + 1
      }
    } else {
      // Vertical: category labels at bottom (X axis)
      let mut ci = 0
      while ci < nc {
        let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
        svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
          int_to_str(py + ph + 14) +
          "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
          chart_axis_label_color + "\">" +
          svg_escape(cats[ci]) + "</text>"
        ci = ci + 1
      }
    }
  }
  // Value axis labels + gridlines
  let na = axes.length()
  let mut ai = 0
  while ai < na {
    let ax = axes[ai]
    if ax.is_val && not(ax.delete) {
      let (_, max_val) = if not(groups.is_empty()) {
        get_series_range(groups[0].series)
      } else {
        ([], 100)
      }
      let mut ti = 0
      while ti <= chart_axis_ticks {
        if h_bar {
          // Horizontal bar: value labels at bottom (X axis), vertical gridlines
          let x = px_ + ti * pw / chart_axis_ticks
          if ax.major_gridlines && ti > 0 && ti < chart_axis_ticks {
            svg = svg + "<line x1=\"" + int_to_str(x) + "\" y1=\"" +
              int_to_str(py) + "\" x2=\"" + int_to_str(x) +
              "\" y2=\"" + int_to_str(py + ph) +
              "\" stroke=\"" + chart_gridline_color +
              "\" stroke-width=\"0.5\"/>"
          }
          let label_val = ti * max_val / chart_axis_ticks
          svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
            int_to_str(py + ph + 14) +
            "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
            chart_axis_label_color + "\">" +
            int_to_str(label_val) + "</text>"
        } else {
          // Vertical: value labels on left (Y axis), horizontal gridlines
          let y = py + ph - ti * ph / chart_axis_ticks
          if ax.major_gridlines && ti > 0 && ti < chart_axis_ticks {
            svg = svg + "<line x1=\"" + int_to_str(px_) + "\" y1=\"" +
              int_to_str(y) + "\" x2=\"" + int_to_str(px_ + pw) +
              "\" y2=\"" + int_to_str(y) +
              "\" stroke=\"" + chart_gridline_color +
              "\" stroke-width=\"0.5\"/>"
          }
          let label_val = ti * max_val / chart_axis_ticks
          svg = svg + "<text x=\"" + int_to_str(px_ - 4) + "\" y=\"" +
            int_to_str(y + 3) +
            "\" text-anchor=\"end\" font-size=\"10\" fill=\"" +
            chart_axis_label_color + "\">" +
            int_to_str(label_val) + "</text>"
        }
        ti = ti + 1
      }
    }
    ai = ai + 1
  }
  svg
}

// ── Legend rendering ─────────────────────────────────────────────────────────

///|
fn render_legend(
  groups : Array[@ooxml.ChartGroup],
  lx : Int,
  ly : Int,
) -> String {
  let mut svg = ""
  let mut y_off = 0
  let ng = groups.length()
  let mut gi = 0
  while gi < ng {
    let group = groups[gi]
    // For pie/donut charts, legend shows category labels (one entry per slice)
    let is_pie = match group.chart_type {
      @ooxml.ChartKind::PieChart => true
      @ooxml.ChartKind::DoughnutChart => true
      _ => false
    }
    if is_pie {
      // Show category labels with per-slice colors
      if not(group.series.is_empty()) {
        let cats = get_cat_labels(group.series[0].cat)
        let nc = cats.length()
        let mut ci = 0
        while ci < nc {
          let color = chart_colors[ci % chart_colors.length()]
          svg = svg + "<rect x=\"" + int_to_str(lx) + "\" y=\"" +
            int_to_str(ly + y_off) + "\" width=\"10\" height=\"10\" fill=\"" +
            color + "\"/>"
          svg = svg + "<text x=\"" + int_to_str(lx + 14) + "\" y=\"" +
            int_to_str(ly + y_off + 9) +
            "\" font-size=\"10\" fill=\"" + chart_title_color + "\">" + svg_escape(cats[ci]) +
            "</text>"
          y_off = y_off + 16
          ci = ci + 1
        }
      }
    } else {
      // Show series titles
      let ns = group.series.length()
      let mut si = 0
      while si < ns {
        let ser = group.series[si]
        let color = series_color(ser, si)
        let label = if ser.title.is_empty() {
          "Series " + int_to_str(si + 1)
        } else {
          ser.title
        }
        svg = svg + "<rect x=\"" + int_to_str(lx) + "\" y=\"" +
          int_to_str(ly + y_off) + "\" width=\"10\" height=\"10\" fill=\"" +
          color + "\"/>"
        svg = svg + "<text x=\"" + int_to_str(lx + 14) + "\" y=\"" +
          int_to_str(ly + y_off + 9) +
          "\" font-size=\"10\" fill=\"" + chart_title_color + "\">" + svg_escape(label) + "</text>"
        y_off = y_off + 16
        si = si + 1
      }
    }
    gi = gi + 1
  }
  svg
}

// ── Helper functions ────────────────────────────────────────────────────────

///|
/// Get series color — from spPr if set, otherwise from default palette.
fn series_color(ser : @ooxml.ChartSeries, idx : Int) -> String {
  if not(ser.sp_pr.fill.is_none()) {
    ser.sp_pr.fill.to_css()
  } else {
    chart_colors[idx % chart_colors.length()]
  }
}

///|
/// Extract numeric values from an AxisDataSource as integers (scaled ×1).
fn get_num_values(src : @ooxml.AxisDataSource) -> Array[Int] {
  let result : Array[Int] = []
  match src {
    @ooxml.AxisDataSource::NumSource(_, data) => {
      let n = data.points.length()
      let mut i = 0
      while i < n {
        result.push(parse_chart_num(data.points[i].value))
        i = i + 1
      }
    }
    _ => ()
  }
  result
}

///|
/// Get category labels from AxisDataSource.
fn get_cat_labels(src : @ooxml.AxisDataSource) -> Array[String] {
  let result : Array[String] = []
  match src {
    @ooxml.AxisDataSource::StrSource(_, data) => {
      let n = data.points.length()
      let mut i = 0
      while i < n {
        result.push(data.points[i].value)
        i = i + 1
      }
    }
    @ooxml.AxisDataSource::NumSource(_, data) => {
      let n = data.points.length()
      let mut i = 0
      while i < n {
        result.push(data.points[i].value)
        i = i + 1
      }
    }
    _ => ()
  }
  result
}

///|
/// Get max data points across all series.
fn get_max_points(series : Array[@ooxml.ChartSeries]) -> Int {
  let mut max_n = 0
  let ns = series.length()
  let mut i = 0
  while i < ns {
    let vals = get_num_values(series[i].val)
    if vals.length() > max_n {
      max_n = vals.length()
    }
    i = i + 1
  }
  max_n
}

///|
/// Get category count and max value across all series.
fn get_series_range(
  series : Array[@ooxml.ChartSeries],
) -> (Array[String], Int) {
  let cats = if not(series.is_empty()) {
    get_cat_labels(series[0].cat)
  } else {
    []
  }
  let mut max_val = 0
  let ns = series.length()
  let mut si = 0
  while si < ns {
    let vals = get_num_values(series[si].val)
    let nv = vals.length()
    let mut vi = 0
    while vi < nv {
      if vals[vi] > max_val {
        max_val = vals[vi]
      }
      vi = vi + 1
    }
    si = si + 1
  }
  // Ensure max_val is at least 1 to avoid division by zero
  if max_val == 0 {
    max_val = 1
  }
  (cats, max_val)
}

///|
/// Parse a numeric string to integer (truncating decimal part).
fn parse_chart_num(s : String) -> Int {
  if s.is_empty() {
    return 0
  }
  // Find decimal point position
  let len = s.length()
  let mut end = len
  let mut i = 0
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '.' {
      end = i
      i = len // skip to end (MoonBit has no break)
    }
    i = i + 1
  }
  // Handle negative
  let start = if len > 0 && s.get_char(0).unwrap() == '-' { 1 } else { 0 }
  let mut result = 0
  i = start
  while i < end {
    let c = s.get_char(i).unwrap()
    if c >= '0' && c <= '9' {
      result = result * 10 + (Char::to_int(c) - 48)
    }
    i = i + 1
  }
  if start == 1 { -result } else { result }
}

// ── Trigonometry (integer, hundredths of degrees) ───────────────────────────

// cos(0°..90°) × 1000, lookup table for accurate integer trig
///|
let cos_table : Array[Int] = [
  1000, 1000, 999, 999, 998, 996, 995, 993, 990, 988, // 0-9
  985, 982, 978, 974, 970, 966, 961, 956, 951, 946, // 10-19
  940, 934, 927, 921, 914, 906, 899, 891, 883, 875, // 20-29
  866, 857, 848, 839, 829, 819, 809, 799, 788, 777, // 30-39
  766, 755, 743, 731, 719, 707, 695, 682, 669, 656, // 40-49
  643, 629, 616, 602, 588, 574, 559, 545, 530, 515, // 50-59
  500, 485, 469, 454, 438, 423, 407, 391, 375, 358, // 60-69
  342, 326, 309, 292, 276, 259, 242, 225, 208, 191, // 70-79
  174, 156, 139, 122, 105, 87, 70, 52, 35, 17, // 80-89
  0, // 90
]

///|
/// Cosine for angle in hundredths of degrees.
/// Returns result × 1000 (integer approximation using lookup table).
fn cos1000(angle_hundredths : Int) -> Int {
  let deg = ((angle_hundredths / 100) % 360 + 360) % 360
  if deg <= 90 {
    cos_table[deg]
  } else if deg <= 180 {
    -cos_table[180 - deg]
  } else if deg <= 270 {
    -cos_table[deg - 180]
  } else {
    cos_table[360 - deg]
  }
}

///|
/// Sine for angle in hundredths of degrees.
/// Returns result × 1000.
fn sin1000(angle_hundredths : Int) -> Int {
  cos1000(angle_hundredths - 9000) // sin(x) = cos(x - 90°)
}
