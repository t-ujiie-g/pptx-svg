// SVG renderer — Chart rendering
//
// Renders ChartData (parsed from ChartML / ECMA-376 Part 1 Chapter 21) as SVG.
// Supports: bar/column, line, pie, donut, scatter, area, radar, bubble, stock charts.
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

///|
let chart_bubble_max_r : Int = 20 // max bubble radius in px

///|
let chart_err_bar_cap_w : Int = 4 // half-width of error bar cap in px

///|
let chart_heatmap_min_sentinel : Int = 2147483647 // Int max sentinel for heatmap min detection

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

// Fallback series colors (used when theme is unavailable)
///|
let chart_colors_fallback : Array[String] = [
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
  "#264478", "#9B5700", "#636363", "#997300", "#255E91", "#43682B",
]

///|
let chart_data_label_font_size : Int = 9

// ── Theme color palette ──────────────────────────────────────────────────────

///|
/// Build chart color palette from theme accent colors (accent1..accent6 + darker variants).
fn build_chart_colors(theme : @ooxml.ThemeData) -> Array[String] {
  let accent1 = theme.accent1
  let accent2 = theme.accent2
  let accent3 = theme.accent3
  let accent4 = theme.accent4
  let accent5 = theme.accent5
  let accent6 = theme.accent6
  // Check if theme has real accent colors (not all none)
  if accent1.is_none() && accent2.is_none() {
    return chart_colors_fallback
  }
  let colors : Array[String] = []
  // Primary 6 accent colors
  let accents = [accent1, accent2, accent3, accent4, accent5, accent6]
  for i in 0..<6 {
    if not(accents[i].is_none()) {
      colors.push(accents[i].to_css())
    } else {
      colors.push(chart_colors_fallback[i])
    }
  }
  // Darker variants (lumMod=75000 approximation: multiply RGB by 0.75)
  for i in 0..<6 {
    if not(accents[i].is_none()) {
      let c = accents[i]
      let dr = c.r * 3 / 4
      let dg = c.g * 3 / 4
      let db = c.b * 3 / 4
      colors.push(({ r: dr, g: dg, b: db, alpha: 255 } : @ooxml.Color).to_css())
    } else {
      colors.push(chart_colors_fallback[6 + i])
    }
  }
  colors
}

// ── Main chart renderer ─────────────────────────────────────────────────────

///|
/// Detect if any chart group is a pie or donut chart (no axes needed).
fn is_pie_or_donut_chart(groups : Array[@ooxml.ChartGroup]) -> Bool {
  if groups.is_empty() {
    return false
  }
  match groups[0].chart_type {
    @ooxml.ChartKind::PieChart
    | @ooxml.ChartKind::DoughnutChart
    | @ooxml.ChartKind::OfPieChart => true
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
  theme : @ooxml.ThemeData,
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
  // Build color palette from theme
  let colors = build_chart_colors(theme)
  // Render chart groups
  for gi in 0..<chart.groups.length() {
    svg = svg + render_chart_group(
      chart.groups[gi], plot_x, plot_y, plot_w, plot_h, colors,
    )
  }
  // Render axes
  svg = svg + render_axes(chart.axes, plot_x, plot_y, plot_w, plot_h, chart.groups)
  // Render legend
  if chart.legend.show {
    svg = svg + render_legend(chart.groups, w - legend_w + 4, plot_y, colors)
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
  colors : Array[String],
) -> String {
  match group.chart_type {
    @ooxml.ChartKind::BarChart => render_bar_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::LineChart => render_line_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::PieChart => render_pie_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::DoughnutChart =>
      render_donut_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::ScatterChart =>
      render_scatter_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::AreaChart => render_area_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::RadarChart => render_line_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::BubbleChart =>
      render_bubble_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::StockChart =>
      render_stock_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::SurfaceChart =>
      render_surface_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::OfPieChart =>
      render_pie_chart(group, px_, py, pw, ph, colors)
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
  colors : Array[String],
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
  for si in 0..<ns {
    let ser = group.series[si]
    let color = series_color(ser, si, colors)
    let values = get_num_values(ser.val)
    let nv = values.length()
    let n_pts = if nv < nc { nv } else { nc }
    for ci in 0..<n_pts {
      let val = values[ci]
      if val > 0 {
        let pt_color = point_color(ser, ci, color)
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
            "\" fill=\"" + pt_color + "\"/>"
        } else {
          svg = svg + "<rect x=\"" + int_to_str(px_ + bar_off) + "\" y=\"" +
            int_to_str(py + ph - bar_len) + "\" width=\"" +
            int_to_str(if bar_w > 1 { bar_w } else { 1 }) + "\" height=\"" +
            int_to_str(bar_len) + "\" fill=\"" + pt_color + "\"/>"
        }
      }
    }
  }
  // Data labels
  if group.data_labels.is_visible() {
    svg = svg + render_bar_data_labels(
      group, px_, py, pw, ph, max_val, is_horizontal, units_per_cat, gap_pct,
    )
  }
  // Trendlines + error bars (vertical column charts only)
  if not(is_horizontal) {
    svg = svg + render_trendlines(group, px_, py, pw, ph, max_val, nc, colors)
    svg = svg + render_err_bars(group, px_, py, pw, ph, max_val, nc, colors)
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
  colors : Array[String],
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
  for si in 0..<ns {
    let ser = group.series[si]
    let color = series_color(ser, si, colors)
    let values = get_num_values(ser.val)
    let nv = values.length()
    // Build polyline points
    let mut points = ""
    let n_pts = if nv < nc { nv } else { nc }
    for ci in 0..<n_pts {
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y = py + ph - values[ci] * ph / max_val
      if not(points.is_empty()) {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
    }
    if not(points.is_empty()) {
      svg = svg + "<polyline points=\"" + points +
        "\" fill=\"none\" stroke=\"" + color +
        "\" stroke-width=\"2\"/>"
      // Draw markers
      for ci in 0..<n_pts {
        let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
        let y = py + ph - values[ci] * ph / max_val
        svg = svg + "<circle cx=\"" + int_to_str(x) + "\" cy=\"" +
          int_to_str(y) + "\" r=\"" + int_to_str(chart_line_marker_r) +
          "\" fill=\"" + color + "\"/>"
      }
    }
  }
  // Data labels
  if group.data_labels.is_visible() {
    svg = svg + render_xy_data_labels(group, px_, py, pw, ph, max_val, nc)
  }
  // Trendlines + error bars
  svg = svg + render_trendlines(group, px_, py, pw, ph, max_val, nc, colors)
  svg = svg + render_err_bars(group, px_, py, pw, ph, max_val, nc, colors)
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
  colors : Array[String],
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
  let mut svg = render_pie_slices(values, nv, cx, cy, r, 0, ser.data_points, colors)
  if group.data_labels.is_visible() {
    svg = svg + render_pie_data_labels(
      group, values, nv, cx, cy, r, 0, ser,
    )
  }
  svg
}

///|
fn render_donut_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
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
  let mut svg = render_pie_slices(values, nv, cx, cy, r, inner_r, ser.data_points, colors)
  if group.data_labels.is_visible() {
    svg = svg + render_pie_data_labels(
      group, values, nv, cx, cy, r, inner_r, ser,
    )
  }
  svg
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
  data_points : Array[@ooxml.ChartDataPoint],
  colors : Array[String],
) -> String {
  // Sum values
  let mut total = 0
  for i in 0..<nv {
    total = total + values[i]
  }
  if total == 0 {
    return ""
  }
  let mut svg = ""
  let mut start_angle = -90 * 100 // in hundredths of degrees, start from top
  for i in 0..<nv {
    let sweep = values[i] * 36000 / total // hundredths of degrees
    if sweep > 0 {
      let color = get_dpt_color(data_points, i, colors)
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
  colors : Array[String],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let (x_max, y_max, _) = get_xy_range(group.series)
  let mut svg = ""
  for si in 0..<group.series.length() {
    let ser = group.series[si]
    let color = series_color(ser, si, colors)
    let xv = get_num_values(ser.x_val)
    let yv = get_num_values(ser.y_val)
    let n = if xv.length() < yv.length() { xv.length() } else { yv.length() }
    for ci in 0..<n {
      let x = px_ + xv[ci] * pw / x_max
      let y = py + ph - yv[ci] * ph / y_max
      svg = svg + "<circle cx=\"" + int_to_str(x) + "\" cy=\"" +
        int_to_str(y) + "\" r=\"" + int_to_str(chart_scatter_marker_r) +
        "\" fill=\"" + color + "\"/>"
    }
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
  colors : Array[String],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let (_, max_val) = get_series_range(group.series)
  let nc = get_max_points(group.series)
  if nc == 0 || max_val == 0 {
    return ""
  }
  let mut svg = ""
  for si in 0..<group.series.length() {
    let ser = group.series[si]
    let color = series_color(ser, si, colors)
    let values = get_num_values(ser.val)
    let nv = values.length()
    let actual = if nv < nc { nv } else { nc }
    // Build polygon points: top line + baseline
    let mut points = ""
    for ci in 0..<actual {
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y = py + ph - values[ci] * ph / max_val
      if not(points.is_empty()) {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
    }
    // Close to baseline
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
  }
  svg
}

// ── Bubble chart ─────────────────────────────────────────────────────────

///|
fn render_bubble_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let (x_max, y_max, bub_max) = get_xy_range(group.series)
  let mut svg = ""
  for si in 0..<group.series.length() {
    let ser = group.series[si]
    let color = series_color(ser, si, colors)
    let xv = get_num_values(ser.x_val)
    let yv = get_num_values(ser.y_val)
    let bv = get_num_values(ser.bubble_size)
    let n = if xv.length() < yv.length() { xv.length() } else { yv.length() }
    for ci in 0..<n {
      let x = px_ + xv[ci] * pw / x_max
      let y = py + ph - yv[ci] * ph / y_max
      // Bubble radius proportional to bubble_size / bub_max (linear approximation)
      let bsize = if ci < bv.length() { bv[ci] } else { bub_max / 2 }
      let r = if bub_max > 0 {
        bsize * chart_bubble_max_r / bub_max
      } else {
        chart_scatter_marker_r
      }
      let r_clamped = if r < 3 { 3 } else { r }
      svg = svg + "<circle cx=\"" + int_to_str(x) + "\" cy=\"" +
        int_to_str(y) + "\" r=\"" + int_to_str(r_clamped) +
        "\" fill=\"" + color + "\" fill-opacity=\"0.6\" stroke=\"" +
        color + "\" stroke-width=\"1\"/>"
    }
  }
  svg
}

// ── Stock chart ──────────────────────────────────────────────────────────

///|
/// Render stock chart (high-low-close).
/// Expects 3 or 4 series: [open], high, low, close.
/// Renders as vertical high-low lines with close markers.
fn render_stock_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
) -> String {
  let ns = group.series.length()
  if ns < 3 {
    return ""
  }
  // Determine series indices: 3-series = H/L/C, 4-series = O/H/L/C
  let (high_idx, low_idx, close_idx, open_idx) = if ns >= 4 {
    (1, 2, 3, 0)
  } else {
    (0, 1, 2, -1)
  }
  let high_vals = get_num_values(group.series[high_idx].val)
  let low_vals = get_num_values(group.series[low_idx].val)
  let close_vals = get_num_values(group.series[close_idx].val)
  let open_vals = if open_idx >= 0 {
    get_num_values(group.series[open_idx].val)
  } else {
    []
  }
  let nc = high_vals.length()
  if nc == 0 {
    return ""
  }
  // Find max value for Y axis
  let mut max_val = 1
  for i in 0..<nc {
    if i < high_vals.length() && high_vals[i] > max_val {
      max_val = high_vals[i]
    }
  }
  // Add 10% headroom
  max_val = max_val * 11 / 10
  let mut svg = ""
  let bar_w = pw / nc / 3 // width for open/close ticks
  for ci in 0..<nc {
    let cx = px_ + (ci * 2 + 1) * pw / (nc * 2)
    let h_val = if ci < high_vals.length() { high_vals[ci] } else { 0 }
    let l_val = if ci < low_vals.length() { low_vals[ci] } else { 0 }
    let c_val = if ci < close_vals.length() { close_vals[ci] } else { 0 }
    let hy = py + ph - h_val * ph / max_val
    let ly = py + ph - l_val * ph / max_val
    let cy_ = py + ph - c_val * ph / max_val
    // High-low vertical line
    svg = svg + "<line x1=\"" + int_to_str(cx) + "\" y1=\"" +
      int_to_str(hy) + "\" x2=\"" + int_to_str(cx) + "\" y2=\"" +
      int_to_str(ly) + "\" stroke=\"" + colors[0 % colors.length()] +
      "\" stroke-width=\"1.5\"/>"
    // Close tick (right)
    svg = svg + "<line x1=\"" + int_to_str(cx) + "\" y1=\"" +
      int_to_str(cy_) + "\" x2=\"" + int_to_str(cx + bar_w) +
      "\" y2=\"" + int_to_str(cy_) + "\" stroke=\"" +
      colors[2 % colors.length()] + "\" stroke-width=\"2\"/>"
    // Open tick (left) if 4 series
    if open_idx >= 0 && ci < open_vals.length() {
      let o_val = open_vals[ci]
      let oy = py + ph - o_val * ph / max_val
      svg = svg + "<line x1=\"" + int_to_str(cx - bar_w) + "\" y1=\"" +
        int_to_str(oy) + "\" x2=\"" + int_to_str(cx) + "\" y2=\"" +
        int_to_str(oy) + "\" stroke=\"" + colors[1 % colors.length()] +
        "\" stroke-width=\"2\"/>"
    }
  }
  svg
}

// ── Surface chart (2D heatmap approximation) ────────────────────────────────

///|
/// Render surface chart as a 2D heatmap grid.
/// Each series represents one row; data points are columns.
/// Colors interpolate from cool (blue) to warm (red) based on value.
fn render_surface_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  _colors : Array[String],
) -> String {
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  // Collect all values and find global min/max
  let all_values : Array[Array[Int]] = []
  let mut global_max = 0
  let mut global_min = chart_heatmap_min_sentinel
  let mut max_cols = 0
  for si in 0..<ns {
    let vals = get_num_values(group.series[si].val)
    if vals.length() > max_cols {
      max_cols = vals.length()
    }
    for vi in 0..<vals.length() {
      if vals[vi] > global_max { global_max = vals[vi] }
      if vals[vi] < global_min { global_min = vals[vi] }
    }
    all_values.push(vals)
  }
  if max_cols == 0 || global_max == global_min {
    return ""
  }
  let cell_w = pw / max_cols
  let cell_h = ph / ns
  let range = global_max - global_min
  let mut svg = ""
  for si in 0..<ns {
    let vals = all_values[si]
    for ci in 0..<vals.length() {
      // Interpolate: 0=blue(cold), 1=red(hot)
      let t = (vals[ci] - global_min) * 255 / range
      let r = t
      let b = 255 - t
      let g = if t < 128 { t * 2 } else { (255 - t) * 2 }
      let color = "rgb(" + int_to_str(r) + "," + int_to_str(g) + "," +
        int_to_str(b) + ")"
      let x = px_ + ci * cell_w
      let y = py + si * cell_h
      svg = svg + "<rect x=\"" + int_to_str(x) + "\" y=\"" +
        int_to_str(y) + "\" width=\"" + int_to_str(cell_w) +
        "\" height=\"" + int_to_str(cell_h) + "\" fill=\"" + color +
        "\" stroke=\"" + chart_gridline_color + "\" stroke-width=\"0.5\"/>"
    }
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
  // Category / X-axis labels
  let is_xy = if not(groups.is_empty()) {
    match groups[0].chart_type {
      @ooxml.ChartKind::ScatterChart | @ooxml.ChartKind::BubbleChart => true
      _ => false
    }
  } else {
    false
  }
  if is_xy {
    // Scatter/bubble: numeric X-axis tick labels
    let x_max = get_x_axis_max(groups)
    for ti in 0..=chart_axis_ticks {
      let x = px_ + ti * pw / chart_axis_ticks
      let label_val = ti * x_max / chart_axis_ticks
      svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
        int_to_str(py + ph + 14) +
        "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
        chart_axis_label_color + "\">" +
        int_to_str(label_val) + "</text>"
    }
  } else {
    let cats = if not(groups.is_empty()) && not(groups[0].series.is_empty()) {
      get_cat_labels(groups[0].series[0].cat)
    } else {
      []
    }
    let nc = cats.length()
    if nc > 0 {
      if h_bar {
        // Horizontal bar: category labels on left (Y axis)
        for ci in 0..<nc {
          let y = py + (ci * 2 + 1) * ph / (nc * 2)
          svg = svg + "<text x=\"" + int_to_str(px_ - 4) + "\" y=\"" +
            int_to_str(y + 4) +
            "\" text-anchor=\"end\" font-size=\"10\" fill=\"" +
            chart_axis_label_color + "\">" +
            svg_escape(cats[ci]) + "</text>"
        }
      } else {
        // Vertical: category labels at bottom (X axis)
        for ci in 0..<nc {
          let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
          svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
            int_to_str(py + ph + 14) +
            "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
            chart_axis_label_color + "\">" +
            svg_escape(cats[ci]) + "</text>"
        }
      }
    }
  }
  // Value axis labels + gridlines
  for ai in 0..<axes.length() {
    let ax = axes[ai]
    if ax.is_val && not(ax.delete) {
      let max_val = get_value_axis_max(groups)
      for ti in 0..=chart_axis_ticks {
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
      }
    }
  }
  svg
}

// ── Legend rendering ─────────────────────────────────────────────────────────

///|
fn render_legend(
  groups : Array[@ooxml.ChartGroup],
  lx : Int,
  ly : Int,
  colors : Array[String],
) -> String {
  let mut svg = ""
  let mut y_off = 0
  for gi in 0..<groups.length() {
    let group = groups[gi]
    let is_pie = match group.chart_type {
      @ooxml.ChartKind::PieChart
      | @ooxml.ChartKind::DoughnutChart
      | @ooxml.ChartKind::OfPieChart => true
      _ => false
    }
    if is_pie {
      // Show category labels with per-slice colors (respecting c:dPt overrides)
      if not(group.series.is_empty()) {
        let cats = get_cat_labels(group.series[0].cat)
        let dpts = group.series[0].data_points
        for ci in 0..<cats.length() {
          let color = get_dpt_color(dpts, ci, colors)
          svg = svg + render_legend_entry(lx, ly + y_off, color, cats[ci])
          y_off = y_off + 16
        }
      }
    } else {
      for si in 0..<group.series.length() {
        let ser = group.series[si]
        let color = series_color(ser, si, colors)
        let label = if ser.title.is_empty() {
          "Series " + int_to_str(si + 1)
        } else {
          ser.title
        }
        svg = svg + render_legend_entry(lx, ly + y_off, color, label)
        y_off = y_off + 16
      }
    }
  }
  svg
}

///|
/// Render a single legend entry (color swatch + label text).
fn render_legend_entry(
  x : Int,
  y : Int,
  color : String,
  label : String,
) -> String {
  "<rect x=\"" + int_to_str(x) + "\" y=\"" + int_to_str(y) +
  "\" width=\"10\" height=\"10\" fill=\"" + color + "\"/>" +
  "<text x=\"" + int_to_str(x + 14) + "\" y=\"" + int_to_str(y + 9) +
  "\" font-size=\"10\" fill=\"" + chart_title_color + "\">" +
  svg_escape(label) + "</text>"
}

// ── Helper functions ────────────────────────────────────────────────────────

///|
/// Get series color — from spPr fill or stroke (line charts) if set, otherwise from theme palette.
fn series_color(
  ser : @ooxml.ChartSeries,
  idx : Int,
  colors : Array[String],
) -> String {
  if not(ser.sp_pr.fill.is_none()) {
    ser.sp_pr.fill.to_css()
  } else if not(ser.sp_pr.stroke.is_none()) {
    // Line charts set color on a:ln (stroke), not fill
    ser.sp_pr.stroke.to_css()
  } else {
    colors[idx % colors.length()]
  }
}

///|
/// Get per-data-point color — from c:dPt spPr if set, otherwise use series color.
fn point_color(
  ser : @ooxml.ChartSeries,
  pt_idx : Int,
  default_color : String,
) -> String {
  for i in 0..<ser.data_points.length() {
    if ser.data_points[i].idx == pt_idx {
      let sp = ser.data_points[i].sp_pr
      if not(sp.fill.is_none()) {
        return sp.fill.to_css()
      }
    }
  }
  default_color
}

///|
/// Get color for a pie/donut slice — from c:dPt if set, otherwise from palette.
fn get_dpt_color(
  data_points : Array[@ooxml.ChartDataPoint],
  idx : Int,
  colors : Array[String],
) -> String {
  for i in 0..<data_points.length() {
    if data_points[i].idx == idx {
      let sp = data_points[i].sp_pr
      if not(sp.fill.is_none()) {
        return sp.fill.to_css()
      }
    }
  }
  colors[idx % colors.length()]
}

///|
/// Extract numeric values from an AxisDataSource as integers (scaled ×1).
fn get_num_values(src : @ooxml.AxisDataSource) -> Array[Int] {
  let result : Array[Int] = []
  match src {
    @ooxml.AxisDataSource::NumSource(_, data) =>
      for i in 0..<data.points.length() {
        result.push(parse_chart_num(data.points[i].value))
      }
    _ => ()
  }
  result
}

///|
/// Get category labels from AxisDataSource.
fn get_cat_labels(src : @ooxml.AxisDataSource) -> Array[String] {
  let result : Array[String] = []
  // Both StrSource and NumSource have points with .value
  let points = match src {
    @ooxml.AxisDataSource::StrSource(_, data) => data.points
    @ooxml.AxisDataSource::NumSource(_, data) => data.points
    _ => return result
  }
  for i in 0..<points.length() {
    result.push(points[i].value)
  }
  result
}

///|
/// Get max data points across all series.
fn get_max_points(series : Array[@ooxml.ChartSeries]) -> Int {
  let mut max_n = 0
  for i in 0..<series.length() {
    let n = get_num_values(series[i].val).length()
    if n > max_n { max_n = n }
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
  for si in 0..<series.length() {
    let vals = get_num_values(series[si].val)
    for vi in 0..<vals.length() {
      if vals[vi] > max_val { max_val = vals[vi] }
    }
  }
  if max_val == 0 { max_val = 1 }
  (cats, max_val)
}

///|
/// Get X/Y/bubble-size max values across all series (scatter/bubble charts).
/// Returns (x_max, y_max, bub_max), each at least 1.
fn get_xy_range(
  series : Array[@ooxml.ChartSeries],
) -> (Int, Int, Int) {
  let mut x_max = 0
  let mut y_max = 0
  let mut bub_max = 0
  for si in 0..<series.length() {
    let xv = get_num_values(series[si].x_val)
    let yv = get_num_values(series[si].y_val)
    let bv = get_num_values(series[si].bubble_size)
    for j in 0..<xv.length() {
      if xv[j] > x_max { x_max = xv[j] }
    }
    for j in 0..<yv.length() {
      if yv[j] > y_max { y_max = yv[j] }
    }
    for j in 0..<bv.length() {
      if bv[j] > bub_max { bub_max = bv[j] }
    }
  }
  (if x_max == 0 { 1 } else { x_max }, if y_max == 0 { 1 } else { y_max }, if bub_max == 0 { 1 } else { bub_max })
}

///|
/// Get the max value for the value axis, considering chart type.
fn get_value_axis_max(groups : Array[@ooxml.ChartGroup]) -> Int {
  if groups.is_empty() {
    return 100
  }
  let group = groups[0]
  match group.chart_type {
    @ooxml.ChartKind::ScatterChart | @ooxml.ChartKind::BubbleChart => {
      let (_, y_max, _) = get_xy_range(group.series)
      y_max
    }
    @ooxml.ChartKind::StockChart => {
      let mut max_val = 0
      for si in 0..<group.series.length() {
        let vals = get_num_values(group.series[si].val)
        for vi in 0..<vals.length() {
          if vals[vi] > max_val { max_val = vals[vi] }
        }
      }
      // Add 10% headroom for stock charts
      max_val = max_val * 11 / 10
      if max_val == 0 { 1 } else { max_val }
    }
    _ => {
      // Consider all groups for composite charts
      let mut max_val = 0
      for gi in 0..<groups.length() {
        let (_, gmax) = get_series_range(groups[gi].series)
        if gmax > max_val { max_val = gmax }
      }
      if max_val == 0 { 1 } else { max_val }
    }
  }
}

///|
/// Get the max X value for scatter/bubble charts.
fn get_x_axis_max(groups : Array[@ooxml.ChartGroup]) -> Int {
  if groups.is_empty() {
    return 100
  }
  let (x_max, _, _) = get_xy_range(groups[0].series)
  x_max
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

// ── Data label rendering ─────────────────────────────────────────────────────

///|
/// Build data label text for a single data point.
fn build_label_text(
  dl : @ooxml.ChartDataLabels,
  val : Int,
  total : Int,
  cat_name : String,
  ser_name : String,
) -> String {
  let parts : Array[String] = []
  if dl.show_ser_name {
    parts.push(ser_name)
  }
  if dl.show_cat_name && not(cat_name.is_empty()) {
    parts.push(cat_name)
  }
  if dl.show_val {
    parts.push(int_to_str(val))
  }
  if dl.show_percent && total > 0 {
    let pct = val * 100 / total
    parts.push(int_to_str(pct) + "%")
  }
  let mut result = ""
  for i in 0..<parts.length() {
    if i > 0 {
      result = result + dl.separator
    }
    result = result + parts[i]
  }
  result
}

///|
/// Render data labels for bar/column charts.
fn render_bar_data_labels(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  max_val : Int,
  is_horizontal : Bool,
  units_per_cat : Int,
  gap_pct : Int,
) -> String {
  let ns = group.series.length()
  let cats = get_cat_labels(if ns > 0 { group.series[0].cat } else { @ooxml.AxisDataSource::NoData })
  let nc = cats.length()
  let dl = group.data_labels
  let cat_size = if is_horizontal { ph } else { pw }
  let val_size = if is_horizontal { pw } else { ph }
  let mut svg = ""
  for si in 0..<ns {
    let ser = group.series[si]
    let values = get_num_values(ser.val)
    let nv = values.length()
    for ci in 0..<nc {
      if ci < nv && values[ci] > 0 {
        let label = build_label_text(
          dl, values[ci], 0, if ci < cats.length() { cats[ci] } else { "" },
          ser.title,
        )
        if not(label.is_empty()) {
          let bar_len = values[ci] * val_size / max_val
          let cat_start = ci * cat_size / nc
          let bar_off = cat_start + (gap_pct / 2 + si * 100) * cat_size /
            nc / units_per_cat
          let bar_w = 100 * cat_size / nc / units_per_cat
          if is_horizontal {
            let x = px_ + bar_len + 3
            let y = py + bar_off + bar_w / 2 + 3
            svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
              int_to_str(y) + "\" font-size=\"" +
              int_to_str(chart_data_label_font_size) +
              "\" fill=\"" + chart_title_color + "\">" +
              svg_escape(label) + "</text>"
          } else {
            let x = px_ + bar_off + bar_w / 2
            let y = py + ph - bar_len - 3
            svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
              int_to_str(y) +
              "\" text-anchor=\"middle\" font-size=\"" +
              int_to_str(chart_data_label_font_size) +
              "\" fill=\"" + chart_title_color + "\">" +
              svg_escape(label) + "</text>"
          }
        }
      }
    }
  }
  svg
}

///|
/// Render data labels for line/area charts (positioned above data points).
fn render_xy_data_labels(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  max_val : Int,
  nc : Int,
) -> String {
  let ns = group.series.length()
  let cats = get_cat_labels(if ns > 0 { group.series[0].cat } else { @ooxml.AxisDataSource::NoData })
  let dl = group.data_labels
  let mut svg = ""
  for si in 0..<ns {
    let ser = group.series[si]
    let values = get_num_values(ser.val)
    let nv = values.length()
    for ci in 0..<nc {
      if ci < nv {
        let label = build_label_text(
          dl, values[ci], 0,
          if ci < cats.length() { cats[ci] } else { "" }, ser.title,
        )
        if not(label.is_empty()) {
          let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
          let y = py + ph - values[ci] * ph / max_val - 6
          svg = svg + "<text x=\"" + int_to_str(x) + "\" y=\"" +
            int_to_str(y) +
            "\" text-anchor=\"middle\" font-size=\"" +
            int_to_str(chart_data_label_font_size) +
            "\" fill=\"" + chart_title_color + "\">" +
            svg_escape(label) + "</text>"
        }
      }
    }
  }
  svg
}

///|
/// Render data labels for pie/donut charts (positioned at mid-angle of each slice).
fn render_pie_data_labels(
  group : @ooxml.ChartGroup,
  values : Array[Int],
  nv : Int,
  cx : Int,
  cy : Int,
  r : Int,
  inner_r : Int,
  ser : @ooxml.ChartSeries,
) -> String {
  let mut total = 0
  for i in 0..<nv {
    total = total + values[i]
  }
  if total == 0 {
    return ""
  }
  let dl = group.data_labels
  let cats = get_cat_labels(ser.cat)
  let label_r = if inner_r > 0 { (r + inner_r) / 2 } else { r * 2 / 3 }
  let mut svg = ""
  let mut start_angle = -90 * 100
  for i in 0..<nv {
    let sweep = values[i] * 36000 / total
    if sweep > 0 {
      let mid_angle = start_angle + sweep / 2
      let lx = cx + label_r * cos1000(mid_angle) / 1000
      let ly = cy + label_r * sin1000(mid_angle) / 1000
      let label = build_label_text(
        dl, values[i], total,
        if i < cats.length() { cats[i] } else { "" }, ser.title,
      )
      if not(label.is_empty()) {
        svg = svg + "<text x=\"" + int_to_str(lx) + "\" y=\"" +
          int_to_str(ly + 3) +
          "\" text-anchor=\"middle\" font-size=\"" +
          int_to_str(chart_data_label_font_size) +
          "\" fill=\"" + chart_title_color + "\">" +
          svg_escape(label) + "</text>"
      }
      start_angle = start_angle + sweep
    }
  }
  svg
}

// ── Trendline rendering ─────────────────────────────────────────────────────

///|
/// Render trendlines for a chart group (category-based: bar/line/area).
/// Currently supports linear trendlines (least-squares fit).
fn render_trendlines(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  max_val : Int,
  nc : Int,
  colors : Array[String],
) -> String {
  let mut svg = ""
  for si in 0..<group.series.length() {
    let ser = group.series[si]
    if ser.trendlines.is_empty() {
      continue si + 1
    }
    let values = get_num_values(ser.val)
    let nv = values.length()
    if nv < 2 {
      continue si + 1
    }
    for ti in 0..<ser.trendlines.length() {
      let tl = ser.trendlines[ti]
      let color = if not(tl.sp_pr.stroke.is_none()) {
        tl.sp_pr.stroke.to_css()
      } else {
        series_color(ser, si, colors)
      }
      match tl.trendline_type {
        "linear" => {
          // Linear regression: y = mx + b
          // Using least squares: m = (n*Σxy - Σx*Σy) / (n*Σx² - (Σx)²)
          let n = if nv < nc { nv } else { nc }
          let mut sum_x = 0
          let mut sum_y = 0
          let mut sum_xy = 0
          let mut sum_x2 = 0
          for i in 0..<n {
            sum_x = sum_x + i
            sum_y = sum_y + values[i]
            sum_xy = sum_xy + i * values[i]
            sum_x2 = sum_x2 + i * i
          }
          let denom = n * sum_x2 - sum_x * sum_x
          if denom != 0 {
            // Compute endpoints using integer math (scale by 1000 for precision)
            let m_1000 = (n * sum_xy - sum_x * sum_y) * 1000 / denom
            let b_1000 = (sum_y * 1000 - m_1000 * sum_x) / n
            // Start point (x=0)
            let y0 = b_1000
            let x_start = px_ + pw / (nc * 2)
            let y_start = py + ph - y0 * ph / (max_val * 1000)
            // End point (x=n-1)
            let yn = m_1000 * (n - 1) + b_1000
            let x_end = px_ + ((n - 1) * 2 + 1) * pw / (nc * 2)
            let y_end = py + ph - yn * ph / (max_val * 1000)
            svg = svg + "<line x1=\"" + int_to_str(x_start) + "\" y1=\"" +
              int_to_str(y_start) + "\" x2=\"" + int_to_str(x_end) +
              "\" y2=\"" + int_to_str(y_end) +
              "\" stroke=\"" + color +
              "\" stroke-width=\"1.5\" stroke-dasharray=\"6,3\"/>"
          }
        }
        // Other trendline types (exp/log/poly/power/movingAvg)
        // rendered as a simple straight line from first to last point
        _ => {
          let n = if nv < nc { nv } else { nc }
          if n >= 2 {
            let x_start = px_ + pw / (nc * 2)
            let y_start = py + ph - values[0] * ph / max_val
            let x_end = px_ + ((n - 1) * 2 + 1) * pw / (nc * 2)
            let y_end = py + ph - values[n - 1] * ph / max_val
            svg = svg + "<line x1=\"" + int_to_str(x_start) + "\" y1=\"" +
              int_to_str(y_start) + "\" x2=\"" + int_to_str(x_end) +
              "\" y2=\"" + int_to_str(y_end) +
              "\" stroke=\"" + color +
              "\" stroke-width=\"1.5\" stroke-dasharray=\"6,3\"/>"
          }
        }
      }
    }
  }
  svg
}

// ── Error bar rendering ──────────────────────────────────────────────────────

///|
/// Render error bars for a chart group (category-based: bar/line).
fn render_err_bars(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  max_val : Int,
  nc : Int,
  colors : Array[String],
) -> String {
  let mut svg = ""
  for si in 0..<group.series.length() {
    let ser = group.series[si]
    if ser.err_bars.is_none() {
      continue si + 1
    }
    let eb = ser.err_bars
    let values = get_num_values(ser.val)
    let nv = values.length()
    let color = series_color(ser, si, colors)
    for ci in 0..<nc {
      if ci >= nv {
        continue ci + 1
      }
      let val = values[ci]
      // Calculate error magnitude
      let err = match eb.err_val_type {
        "fixedVal" => eb.val
        "percentage" => val * eb.val / 100
        _ => val / 10 // fallback: 10% of value
      }
      if err <= 0 {
        continue ci + 1
      }
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y_center = py + ph - val * ph / max_val
      let show_plus = eb.err_bar_type == "both" || eb.err_bar_type == "plus"
      let show_minus = eb.err_bar_type == "both" || eb.err_bar_type == "minus"
      let y_top = if show_plus {
        py + ph - (val + err) * ph / max_val
      } else {
        y_center
      }
      let y_bottom = if show_minus {
        py + ph - (val - err) * ph / max_val
      } else {
        y_center
      }
      // Vertical line
      svg = svg + "<line x1=\"" + int_to_str(x) + "\" y1=\"" +
        int_to_str(y_top) + "\" x2=\"" + int_to_str(x) + "\" y2=\"" +
        int_to_str(y_bottom) +
        "\" stroke=\"" + color + "\" stroke-width=\"1\"/>"
      // Top cap
      if show_plus {
        svg = svg + "<line x1=\"" + int_to_str(x - chart_err_bar_cap_w) + "\" y1=\"" +
          int_to_str(y_top) + "\" x2=\"" + int_to_str(x + chart_err_bar_cap_w) +
          "\" y2=\"" + int_to_str(y_top) +
          "\" stroke=\"" + color + "\" stroke-width=\"1\"/>"
      }
      // Bottom cap
      if show_minus {
        svg = svg + "<line x1=\"" + int_to_str(x - chart_err_bar_cap_w) + "\" y1=\"" +
          int_to_str(y_bottom) + "\" x2=\"" + int_to_str(x + chart_err_bar_cap_w) +
          "\" y2=\"" + int_to_str(y_bottom) +
          "\" stroke=\"" + color + "\" stroke-width=\"1\"/>"
      }
    }
  }
  svg
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
