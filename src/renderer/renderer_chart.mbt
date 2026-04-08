// SVG renderer — Chart rendering
//
// Renders ChartData (parsed from ChartML / ECMA-376 Part 1 Chapter 21) as SVG.
// Supports: bar/column, line, pie, donut, scatter, area, radar, bubble, stock,
// surface, ofPie charts (c:chart), and waterfall, treemap, sunburst, histogram,
// box & whisker, funnel charts (cx:chart / Office 2016+).
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

///|
let chart_val_scale : Int = 10000 // fixed-point scale for chart numeric values (4 decimal places)

///|
let chart_waterfall_gap_pct : Int = 20 // gap percentage for waterfall bars

///|
let chart_box_whisker_gap_pct : Int = 40 // gap percentage for box-whisker boxes

///|
let chart_funnel_gap : Int = 2 // gap between funnel bars in px

///|
let chart_sunburst_inner_pct : Int = 40 // inner radius as % of outer for sunburst

///|
let chart_pie_start_angle : Int = -9000 // -90° in hundredths (start from top)

///|
let chart_label_offset_y : Int = 12 // vertical offset for bottom axis labels

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
  "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9B5700",
  "#636363", "#997300", "#255E91", "#43682B",
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
    if !accents[i].is_none() {
      colors.push(accents[i].to_css())
    } else {
      colors.push(chart_colors_fallback[i])
    }
  }
  // Darker variants (lumMod=75000 approximation: multiply RGB by 0.75)
  for i in 0..<6 {
    if !accents[i].is_none() {
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
/// Detect if a chart needs no axes (pie, donut, treemap, sunburst, funnel).
fn is_pie_or_donut_chart(groups : Array[@ooxml.ChartGroup]) -> Bool {
  if groups.is_empty() {
    return false
  }
  match groups[0].chart_type {
    @ooxml.ChartKind::PieChart
    | @ooxml.ChartKind::DoughnutChart
    | @ooxml.ChartKind::OfPieChart
    | @ooxml.ChartKind::TreemapChart
    | @ooxml.ChartKind::SunburstChart
    | @ooxml.ChartKind::FunnelChart => true
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
  let mut svg = "<g transform=\"translate(" +
    int_to_str(ox) +
    "," +
    int_to_str(oy) +
    ")\">"
  // Background rect
  svg = svg +
    "<rect width=\"" +
    int_to_str(w) +
    "\" height=\"" +
    int_to_str(h) +
    "\" fill=\"" +
    chart_bg_color +
    "\" stroke=\"" +
    chart_border_color +
    "\" stroke-width=\"1\"/>"
  // Title
  let title_h = if chart.title.is_empty() { 0 } else { chart_title_h }
  if !chart.title.is_empty() {
    svg = svg +
      "<text x=\"" +
      int_to_str(w / 2) +
      "\" y=\"" +
      int_to_str(title_h - chart_title_y_offset) +
      "\" text-anchor=\"middle\" font-size=\"" +
      int_to_str(chart_title_font_size) +
      "\" font-weight=\"bold\" fill=\"" +
      chart_title_color +
      "\">" +
      svg_escape(chart.title) +
      "</text>"
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
  // Detect if X-axis labels need rotation (many categories)
  let has_many_cats = if !no_axes &&
    !h_bar &&
    !chart.groups.is_empty() &&
    !chart.groups[0].series.is_empty() {
    let nc = get_cat_labels(chart.groups[0].series[0].cat).length()
    nc > 8
  } else {
    false
  }
  let bottom_pad = if no_axes {
    chart_pad
  } else if has_many_cats {
    chart_pad + chart_axis_label_h + 40 // extra space for rotated labels
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
    svg = svg +
      render_chart_group(
        chart.groups[gi],
        plot_x,
        plot_y,
        plot_w,
        plot_h,
        colors,
        chart.axes,
      )
  }
  // Render axes
  svg = svg +
    render_axes(chart.axes, plot_x, plot_y, plot_w, plot_h, chart.groups)
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
  axes : Array[@ooxml.ChartAxis],
) -> String {
  match group.chart_type {
    @ooxml.ChartKind::BarChart =>
      render_bar_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::LineChart =>
      render_line_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::PieChart =>
      render_pie_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::DoughnutChart =>
      render_donut_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::ScatterChart =>
      render_scatter_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::AreaChart =>
      render_area_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::RadarChart =>
      render_line_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::BubbleChart =>
      render_bubble_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::StockChart =>
      render_stock_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::SurfaceChart =>
      render_surface_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::OfPieChart =>
      render_pie_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::WaterfallChart =>
      render_waterfall_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::TreemapChart =>
      render_treemap_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::SunburstChart =>
      render_sunburst_chart(group, px_, py, pw, ph, colors)
    @ooxml.ChartKind::HistogramChart =>
      render_bar_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::BoxWhiskerChart =>
      render_box_whisker_chart(group, px_, py, pw, ph, colors, axes)
    @ooxml.ChartKind::FunnelChart =>
      render_funnel_chart(group, px_, py, pw, ph, colors)
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
  axes : Array[@ooxml.ChartAxis],
) -> String {
  let is_horizontal = group.bar_dir == "bar"
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  let is_stacked = group.grouping == "stacked" ||
    group.grouping == "percentStacked"
  let is_pct_stacked = group.grouping == "percentStacked"
  // Get data range
  let (cats, raw_max) = if is_stacked {
    get_stacked_series_range(group.series)
  } else {
    get_series_range(group.series)
  }
  let max_val = if is_pct_stacked {
    // percentStacked: max is always 100% = 1.0 × chart_val_scale
    chart_val_scale
  } else {
    compute_axis_max(raw_max, axes)
  }
  let nc = cats.length()
  if nc == 0 || max_val == 0 {
    return ""
  }
  let gap_pct = group.gap_width // percentage, e.g. 150 = 150% gap
  let mut svg = ""
  let cat_size = if is_horizontal { ph } else { pw }
  let val_size = if is_horizontal { pw } else { ph }
  if is_stacked {
    // Stacked / percentStacked: one bar per category, series stacked
    let units_per_cat = 100 + gap_pct
    let cat_totals = compute_cat_totals(group, nc)
    // Accumulate stacked offsets per category
    let stack_offset : Array[Int] = Array::make(nc, 0)
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
          let draw_val = pct_stacked_val(val, cat_totals[ci], is_pct_stacked)
          let bar_len = mul_div(draw_val, val_size, max_val)
          let base_offset = stack_offset[ci]
          let cat_start = ci * cat_size / nc
          let bar_off = cat_start + gap_pct / 2 * cat_size / nc / units_per_cat
          let bar_w = 100 * cat_size / nc / units_per_cat
          if is_horizontal {
            svg = svg +
              "<rect x=\"" +
              int_to_str(px_ + mul_div(base_offset, val_size, max_val)) +
              "\" y=\"" +
              int_to_str(py + bar_off) +
              "\" width=\"" +
              int_to_str(bar_len) +
              "\" height=\"" +
              int_to_str(if bar_w > 1 { bar_w } else { 1 }) +
              "\" fill=\"" +
              pt_color +
              "\"/>"
          } else {
            svg = svg +
              "<rect x=\"" +
              int_to_str(px_ + bar_off) +
              "\" y=\"" +
              int_to_str(
                py + ph - mul_div(base_offset, val_size, max_val) - bar_len,
              ) +
              "\" width=\"" +
              int_to_str(if bar_w > 1 { bar_w } else { 1 }) +
              "\" height=\"" +
              int_to_str(bar_len) +
              "\" fill=\"" +
              pt_color +
              "\"/>"
          }
          stack_offset[ci] = base_offset +
            pct_stacked_val(val, cat_totals[ci], is_pct_stacked)
        }
      }
    }
  } else {
    // Clustered: each category has ns bars side by side + gap
    let units_per_cat = ns * 100 + gap_pct
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
          let bar_len = mul_div(val, val_size, max_val)
          let cat_start = ci * cat_size / nc
          let bar_off = cat_start +
            (gap_pct / 2 + si * 100) * cat_size / nc / units_per_cat
          let bar_w = 100 * cat_size / nc / units_per_cat
          if is_horizontal {
            svg = svg +
              "<rect x=\"" +
              int_to_str(px_) +
              "\" y=\"" +
              int_to_str(py + bar_off) +
              "\" width=\"" +
              int_to_str(bar_len) +
              "\" height=\"" +
              int_to_str(if bar_w > 1 { bar_w } else { 1 }) +
              "\" fill=\"" +
              pt_color +
              "\"/>"
          } else {
            svg = svg +
              "<rect x=\"" +
              int_to_str(px_ + bar_off) +
              "\" y=\"" +
              int_to_str(py + ph - bar_len) +
              "\" width=\"" +
              int_to_str(if bar_w > 1 { bar_w } else { 1 }) +
              "\" height=\"" +
              int_to_str(bar_len) +
              "\" fill=\"" +
              pt_color +
              "\"/>"
          }
        }
      }
    }
  }
  // Data labels (check group-level and series-level)
  let has_data_labels = group.data_labels.is_visible() ||
    ({
      let mut found = false
      for si2 in 0..<ns {
        if group.series[si2].data_labels.is_visible() {
          found = true
        }
      }
      found
    })
  if has_data_labels {
    svg = svg +
      render_bar_data_labels(group, px_, py, pw, ph, max_val, is_horizontal)
  }
  // Trendlines + error bars (vertical column charts only)
  if !is_horizontal {
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
  axes : Array[@ooxml.ChartAxis],
) -> String {
  let ns = group.series.length()
  if ns == 0 {
    return ""
  }
  let (cats, raw_max) = get_series_range(group.series)
  let max_val = compute_axis_max(raw_max, axes)
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
      let y = py + ph - mul_div(values[ci], ph, max_val)
      if !points.is_empty() {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
    }
    if !points.is_empty() {
      svg = svg +
        "<polyline points=\"" +
        points +
        "\" fill=\"none\" stroke=\"" +
        color +
        "\" stroke-width=\"2\"/>"
      // Draw markers
      for ci in 0..<n_pts {
        let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
        let y = py + ph - mul_div(values[ci], ph, max_val)
        svg = svg +
          "<circle cx=\"" +
          int_to_str(x) +
          "\" cy=\"" +
          int_to_str(y) +
          "\" r=\"" +
          int_to_str(chart_line_marker_r) +
          "\" fill=\"" +
          color +
          "\"/>"
      }
    }
  }
  // Data labels (check group-level and series-level)
  let has_xy_dl = group.data_labels.is_visible() ||
    ({
      let mut found = false
      for si2 in 0..<ns {
        if group.series[si2].data_labels.is_visible() {
          found = true
        }
      }
      found
    })
  if has_xy_dl {
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
  let mut svg = render_pie_slices(
    values,
    nv,
    cx,
    cy,
    r,
    0,
    ser.data_points,
    colors,
  )
  if group.data_labels.is_visible() || ser.data_labels.is_visible() {
    svg = svg + render_pie_data_labels(group, values, nv, cx, cy, r, 0, ser)
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
  let mut svg = render_pie_slices(
    values,
    nv,
    cx,
    cy,
    r,
    inner_r,
    ser.data_points,
    colors,
  )
  if group.data_labels.is_visible() || ser.data_labels.is_visible() {
    svg = svg +
      render_pie_data_labels(group, values, nv, cx, cy, r, inner_r, ser)
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
  let mut start_angle = chart_pie_start_angle
  for i in 0..<nv {
    let sweep = mul_div(values[i], 36000, total) // hundredths of degrees
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
        svg = svg +
          "<path d=\"M" +
          int_to_str(sx) +
          " " +
          int_to_str(sy) +
          " A" +
          int_to_str(r) +
          " " +
          int_to_str(r) +
          " 0 " +
          large_arc +
          " 1 " +
          int_to_str(ex) +
          " " +
          int_to_str(ey) +
          " L" +
          int_to_str(iex) +
          " " +
          int_to_str(iey) +
          " A" +
          int_to_str(inner_r) +
          " " +
          int_to_str(inner_r) +
          " 0 " +
          large_arc +
          " 0 " +
          int_to_str(isx) +
          " " +
          int_to_str(isy) +
          " Z\" fill=\"" +
          color +
          "\" stroke=\"" +
          chart_slice_stroke +
          "\" stroke-width=\"1\"/>"
      } else {
        // Pie slice
        svg = svg +
          "<path d=\"M" +
          int_to_str(cx) +
          " " +
          int_to_str(cy) +
          " L" +
          int_to_str(sx) +
          " " +
          int_to_str(sy) +
          " A" +
          int_to_str(r) +
          " " +
          int_to_str(r) +
          " 0 " +
          large_arc +
          " 1 " +
          int_to_str(ex) +
          " " +
          int_to_str(ey) +
          " Z\" fill=\"" +
          color +
          "\" stroke=\"" +
          chart_slice_stroke +
          "\" stroke-width=\"1\"/>"
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
      let x = px_ + mul_div(xv[ci], pw, x_max)
      let y = py + ph - mul_div(yv[ci], ph, y_max)
      svg = svg +
        "<circle cx=\"" +
        int_to_str(x) +
        "\" cy=\"" +
        int_to_str(y) +
        "\" r=\"" +
        int_to_str(chart_scatter_marker_r) +
        "\" fill=\"" +
        color +
        "\"/>"
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
  axes : Array[@ooxml.ChartAxis],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let (_, raw_max) = get_series_range(group.series)
  let max_val = compute_axis_max(raw_max, axes)
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
      let y = py + ph - mul_div(values[ci], ph, max_val)
      if !points.is_empty() {
        points = points + " "
      }
      points = points + int_to_str(x) + "," + int_to_str(y)
    }
    // Close to baseline
    if actual > 0 {
      let last_x = px_ + ((actual - 1) * 2 + 1) * pw / (nc * 2)
      let first_x = px_ + pw / (nc * 2)
      points = points +
        " " +
        int_to_str(last_x) +
        "," +
        int_to_str(py + ph) +
        " " +
        int_to_str(first_x) +
        "," +
        int_to_str(py + ph)
      svg = svg +
        "<polygon points=\"" +
        points +
        "\" fill=\"" +
        color +
        "\" fill-opacity=\"0.5\" stroke=\"" +
        color +
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
      let x = px_ + mul_div(xv[ci], pw, x_max)
      let y = py + ph - mul_div(yv[ci], ph, y_max)
      // Bubble radius proportional to bubble_size / bub_max (linear approximation)
      let bsize = if ci < bv.length() { bv[ci] } else { bub_max / 2 }
      let r = if bub_max > 0 {
        bsize * chart_bubble_max_r / bub_max
      } else {
        chart_scatter_marker_r
      }
      let r_clamped = if r < 3 { 3 } else { r }
      svg = svg +
        "<circle cx=\"" +
        int_to_str(x) +
        "\" cy=\"" +
        int_to_str(y) +
        "\" r=\"" +
        int_to_str(r_clamped) +
        "\" fill=\"" +
        color +
        "\" fill-opacity=\"0.6\" stroke=\"" +
        color +
        "\" stroke-width=\"1\"/>"
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
    let hy = py + ph - mul_div(h_val, ph, max_val)
    let ly = py + ph - mul_div(l_val, ph, max_val)
    let cy_ = py + ph - mul_div(c_val, ph, max_val)
    // High-low vertical line
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx) +
      "\" y1=\"" +
      int_to_str(hy) +
      "\" x2=\"" +
      int_to_str(cx) +
      "\" y2=\"" +
      int_to_str(ly) +
      "\" stroke=\"" +
      colors[0 % colors.length()] +
      "\" stroke-width=\"1.5\"/>"
    // Close tick (right)
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx) +
      "\" y1=\"" +
      int_to_str(cy_) +
      "\" x2=\"" +
      int_to_str(cx + bar_w) +
      "\" y2=\"" +
      int_to_str(cy_) +
      "\" stroke=\"" +
      colors[2 % colors.length()] +
      "\" stroke-width=\"2\"/>"
    // Open tick (left) if 4 series
    if open_idx >= 0 && ci < open_vals.length() {
      let o_val = open_vals[ci]
      let oy = py + ph - mul_div(o_val, ph, max_val)
      svg = svg +
        "<line x1=\"" +
        int_to_str(cx - bar_w) +
        "\" y1=\"" +
        int_to_str(oy) +
        "\" x2=\"" +
        int_to_str(cx) +
        "\" y2=\"" +
        int_to_str(oy) +
        "\" stroke=\"" +
        colors[1 % colors.length()] +
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
      if vals[vi] > global_max {
        global_max = vals[vi]
      }
      if vals[vi] < global_min {
        global_min = vals[vi]
      }
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
      let t = mul_div(vals[ci] - global_min, 255, range)
      let r = t
      let b = 255 - t
      let g = if t < 128 { t * 2 } else { (255 - t) * 2 }
      let color = "rgb(" +
        int_to_str(r) +
        "," +
        int_to_str(g) +
        "," +
        int_to_str(b) +
        ")"
      let x = px_ + ci * cell_w
      let y = py + si * cell_h
      svg = svg +
        "<rect x=\"" +
        int_to_str(x) +
        "\" y=\"" +
        int_to_str(y) +
        "\" width=\"" +
        int_to_str(cell_w) +
        "\" height=\"" +
        int_to_str(cell_h) +
        "\" fill=\"" +
        color +
        "\" stroke=\"" +
        chart_gridline_color +
        "\" stroke-width=\"0.5\"/>"
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
  svg = svg +
    "<rect x=\"" +
    int_to_str(px_) +
    "\" y=\"" +
    int_to_str(py) +
    "\" width=\"" +
    int_to_str(pw) +
    "\" height=\"" +
    int_to_str(ph) +
    "\" fill=\"none\" stroke=\"" +
    chart_border_color +
    "\" stroke-width=\"0.5\"/>"
  // Category / X-axis labels
  let is_xy = if !groups.is_empty() {
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
    for ti in 0..<=chart_axis_ticks {
      let x = px_ + ti * pw / chart_axis_ticks
      let label_val = mul_div(ti, x_max, chart_axis_ticks)
      svg = svg +
        "<text x=\"" +
        int_to_str(x) +
        "\" y=\"" +
        int_to_str(py + ph + 14) +
        "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
        chart_axis_label_color +
        "\">" +
        format_chart_val(label_val, "General") +
        "</text>"
    }
  } else {
    let cats = if !groups.is_empty() && !groups[0].series.is_empty() {
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
          svg = svg +
            "<text x=\"" +
            int_to_str(px_ - 4) +
            "\" y=\"" +
            int_to_str(y + 4) +
            "\" text-anchor=\"end\" font-size=\"10\" fill=\"" +
            chart_axis_label_color +
            "\">" +
            svg_escape(cats[ci]) +
            "</text>"
        }
      } else {
        // Vertical: category labels at bottom (X axis)
        // Rotate labels when there are many categories to avoid overlap
        let rotate = nc > 8
        for ci in 0..<nc {
          let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
          if rotate {
            svg = svg +
              "<text x=\"" +
              int_to_str(x) +
              "\" y=\"" +
              int_to_str(py + ph + 10) +
              "\" text-anchor=\"end\" font-size=\"8\" fill=\"" +
              chart_axis_label_color +
              "\" transform=\"rotate(-45," +
              int_to_str(x) +
              "," +
              int_to_str(py + ph + 10) +
              ")\">" +
              svg_escape(cats[ci]) +
              "</text>"
          } else {
            svg = svg +
              "<text x=\"" +
              int_to_str(x) +
              "\" y=\"" +
              int_to_str(py + ph + 14) +
              "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
              chart_axis_label_color +
              "\">" +
              svg_escape(cats[ci]) +
              "</text>"
          }
        }
      }
    }
  }
  // Value axis labels + gridlines
  for ai in 0..<axes.length() {
    let ax = axes[ai]
    if ax.is_val && !ax.delete {
      let raw_max = get_value_axis_max(groups)
      let val_fmt = ax.num_fmt
      let max_val = compute_axis_max(raw_max, axes)
      // Determine tick count from majorUnit if specified
      let major_unit = if ax.major_unit.is_empty() {
        0
      } else {
        parse_chart_num(ax.major_unit)
      }
      let n_ticks = if major_unit > 0 && max_val > 0 {
        let n = max_val / major_unit
        if n > 20 {
          20
        } else if n < 1 {
          1
        } else {
          n
        }
      } else {
        chart_axis_ticks
      }
      for ti in 0..<=n_ticks {
        if h_bar {
          // Horizontal bar: value labels at bottom (X axis), vertical gridlines
          let x = px_ + ti * pw / n_ticks
          if ax.major_gridlines && ti > 0 && ti < n_ticks {
            svg = svg +
              "<line x1=\"" +
              int_to_str(x) +
              "\" y1=\"" +
              int_to_str(py) +
              "\" x2=\"" +
              int_to_str(x) +
              "\" y2=\"" +
              int_to_str(py + ph) +
              "\" stroke=\"" +
              chart_gridline_color +
              "\" stroke-width=\"0.5\"/>"
          }
          let label_val = mul_div(ti, max_val, n_ticks)
          svg = svg +
            "<text x=\"" +
            int_to_str(x) +
            "\" y=\"" +
            int_to_str(py + ph + 14) +
            "\" text-anchor=\"middle\" font-size=\"10\" fill=\"" +
            chart_axis_label_color +
            "\">" +
            format_chart_val(label_val, val_fmt) +
            "</text>"
        } else {
          // Vertical: value labels on left (Y axis), horizontal gridlines
          let y = py + ph - ti * ph / n_ticks
          if ax.major_gridlines && ti > 0 && ti < n_ticks {
            svg = svg +
              "<line x1=\"" +
              int_to_str(px_) +
              "\" y1=\"" +
              int_to_str(y) +
              "\" x2=\"" +
              int_to_str(px_ + pw) +
              "\" y2=\"" +
              int_to_str(y) +
              "\" stroke=\"" +
              chart_gridline_color +
              "\" stroke-width=\"0.5\"/>"
          }
          let label_val = mul_div(ti, max_val, n_ticks)
          svg = svg +
            "<text x=\"" +
            int_to_str(px_ - 4) +
            "\" y=\"" +
            int_to_str(y + 3) +
            "\" text-anchor=\"end\" font-size=\"10\" fill=\"" +
            chart_axis_label_color +
            "\">" +
            format_chart_val(label_val, val_fmt) +
            "</text>"
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
      if !group.series.is_empty() {
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
  "<rect x=\"" +
  int_to_str(x) +
  "\" y=\"" +
  int_to_str(y) +
  "\" width=\"10\" height=\"10\" fill=\"" +
  color +
  "\"/>" +
  "<text x=\"" +
  int_to_str(x + 14) +
  "\" y=\"" +
  int_to_str(y + 9) +
  "\" font-size=\"10\" fill=\"" +
  chart_title_color +
  "\">" +
  svg_escape(label) +
  "</text>"
}

// ── Helper functions ────────────────────────────────────────────────────────

///|
/// Get series color — from spPr fill or stroke (line charts) if set, otherwise from theme palette.
fn series_color(
  ser : @ooxml.ChartSeries,
  idx : Int,
  colors : Array[String],
) -> String {
  if !ser.sp_pr.fill.is_none() {
    ser.sp_pr.fill.to_css()
  } else if !ser.sp_pr.stroke.is_none() {
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
      if !sp.fill.is_none() {
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
      if !sp.fill.is_none() {
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
/// Get per-category stacked totals and max across categories.
fn get_stacked_series_range(
  series : Array[@ooxml.ChartSeries],
) -> (Array[String], Int) {
  let cats = if !series.is_empty() { get_cat_labels(series[0].cat) } else { [] }
  let nc = cats.length()
  let totals : Array[Int] = Array::make(nc, 0)
  for si in 0..<series.length() {
    let vals = get_num_values(series[si].val)
    for ci in 0..<nc {
      if ci < vals.length() {
        totals[ci] = totals[ci] + vals[ci]
      }
    }
  }
  let mut max_val = 0
  for ci in 0..<nc {
    if totals[ci] > max_val {
      max_val = totals[ci]
    }
  }
  if max_val == 0 {
    max_val = 1
  }
  (cats, max_val)
}

///|
/// Compute per-category totals for percentStacked normalization.
fn compute_cat_totals(group : @ooxml.ChartGroup, nc : Int) -> Array[Int] {
  let totals : Array[Int] = Array::make(nc, 0)
  let ns = group.series.length()
  for si in 0..<ns {
    let vals = get_num_values(group.series[si].val)
    for ci in 0..<nc {
      if ci < vals.length() {
        totals[ci] = totals[ci] + vals[ci]
      }
    }
  }
  totals
}

///|
/// Normalize a value for percentStacked: val * chart_val_scale / total.
/// Returns the raw val if not percentStacked or total is 0.
fn pct_stacked_val(val : Int, cat_total : Int, is_pct_stacked : Bool) -> Int {
  if is_pct_stacked && cat_total > 0 {
    mul_div(val, chart_val_scale, cat_total)
  } else {
    val
  }
}

///|
/// Get the numeric format code from an AxisDataSource (e.g. "0%", "General").
fn get_num_format(src : @ooxml.AxisDataSource) -> String {
  match src {
    @ooxml.AxisDataSource::NumSource(_, data) =>
      if data.format_code.is_empty() {
        "General"
      } else {
        data.format_code
      }
    _ => "General"
  }
}

///|
/// Get max data points across all series.
fn get_max_points(series : Array[@ooxml.ChartSeries]) -> Int {
  let mut max_n = 0
  for i in 0..<series.length() {
    let n = get_num_values(series[i].val).length()
    if n > max_n {
      max_n = n
    }
  }
  max_n
}

///|
/// Round up a tick interval to a nice number from {1,2,5}×10^n series.
/// E.g. 762 → 1000, 350 → 500, 18 → 20.
fn nice_ceil(val : Int) -> Int {
  if val <= 0 {
    return 1
  }
  // Find magnitude
  let mut mag = 1
  while mag * 10 <= val {
    mag = mag * 10
  }
  // val / mag is 1..9; pick the next nice number
  let leading = (val + mag - 1) / mag // ceil(val/mag)
  if leading <= 1 {
    mag
  } else if leading <= 2 {
    mag * 2
  } else if leading <= 5 {
    mag * 5
  } else {
    mag * 10
  }
}

///|
/// Round up max value to produce clean axis tick labels.
/// Uses nice_ceil on the tick interval so labels are round numbers.
fn round_up_nice(val : Int) -> Int {
  if val <= 0 {
    return chart_val_scale
  }
  let raw_interval = (val + chart_axis_ticks - 1) / chart_axis_ticks
  let nice_interval = nice_ceil(raw_interval)
  nice_interval * chart_axis_ticks
}

///|
/// Get category count and max value across all series.
fn get_series_range(series : Array[@ooxml.ChartSeries]) -> (Array[String], Int) {
  let cats = if !series.is_empty() { get_cat_labels(series[0].cat) } else { [] }
  let mut max_val = 0
  for si in 0..<series.length() {
    let vals = get_num_values(series[si].val)
    for vi in 0..<vals.length() {
      if vals[vi] > max_val {
        max_val = vals[vi]
      }
    }
  }
  if max_val == 0 {
    max_val = 1
  }
  (cats, max_val)
}

///|
/// Get X/Y/bubble-size max values across all series (scatter/bubble charts).
/// Returns (x_max, y_max, bub_max), each at least 1.
fn get_xy_range(series : Array[@ooxml.ChartSeries]) -> (Int, Int, Int) {
  let mut x_max = 0
  let mut y_max = 0
  let mut bub_max = 0
  for si in 0..<series.length() {
    let xv = get_num_values(series[si].x_val)
    let yv = get_num_values(series[si].y_val)
    let bv = get_num_values(series[si].bubble_size)
    for j in 0..<xv.length() {
      if xv[j] > x_max {
        x_max = xv[j]
      }
    }
    for j in 0..<yv.length() {
      if yv[j] > y_max {
        y_max = yv[j]
      }
    }
    for j in 0..<bv.length() {
      if bv[j] > bub_max {
        bub_max = bv[j]
      }
    }
  }
  (
    if x_max == 0 {
      1
    } else {
      x_max
    },
    if y_max == 0 {
      1
    } else {
      y_max
    },
    if bub_max == 0 {
      1
    } else {
      bub_max
    },
  )
}

///|
/// Compute effective axis max from raw data max, using majorUnit if available,
/// otherwise round_up_nice.
fn compute_axis_max(raw_max : Int, axes : Array[@ooxml.ChartAxis]) -> Int {
  // Find the value axis majorUnit
  let mut major_unit = 0
  for ai in 0..<axes.length() {
    if axes[ai].is_val && !axes[ai].major_unit.is_empty() {
      major_unit = parse_chart_num(axes[ai].major_unit)
    }
  }
  if major_unit > 0 {
    let n = (raw_max + major_unit - 1) / major_unit
    let n_capped = if n > 20 { 20 } else if n < 1 { 1 } else { n }
    n_capped * major_unit
  } else {
    round_up_nice(raw_max)
  }
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
          if vals[vi] > max_val {
            max_val = vals[vi]
          }
        }
      }
      // Add 10% headroom for stock charts
      max_val = max_val * 11 / 10
      if max_val == 0 {
        1
      } else {
        max_val
      }
    }
    _ => {
      // Consider all groups for composite charts
      let mut max_val = 0
      for gi in 0..<groups.length() {
        let g = groups[gi]
        if g.grouping == "percentStacked" {
          // percentStacked: max is always 100%
          max_val = chart_val_scale
        } else if g.grouping == "stacked" {
          let (_, gmax) = get_stacked_series_range(g.series)
          if gmax > max_val {
            max_val = gmax
          }
        } else {
          let (_, gmax) = get_series_range(g.series)
          if gmax > max_val {
            max_val = gmax
          }
        }
      }
      if max_val == 0 {
        1
      } else {
        max_val
      }
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
/// Parse a numeric string to fixed-point integer (×chart_val_scale).
/// Handles: "0.381" → 3810, "38.1" → 381000, "500" → 5000000.
/// Also handles scientific notation: "1.77E-2" → 177.
fn parse_chart_num(s : String) -> Int {
  if s.is_empty() {
    return 0
  }
  let len = s.length()
  // Handle negative
  let start = if len > 0 && s.get_char(0).unwrap() == '-' { 1 } else { 0 }
  // Detect scientific notation (E or e)
  let mut e_pos = -1
  let mut si = start
  while si < len {
    let c = s.get_char(si).unwrap()
    if c == 'E' || c == 'e' {
      e_pos = si
      si = len // break
    }
    si = si + 1
  }
  if e_pos >= 0 {
    // Scientific notation: parse mantissa and exponent separately
    // Parse mantissa digits (ignoring dot, max 8 significant digits to avoid overflow)
    let mut mantissa = 0
    let mut mant_dec_digits = 0
    let mut found_dot = false
    let mut sig_digits = 0
    si = start
    while si < e_pos {
      let c = s.get_char(si).unwrap()
      if c == '.' {
        found_dot = true
      } else if c >= '0' && c <= '9' && sig_digits < 8 {
        mantissa = mantissa * 10 + (Char::to_int(c) - 48)
        sig_digits = sig_digits + 1
        if found_dot {
          mant_dec_digits = mant_dec_digits + 1
        }
      }
      si = si + 1
    }
    // Parse exponent (may be negative)
    let mut exp = 0
    let mut exp_neg = false
    si = e_pos + 1
    if si < len {
      let ec = s.get_char(si).unwrap()
      if ec == '-' {
        exp_neg = true
        si = si + 1
      } else if ec == '+' {
        si = si + 1
      }
    }
    while si < len {
      let c = s.get_char(si).unwrap()
      if c >= '0' && c <= '9' {
        exp = exp * 10 + (Char::to_int(c) - 48)
      }
      si = si + 1
    }
    if exp_neg {
      exp = -exp
    }
    // Net decimal shift: number of decimal places after applying exponent
    // E.g. 1.7738E-2: mantissa=17738, mant_dec_digits=4, exp=-2
    // Actual value = 17738 × 10^(-4-2) = 17738 × 10^-6 = 0.017738
    // We need result × 10^-4 (chart_val_scale = 10^4)
    // So result = mantissa × 10^(exp - mant_dec_digits + 4)
    let shift = exp - mant_dec_digits + 4
    let mut result = mantissa
    if shift >= 0 {
      let mut p = 0
      while p < shift {
        result = result * 10
        p = p + 1
      }
    } else {
      let mut p = 0
      while p < -shift {
        result = result / 10
        p = p + 1
      }
    }
    return if start == 1 { -result } else { result }
  }
  // Standard decimal parsing
  // Find decimal point position
  let mut dot_pos = len
  let mut i = start
  while i < len {
    let c = s.get_char(i).unwrap()
    if c == '.' {
      dot_pos = i
      i = len // break
    }
    i = i + 1
  }
  // Parse integer part
  let mut int_part = 0
  i = start
  while i < dot_pos {
    let c = s.get_char(i).unwrap()
    if c >= '0' && c <= '9' {
      int_part = int_part * 10 + (Char::to_int(c) - 48)
    }
    i = i + 1
  }
  // Parse decimal part (up to 4 digits for scale=10000)
  let mut dec_part = 0
  let mut dec_digits = 0
  i = dot_pos + 1
  while i < len && dec_digits < 4 {
    let c = s.get_char(i).unwrap()
    if c >= '0' && c <= '9' {
      dec_part = dec_part * 10 + (Char::to_int(c) - 48)
      dec_digits = dec_digits + 1
    }
    i = i + 1
  }
  // Pad remaining decimal places with zeros
  while dec_digits < 4 {
    dec_part = dec_part * 10
    dec_digits = dec_digits + 1
  }
  let result = int_part * chart_val_scale + dec_part
  if start == 1 {
    -result
  } else {
    result
  }
}

///|
/// Safe a*b/c avoiding 32-bit overflow.
/// Uses: a*b/c = (a/c)*b + (a%c)*b/c.
fn gcd(a : Int, b : Int) -> Int {
  let mut x = if a < 0 { -a } else { a }
  let mut y = if b < 0 { -b } else { b }
  while y != 0 {
    let t = y
    y = x % y
    x = t
  }
  x
}

///|
fn mul_div(a : Int, b : Int, c : Int) -> Int {
  if c == 0 {
    return 0
  }
  // Reduce a/c and b/c by GCD to avoid overflow in a*b
  let g1 = gcd(a, c)
  let a1 = a / g1
  let c1 = c / g1
  let g2 = gcd(b, c1)
  let b1 = b / g2
  let c2 = c1 / g2
  // Now a1 * b1 / c2 with minimal overflow risk
  if c2 == 0 {
    return 0
  }
  a1 * b1 / c2
}

///|
/// Format a scaled chart value for axis/data label display.
/// Handles "0%", "0.0%", "0.00%", "General", etc.
fn format_chart_val(val : Int, num_fmt : String) -> String {
  // Check if percentage format
  let is_pct = num_fmt.length() > 0 &&
    ({
      let last = num_fmt.get_char(num_fmt.length() - 1).unwrap()
      last == '%'
    })
  if is_pct {
    // Percentage: multiply by 100 then format
    // val is scaled ×10000, so 0.381 = 3810.
    // ×100 = 381000 / 10000 = 38.1
    let pct_scaled = val * 100 // now scaled ×10000 of percentage value
    // Count decimal places in format (e.g. "0%" = 0, "0.0%" = 1, "0.00%" = 2)
    let mut dec_places = 0
    let mut found_dot = false
    let mut fi = 0
    while fi < num_fmt.length() {
      let c = num_fmt.get_char(fi).unwrap()
      if c == '.' {
        found_dot = true
      } else if found_dot && c == '0' {
        dec_places = dec_places + 1
      } else if c == '%' {
        fi = num_fmt.length() // break
      }
      fi = fi + 1
    }
    if dec_places == 0 {
      // Round to integer percentage
      let int_pct = (pct_scaled + chart_val_scale / 2) / chart_val_scale
      int_to_str(int_pct) + "%"
    } else {
      // Show with decimal places, properly rounded
      // For N decimal places, round to 10^N resolution then split
      let round_div = if dec_places == 1 {
        1000 // chart_val_scale / 10
      } else if dec_places == 2 {
        100 // chart_val_scale / 100
      } else {
        10
      }
      let rounded = (pct_scaled + round_div / 2) / round_div
      let pows = if dec_places == 1 {
        10
      } else if dec_places == 2 {
        100
      } else {
        1000
      }
      let int_pct = rounded / pows
      let frac = if rounded >= 0 {
        rounded - int_pct * pows
      } else {
        -(rounded - int_pct * pows)
      }
      let frac_str = if dec_places == 1 {
        int_to_str(frac)
      } else if dec_places == 2 {
        (if frac < 10 { "0" } else { "" }) + int_to_str(frac)
      } else {
        (if frac < 100 { "0" } else { "" }) +
        (if frac < 10 { "0" } else { "" }) +
        int_to_str(frac)
      }
      int_to_str(int_pct) + "." + frac_str + "%"
    }
  } else {
    // General format: just show the integer part
    let int_val = val / chart_val_scale
    int_to_str(int_val)
  }
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
  num_fmt : String,
) -> String {
  let parts : Array[String] = []
  if dl.show_ser_name {
    parts.push(ser_name)
  }
  if dl.show_cat_name && !cat_name.is_empty() {
    parts.push(cat_name)
  }
  if dl.show_val {
    parts.push(format_chart_val(val, num_fmt))
  }
  if dl.show_percent && total > 0 {
    let pct = mul_div(val, 100 * chart_val_scale, total)
    parts.push(format_chart_val(pct, "0%"))
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
) -> String {
  let ns = group.series.length()
  let cats = get_cat_labels(
    if ns > 0 {
      group.series[0].cat
    } else {
      @ooxml.AxisDataSource::NoData
    },
  )
  let nc = cats.length()
  let group_dl = group.data_labels
  let cat_size = if is_horizontal { ph } else { pw }
  let val_size = if is_horizontal { pw } else { ph }
  let is_stacked = group.grouping == "stacked" ||
    group.grouping == "percentStacked"
  let is_pct_stacked = group.grouping == "percentStacked"
  let gap_pct = group.gap_width
  let units_per_cat = if is_stacked {
    100 + gap_pct
  } else {
    ns * 100 + gap_pct
  }
  let cat_totals = compute_cat_totals(group, nc)
  // Track stacked offset for label positioning
  let stack_offset : Array[Int] = Array::make(nc, 0)
  let mut svg = ""
  for si in 0..<ns {
    let ser = group.series[si]
    let dl = if ser.data_labels.is_visible() {
      ser.data_labels
    } else {
      group_dl
    }
    if !dl.is_visible() {
      continue
    }
    let values = get_num_values(ser.val)
    let val_fmt = get_num_format(ser.val)
    let nv = values.length()
    for ci in 0..<nc {
      if ci < nv && values[ci] > 0 {
        let label = build_label_text(
          dl,
          values[ci],
          0,
          if ci < cats.length() {
            cats[ci]
          } else {
            ""
          },
          ser.title,
          val_fmt,
        )
        if !label.is_empty() {
          let draw_val = pct_stacked_val(
            values[ci],
            cat_totals[ci],
            is_pct_stacked,
          )
          let bar_len = mul_div(draw_val, val_size, max_val)
          let cat_start = ci * cat_size / nc
          if is_stacked {
            let base_px = mul_div(stack_offset[ci], val_size, max_val)
            let bar_off = cat_start +
              gap_pct / 2 * cat_size / nc / units_per_cat
            let bar_w = 100 * cat_size / nc / units_per_cat
            if is_horizontal {
              let x = px_ + base_px + bar_len / 2
              let y = py + bar_off + bar_w / 2 + 3
              svg = svg +
                "<text x=\"" +
                int_to_str(x) +
                "\" y=\"" +
                int_to_str(y) +
                "\" text-anchor=\"middle\" font-size=\"" +
                int_to_str(chart_data_label_font_size) +
                "\" fill=\"" +
                chart_title_color +
                "\">" +
                svg_escape(label) +
                "</text>"
            } else {
              let x = px_ + bar_off + bar_w / 2
              let y = py + ph - base_px - bar_len / 2 + 3
              svg = svg +
                "<text x=\"" +
                int_to_str(x) +
                "\" y=\"" +
                int_to_str(y) +
                "\" text-anchor=\"middle\" font-size=\"" +
                int_to_str(chart_data_label_font_size) +
                "\" fill=\"" +
                chart_title_color +
                "\">" +
                svg_escape(label) +
                "</text>"
            }
          } else {
            let bar_off = cat_start +
              (gap_pct / 2 + si * 100) * cat_size / nc / units_per_cat
            let bar_w = 100 * cat_size / nc / units_per_cat
            if is_horizontal {
              let x = px_ + bar_len + 3
              let y = py + bar_off + bar_w / 2 + 3
              svg = svg +
                "<text x=\"" +
                int_to_str(x) +
                "\" y=\"" +
                int_to_str(y) +
                "\" font-size=\"" +
                int_to_str(chart_data_label_font_size) +
                "\" fill=\"" +
                chart_title_color +
                "\">" +
                svg_escape(label) +
                "</text>"
            } else {
              let x = px_ + bar_off + bar_w / 2
              let y = py + ph - bar_len - 3
              svg = svg +
                "<text x=\"" +
                int_to_str(x) +
                "\" y=\"" +
                int_to_str(y) +
                "\" text-anchor=\"middle\" font-size=\"" +
                int_to_str(chart_data_label_font_size) +
                "\" fill=\"" +
                chart_title_color +
                "\">" +
                svg_escape(label) +
                "</text>"
            }
          }
        }
        // Update stacked offset (must be outside label visibility check)
        if is_stacked {
          stack_offset[ci] = stack_offset[ci] +
            pct_stacked_val(values[ci], cat_totals[ci], is_pct_stacked)
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
  let cats = get_cat_labels(
    if ns > 0 {
      group.series[0].cat
    } else {
      @ooxml.AxisDataSource::NoData
    },
  )
  let group_dl = group.data_labels
  let mut svg = ""
  for si in 0..<ns {
    let ser = group.series[si]
    let dl = if ser.data_labels.is_visible() {
      ser.data_labels
    } else {
      group_dl
    }
    if !dl.is_visible() {
      continue
    }
    let values = get_num_values(ser.val)
    let val_fmt = get_num_format(ser.val)
    let nv = values.length()
    for ci in 0..<nc {
      if ci < nv {
        let label = build_label_text(
          dl,
          values[ci],
          0,
          if ci < cats.length() {
            cats[ci]
          } else {
            ""
          },
          ser.title,
          val_fmt,
        )
        if !label.is_empty() {
          let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
          let y = py + ph - mul_div(values[ci], ph, max_val) - 6
          svg = svg +
            "<text x=\"" +
            int_to_str(x) +
            "\" y=\"" +
            int_to_str(y) +
            "\" text-anchor=\"middle\" font-size=\"" +
            int_to_str(chart_data_label_font_size) +
            "\" fill=\"" +
            chart_title_color +
            "\">" +
            svg_escape(label) +
            "</text>"
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
  let dl = if ser.data_labels.is_visible() {
    ser.data_labels
  } else {
    group.data_labels
  }
  let cats = get_cat_labels(ser.cat)
  let label_r = if inner_r > 0 { (r + inner_r) / 2 } else { r * 2 / 3 }
  let mut svg = ""
  let mut start_angle = chart_pie_start_angle
  for i in 0..<nv {
    let sweep = mul_div(values[i], 36000, total)
    if sweep > 0 {
      let mid_angle = start_angle + sweep / 2
      let lx = cx + label_r * cos1000(mid_angle) / 1000
      let ly = cy + label_r * sin1000(mid_angle) / 1000
      let label = build_label_text(
        dl,
        values[i],
        total,
        if i < cats.length() {
          cats[i]
        } else {
          ""
        },
        ser.title,
        get_num_format(ser.val),
      )
      if !label.is_empty() {
        svg = svg +
          "<text x=\"" +
          int_to_str(lx) +
          "\" y=\"" +
          int_to_str(ly + 3) +
          "\" text-anchor=\"middle\" font-size=\"" +
          int_to_str(chart_data_label_font_size) +
          "\" fill=\"" +
          chart_title_color +
          "\">" +
          svg_escape(label) +
          "</text>"
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
      continue
    }
    let values = get_num_values(ser.val)
    let nv = values.length()
    if nv < 2 {
      continue
    }
    for ti in 0..<ser.trendlines.length() {
      let tl = ser.trendlines[ti]
      let color = if !tl.sp_pr.stroke.is_none() {
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
            // Compute slope and intercept directly
            // Values are already scaled by chart_val_scale, enough precision
            let numer = n * sum_xy - sum_x * sum_y
            let m = numer / denom // slope in scaled units per x index
            let b = (sum_y - m * sum_x) / n // y-intercept in scaled units
            // Start point (x=0)
            let y0 = b
            let x_start = px_ + pw / (nc * 2)
            let y_start = py + ph - mul_div(y0, ph, max_val)
            // End point (x=n-1)
            let yn = m * (n - 1) + b
            let x_end = px_ + ((n - 1) * 2 + 1) * pw / (nc * 2)
            let y_end = py + ph - mul_div(yn, ph, max_val)
            svg = svg +
              "<line x1=\"" +
              int_to_str(x_start) +
              "\" y1=\"" +
              int_to_str(y_start) +
              "\" x2=\"" +
              int_to_str(x_end) +
              "\" y2=\"" +
              int_to_str(y_end) +
              "\" stroke=\"" +
              color +
              "\" stroke-width=\"1.5\" stroke-dasharray=\"6,3\"/>"
          }
        }
        // Other trendline types (exp/log/poly/power/movingAvg)
        // rendered as a simple straight line from first to last point
        _ => {
          let n = if nv < nc { nv } else { nc }
          if n >= 2 {
            let x_start = px_ + pw / (nc * 2)
            let y_start = py + ph - mul_div(values[0], ph, max_val)
            let x_end = px_ + ((n - 1) * 2 + 1) * pw / (nc * 2)
            let y_end = py + ph - mul_div(values[n - 1], ph, max_val)
            svg = svg +
              "<line x1=\"" +
              int_to_str(x_start) +
              "\" y1=\"" +
              int_to_str(y_start) +
              "\" x2=\"" +
              int_to_str(x_end) +
              "\" y2=\"" +
              int_to_str(y_end) +
              "\" stroke=\"" +
              color +
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
      continue
    }
    let eb = ser.err_bars
    let values = get_num_values(ser.val)
    let nv = values.length()
    let color = series_color(ser, si, colors)
    for ci in 0..<nc {
      if ci >= nv {
        continue
      }
      let val = values[ci]
      // Calculate error magnitude
      let err = match eb.err_val_type {
        "fixedVal" => eb.val * chart_val_scale
        "percentage" => mul_div(val, eb.val, 100)
        _ => val / 10 // fallback: 10% of value
      }
      if err <= 0 {
        continue
      }
      let x = px_ + (ci * 2 + 1) * pw / (nc * 2)
      let y_center = py + ph - mul_div(val, ph, max_val)
      let show_plus = eb.err_bar_type == "both" || eb.err_bar_type == "plus"
      let show_minus = eb.err_bar_type == "both" || eb.err_bar_type == "minus"
      let y_top = if show_plus {
        py + ph - mul_div(val + err, ph, max_val)
      } else {
        y_center
      }
      let y_bottom = if show_minus {
        py + ph - mul_div(val - err, ph, max_val)
      } else {
        y_center
      }
      // Vertical line
      svg = svg +
        "<line x1=\"" +
        int_to_str(x) +
        "\" y1=\"" +
        int_to_str(y_top) +
        "\" x2=\"" +
        int_to_str(x) +
        "\" y2=\"" +
        int_to_str(y_bottom) +
        "\" stroke=\"" +
        color +
        "\" stroke-width=\"1\"/>"
      // Top cap
      if show_plus {
        svg = svg +
          "<line x1=\"" +
          int_to_str(x - chart_err_bar_cap_w) +
          "\" y1=\"" +
          int_to_str(y_top) +
          "\" x2=\"" +
          int_to_str(x + chart_err_bar_cap_w) +
          "\" y2=\"" +
          int_to_str(y_top) +
          "\" stroke=\"" +
          color +
          "\" stroke-width=\"1\"/>"
      }
      // Bottom cap
      if show_minus {
        svg = svg +
          "<line x1=\"" +
          int_to_str(x - chart_err_bar_cap_w) +
          "\" y1=\"" +
          int_to_str(y_bottom) +
          "\" x2=\"" +
          int_to_str(x + chart_err_bar_cap_w) +
          "\" y2=\"" +
          int_to_str(y_bottom) +
          "\" stroke=\"" +
          color +
          "\" stroke-width=\"1\"/>"
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
  let deg = (angle_hundredths / 100 % 360 + 360) % 360
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

// ── ChartEx (cx: namespace) renderers ───────────────────────────────────────

///|
/// Check if an index is in the subtotals array.
fn is_subtotal(subtotals : Array[Int], idx : Int) -> Bool {
  for i in 0..<subtotals.length() {
    if subtotals[i] == idx {
      return true
    }
  }
  false
}

///|
/// Waterfall chart: cumulative bars with connector lines.
fn render_waterfall_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
  axes : Array[@ooxml.ChartAxis],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let ser = group.series[0]
  let values = get_num_values(ser.val)
  let cats = get_cat_labels(ser.cat)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  // Compute cumulative positions and find global min/max for axis scaling
  let cumulative : Array[Int] = Array::make(nv, 0)
  let bar_bottom : Array[Int] = Array::make(nv, 0)
  let bar_top : Array[Int] = Array::make(nv, 0)
  let mut running = 0
  let mut global_min = 0
  let mut global_max = 0
  for i in 0..<nv {
    if is_subtotal(group.subtotals, i) {
      // Subtotal: bar from 0 to running total
      bar_bottom[i] = 0
      bar_top[i] = running
    } else {
      bar_bottom[i] = running
      running = running + values[i]
      bar_top[i] = running
    }
    cumulative[i] = running
    if bar_top[i] > global_max {
      global_max = bar_top[i]
    }
    if bar_bottom[i] > global_max {
      global_max = bar_bottom[i]
    }
    if bar_top[i] < global_min {
      global_min = bar_top[i]
    }
    if bar_bottom[i] < global_min {
      global_min = bar_bottom[i]
    }
  }
  let val_range = global_max - global_min
  let max_val = if val_range == 0 {
    1
  } else {
    compute_axis_max(val_range, axes)
  }
  // Gap between bars
  let gap_pct = chart_waterfall_gap_pct
  let total_units = nv * 100 + (nv + 1) * gap_pct
  let mut svg = ""
  // Connector color
  let conn_color = chart_gridline_color
  for i in 0..<nv {
    let bar_x = px_ + mul_div(i * 100 + (i + 1) * gap_pct, pw, total_units)
    let bar_w = mul_div(100, pw, total_units)
    // Y coordinates (inverted: higher values = lower Y)
    let top_val = if bar_top[i] > bar_bottom[i] {
      bar_top[i]
    } else {
      bar_bottom[i]
    }
    let bot_val = if bar_top[i] > bar_bottom[i] {
      bar_bottom[i]
    } else {
      bar_top[i]
    }
    let top_y = py + ph - mul_div(top_val - global_min, ph, max_val)
    let bot_y = py + ph - mul_div(bot_val - global_min, ph, max_val)
    let bar_h = bot_y - top_y
    let bar_h_clamped = if bar_h < 1 { 1 } else { bar_h }
    // Color: subtotal = accent1, positive = accent2, negative = accent3
    let color = if is_subtotal(group.subtotals, i) {
      colors[0 % colors.length()]
    } else {
      let pt_color = get_dpt_color(ser.data_points, i, colors)
      if pt_color != colors[i % colors.length()] {
        pt_color
      } else if values[i] >= 0 {
        colors[1 % colors.length()]
      } else {
        colors[2 % colors.length()]
      }
    }
    svg = svg +
      "<rect x=\"" +
      int_to_str(bar_x) +
      "\" y=\"" +
      int_to_str(top_y) +
      "\" width=\"" +
      int_to_str(bar_w) +
      "\" height=\"" +
      int_to_str(bar_h_clamped) +
      "\" fill=\"" +
      color +
      "\"/>"
    // Connector line from this bar's top to next bar's start
    if i < nv - 1 {
      let next_bar_x = px_ +
        mul_div((i + 1) * 100 + (i + 2) * gap_pct, pw, total_units)
      let conn_y = py + ph - mul_div(cumulative[i] - global_min, ph, max_val)
      svg = svg +
        "<line x1=\"" +
        int_to_str(bar_x + bar_w) +
        "\" y1=\"" +
        int_to_str(conn_y) +
        "\" x2=\"" +
        int_to_str(next_bar_x) +
        "\" y2=\"" +
        int_to_str(conn_y) +
        "\" stroke=\"" +
        conn_color +
        "\" stroke-width=\"1\" stroke-dasharray=\"3,2\"/>"
    }
    // Category label
    if i < cats.length() {
      svg = svg +
        "<text x=\"" +
        int_to_str(bar_x + bar_w / 2) +
        "\" y=\"" +
        int_to_str(py + ph + chart_label_offset_y) +
        "\" text-anchor=\"middle\" font-size=\"9\" fill=\"" +
        chart_axis_label_color +
        "\">" +
        svg_escape(cats[i]) +
        "</text>"
    }
  }
  svg
}

///|
/// Treemap chart: nested rectangles proportional to values.
fn render_treemap_chart(
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
  let cats = get_cat_labels(ser.cat)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  // Sort indices by value descending (for better layout)
  let indices : Array[Int] = Array::make(nv, 0)
  for i in 0..<nv {
    indices[i] = i
  }
  // Simple selection sort (small n)
  for i in 0..<nv {
    let mut max_idx = i
    for j in (i + 1)..<nv {
      if values[indices[j]] > values[indices[max_idx]] {
        max_idx = j
      }
    }
    if max_idx != i {
      let tmp = indices[i]
      indices[i] = indices[max_idx]
      indices[max_idx] = tmp
    }
  }
  let mut total = 0
  for i in 0..<nv {
    total = total + values[i]
  }
  if total == 0 {
    total = 1
  }
  // Slice-and-dice layout: alternate horizontal/vertical splits
  treemap_layout(
    indices, values, cats, colors, px_, py, pw, ph, total, true, "",
  )
}

///|
/// Recursive slice-and-dice treemap layout.
fn treemap_layout(
  indices : Array[Int],
  values : Array[Int],
  cats : Array[String],
  colors : Array[String],
  x : Int,
  y : Int,
  w : Int,
  h : Int,
  total : Int,
  horizontal : Bool,
  svg : String,
) -> String {
  let n = indices.length()
  if n == 0 || total == 0 {
    return svg
  }
  let mut result = svg
  if n == 1 {
    let idx = indices[0]
    let color = get_dpt_color([], idx, colors)
    result = result +
      "<rect x=\"" +
      int_to_str(x) +
      "\" y=\"" +
      int_to_str(y) +
      "\" width=\"" +
      int_to_str(w) +
      "\" height=\"" +
      int_to_str(h) +
      "\" fill=\"" +
      color +
      "\" stroke=\"white\" stroke-width=\"1\"/>"
    // Label if large enough
    if w > 30 && h > 14 && idx < cats.length() {
      result = result +
        "<text x=\"" +
        int_to_str(x + 4) +
        "\" y=\"" +
        int_to_str(y + 13) +
        "\" font-size=\"10\" fill=\"white\" clip-path=\"inset(0)\">" +
        svg_escape(cats[idx]) +
        "</text>"
    }
    return result
  }
  // Split roughly in half by value
  let mut half_total = 0
  let mut split = 0
  let half_target = total / 2
  for i in 0..<n {
    half_total = half_total + values[indices[i]]
    split = i + 1
    if half_total >= half_target {
      break
    }
  }
  if split == 0 {
    split = 1
  }
  if split >= n {
    split = n - 1
  }
  let first : Array[Int] = []
  let second : Array[Int] = []
  let mut first_total = 0
  for i in 0..<split {
    first.push(indices[i])
    first_total = first_total + values[indices[i]]
  }
  for i in split..<n {
    second.push(indices[i])
  }
  let second_total = total - first_total
  if horizontal {
    let w1 = mul_div(first_total, w, total)
    let w2 = w - w1
    result = treemap_layout(
      first,
      values,
      cats,
      colors,
      x,
      y,
      w1,
      h,
      first_total,
      !horizontal,
      result,
    )
    result = treemap_layout(
      second,
      values,
      cats,
      colors,
      x + w1,
      y,
      w2,
      h,
      second_total,
      !horizontal,
      result,
    )
  } else {
    let h1 = mul_div(first_total, h, total)
    let h2 = h - h1
    result = treemap_layout(
      first,
      values,
      cats,
      colors,
      x,
      y,
      w,
      h1,
      first_total,
      !horizontal,
      result,
    )
    result = treemap_layout(
      second,
      values,
      cats,
      colors,
      x,
      y + h1,
      w,
      h2,
      second_total,
      !horizontal,
      result,
    )
  }
  result
}

///|
/// Sunburst chart: single-level concentric pie chart.
fn render_sunburst_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
) -> String {
  // Sunburst with single level is equivalent to a pie chart
  if group.series.is_empty() {
    return ""
  }
  let ser = group.series[0]
  let values = get_num_values(ser.val)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  let mut total = 0
  for i in 0..<nv {
    total = total + values[i]
  }
  if total == 0 {
    return ""
  }
  let cx = px_ + pw / 2
  let cy = py + ph / 2
  let r = (if pw < ph { pw } else { ph }) / 2 - chart_pie_pad
  let inner_r = r * chart_sunburst_inner_pct / 100
  // Render as donut slices
  let mut svg = ""
  let mut start_angle = chart_pie_start_angle
  for i in 0..<nv {
    let sweep = mul_div(values[i], 36000, total)
    if sweep == 0 {
      continue
    }
    let end_angle = start_angle + sweep
    let color = get_dpt_color(ser.data_points, i, colors)
    let large_arc = if sweep > 18000 { "1" } else { "0" }
    // Outer arc points
    let sx = cx + mul_div(r, cos1000(start_angle), 1000)
    let sy = cy + mul_div(r, sin1000(start_angle), 1000)
    let ex = cx + mul_div(r, cos1000(end_angle), 1000)
    let ey = cy + mul_div(r, sin1000(end_angle), 1000)
    // Inner arc points
    let isx = cx + mul_div(inner_r, cos1000(start_angle), 1000)
    let isy = cy + mul_div(inner_r, sin1000(start_angle), 1000)
    let iex = cx + mul_div(inner_r, cos1000(end_angle), 1000)
    let iey = cy + mul_div(inner_r, sin1000(end_angle), 1000)
    let r_str = int_to_str(r)
    let ir_str = int_to_str(inner_r)
    svg = svg +
      "<path d=\"M" +
      int_to_str(sx) +
      " " +
      int_to_str(sy) +
      "A" +
      r_str +
      " " +
      r_str +
      " 0 " +
      large_arc +
      " 1 " +
      int_to_str(ex) +
      " " +
      int_to_str(ey) +
      "L" +
      int_to_str(iex) +
      " " +
      int_to_str(iey) +
      "A" +
      ir_str +
      " " +
      ir_str +
      " 0 " +
      large_arc +
      " 0 " +
      int_to_str(isx) +
      " " +
      int_to_str(isy) +
      "Z\" fill=\"" +
      color +
      "\" stroke=\"white\" stroke-width=\"1\"/>"
    start_angle = end_angle
  }
  svg
}

///|
/// Funnel chart: centered bars decreasing in width.
fn render_funnel_chart(
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
  let cats = get_cat_labels(ser.cat)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  let mut max_val = 0
  for i in 0..<nv {
    if values[i] > max_val {
      max_val = values[i]
    }
  }
  if max_val == 0 {
    max_val = 1
  }
  let gap = chart_funnel_gap
  let bar_h = (ph - gap * (nv - 1)) / nv
  let bar_h_final = if bar_h < 4 { 4 } else { bar_h }
  let cx = px_ + pw / 2
  let mut svg = ""
  for i in 0..<nv {
    let bar_w = mul_div(values[i], pw, max_val)
    let bar_x = cx - bar_w / 2
    let bar_y = py + i * (bar_h_final + gap)
    let color = get_dpt_color(ser.data_points, i, colors)
    svg = svg +
      "<rect x=\"" +
      int_to_str(bar_x) +
      "\" y=\"" +
      int_to_str(bar_y) +
      "\" width=\"" +
      int_to_str(bar_w) +
      "\" height=\"" +
      int_to_str(bar_h_final) +
      "\" fill=\"" +
      color +
      "\" rx=\"2\"/>"
    // Label
    if i < cats.length() {
      svg = svg +
        "<text x=\"" +
        int_to_str(cx) +
        "\" y=\"" +
        int_to_str(bar_y + bar_h_final / 2 + 4) +
        "\" text-anchor=\"middle\" font-size=\"10\" fill=\"white\">" +
        svg_escape(cats[i]) +
        "</text>"
    }
  }
  svg
}

///|
/// Box & Whisker chart: statistical box plots.
fn render_box_whisker_chart(
  group : @ooxml.ChartGroup,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  colors : Array[String],
  axes : Array[@ooxml.ChartAxis],
) -> String {
  if group.series.is_empty() {
    return ""
  }
  let ser = group.series[0]
  let values = get_num_values(ser.val)
  let cats = get_cat_labels(ser.cat)
  let nv = values.length()
  if nv == 0 {
    return ""
  }
  // For box-whisker we expect data in groups:
  // Simple approach: if only 1 series, sort values and compute quartiles
  // per category. If multiple series, treat each series as one box.
  let ns = group.series.length()
  if ns >= 2 {
    // Multiple series: each series is one box category
    // Gather all values per series to compute quartiles
    let mut global_max = 0
    let all_quartiles : Array[Array[Int]] = []
    for si in 0..<ns {
      let vals = get_num_values(group.series[si].val)
      let q = compute_quartiles(vals)
      all_quartiles.push(q)
      if q.length() >= 5 && q[4] > global_max {
        global_max = q[4]
      }
    }
    let max_val = compute_axis_max(global_max, axes)
    if max_val == 0 {
      return ""
    }
    return render_box_whisker_boxes(
      all_quartiles, cats, ns, px_, py, pw, ph, max_val, colors,
    )
  }
  // Single series: treat values directly as raw data for one box
  let q = compute_quartiles(values)
  if q.length() < 5 {
    return ""
  }
  let max_val = compute_axis_max(q[4], axes)
  if max_val == 0 {
    return ""
  }
  render_box_whisker_boxes([q], cats, 1, px_, py, pw, ph, max_val, colors)
}

///|
/// Render box-whisker boxes given pre-computed quartiles.
fn render_box_whisker_boxes(
  quartiles : Array[Array[Int]],
  cats : Array[String],
  nc : Int,
  px_ : Int,
  py : Int,
  pw : Int,
  ph : Int,
  max_val : Int,
  colors : Array[String],
) -> String {
  let mut svg = ""
  let gap_pct = chart_box_whisker_gap_pct
  let total_units = nc * 100 + (nc + 1) * gap_pct
  for ci in 0..<nc {
    if ci >= quartiles.length() || quartiles[ci].length() < 5 {
      continue
    }
    let q = quartiles[ci]
    let min_v = q[0]
    let q1 = q[1]
    let med = q[2]
    let q3 = q[3]
    let max_v = q[4]
    let box_x = px_ + mul_div(ci * 100 + (ci + 1) * gap_pct, pw, total_units)
    let box_w = mul_div(100, pw, total_units)
    let cx = box_x + box_w / 2
    let color = colors[ci % colors.length()]
    // Y coordinates (inverted)
    let min_y = py + ph - mul_div(min_v, ph, max_val)
    let q1_y = py + ph - mul_div(q1, ph, max_val)
    let med_y = py + ph - mul_div(med, ph, max_val)
    let q3_y = py + ph - mul_div(q3, ph, max_val)
    let max_y = py + ph - mul_div(max_v, ph, max_val)
    // Whisker: vertical line from min to max
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx) +
      "\" y1=\"" +
      int_to_str(max_y) +
      "\" x2=\"" +
      int_to_str(cx) +
      "\" y2=\"" +
      int_to_str(q3_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"1\"/>"
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx) +
      "\" y1=\"" +
      int_to_str(q1_y) +
      "\" x2=\"" +
      int_to_str(cx) +
      "\" y2=\"" +
      int_to_str(min_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"1\"/>"
    // Whisker caps (horizontal lines at min and max)
    let cap_w = box_w / 3
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx - cap_w) +
      "\" y1=\"" +
      int_to_str(max_y) +
      "\" x2=\"" +
      int_to_str(cx + cap_w) +
      "\" y2=\"" +
      int_to_str(max_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"1\"/>"
    svg = svg +
      "<line x1=\"" +
      int_to_str(cx - cap_w) +
      "\" y1=\"" +
      int_to_str(min_y) +
      "\" x2=\"" +
      int_to_str(cx + cap_w) +
      "\" y2=\"" +
      int_to_str(min_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"1\"/>"
    // Box: Q1 to Q3
    let box_h = q1_y - q3_y
    let box_h_clamped = if box_h < 1 { 1 } else { box_h }
    svg = svg +
      "<rect x=\"" +
      int_to_str(box_x) +
      "\" y=\"" +
      int_to_str(q3_y) +
      "\" width=\"" +
      int_to_str(box_w) +
      "\" height=\"" +
      int_to_str(box_h_clamped) +
      "\" fill=\"" +
      color +
      "\" fill-opacity=\"0.3\" stroke=\"" +
      color +
      "\" stroke-width=\"1.5\"/>"
    // Median line
    svg = svg +
      "<line x1=\"" +
      int_to_str(box_x) +
      "\" y1=\"" +
      int_to_str(med_y) +
      "\" x2=\"" +
      int_to_str(box_x + box_w) +
      "\" y2=\"" +
      int_to_str(med_y) +
      "\" stroke=\"" +
      color +
      "\" stroke-width=\"2\"/>"
    // Category label
    if ci < cats.length() {
      svg = svg +
        "<text x=\"" +
        int_to_str(cx) +
        "\" y=\"" +
        int_to_str(py + ph + chart_label_offset_y) +
        "\" text-anchor=\"middle\" font-size=\"9\" fill=\"" +
        chart_axis_label_color +
        "\">" +
        svg_escape(cats[ci]) +
        "</text>"
    }
  }
  svg
}

///|
/// Compute quartiles [min, Q1, median, Q3, max] from an array of values.
fn compute_quartiles(values : Array[Int]) -> Array[Int] {
  let n = values.length()
  if n == 0 {
    return []
  }
  // Copy and sort
  let sorted : Array[Int] = Array::make(n, 0)
  for i in 0..<n {
    sorted[i] = values[i]
  }
  // Simple insertion sort (typically small n for chart data)
  for i in 1..<n {
    let key = sorted[i]
    let mut j = i - 1
    while j >= 0 && sorted[j] > key {
      sorted[j + 1] = sorted[j]
      j = j - 1
    }
    sorted[j + 1] = key
  }
  let min_v = sorted[0]
  let max_v = sorted[n - 1]
  let median = sorted[n / 2]
  let q1 = sorted[n / 4]
  let q3 = sorted[n * 3 / 4]
  [min_v, q1, median, q3, max_v]
}
