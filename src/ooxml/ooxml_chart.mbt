// OOXML Chart parsing — ChartML (ECMA-376 Part 1 Chapter 21)
//
// Parses c:chartSpace XML from ppt/charts/*.xml into ChartData.
// Supports: bar/column, line, pie, donut, scatter, area, radar, bubble, stock charts.
// Series data is read from numCache/strCache (not from external xlsx).
//
// Lazy resolution: parse_graphic_frame stores chart r:id in chart_xml field;
// main.mbt resolves the actual chart XML file via rels + ffi_get_file.
//
// Same package as ooxml.mbt — all types and helpers are shared.

// ── Chart defaults (ECMA-376) ───────────────────────────────────────────────

///|
let default_bar_gap_width : Int = 150 // ECMA-376 default gap width percentage

///|
let default_donut_hole_size : Int = 50 // ECMA-376 default donut hole size percentage

// ── Chart data labels ─────────────────────────────────────────────────────

///|
/// Data label configuration (c:dLbls).
pub(all) struct ChartDataLabels {
  show_val : Bool // c:showVal
  show_cat_name : Bool // c:showCatName
  show_ser_name : Bool // c:showSerName
  show_percent : Bool // c:showPercent
  show_leader_lines : Bool // c:showLeaderLines
  separator : String // c:separator (default ",")
}

///|
pub fn ChartDataLabels::none() -> ChartDataLabels {
  {
    show_val: false,
    show_cat_name: false,
    show_ser_name: false,
    show_percent: false,
    show_leader_lines: false,
    separator: ", ",
  }
}

///|
pub fn ChartDataLabels::is_visible(self : ChartDataLabels) -> Bool {
  self.show_val || self.show_cat_name || self.show_ser_name || self.show_percent
}

///|
/// Per-data-point formatting override (c:dPt).
pub(all) struct ChartDataPoint {
  idx : Int // c:idx/@val
  sp_pr : ChartSpPr // point-level formatting
}

// ── Trendline ───────────────────────────────────────────────────────────────

///|
/// Trendline type (c:trendlineType).
pub(all) struct ChartTrendline {
  trendline_type : String // "linear"|"exp"|"log"|"poly"|"power"|"movingAvg"
  order : Int // polynomial order (c:order, default 2)
  period : Int // moving average period (c:period, default 2)
  forward : Int // forward extrapolation (c:forward, in category units × 100)
  backward : Int // backward extrapolation (c:backward, in category units × 100)
  sp_pr : ChartSpPr // line formatting
  name : String // custom name (c:name)
}

///|
pub fn ChartTrendline::none() -> ChartTrendline {
  {
    trendline_type: "",
    order: 2,
    period: 2,
    forward: 0,
    backward: 0,
    sp_pr: ChartSpPr::none(),
    name: "",
  }
}

// ── Error bars ──────────────────────────────────────────────────────────────

///|
/// Error bar configuration (c:errBars).
pub(all) struct ChartErrBars {
  err_dir : String // "x"|"y" (c:errDir)
  err_bar_type : String // "both"|"plus"|"minus" (c:errBarType)
  err_val_type : String // "fixedVal"|"percentage"|"stdDev"|"stdErr"|"cust" (c:errValType)
  val : Int // fixed value or percentage (c:val × 1, scaled)
  sp_pr : ChartSpPr // bar line formatting
}

///|
pub fn ChartErrBars::none() -> ChartErrBars {
  {
    err_dir: "y",
    err_bar_type: "both",
    err_val_type: "",
    val: 0,
    sp_pr: ChartSpPr::none(),
  }
}

///|
pub fn ChartErrBars::is_none(self : ChartErrBars) -> Bool {
  self.err_val_type.is_empty()
}

// ── 3D view ─────────────────────────────────────────────────────────────────

///|
/// 3D view settings (c:view3D) — preserved for round-trip, not rendered.
pub(all) struct ChartView3D {
  rot_x : Int // c:rotX (degrees × 1)
  rot_y : Int // c:rotY (degrees × 1)
  depth_percent : Int // c:depthPercent (default 100)
  r_ang_ax : Bool // c:rAngAx (right-angle axes)
  perspective : Int // c:perspective (0-240, default 30)
}

///|
pub fn ChartView3D::none() -> ChartView3D {
  { rot_x: 15, rot_y: 20, depth_percent: 100, r_ang_ax: true, perspective: 30 }
}

// ── Chart data types ────────────────────────────────────────────────────────

///|
/// Chart type — determines rendering strategy.
pub(all) enum ChartKind {
  BarChart // c:barChart (includes column)
  LineChart // c:lineChart
  PieChart // c:pieChart
  DoughnutChart // c:doughnutChart
  ScatterChart // c:scatterChart
  AreaChart // c:areaChart
  RadarChart // c:radarChart
  BubbleChart // c:bubbleChart
  StockChart // c:stockChart
  SurfaceChart // c:surfaceChart / c:surface3DChart — 2D heatmap approximation
  OfPieChart // c:ofPieChart — pie/bar of pie, rendered as standard pie
}

///|
/// A single data point (from c:pt).
pub(all) struct ChartPoint {
  idx : Int // c:pt/@idx
  value : String // c:v text
}

///|
/// Cached numeric data (c:numCache or c:numLit).
pub(all) struct NumData {
  format_code : String // c:formatCode
  points : Array[ChartPoint] // c:pt list
}

///|
pub fn NumData::empty() -> NumData {
  { format_code: "", points: [] }
}

///|
/// Cached string data (c:strCache or c:strLit).
pub(all) struct StrData {
  points : Array[ChartPoint]
}

///|
pub fn StrData::empty() -> StrData {
  { points: [] }
}

///|
/// Data source — either numeric or string reference with cache.
pub(all) enum AxisDataSource {
  NumSource(String, NumData) // formula, cache
  StrSource(String, StrData) // formula, cache
  NoData
}

///|
/// Simplified shape properties for chart elements.
pub(all) struct ChartSpPr {
  fill : Color
  grad_fill : GradientFill
  patt_fill : PatternFill
  stroke : Color
  stroke_w : Int
  no_fill : Bool
}

///|
pub fn ChartSpPr::none() -> ChartSpPr {
  {
    fill: Color::none(),
    grad_fill: GradientFill::none(),
    patt_fill: PatternFill::none(),
    stroke: Color::none(),
    stroke_w: 0,
    no_fill: false,
  }
}

///|
/// A single chart series (c:ser).
pub(all) struct ChartSeries {
  idx : Int // c:idx/@val
  order : Int // c:order/@val
  title : String // series name (from c:tx strCache)
  sp_pr : ChartSpPr // series formatting
  // Category-based (bar/line/area/pie)
  cat : AxisDataSource // c:cat
  val : AxisDataSource // c:val
  // Scatter/bubble
  x_val : AxisDataSource // c:xVal
  y_val : AxisDataSource // c:yVal
  bubble_size : AxisDataSource // c:bubbleSize (bubble chart only)
  // Line/scatter
  smooth : Bool // c:smooth
  // Pie
  explosion : Int // c:explosion (0 = none)
  // Per-point formatting
  data_points : Array[ChartDataPoint] // c:dPt list
  // Trendlines + error bars
  trendlines : Array[ChartTrendline] // c:trendline list
  err_bars : ChartErrBars // c:errBars
}

///|
/// A chart type group in the plot area.
pub(all) struct ChartGroup {
  chart_type : ChartKind
  series : Array[ChartSeries]
  bar_dir : String // "bar"|"col" (bar chart only)
  grouping : String // "clustered"|"stacked"|"percentStacked"|"standard"
  gap_width : Int // percentage (default 150)
  overlap : Int // -100 to 100 (default 0)
  vary_colors : Bool
  hole_size : Int // donut: 1-90 percentage
  scatter_style : String // scatter: "lineMarker" etc.
  ax_ids : Array[Int] // axis ID references
  data_labels : ChartDataLabels // c:dLbls
  of_pie_type : String // ofPieChart: "pie"|"bar" (c:ofPieType)
  split_pos : Int // ofPieChart: split position (c:splitPos)
  wireframe : Bool // surface chart: wireframe mode
}

///|
pub fn ChartGroup::default(ct : ChartKind) -> ChartGroup {
  {
    chart_type: ct,
    series: [],
    bar_dir: "col",
    grouping: "clustered",
    gap_width: default_bar_gap_width,
    overlap: 0,
    vary_colors: false,
    hole_size: default_donut_hole_size,
    scatter_style: "",
    ax_ids: [],
    data_labels: ChartDataLabels::none(),
    of_pie_type: "",
    split_pos: 0,
    wireframe: false,
  }
}

///|
/// Chart axis.
pub(all) struct ChartAxis {
  ax_id : Int
  cross_ax : Int
  ax_pos : String // "b"|"l"|"r"|"t"
  delete : Bool // hidden axis
  is_val : Bool // true=valAx, false=catAx/dateAx
  major_gridlines : Bool
  minor_gridlines : Bool
  title : String // axis title text (simplified)
  orientation : String // "minMax"|"maxMin"
  min_val : String // "" = auto
  max_val : String // "" = auto
  major_unit : String // "" = auto
  num_fmt : String // format code
  tick_lbl_pos : String // "high"|"low"|"nextTo"|"none"
  cross_between : String // "between"|"midCat" (valAx only)
  sp_pr : ChartSpPr
}

///|
pub fn ChartAxis::default() -> ChartAxis {
  {
    ax_id: 0,
    cross_ax: 0,
    ax_pos: "b",
    delete: false,
    is_val: false,
    major_gridlines: false,
    minor_gridlines: false,
    title: "",
    orientation: "minMax",
    min_val: "",
    max_val: "",
    major_unit: "",
    num_fmt: "General",
    tick_lbl_pos: "nextTo",
    cross_between: "between",
    sp_pr: ChartSpPr::none(),
  }
}

///|
/// Chart legend.
pub(all) struct ChartLegend {
  position : String // "b"|"l"|"r"|"t"|"tr"
  overlay : Bool
  show : Bool // false = no legend
}

///|
pub fn ChartLegend::none() -> ChartLegend {
  { position: "r", overlay: false, show: false }
}

///|
/// Complete chart data (from c:chartSpace).
pub(all) struct ChartData {
  groups : Array[ChartGroup] // chart type groups in plotArea
  axes : Array[ChartAxis] // all axes
  title : String // chart title text (simplified)
  legend : ChartLegend
  style : Int // c:style (0 = none)
  chart_xml : String // raw chart XML for round-trip preservation
  view_3d : ChartView3D // c:view3D — preserved for round-trip
}

///|
pub fn ChartData::empty() -> ChartData {
  {
    groups: [],
    axes: [],
    title: "",
    legend: ChartLegend::none(),
    style: 0,
    chart_xml: "",
    view_3d: ChartView3D::none(),
  }
}

// ── Chart XML parser ────────────────────────────────────────────────────────

///|
/// Parse chart XML (c:chartSpace) into ChartData.
pub fn parse_chart_xml(chart_xml : String, theme : ThemeData) -> ChartData {
  let nodes = @xml.parse_xml(chart_xml)
  let cs = match @xml.find_child(nodes, "c:chartSpace") {
    None => return ChartData::empty()
    Some(n) => n
  }
  let cs_ch = @xml.get_children(cs)
  // c:style
  let style = match @xml.find_child(cs_ch, "c:style") {
    None => 0
    Some(s) => attr_int(s, "val")
  }
  // c:chart
  let chart_node = match @xml.find_child(cs_ch, "c:chart") {
    None => return { ..ChartData::empty(), style, chart_xml }
    Some(n) => n
  }
  let chart_ch = @xml.get_children(chart_node)
  // Chart title
  let title = parse_chart_title_text(chart_ch)
  // Plot area
  let (groups, axes) = match @xml.find_child(chart_ch, "c:plotArea") {
    None => ([], [])
    Some(pa) => parse_plot_area(@xml.get_children(pa), theme)
  }
  // Legend
  let legend = match @xml.find_child(chart_ch, "c:legend") {
    None => ChartLegend::none()
    Some(lg) => parse_chart_legend(@xml.get_children(lg))
  }
  // 3D view (preserved for round-trip)
  let view_3d = match @xml.find_child(chart_ch, "c:view3D") {
    None => ChartView3D::none()
    Some(v) => parse_view_3d(@xml.get_children(v))
  }
  { groups, axes, title, legend, style, chart_xml, view_3d }
}

///|
/// Extract plain text from c:title.
fn parse_chart_title_text(chart_ch : Array[@xml.XmlNode]) -> String {
  let title_node = match @xml.find_child(chart_ch, "c:title") {
    None => return ""
    Some(n) => n
  }
  let title_ch = @xml.get_children(title_node)
  let tx = match @xml.find_child(title_ch, "c:tx") {
    None => return ""
    Some(n) => n
  }
  let tx_ch = @xml.get_children(tx)
  // c:rich → a:p → a:r → a:t
  match @xml.find_child(tx_ch, "c:rich") {
    None => ""
    Some(rich) => extract_all_text(@xml.get_children(rich))
  }
}

///|
/// Extract all text from a:p/a:r/a:t nodes.
fn extract_all_text(nodes : Array[@xml.XmlNode]) -> String {
  let mut result = ""
  for i in 0..<nodes.length() {
    match nodes[i] {
      @xml.XmlNode::Element(tag, _, _) =>
        if tag == "a:p" {
          let p_ch = @xml.get_children(nodes[i])
          for j in 0..<p_ch.length() {
            match p_ch[j] {
              @xml.XmlNode::Element(rtag, _, _) =>
                if rtag == "a:r" {
                  match @xml.find_child(
                    @xml.get_children(p_ch[j]),
                    "a:t",
                  ) {
                    None => ()
                    Some(t) => result = result + @xml.get_text_content(t)
                  }
                }
              _ => ()
            }
          }
        }
      _ => ()
    }
  }
  result
}

///|
/// Parse plot area — chart groups + axes.
fn parse_plot_area(
  pa_ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> (Array[ChartGroup], Array[ChartAxis]) {
  let groups : Array[ChartGroup] = []
  let axes : Array[ChartAxis] = []
  for i in 0..<pa_ch.length() {
    match pa_ch[i] {
      @xml.XmlNode::Element(tag, _, _) => {
        let ch = @xml.get_children(pa_ch[i])
        if tag == "c:barChart" || tag == "c:bar3DChart" {
          groups.push(parse_bar_chart(ch, theme))
        } else if tag == "c:lineChart" || tag == "c:line3DChart" {
          groups.push(parse_line_chart(ch, theme))
        } else if tag == "c:pieChart" || tag == "c:pie3DChart" {
          groups.push(parse_pie_chart(ch, theme))
        } else if tag == "c:doughnutChart" {
          groups.push(parse_donut_chart(ch, theme))
        } else if tag == "c:scatterChart" {
          groups.push(parse_scatter_chart(ch, theme))
        } else if tag == "c:areaChart" || tag == "c:area3DChart" {
          groups.push(parse_area_chart(ch, theme))
        } else if tag == "c:radarChart" {
          groups.push(parse_radar_chart(ch, theme))
        } else if tag == "c:bubbleChart" {
          groups.push(parse_bubble_chart(ch, theme))
        } else if tag == "c:stockChart" {
          groups.push(parse_stock_chart(ch, theme))
        } else if tag == "c:surfaceChart" || tag == "c:surface3DChart" {
          groups.push(parse_surface_chart(ch, theme))
        } else if tag == "c:ofPieChart" {
          groups.push(parse_of_pie_chart(ch, theme))
        } else if tag == "c:catAx" || tag == "c:dateAx" || tag == "c:serAx" {
          axes.push(parse_chart_axis(ch, false, theme))
        } else if tag == "c:valAx" {
          axes.push(parse_chart_axis(ch, true, theme))
        }
      }
      _ => ()
    }
  }
  (groups, axes)
}

///|
fn parse_bar_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  let bar_dir_raw = child_attr_val(ch, "c:barDir")
  let grouping_raw = child_attr_val(ch, "c:grouping")
  let gw = child_attr_int(ch, "c:gapWidth")
  {
    ..ChartGroup::default(BarChart),
    bar_dir: if bar_dir_raw.is_empty() { "col" } else { bar_dir_raw },
    grouping: if grouping_raw.is_empty() { "clustered" } else { grouping_raw },
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    gap_width: if gw == 0 { default_bar_gap_width } else { gw },
    overlap: child_attr_int(ch, "c:overlap"),
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_line_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  let grouping_raw = child_attr_val(ch, "c:grouping")
  {
    ..ChartGroup::default(LineChart),
    grouping: if grouping_raw.is_empty() { "standard" } else { grouping_raw },
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_pie_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(PieChart),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_donut_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  let hs = child_attr_int(ch, "c:holeSize")
  {
    ..ChartGroup::default(DoughnutChart),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    hole_size: if hs == 0 { default_donut_hole_size } else { hs },
    series: parse_series_list(ch, theme),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_scatter_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(ScatterChart),
    scatter_style: child_attr_val(ch, "c:scatterStyle"),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_scatter_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_area_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  let grouping_raw = child_attr_val(ch, "c:grouping")
  {
    ..ChartGroup::default(AreaChart),
    grouping: if grouping_raw.is_empty() { "standard" } else { grouping_raw },
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_radar_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(RadarChart),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_bubble_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(BubbleChart),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_scatter_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_stock_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(StockChart),
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_surface_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  {
    ..ChartGroup::default(SurfaceChart),
    wireframe: child_attr_val(ch, "c:wireframe") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
fn parse_of_pie_chart(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartGroup {
  let of_pie_type = child_attr_val(ch, "c:ofPieType")
  let gw = child_attr_int(ch, "c:gapWidth")
  let sp = child_attr_int(ch, "c:splitPos")
  {
    ..ChartGroup::default(OfPieChart),
    of_pie_type: if of_pie_type.is_empty() { "pie" } else { of_pie_type },
    split_pos: sp,
    gap_width: if gw == 0 { default_bar_gap_width } else { gw },
    vary_colors: child_attr_val(ch, "c:varyColors") == "1",
    series: parse_series_list(ch, theme),
    ax_ids: parse_ax_ids(ch),
    data_labels: parse_chart_data_labels(ch),
  }
}

///|
/// Get val attribute of a child element.
fn child_attr_val(ch : Array[@xml.XmlNode], tag : String) -> String {
  match @xml.find_child(ch, tag) {
    None => ""
    Some(n) =>
      match @xml.get_attr(n, "val") {
        None => ""
        Some(v) => v
      }
  }
}

///|
/// Get val attribute as int.
fn child_attr_int(ch : Array[@xml.XmlNode], tag : String) -> Int {
  match @xml.find_child(ch, tag) {
    None => 0
    Some(n) => attr_int(n, "val")
  }
}

///|
/// Parse axId references.
fn parse_ax_ids(ch : Array[@xml.XmlNode]) -> Array[Int] {
  let ids : Array[Int] = []
  let nodes = @xml.find_children(ch, "c:axId")
  for i in 0..<nodes.length() {
    ids.push(attr_int(nodes[i], "val"))
  }
  ids
}

// ── Data labels parsing ──────────────────────────────────────────────────────

///|
/// Parse c:dLbls element.
fn parse_chart_data_labels(ch : Array[@xml.XmlNode]) -> ChartDataLabels {
  match @xml.find_child(ch, "c:dLbls") {
    None => ChartDataLabels::none()
    Some(n) => {
      let dc = @xml.get_children(n)
      {
        show_val: child_attr_val(dc, "c:showVal") == "1",
        show_cat_name: child_attr_val(dc, "c:showCatName") == "1",
        show_ser_name: child_attr_val(dc, "c:showSerName") == "1",
        show_percent: child_attr_val(dc, "c:showPercent") == "1",
        show_leader_lines: child_attr_val(dc, "c:showLeaderLines") == "1",
        separator: match @xml.find_child(dc, "c:separator") {
          None => ", "
          Some(sep) => {
            let t = @xml.get_text_content(sep)
            if t.is_empty() { ", " } else { t }
          }
        },
      }
    }
  }
}

// ── Data point parsing ──────────────────────────────────────────────────────

///|
/// Parse c:dPt elements from series children.
fn parse_data_points(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Array[ChartDataPoint] {
  let pts : Array[ChartDataPoint] = []
  let dpt_nodes = @xml.find_children(ch, "c:dPt")
  for i in 0..<dpt_nodes.length() {
    let dc = @xml.get_children(dpt_nodes[i])
    let idx = child_attr_int(dc, "c:idx")
    let sp_pr = parse_chart_sp_pr(dc, theme)
    pts.push({ idx, sp_pr })
  }
  pts
}

// ── Trendline parsing ────────────────────────────────────────────────────────

///|
/// Parse c:trendline elements from series children.
fn parse_trendlines(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Array[ChartTrendline] {
  let result : Array[ChartTrendline] = []
  let nodes = @xml.find_children(ch, "c:trendline")
  for i in 0..<nodes.length() {
    let tc = @xml.get_children(nodes[i])
    let ttype = child_attr_val(tc, "c:trendlineType")
    if not(ttype.is_empty()) {
      let name = match @xml.find_child(tc, "c:name") {
        None => ""
        Some(n) => @xml.get_text_content(n)
      }
      result.push({
        trendline_type: ttype,
        order: {
          let v = child_attr_int(tc, "c:order")
          if v == 0 { 2 } else { v }
        },
        period: {
          let v = child_attr_int(tc, "c:period")
          if v == 0 { 2 } else { v }
        },
        forward: child_attr_int(tc, "c:forward"),
        backward: child_attr_int(tc, "c:backward"),
        sp_pr: parse_chart_sp_pr(tc, theme),
        name,
      })
    }
  }
  result
}

// ── Error bar parsing ────────────────────────────────────────────────────────

///|
/// Parse c:errBars element from series children.
fn parse_err_bars(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartErrBars {
  match @xml.find_child(ch, "c:errBars") {
    None => ChartErrBars::none()
    Some(n) => {
      let ec = @xml.get_children(n)
      let dir = child_attr_val(ec, "c:errDir")
      let bar_type = child_attr_val(ec, "c:errBarType")
      let val_type = child_attr_val(ec, "c:errValType")
      let val = child_attr_int(ec, "c:val")
      {
        err_dir: if dir.is_empty() { "y" } else { dir },
        err_bar_type: if bar_type.is_empty() { "both" } else { bar_type },
        err_val_type: val_type,
        val,
        sp_pr: parse_chart_sp_pr(ec, theme),
      }
    }
  }
}

// ── Series parsing ──────────────────────────────────────────────────────────

///|
/// Parse all c:ser elements (category-based charts).
fn parse_series_list(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Array[ChartSeries] {
  let series : Array[ChartSeries] = []
  let ser_nodes = @xml.find_children(ch, "c:ser")
  for i in 0..<ser_nodes.length() {
    series.push(parse_cat_series(@xml.get_children(ser_nodes[i]), theme))
  }
  series
}

///|
/// Parse a single category-based series (bar/line/area/pie/donut).
fn parse_cat_series(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartSeries {
  let idx = child_attr_int(ch, "c:idx")
  let order = child_attr_int(ch, "c:order")
  let title = parse_series_title(ch)
  let sp_pr = parse_chart_sp_pr(ch, theme)
  let cat = match @xml.find_child(ch, "c:cat") {
    None => AxisDataSource::NoData
    Some(n) => parse_axis_data_source(@xml.get_children(n))
  }
  let val = match @xml.find_child(ch, "c:val") {
    None => AxisDataSource::NoData
    Some(n) => parse_axis_data_source(@xml.get_children(n))
  }
  let smooth = child_attr_val(ch, "c:smooth") == "1"
  let explosion = child_attr_int(ch, "c:explosion")
  let data_points = parse_data_points(ch, theme)
  let trendlines = parse_trendlines(ch, theme)
  let err_bars = parse_err_bars(ch, theme)
  {
    idx,
    order,
    title,
    sp_pr,
    cat,
    val,
    x_val: NoData,
    y_val: NoData,
    bubble_size: NoData,
    smooth,
    explosion,
    data_points,
    trendlines,
    err_bars,
  }
}

///|
/// Parse scatter series (xVal/yVal instead of cat/val).
fn parse_scatter_series_list(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> Array[ChartSeries] {
  let series : Array[ChartSeries] = []
  let ser_nodes = @xml.find_children(ch, "c:ser")
  for i in 0..<ser_nodes.length() {
    let sc = @xml.get_children(ser_nodes[i])
    let x_val = match @xml.find_child(sc, "c:xVal") {
      None => AxisDataSource::NoData
      Some(n) => parse_axis_data_source(@xml.get_children(n))
    }
    let y_val = match @xml.find_child(sc, "c:yVal") {
      None => AxisDataSource::NoData
      Some(n) => parse_axis_data_source(@xml.get_children(n))
    }
    let bubble_size = match @xml.find_child(sc, "c:bubbleSize") {
      None => AxisDataSource::NoData
      Some(n) => parse_axis_data_source(@xml.get_children(n))
    }
    series.push({
      idx: child_attr_int(sc, "c:idx"),
      order: child_attr_int(sc, "c:order"),
      title: parse_series_title(sc),
      sp_pr: parse_chart_sp_pr(sc, theme),
      cat: NoData,
      val: NoData,
      x_val,
      y_val,
      bubble_size,
      smooth: child_attr_val(sc, "c:smooth") == "1",
      explosion: 0,
      data_points: parse_data_points(sc, theme),
      trendlines: parse_trendlines(sc, theme),
      err_bars: parse_err_bars(sc, theme),
    })
  }
  series
}

///|
/// Extract series title from c:tx.
fn parse_series_title(ch : Array[@xml.XmlNode]) -> String {
  let tx = match @xml.find_child(ch, "c:tx") {
    None => return ""
    Some(n) => n
  }
  let tx_ch = @xml.get_children(tx)
  // Try c:strRef/c:strCache first
  match @xml.find_child(tx_ch, "c:strRef") {
    Some(sr) => {
      let sr_ch = @xml.get_children(sr)
      match @xml.find_child(sr_ch, "c:strCache") {
        Some(cache) => {
          let pts = @xml.find_children(@xml.get_children(cache), "c:pt")
          if pts.length() > 0 {
            match @xml.find_child(@xml.get_children(pts[0]), "c:v") {
              Some(v) => return @xml.get_text_content(v)
              None => ()
            }
          }
        }
        None => ()
      }
    }
    None => ()
  }
  // Try c:v directly
  match @xml.find_child(tx_ch, "c:v") {
    Some(v) => @xml.get_text_content(v)
    None => ""
  }
}

///|
/// Parse axis data source (c:cat, c:val, c:xVal, c:yVal children).
fn parse_axis_data_source(ch : Array[@xml.XmlNode]) -> AxisDataSource {
  // Try c:numRef
  match @xml.find_child(ch, "c:numRef") {
    Some(nr) => {
      let nr_ch = @xml.get_children(nr)
      let formula = match @xml.find_child(nr_ch, "c:f") {
        None => ""
        Some(f) => @xml.get_text_content(f)
      }
      let cache = match @xml.find_child(nr_ch, "c:numCache") {
        None => NumData::empty()
        Some(nc) => parse_num_data(@xml.get_children(nc))
      }
      return NumSource(formula, cache)
    }
    None => ()
  }
  // Try c:numLit
  match @xml.find_child(ch, "c:numLit") {
    Some(nl) => return NumSource("", parse_num_data(@xml.get_children(nl)))
    None => ()
  }
  // Try c:strRef
  match @xml.find_child(ch, "c:strRef") {
    Some(sr) => {
      let sr_ch = @xml.get_children(sr)
      let formula = match @xml.find_child(sr_ch, "c:f") {
        None => ""
        Some(f) => @xml.get_text_content(f)
      }
      let cache = match @xml.find_child(sr_ch, "c:strCache") {
        None => StrData::empty()
        Some(sc2) => parse_str_data(@xml.get_children(sc2))
      }
      return StrSource(formula, cache)
    }
    None => ()
  }
  // Try c:strLit
  match @xml.find_child(ch, "c:strLit") {
    Some(sl) => return StrSource("", parse_str_data(@xml.get_children(sl)))
    None => ()
  }
  NoData
}

///|
/// Parse numeric cache data (c:numCache or c:numLit children).
fn parse_num_data(ch : Array[@xml.XmlNode]) -> NumData {
  let format_code = match @xml.find_child(ch, "c:formatCode") {
    None => ""
    Some(fc) => @xml.get_text_content(fc)
  }
  let points = parse_pt_list(ch)
  { format_code, points }
}

///|
/// Parse string cache data (c:strCache or c:strLit children).
fn parse_str_data(ch : Array[@xml.XmlNode]) -> StrData {
  { points: parse_pt_list(ch) }
}

///|
/// Parse c:pt list from cache children.
fn parse_pt_list(ch : Array[@xml.XmlNode]) -> Array[ChartPoint] {
  let pts : Array[ChartPoint] = []
  let pt_nodes = @xml.find_children(ch, "c:pt")
  for i in 0..<pt_nodes.length() {
    let idx = attr_int(pt_nodes[i], "idx")
    let value = match @xml.find_child(@xml.get_children(pt_nodes[i]), "c:v") {
      None => ""
      Some(v) => @xml.get_text_content(v)
    }
    pts.push({ idx, value })
  }
  pts
}

// ── Chart shape properties ──────────────────────────────────────────────────

///|
/// Parse c:spPr from series/chart element children.
fn parse_chart_sp_pr(
  ch : Array[@xml.XmlNode],
  theme : ThemeData,
) -> ChartSpPr {
  match @xml.find_child(ch, "c:spPr") {
    None => ChartSpPr::none()
    Some(sppr) => parse_chart_sp_pr_node(sppr, theme)
  }
}

///|
fn parse_chart_sp_pr_node(
  sppr : @xml.XmlNode,
  theme : ThemeData,
) -> ChartSpPr {
  let sppr_ch = @xml.get_children(sppr)
  let fill = parse_solid_fill(sppr_ch, theme)
  let grad_fill = parse_gradient_fill(sppr_ch, theme)
  let patt_fill = parse_pattern_fill(sppr_ch, theme)
  let no_fill = match @xml.find_child(sppr_ch, "a:noFill") {
    Some(_) => true
    None => false
  }
  let (stroke, stroke_w) = match @xml.find_child(sppr_ch, "a:ln") {
    None => (Color::none(), 0)
    Some(ln) => {
      let ln_ch = @xml.get_children(ln)
      let w = attr_int(ln, "w")
      let c = parse_solid_fill(ln_ch, theme)
      (c, w)
    }
  }
  { fill, grad_fill, patt_fill, stroke, stroke_w, no_fill }
}

// ── Axis parsing ────────────────────────────────────────────────────────────

///|
fn parse_chart_axis(
  ch : Array[@xml.XmlNode],
  is_val : Bool,
  _theme : ThemeData,
) -> ChartAxis {
  let ax_pos_raw = child_attr_val(ch, "c:axPos")
  let tick_raw = child_attr_val(ch, "c:tickLblPos")
  // Scaling
  let (orientation, min_val, max_val) = match @xml.find_child(ch, "c:scaling") {
    None => ("minMax", "", "")
    Some(sc) => {
      let sc_ch = @xml.get_children(sc)
      let o = child_attr_val(sc_ch, "c:orientation")
      (
        if o.is_empty() { "minMax" } else { o },
        child_attr_val(sc_ch, "c:min"),
        child_attr_val(sc_ch, "c:max"),
      )
    }
  }
  // Number format
  let num_fmt = match @xml.find_child(ch, "c:numFmt") {
    None => "General"
    Some(nf) =>
      match @xml.get_attr(nf, "formatCode") {
        None => "General"
        Some(v) => v
      }
  }
  // Value axis specifics
  let cb_raw = if is_val { child_attr_val(ch, "c:crossBetween") } else { "" }
  let mu = if is_val { child_attr_val(ch, "c:majorUnit") } else { "" }
  // Axis title
  let title = match @xml.find_child(ch, "c:title") {
    None => ""
    Some(title_node) => {
      let tc = @xml.get_children(title_node)
      if tc.length() > 0 {
        extract_all_text(@xml.get_children(tc[0]))
      } else {
        ""
      }
    }
  }
  {
    is_val,
    ax_id: child_attr_int(ch, "c:axId"),
    cross_ax: child_attr_int(ch, "c:crossAx"),
    delete: child_attr_val(ch, "c:delete") == "1",
    ax_pos: if ax_pos_raw.is_empty() {
      if is_val { "l" } else { "b" }
    } else {
      ax_pos_raw
    },
    tick_lbl_pos: if tick_raw.is_empty() { "nextTo" } else { tick_raw },
    orientation,
    min_val,
    max_val,
    major_gridlines: match @xml.find_child(ch, "c:majorGridlines") {
      Some(_) => true
      None => false
    },
    minor_gridlines: match @xml.find_child(ch, "c:minorGridlines") {
      Some(_) => true
      None => false
    },
    num_fmt,
    cross_between: if cb_raw.is_empty() && is_val { "between" } else { cb_raw },
    major_unit: mu,
    title,
    sp_pr: ChartSpPr::none(),
  }
}

// ── Legend parsing ───────────────────────────────────────────────────────────

///|
fn parse_chart_legend(ch : Array[@xml.XmlNode]) -> ChartLegend {
  let pos = child_attr_val(ch, "c:legendPos")
  let overlay = child_attr_val(ch, "c:overlay") == "1"
  { position: if pos.is_empty() { "r" } else { pos }, overlay, show: true }
}

///|
/// Parse c:view3D element — 3D rotation/perspective for round-trip preservation.
/// Uses child_attr_val to distinguish missing elements (default) from val="0".
fn parse_view_3d(ch : Array[@xml.XmlNode]) -> ChartView3D {
  let rx_s = child_attr_val(ch, "c:rotX")
  let ry_s = child_attr_val(ch, "c:rotY")
  let dp_s = child_attr_val(ch, "c:depthPercent")
  let p_s = child_attr_val(ch, "c:perspective")
  {
    rot_x: if rx_s.is_empty() { 15 } else { @xml.parse_int(rx_s) },
    rot_y: if ry_s.is_empty() { 20 } else { @xml.parse_int(ry_s) },
    depth_percent: if dp_s.is_empty() { 100 } else { @xml.parse_int(dp_s) },
    r_ang_ax: child_attr_val(ch, "c:rAngAx") != "0",
    perspective: if p_s.is_empty() { 30 } else { @xml.parse_int(p_s) },
  }
}
