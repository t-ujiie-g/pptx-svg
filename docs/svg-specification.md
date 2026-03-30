# SVG Output Specification

pptx-svg converts each PPTX slide into an SVG element with `data-ooxml-*` attributes that preserve the original OOXML metadata. This enables lossless round-trip: SVG can be edited and converted back to PPTX without data loss.

## Overview

```xml
<svg xmlns="http://www.w3.org/2000/svg"
     width="960" height="540" viewBox="0 0 960 540"
     data-ooxml-slide-cx="9144000" data-ooxml-slide-cy="5143500"
     data-ooxml-bg="FFFFFF" data-ooxml-scale="...">
  <!-- Each shape is a <g> element -->
  <g data-ooxml-shape-type="autoshape" data-ooxml-geom="rect" ...>
    <!-- Visual SVG elements (rect, path, text, etc.) -->
  </g>
</svg>
```

## Units

| Unit | Description | Conversion |
|------|-------------|------------|
| EMU (English Metric Unit) | OOXML's base unit | 1 inch = 914,400 EMU |
| 60,000ths of a degree | Rotation/angle unit | 360 degrees = 21,600,000 |
| Hundredths-of-a-point | Font size unit | 18pt = 1800 |
| Percentage (0-100,000) | Used for gradient stops, alpha, etc. | 100% = 100,000 |

## Slide-Level Attributes

Set on the root `<svg>` element.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-slide-cx` | int | Slide width in EMU |
| `data-ooxml-slide-cy` | int | Slide height in EMU |
| `data-ooxml-bg` | string | Background color (`RRGGBB` hex or `"none"`) |
| `data-ooxml-scale` | int | EMU-to-pixel conversion factor |
| `data-ooxml-hidden` | string | `"1"` if slide is hidden (`<p:sld show="0">`); absent otherwise |
| `data-ooxml-transition-xml` | string | Raw `<p:transition>` XML (XML-escaped); absent if no transition |
| `data-ooxml-timing-xml` | string | Raw `<p:timing>` XML (XML-escaped); absent if no timing |

## Shape Wrapper Attributes

Each shape is rendered as a `<g>` element with these attributes.

### Identification

| Attribute | Type | Description | Values |
|-----------|------|-------------|--------|
| `data-ooxml-shape-type` | string | Shape kind | `"autoshape"`, `"picture"`, `"chart"`, `"table"`, `"group"` |
| `data-ooxml-geom` | string | Geometry preset | `"rect"`, `"ellipse"`, `"roundRect"`, `"line"`, `"custGeom"`, connector names, etc. |
| `data-ooxml-shape-idx` | int | Shape index in slide | 0-based |
| `data-ooxml-ph-type` | string | Placeholder type | `"ctrTitle"`, `"body"`, `"sldNum"`, `"ftr"`, `"dt"`, etc. |
| `data-ooxml-ph-idx` | int | Placeholder index | integer |

### Transform (Position & Size)

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-x` | int | Left position (EMU) |
| `data-ooxml-y` | int | Top position (EMU) |
| `data-ooxml-cx` | int | Width (EMU) |
| `data-ooxml-cy` | int | Height (EMU) |
| `data-ooxml-rot` | int | Rotation (60,000ths of degree) |
| `data-ooxml-flip-h` | string | Horizontal flip (`"0"` or `"1"`) |
| `data-ooxml-flip-v` | string | Vertical flip (`"0"` or `"1"`) |

### Solid Fill

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-fill` | string | Fill color (`RRGGBB` hex or `"none"`) |

### Gradient Fill

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-grad-stops` | string | Color stops: `"pos,RRGGBB;pos,RRGGBB;..."` (pos: 0-100000) |
| `data-ooxml-grad-angle` | int | Gradient angle (60,000ths of degree) |
| `data-ooxml-grad-path` | string | `"path"` for radial, empty for linear |
| `data-ooxml-grad-rot` | string | Rotate with shape (`"0"` or `"1"`) |
| `data-ooxml-grad-ftl` | int | Radial fill-to left (0-100000) |
| `data-ooxml-grad-ftt` | int | Radial fill-to top |
| `data-ooxml-grad-ftr` | int | Radial fill-to right |
| `data-ooxml-grad-ftb` | int | Radial fill-to bottom |
| `data-ooxml-grad-tile-flip` | string | Tile flip mode (`"x"`, `"y"`, `"xy"`, or empty) |

### Blip Fill (Image)

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-blip-rid` | string | Image relationship ID |
| `data-ooxml-blip-stretch` | string | `"1"` for stretch, `"0"` for tile |
| `data-ooxml-blip-src-l/t/r/b` | int | Source crop (0-100000) |
| `data-ooxml-blip-tile-tx/ty` | int | Tile offset (EMU) |
| `data-ooxml-blip-tile-sx/sy` | int | Tile scale (0-100000) |
| `data-ooxml-blip-tile-flip` | string | Tile flip mode |
| `data-ooxml-blip-tile-algn` | string | Tile alignment (`"tl"`, `"ctr"`, `"br"`, etc.) |
| `data-ooxml-blip-alpha` | int | Alpha (0-100000, 100000=opaque) |
| `data-ooxml-blip-bright` | int | Brightness (-100000 to 100000) |
| `data-ooxml-blip-contrast` | int | Contrast (-100000 to 100000) |
| `data-ooxml-blip-duotone1/2` | string | Duotone colors (`RRGGBB`) |
| `data-ooxml-blip-clr-from/to` | string | Color shift (`RRGGBB`) |
| `data-ooxml-blip-svg-rid` | string | SVG image relationship ID |

### Pattern Fill

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-patt-prst` | string | Pattern preset (e.g. `"pct5"`, `"ltDnDiag"`, `"cross"`) |
| `data-ooxml-patt-fg` | string | Foreground color (`RRGGBB`) |
| `data-ooxml-patt-bg` | string | Background color (`RRGGBB`) |

### Stroke

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-stroke` | string | Stroke color (`RRGGBB` or `"none"`) |
| `data-ooxml-stroke-w` | int | Stroke width (EMU) |
| `data-ooxml-stroke-dash` | string | `"solid"`, `"dash"`, `"dashDot"`, `"lgDash"`, etc. |
| `data-ooxml-stroke-cap` | string | `"flat"`, `"rnd"`, `"sq"` |
| `data-ooxml-stroke-join` | string | `"miter"`, `"round"`, `"bevel"` |
| `data-ooxml-stroke-miter-lim` | int | Miter limit |
| `data-ooxml-stroke-head-type` | string | Arrow head: `"triangle"`, `"stealth"`, `"diamond"`, `"oval"`, `"none"` |
| `data-ooxml-stroke-head-w/len` | string | `"sm"`, `"med"`, `"lg"` |
| `data-ooxml-stroke-tail-type/w/len` | string | Same as head |
| `data-ooxml-stroke-cmpd` | string | Compound line: `"sng"`, `"dbl"`, `"tri"` |
| `data-ooxml-stroke-no-fill` | string | `"1"` if no stroke fill |

### Hyperlinks

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-sh-link-rid` | string | Shape click hyperlink (relationship ID) |
| `data-ooxml-sh-link-hover-rid` | string | Shape hover hyperlink |

### Connection Points

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-st-cxn-id` | int | Start connection shape ID (-1 = none) |
| `data-ooxml-st-cxn-idx` | int | Start connection point index |
| `data-ooxml-end-cxn-id` | int | End connection shape ID |
| `data-ooxml-end-cxn-idx` | int | End connection point index |

### SmartArt / AlternateContent

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-mc-choice` | string (XML-escaped) | Raw `mc:Choice` XML for SmartArt round-trip. Present only on shapes parsed from `mc:Fallback`. |

### OLE / Embedded Objects

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-ole` | string (XML-escaped) | Raw `p:graphicFrame` XML for OLE round-trip. The shape renders the fallback image from `p:oleObj/p:pic`. |

### Math Equations (OMML)

Present on `<tspan>` elements that represent math equations:

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-math-xml` | string (XML-escaped) | Raw `m:oMath` or `m:oMathPara` XML. Visual SVG rendering of fractions, radicals, integrals, matrices, etc. is generated alongside this attribute for round-trip preservation. |

### Custom Geometry

When `data-ooxml-geom="custGeom"`:

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-cust-gdlst` | string | Guide formulas (comma-separated) |
| `data-ooxml-cust-paths` | string | Path definitions |
| `data-ooxml-cust-pw/ph` | int | Path coordinate space dimensions |
| `data-ooxml-cust-rl/rt/rr/rb` | string | Text rectangle guide formulas |
| `data-ooxml-cust-cxn` | string | Connection points (serialized XML) |

### Connector Adjustments

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-cxn-adj` | string | Adjustment values (comma-separated integers) |

## Text Element Attributes

Set on the `<text>` element within a shape.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-font-face` | string | Theme default font (minor font) applied as fallback |

## Text Body Properties

Set on the shape `<g>` element.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-anchor` | string | Vertical alignment: `"t"`, `"ctr"`, `"b"` |
| `data-ooxml-l-ins/t-ins/r-ins/b-ins` | int | Margins (EMU, -1 = default) |
| `data-ooxml-auto-fit` | string | `"1"` if auto-fit enabled |
| `data-ooxml-font-scale` | int | Font scale for normAutofit (0-100000, -1 = inactive) |
| `data-ooxml-ln-spc-reduction` | int | Line spacing reduction (0-100000, -1 = inactive) |
| `data-ooxml-wrap` | string | `"none"`, `"square"`, `"shrinkToFit"` |
| `data-ooxml-bp-rot` | int | Text body rotation (60,000ths of degree) |
| `data-ooxml-bp-vert` | string | Vertical text: `"vert"`, `"eaVert"`, `"vert270"`, `"wordArtVert"` |
| `data-ooxml-bp-num-cols` | int | Column count |
| `data-ooxml-bp-col-spacing` | int | Column spacing (EMU) |

### Warp Transform

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-warp-prst` | string | Warp preset (e.g. `"textArchDown"`, `"textCircle"`) |
| `data-ooxml-warp-av1/av2` | int | Warp adjustments (-1 = inactive) |

## Paragraph Attributes

Set on paragraph-level `<tspan>` or `<text>` elements.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-para-idx` | string | Paragraph index (0-based) |
| `data-ooxml-para-align` | string | `"l"`, `"ctr"`, `"r"`, `"just"` |
| `data-ooxml-para-lvl` | string | Outline level (0-8) |
| `data-ooxml-bu-font` | string | Bullet font name |
| `data-ooxml-bu-size` | string | Bullet size |
| `data-ooxml-bu-color` | string | Bullet color (`RRGGBB`) |
| `data-ooxml-rtl` | string | Right-to-left (`"1"` or `"0"`) |

## Text Run Attributes

Set on run-level `<tspan>` elements.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-run-idx` | string | Run index (0-based) |
| `data-ooxml-bold` | string | Bold (`"1"` or `"0"`) |
| `data-ooxml-font-size` | string | Font size (hundredths-of-a-point, e.g. `"1800"` = 18pt) |
| `data-ooxml-color` | string | Text color (`RRGGBB` or `"none"`) |
| `data-ooxml-run-font` | string | Latin font face |
| `data-ooxml-ea-font` | string | East Asian font |
| `data-ooxml-cs-font` | string | Complex script font |
| `data-ooxml-sym-font` | string | Symbol font |
| `data-ooxml-underline` | string | `"sng"`, `"dbl"`, `"dotted"`, `"wave"`, `"none"`, etc. |
| `data-ooxml-strike` | string | `"sngStrike"`, `"dblStrike"`, or empty |
| `data-ooxml-baseline` | string | Baseline shift (positive=super, negative=sub) |
| `data-ooxml-char-spacing` | string | Character spacing (EMU) |
| `data-ooxml-kern` | string | Kerning threshold (hundredths-of-a-point) |
| `data-ooxml-cap` | string | `"small"`, `"all"`, or `"none"` |
| `data-ooxml-hlink-rid` | string | Hyperlink relationship ID |
| `data-ooxml-hlink-hover-rid` | string | Hover hyperlink relationship ID |
| `data-ooxml-outline-color` | string | Text outline color (`RRGGBB`) |
| `data-ooxml-outline-w` | string | Text outline width (EMU) |

### Text Run Effects (on `<tspan>`)

These attributes preserve text-run-level effects for round-trip. Prefixed `reff-` (run effects) to distinguish from shape-level effects.

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-reff-os-blur` | int | Outer shadow blur radius (EMU) |
| `data-ooxml-reff-os-dist` | int | Outer shadow distance (EMU) |
| `data-ooxml-reff-os-dir` | int | Outer shadow direction (60,000ths of degree) |
| `data-ooxml-reff-os-clr` | string | Outer shadow color (`RRGGBB`) |
| `data-ooxml-reff-gl-rad` | int | Glow radius (EMU) |
| `data-ooxml-reff-gl-clr` | string | Glow color (`RRGGBB`) |
| `data-ooxml-reff-blur-rad` | int | Blur radius (EMU) |
| `data-ooxml-reff-ps-prst` | string | Preset shadow name (`shdw1`–`shdw20`) |
| `data-ooxml-reff-ps-dist` | int | Preset shadow distance (EMU) |
| `data-ooxml-reff-ps-dir` | int | Preset shadow direction (60,000ths of degree) |
| `data-ooxml-reff-ps-clr` | string | Preset shadow color (`RRGGBB` or `RRGGBBAA`) |

### Text Gradient/Pattern Fill (on `<tspan>`)

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-tgrad-stops` | string | Gradient stops |
| `data-ooxml-tgrad-angle` | string | Gradient angle |
| `data-ooxml-tgrad-path` | string | `"path"` for radial |
| `data-ooxml-tgrad-rot` | string | Rotate with shape |
| `data-ooxml-tpatt-prst` | string | Pattern preset |
| `data-ooxml-tpatt-fg/bg` | string | Pattern colors |

## Effects

Set on the shape `<g>` element.

### Outer Shadow

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-os-blur` | int | Blur radius (EMU) |
| `data-ooxml-eff-os-dist` | int | Shadow distance (EMU) |
| `data-ooxml-eff-os-dir` | int | Direction (60,000ths of degree) |
| `data-ooxml-eff-os-clr` | string | Shadow color (`RRGGBB`) |
| `data-ooxml-eff-os-sx/sy` | int | Scale (0-100000) |
| `data-ooxml-eff-os-algn` | string | Alignment (`"tl"` through `"br"`) |
| `data-ooxml-eff-os-rot` | string | Rotate with shape |

### Inner Shadow

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-is-blur/dist/dir` | int | Same units as outer shadow |
| `data-ooxml-eff-is-clr` | string | Shadow color |

### Glow

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-gl-rad` | int | Glow radius (EMU) |
| `data-ooxml-eff-gl-clr` | string | Glow color |

### Soft Edge

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-se-rad` | int | Edge radius (EMU) |

### Reflection

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-refl-blur/dist/dir` | int | Position/direction |
| `data-ooxml-eff-refl-sta/enda` | int | Start/end alpha (0-100000) |
| `data-ooxml-eff-refl-fadedir` | int | Fade direction |
| `data-ooxml-eff-refl-sx/sy` | int | Scale |
| `data-ooxml-eff-refl-algn` | string | Alignment |
| `data-ooxml-eff-refl-rot` | string | Rotate with shape |

### Blur

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-blur-rad` | int | Blur radius (EMU) |

### Preset Shadow

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-ps-prst` | string | Preset name (`shdw1`..`shdw20`) |
| `data-ooxml-eff-ps-dist` | int | Distance (EMU) |
| `data-ooxml-eff-ps-dir` | int | Direction (60,000ths of degree) |
| `data-ooxml-eff-ps-clr` | string | Shadow color (`RRGGBB` or `RRGGBBAA`) |

### Fill Overlay

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-eff-fo-blend` | string | Blend mode (`over`, `mult`, `screen`, `darken`, `lighten`) |
| `data-ooxml-eff-fo-clr` | string | Overlay solid fill color (`RRGGBB` or `AARRGGBB`) |

## 3D Properties

Data-only preservation for round-trip. Not rendered visually in SVG.

### Scene 3D (on `<svg>`)

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-3d-cam` | string | Camera preset |
| `data-ooxml-3d-rig` | string | Light rig preset |
| `data-ooxml-3d-rig-dir` | string | Light direction |

### Shape 3D (on `<g>`)

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-3d-bt-w/h/prst` | int/string | Top bevel |
| `data-ooxml-3d-bb-w/h/prst` | int/string | Bottom bevel |
| `data-ooxml-3d-ext-h/clr` | int/string | Extrusion height and color |
| `data-ooxml-3d-cnt-w/clr` | int/string | Contour width and color |
| `data-ooxml-3d-mat` | string | Material preset |
| `data-ooxml-3d-z` | int | Z-order |

## Group Shape

| Attribute | Type | Description |
|-----------|------|-------------|
| `data-ooxml-grp-ch-off-x/y` | int | Child offset (EMU) |
| `data-ooxml-grp-ch-ext-cx/cy` | int | Child coordinate space (EMU) |

## SVG Structure Example

```xml
<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540"
     viewBox="0 0 960 540"
     data-ooxml-slide-cx="9144000" data-ooxml-slide-cy="5143500"
     data-ooxml-bg="FFFFFF" data-ooxml-scale="9525">

  <!-- Shape: blue rectangle with text -->
  <g data-ooxml-shape-type="autoshape" data-ooxml-geom="rect"
     data-ooxml-shape-idx="0"
     data-ooxml-x="457200" data-ooxml-y="274638"
     data-ooxml-cx="8229600" data-ooxml-cy="1143000"
     data-ooxml-fill="4472C4" data-ooxml-stroke="2F5496" data-ooxml-stroke-w="12700"
     data-ooxml-anchor="ctr">
    <rect x="48" y="28.8" width="864" height="120" rx="0"
          fill="#4472C4" stroke="#2F5496" stroke-width="1.33"/>
    <text>
      <tspan data-ooxml-para-idx="0" data-ooxml-para-align="ctr">
        <tspan data-ooxml-run-idx="0" data-ooxml-bold="1"
               data-ooxml-font-size="2400" data-ooxml-color="FFFFFF"
               data-ooxml-run-font="Calibri"
               x="480" dy="0" text-anchor="middle"
               fill="#FFFFFF" font-size="24pt" font-weight="bold">
          Hello World
        </tspan>
      </tspan>
    </text>
  </g>
</svg>
```

## Round-Trip Workflow

1. **Render**: `renderer.renderSlideSvg(0)` returns SVG with `data-ooxml-*` attributes
2. **Edit**: Modify the SVG (change text content, move shapes, alter fills)
3. **Update**: `renderer.updateSlideFromSvg(0, editedSvg)` parses `data-ooxml-*` back into internal data
4. **Export**: `renderer.exportPptx()` generates a valid .pptx file

When editing SVG for round-trip:
- **Modify visual properties** (fill color, text, position) freely
- **Preserve `data-ooxml-*` attributes** to maintain OOXML metadata
- **Shape structure** (`<g>` wrappers, `<tspan>` hierarchy) should be maintained
- Attributes not represented visually (3D, custom geometry guides) are preserved through `data-ooxml-*` for lossless round-trip
