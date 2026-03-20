# Interactive Editing Guide

This guide describes how to build an interactive PPTX editor UI using pptx-svg's shape-level APIs.

## API Overview

### Shape-level Wasm APIs

| Method | Description |
|--------|-------------|
| `renderShapeSvg(slideIdx, shapeIdx)` | Render a single shape as SVG fragment |
| `updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot)` | Update position/size/rotation (EMU), returns re-rendered SVG |
| `updateShapeText(slideIdx, shapeIdx, paraIdx, runIdx, text)` | Update text content, returns re-rendered SVG |
| `updateShapeFill(slideIdx, shapeIdx, r, g, b)` | Update solid fill color, returns re-rendered SVG |

All `update*` methods:
- Modify the cached SlideData in-place (no XML re-parse)
- Mark the slide as modified for export
- Return the re-rendered shape SVG with its `<defs>`

### Unit Conversion Helpers

```typescript
import { pxToEmu, emuToPx, ptToHundredths, degreesToOoxml } from 'pptx-svg';

pxToEmu(100)        // 952500 EMU
emuToPx(914400)     // 96 px
ptToHundredths(18)  // 1800
degreesToOoxml(90)  // 5400000
```

### SVG DOM Helpers

```typescript
import { findShapeElement, getShapeTransform, getAllShapes, getSlideScale } from 'pptx-svg';

const shapes = getAllShapes(svgElement);
const g = findShapeElement(svgElement, 0);
const transform = getShapeTransform(g);  // { x, y, cx, cy, rot } in EMU
const scale = getSlideScale(svgElement); // EMU per pixel
```

## Recommended Pattern: Preview + Commit

For smooth drag/resize interactions, use a two-phase approach:

```
mousedown  -> record initial transform
mousemove  -> update SVG transform attribute directly (no Wasm, 60fps)
             + debounce 200ms: updateShapeTransform() for text reflow preview
mouseup    -> updateShapeTransform() to commit -> replace DOM with result
```

### Implementation Example

```typescript
import { PptxRenderer, findShapeElement, getShapeTransform, getSlideScale, pxToEmu } from 'pptx-svg';

const renderer = new PptxRenderer();
await renderer.init();
await renderer.loadPptx(buffer);

// Render slide
const svgString = renderer.renderSlideSvg(0);
container.innerHTML = svgString;
const svg = container.querySelector('svg');
const scale = getSlideScale(svg);

let dragState = null;
let debounceTimer = null;

svg.addEventListener('mousedown', (e) => {
  const g = e.target.closest('g[data-ooxml-shape-idx]');
  if (!g) return;

  const shapeIdx = parseInt(g.dataset.ooxmlShapeIdx);
  const transform = getShapeTransform(g);
  dragState = { g, shapeIdx, startX: e.clientX, startY: e.clientY, transform };
});

svg.addEventListener('mousemove', (e) => {
  if (!dragState) return;

  // Phase 1: Pure SVG transform for 60fps visual feedback
  const dx = e.clientX - dragState.startX;
  const dy = e.clientY - dragState.startY;
  const { transform: t } = dragState;
  const newX = t.x / scale + dx;
  const newY = t.y / scale + dy;

  // Apply CSS transform for instant visual feedback
  dragState.g.style.transform = `translate(${dx}px, ${dy}px)`;

  // Phase 2: Debounced Wasm call for text reflow preview
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    const emuX = t.x + pxToEmu(dx);
    const emuY = t.y + pxToEmu(dy);
    // Preview only — don't replace DOM yet
    renderer.updateShapeTransform(0, dragState.shapeIdx,
      emuX, emuY, t.cx, t.cy, t.rot);
  }, 200);
});

svg.addEventListener('mouseup', (e) => {
  if (!dragState) return;
  clearTimeout(debounceTimer);

  const dx = e.clientX - dragState.startX;
  const dy = e.clientY - dragState.startY;
  const { transform: t } = dragState;

  // Commit: update via Wasm and replace DOM
  const shapeSvg = renderer.updateShapeTransform(0, dragState.shapeIdx,
    t.x + pxToEmu(dx), t.y + pxToEmu(dy), t.cx, t.cy, t.rot);

  // Replace the shape's <g> element and update defs
  dragState.g.style.transform = '';
  dragState.g.outerHTML = shapeSvg;
  dragState = null;
});
```

### Key Benefits

- **60fps drag**: Native SVG transform manipulation, no Wasm calls
- **Text reflow**: Debounced `updateShapeTransform()` recalculates text wrapping at new dimensions
- **Single-shape rendering**: Only the affected shape is re-rendered, not the entire slide
- **Round-trip safe**: All updates go through the SlideData cache, ensuring export correctness

## Text Editing

```typescript
// Update the first run of the first paragraph in shape 0
const newSvg = renderer.updateShapeText(0, 0, 0, 0, 'New text content');
// Replace the shape element in the DOM
shapeElement.outerHTML = newSvg;
```

## Fill Color Editing

```typescript
// Set shape 0 to red
const newSvg = renderer.updateShapeFill(0, 0, 255, 0, 0);
shapeElement.outerHTML = newSvg;
```

## Export

After editing, export the modified PPTX:

```typescript
const pptxBuffer = await renderer.exportPptx();
// Download or send to server
const blob = new Blob([pptxBuffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
```

All shape-level edits are automatically included in the export — the `update*` methods mark slides as modified.
