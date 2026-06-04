# Interactive Editing Guide

This guide describes how to build an interactive PPTX editor UI using pptx-svg's shape-level APIs.

## API Overview

### Shape-level Wasm APIs

| Method | Description |
|--------|-------------|
| `renderShapeSvg(slideIdx, shapeIdx)` | Render a single shape as SVG fragment |
| `updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot)` | Update position/size/rotation (EMU), returns re-rendered SVG |
| `updateShapesTransform(slideIdx, items)` | Atomically update several shapes' transforms as one undo step, returns `OK:<count>` |
| `updateShapeText(slideIdx, shapeIdx, paraIdx, runIdx, text)` | Update text content, returns re-rendered SVG |
| `updateShapeFill(slideIdx, shapeIdx, r, g, b)` | Update solid fill color, returns re-rendered SVG |
| `deleteShape(slideIdx, shapeIdx)` | Delete a shape by index (supports group children via composite index) |
| `addShape(slideIdx, geomType, x, y, cx, cy, fillR, fillG, fillB)` | Add a basic shape (rect/ellipse/roundRect/line), returns `OK:<index>` |
| `duplicateShape(slideIdx, shapeIdx, dxEmu?, dyEmu?)` | Duplicate a shape with offset, returns `OK:<index>` |
| `getShapeSpec(slideIdx, shapeIdx)` | Copy a shape to a portable JSON spec (media inlined), returns spec or `ERROR:...` |
| `insertShapeSpec(slideIdx, spec, dxEmu?, dyEmu?)` | Paste a shape spec onto a slide (re-links media), returns `OK:<index>` |
| `updateShapeGradientFill(slideIdx, shapeIdx, angle, stops)` | Apply linear gradient fill, returns re-rendered SVG |
| `addShapeText(slideIdx, shapeIdx, text, fontSize?, colorR?, colorG?, colorB?)` | Add a text paragraph to a shape, returns `OK:<paraIndex>` |
| `updateShapeStroke(slideIdx, shapeIdx, r, g, b, widthEmu?, dash?)` | Set stroke color/width/dash, returns re-rendered SVG |
| `addParagraph(slideIdx, shapeIdx, text, align?)` | Add paragraph with alignment, returns `OK:<paraIndex>` |
| `deleteParagraph(slideIdx, shapeIdx, paraIdx)` | Delete a paragraph, returns `OK` |
| `addRun(slideIdx, shapeIdx, paraIdx, text)` | Add a text run to a paragraph, returns `OK:<runIndex>` |
| `deleteRun(slideIdx, shapeIdx, paraIdx, runIdx)` | Delete a text run, returns `OK` |
| `updateTextRunStyle(slideIdx, shapeIdx, paraIdx, runIdx, bold?, italic?)` | Set bold/italic (1/0/-1), returns re-rendered SVG |
| `updateTextRunFontSize(slideIdx, shapeIdx, paraIdx, runIdx, fontSize)` | Set font size (hundredths of pt), returns re-rendered SVG |
| `updateTextRunColor(slideIdx, shapeIdx, paraIdx, runIdx, r, g, b)` | Set text color (r=-1 to inherit), returns re-rendered SVG |
| `updateTextRunFont(slideIdx, shapeIdx, paraIdx, runIdx, fontFace?, eaFont?, csFont?)` | Set font family, returns re-rendered SVG |
| `updateParagraphAlign(slideIdx, shapeIdx, paraIdx, align)` | Set paragraph alignment, returns re-rendered SVG |
| `updateTextRunDecoration(slideIdx, shapeIdx, paraIdx, runIdx, underline?, strike?, baseline?)` | Set underline/strike/super-subscript, returns re-rendered SVG |
| `bringToFront(slideIdx, shapeIdx)` / `sendToBack(...)` | Move a shape to the front / back of the z-order, returns `OK:<newShapeIdx>` |
| `bringForward(slideIdx, shapeIdx)` / `sendBackward(...)` | Move a shape one step toward the front / back, returns `OK:<newShapeIdx>` |

All shape `update*` methods:
- Modify the cached SlideData in-place (no XML re-parse)
- Mark the slide as modified for export
- Return the re-rendered shape SVG with its `<defs>`

> **Note (headless / batch use):** Editing APIs parse the target slide on demand,
> so you do **not** need to call `renderSlideSvg(idx)` first. A `loadPptx → updateShapeText → exportPptx`
> flow works without ever producing SVG. (Before 0.5.10, editing a slide that had
> not yet been rendered silently failed with `ERROR:shape index out of range`.)

### Slide Management APIs

| Method | Description |
|--------|-------------|
| `addSlide(afterIdx?, sourceSlideIdx?)` | Add a blank slide at the given position |
| `deleteSlide(slideIdx)` | Remove a slide (minimum 1 must remain) |
| `reorderSlides(newOrder)` | Reorder slides by permutation array |

Slide management methods update package metadata (`presentation.xml`, `.rels`, `[Content_Types].xml`) and re-initialize the Wasm engine automatically.

### Image APIs

| Method | Description |
|--------|-------------|
| `addImage(slideIdx, imageData, mimeType, x, y, cx, cy)` | Add a picture shape with the given image data (Uint8Array). Handles media file, `.rels`, and `[Content_Types].xml` updates. Returns `"OK:<shapeIdx>"` |
| `replaceImage(slideIdx, shapeIdx, imageData, mimeType)` | Replace the image of an existing picture shape. Returns re-rendered SVG |
| `deleteImage(slideIdx, shapeIdx)` | Delete a picture shape and clean up orphaned media files. Returns `"OK"` |

Supported MIME types: `image/png`, `image/jpeg`, `image/gif`, `image/bmp`, `image/tiff`, `image/svg+xml`, `image/x-emf`, `image/x-wmf`.

Coordinates (`x`, `y`, `cx`, `cy`) are in EMU (English Metric Units). Use `pxToEmu()` for conversion.

### History APIs (Undo / Redo)

| Method | Description |
|--------|-------------|
| `undo()` | Revert the most recent edit (or batch). Returns a JSON `HistoryResult` string, or `"ERROR:nothing to undo"` |
| `redo()` | Re-apply the most recently undone edit. Returns a JSON `HistoryResult` string, or `"ERROR:nothing to redo"` |
| `canUndo()` / `canRedo()` | Whether an undo / redo step is available |
| `beginBatch()` / `endBatch()` | Collapse multiple edits into a single undo step (nestable) |
| `clearHistory()` | Discard all undo/redo history |

Every mutating editing method (shape/text/fill/stroke, add/delete/duplicate, image ops, and slide add/delete/reorder) automatically records a history checkpoint. See [Undo / Redo](#undo--redo) below.

### Inline Text Editing APIs

| Method | Description |
|--------|-------------|
| `getTextLayout(slideIdx, shapeIdx)` | Returns JSON text geometry (EMU): box + lines → run boxes → per-character boxes. For drawing carets / selection rectangles. |
| `hitTestText(slideIdx, shapeIdx, xEmu, yEmu)` | Returns JSON `{ paraIdx, runIdx, charOffset, paraOffset }` for a click point (EMU). |
| `replaceTextRange(slideIdx, shapeIdx, startPara, startChar, endPara, endChar, newText)` | Replace a text range, preserving boundary run formatting. Undoable. |

See [Inline Text Editing](#inline-text-editing) below.

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

### Add text to a shape

```typescript
// Add a paragraph with 18pt text to shape 0
const result = renderer.addShapeText(0, 0, 'Hello World', 1800);
// fontSize: hundredths of a point (1800 = 18pt), optional color (0-255)
renderer.addShapeText(0, 0, 'Red text', 1400, 255, 0, 0);
```

### Update existing text

```typescript
// Update the first run of the first paragraph in shape 0
const newSvg = renderer.updateShapeText(0, 0, 0, 0, 'New text content');
// Replace the shape element in the DOM
shapeElement.outerHTML = newSvg;
```

### Paragraph management

```typescript
// Add a centered paragraph with text
const result = renderer.addParagraph(0, shapeIdx, 'New paragraph', 'ctr');
// align: "l" (left), "ctr" (center), "r" (right), "just" (justify), "" (inherit)

// Delete paragraph at index 1
renderer.deleteParagraph(0, shapeIdx, 1);

// Change alignment
renderer.updateParagraphAlign(0, shapeIdx, 0, 'r');
```

### Run management

```typescript
// Add a run to paragraph 0
const result = renderer.addRun(0, shapeIdx, 0, 'appended text');

// Delete run at index 1 from paragraph 0
renderer.deleteRun(0, shapeIdx, 0, 1);
```

### Text formatting

```typescript
// Bold and italic (1 = on, 0 = off, -1 = no change)
renderer.updateTextRunStyle(0, shapeIdx, 0, 0, 1, -1);   // bold on
renderer.updateTextRunStyle(0, shapeIdx, 0, 0, -1, 1);   // italic on

// Font size (hundredths of a point: 1800 = 18pt, 0 = inherit)
renderer.updateTextRunFontSize(0, shapeIdx, 0, 0, 2400);  // 24pt

// Text color (RGB 0-255, r=-1 to inherit from theme)
renderer.updateTextRunColor(0, shapeIdx, 0, 0, 255, 0, 0);  // red

// Font family (empty string = no change)
renderer.updateTextRunFont(0, shapeIdx, 0, 0, 'Arial', 'MS Gothic', '');
// Arguments: fontFace (Latin), eaFont (East Asian), csFont (Complex Script)

// Underline, strikethrough, superscript/subscript
renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, 'sng', '', -1);       // single underline
renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, '', 'sngStrike', -1); // strikethrough
renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, '', '', 30000);       // superscript
renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, '', '', -25000);      // subscript
renderer.updateTextRunDecoration(0, shapeIdx, 0, 0, 'none', 'none', 0);   // remove all
// underline: "sng", "dbl", "" (no change), "none" (remove)
// strike: "sngStrike", "dblStrike", "" (no change), "none" (remove)
// baseline: 30000 (super), -25000 (sub), 0 (normal), -1 (no change)
```

## Fill Color Editing

```typescript
// Set shape 0 to red
const newSvg = renderer.updateShapeFill(0, 0, 255, 0, 0);
shapeElement.outerHTML = newSvg;
```

## Slide Management

Add, delete, and reorder slides programmatically. These methods update `presentation.xml`, `.rels`, and `[Content_Types].xml` automatically.

### Add Slide

```typescript
// Append a blank slide at the end (layout copied from last slide)
const { slideCount, insertedIdx } = await renderer.addSlide();

// Insert after slide 0 (becomes new slide 1)
await renderer.addSlide(0);

// Insert at the beginning
await renderer.addSlide(-1);

// Copy layout from slide 2
await renderer.addSlide(undefined, 2);
```

### Delete Slide

```typescript
// Delete slide at index 1
const { slideCount } = await renderer.deleteSlide(1);
// At least one slide must remain — throws if you try to delete the last one
```

### Reorder Slides

```typescript
// Reverse 2 slides: [1, 0]
await renderer.reorderSlides([1, 0]);

// Rotate 3 slides: slide 1 �� 0, slide 2 → 1, slide 0 → 2
await renderer.reorderSlides([1, 2, 0]);

// Swap slides 0 and 2 (keep 1 in place)
await renderer.reorderSlides([2, 1, 0]);
```

The argument is a permutation array where `newOrder[i]` is the old index of the slide that should appear at position `i`.

## Shape Management

### Add Shape

```typescript
// Add a red rectangle (position and size in EMU)
const result = renderer.addShape(0, 'rect', 914400, 914400, 1828800, 914400, 255, 0, 0);
const shapeIdx = parseInt(result.split(':')[1]);

// Add an ellipse with no fill (pass -1 for fill values)
renderer.addShape(0, 'ellipse', 0, 0, 914400, 914400);
```

Supported geometry types: `rect`, `ellipse`, `roundRect`, `line`.

### Delete Shape

```typescript
renderer.deleteShape(0, shapeIdx);
```

### Duplicate Shape

```typescript
// Duplicate shape with default offset (457200 EMU = 0.5 inch)
const result = renderer.duplicateShape(0, shapeIdx);
const newIdx = parseInt(result.split(':')[1]);

// Duplicate with custom offset
renderer.duplicateShape(0, shapeIdx, 914400, 914400);
```

## Gradient Fill

```typescript
const stops = [
  { pos: 0,      r: 255, g: 0,   b: 0 },   // Red at start
  { pos: 100000, r: 0,   g: 0,   b: 255 },  // Blue at end
];
// angle: 5400000 = 90 degrees (in 60000ths of a degree)
const svg = renderer.updateShapeGradientFill(0, shapeIdx, 5400000, stops);
shapeElement.outerHTML = svg;
```

## Stroke Editing

```typescript
// Set red stroke, 2pt width, dashed
const svg = renderer.updateShapeStroke(0, shapeIdx, 255, 0, 0, 25400, 'dash');

// Remove stroke (pass -1 for color)
renderer.updateShapeStroke(0, shapeIdx, -1, -1, -1, 0);
```

Dash presets: `dash`, `dot`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot`.

## Image Operations

```typescript
// Add a picture shape from image data
const imageData = new Uint8Array(await fetch('photo.png').then(r => r.arrayBuffer()));
const result = renderer.addImage(0, imageData, 'image/png',
  914400, 914400, 3657600, 2743200);  // x, y, width, height in EMU
const shapeIdx = parseInt(result.split(':')[1]);

// Replace the image of an existing picture shape
const newImage = new Uint8Array(await fetch('new-photo.jpg').then(r => r.arrayBuffer()));
renderer.replaceImage(0, shapeIdx, newImage, 'image/jpeg');

// Delete a picture shape (cleans up orphaned media files)
renderer.deleteImage(0, shapeIdx);
```

Supported MIME types: `image/png`, `image/jpeg`, `image/gif`, `image/bmp`, `image/tiff`, `image/svg+xml`, `image/x-emf`, `image/x-wmf`.

## Copy & Paste (cross-slide)

`getShapeSpec` / `insertShapeSpec` implement `Ctrl+C` / `Ctrl+V`, including pasting onto a **different** slide (or even a different presentation). `getShapeSpec` returns a portable, self-contained JSON string — the shape's OOXML plus any referenced images inlined as base64 — so it survives a clipboard round-trip. `insertShapeSpec` re-adds the media to the package and re-links the shape's image relationships to fresh rIds on the target slide.

```typescript
// Copy
const clipboard = renderer.getShapeSpec(0, shapeIdx); // JSON string

// Paste onto another slide, offset by (dx, dy) EMU
const res = renderer.insertShapeSpec(2, clipboard, 457200, 457200);
if (res.startsWith('OK:')) {
  const newIdx = parseInt(res.slice(3));
  renderer.renderSlideSvg(2); // re-render the target slide
}
```

Undoable (integrated with the history). **v1 limitations:** charts (serialized out-of-band) are not copyable (`getShapeSpec` returns `ERROR`); for OLE/SmartArt only inline image media is re-linked, not other external parts.

## Multi-shape Transform

For multi-select move/align, `updateShapesTransform` applies new transforms to several shapes in a single atomic call — every `shapeIdx` is validated before any change is applied (a bad index leaves the slide untouched), and the whole batch becomes **one** undo step:

```typescript
const res = renderer.updateShapesTransform(0, [
  { shapeIdx: 2, x: 1000000, y: 1000000, cx: 914400, cy: 914400, rot: 0 },
  { shapeIdx: 5, x: 2000000, y: 1000000, cx: 914400, cy: 914400, rot: 0 },
]); // → "OK:2"
renderer.renderSlideSvg(0); // re-render the slide afterwards
```

Values are EMU (rotation in 1/60000°). This is equivalent to wrapping individual `updateShapeTransform` calls in `beginBatch()`/`endBatch()`, but additionally guarantees atomicity (no partial application on a bad index) in a single Wasm call. A failed call records no undo step.

## Z-Order

Shapes paint in array order — the last shape in a slide is drawn on top (front-most), the first is at the back. These four methods reorder a shape within its container (the slide, or a group for group-child shapes via composite index):

```typescript
renderer.bringToFront(slideIdx, shapeIdx);   // → "OK:<newShapeIdx>"
renderer.sendToBack(slideIdx, shapeIdx);
renderer.bringForward(slideIdx, shapeIdx);   // one step toward front (no-op if front-most)
renderer.sendBackward(slideIdx, shapeIdx);   // one step toward back  (no-op if back-most)
```

Reordering changes the shape's index, so each call returns `"OK:<newShapeIdx>"` — use it to keep your selection pointing at the same shape:

```typescript
const res = renderer.bringToFront(0, selectedIdx);
if (res.startsWith('OK:')) selectedIdx = parseInt(res.slice(3));
```

All four are undoable (integrated with [Undo / Redo](#undo--redo)).

## Inline Text Editing

These three primitives let you build a PowerPoint/Google-Slides–style direct-typing experience (double-click to edit, IME input) using a `contentEditable` overlay: `getTextLayout` gives the geometry to draw a caret/selection, `hitTestText` maps a click to an insertion point, and `replaceTextRange` applies the edit while preserving formatting. All coordinates are in **EMU**.

```typescript
import type { TextLayout, TextHit } from 'pptx-svg';

// 1. Geometry for caret / selection rendering
const layout: TextLayout = JSON.parse(renderer.getTextLayout(slideIdx, shapeIdx));
// layout.box   → { x, y, cx, cy }            (the text body rect, EMU)
// layout.lines → [{ paraIdx, y, h, runs: [   (each visual line)
//   { paraIdx, runIdx, x, w, runCharStart, paraCharStart,
//     chars: [{ x, w }, ...] }                (per-character advance, EMU)
// ] }]

// 2. Click → caret position
const hit: TextHit = JSON.parse(renderer.hitTestText(slideIdx, shapeIdx, xEmu, yEmu));
// hit → { paraIdx, runIdx, charOffset, paraOffset }
//   charOffset = offset within the run; paraOffset = offset within the paragraph

// 3. Apply a typed/pasted/deleted range (paragraph-level offsets)
renderer.replaceTextRange(slideIdx, shapeIdx,
  hit.paraIdx, hit.paraOffset,   // start
  hit.paraIdx, hit.paraOffset,   // end (== start → insertion)
  'typed text');
```

`replaceTextRange` behavior:

- **Offsets are paragraph-level** (use `paraOffset` from `hitTestText`, not `charOffset`).
- A **collapsed range** (`start == end`) inserts; an **empty `newText`** deletes; a non-empty range with non-empty text replaces.
- The inserted text **inherits the formatting** of the run at the start boundary; runs are split/merged as needed and surrounding formatting is preserved.
- **`\n` in `newText` splits into paragraphs** (e.g. multi-line paste); a range spanning paragraphs **merges** them.
- It is **undoable** (integrated with the history above) and returns the re-rendered shape SVG.

**Scope / limitations of `getTextLayout` (v1):** targets horizontal LTR text (left/center/right/justify). Vertical text, text warp, OMML math, and multi-column bodies return only the bounding box with no per-line geometry. Line and run counts always match the rendered SVG (the layout shares the renderer's wrapping and autofit). Bullets are accounted for approximately on left-aligned lines.

```typescript
// Example: draw a caret at a click point
const hit = JSON.parse(renderer.hitTestText(0, shapeIdx, clickXEmu, clickYEmu));
const layout = JSON.parse(renderer.getTextLayout(0, shapeIdx));
for (const line of layout.lines) {
  for (const run of line.runs) {
    if (run.paraIdx === hit.paraIdx && hit.charOffset >= run.runCharStart &&
        hit.charOffset <= run.runCharStart + run.chars.length) {
      const i = hit.charOffset - run.runCharStart;
      const caretX = i < run.chars.length ? run.chars[i].x : run.x + run.w; // EMU
      drawCaret(caretX, line.y, line.h);
    }
  }
}
```

## Undo / Redo

Every mutating editing method records a checkpoint before it runs, so `undo()` / `redo()` work across all edit types — shape transforms, text, fills, strokes, add/delete/duplicate, image operations, and slide add/delete/reorder. This makes a `Ctrl+Z` / `Ctrl+Y` workflow trivial to wire up.

```typescript
// Configure history depth (default 50; 0 disables history)
const renderer = new PptxRenderer({ maxHistory: 100 });
// ... init + loadPptx ...

renderer.updateShapeTransform(0, 0, 1000000, 1000000, 2000000, 1000000, 0);

if (renderer.canUndo()) {
  const result = JSON.parse(renderer.undo()); // { slides: [0], slideCount: 2 }
  result.slides.forEach(i => repaint(i, renderer.renderSlideSvg(i)));
}

if (renderer.canRedo()) {
  JSON.parse(renderer.redo());
}
```

`undo()` / `redo()` return a JSON-encoded `HistoryResult` (`{ slides: number[]; slideCount: number }`) listing the 0-indexed slides whose content changed (re-render just those) and the slide count after the operation. On an empty stack they return `"ERROR:nothing to undo"` / `"ERROR:nothing to redo"`.

### Batching

Wrap a compound action (e.g. paste = add shape + set text + set fill) so a single `Ctrl+Z` reverts the whole thing. Batches are nestable; only the outermost pair takes effect.

```typescript
renderer.beginBatch();
try {
  const idx = parseInt(renderer.addShape(0, 'rect', 0, 0, 914400, 914400, 255, 0, 0).split(':')[1]);
  renderer.addShapeText(0, idx, 'Pasted', 1800, 255, 255, 255);
  renderer.updateShapeStroke(0, idx, 0, 0, 0, 12700, '');
} finally {
  renderer.endBatch();
}
// One undo() reverts all three edits.
```

### Notes

- History is cleared automatically on `loadPptx()`. Call `clearHistory()` to reset it manually.
- Snapshots are lightweight: they shallow-clone the file overrides (strings/bytes are shared by reference) plus the OOXML of any in-engine modified slides. `undo()` / `redo()` rebuild engine state from the snapshot, so they are O(document) and intended for user-initiated actions, not per-keystroke calls.
- Memory is bounded by `maxHistory` (oldest steps are discarded once exceeded).

## Export

After editing, export the modified PPTX:

```typescript
const pptxBuffer = await renderer.exportPptx();
// Download or send to server
const blob = new Blob([pptxBuffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
```

All edits — shape-level updates and slide management operations — are automatically included in the export.
