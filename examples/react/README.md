# pptx-svg React Editor

A PowerPoint / Google-Slides-style editor built on the published
[`pptx-svg`](https://www.npmjs.com/package/pptx-svg) package (`^0.6.0`). It is a
reference implementation showing how to drive every 0.6.0 editing API from a
React app.

## Run

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # tsc + vite build → dist/
```

Open a `.pptx` file (drag-and-drop or the **Open** button) to start editing, then
**Export** to download the edited deck.

## Features

| Area | What it shows |
|------|---------------|
| **Select / move / resize / rotate** | Click a shape, drag it, use the 8 resize handles or the rotate handle (Shift = 15° snap). |
| **Multi-select** | Shift-click to toggle; drag or arrow-key nudge moves the group in one undo step (`updateShapesTransform`). |
| **Inline text** | Double-click a text shape to type with a live caret; arrow keys move the caret, Enter adds a line (`getTextLayout` / `hitTestText` / `replaceTextRange`). |
| **Z-order** | To front / forward / backward / to back (`bringToFront` / `bringForward` / `sendBackward` / `sendToBack`). |
| **Copy / paste** | ⌘C / ⌘V across slides (`getShapeSpec` / `insertShapeSpec`). |
| **Undo / redo** | ⌘Z / ⌘⇧Z / ⌘Y, with a bounded history (`maxHistory: 100`). |
| **Fill / line** | Solid + 2-stop gradient fill, stroke colour / width / dash. |
| **Text formatting** | Per-run bold/italic/underline/strike/super-sub/size/colour/font, per-paragraph alignment, add/delete paragraph & run. |
| **Tables** | Set cell text, add/delete rows and columns. |
| **Images** | Insert, replace, delete. |
| **Slides** | Add, duplicate, delete, drag-reorder in the rail. |

### Keyboard shortcuts

| Shortcut | Action |
|----------|--------|
| `⌘Z` / `⌘⇧Z` / `⌘Y` | Undo / redo |
| `⌘C` / `⌘V` | Copy / paste shape |
| `⌘D` | Duplicate |
| `Delete` / `Backspace` | Delete selection |
| Arrow keys | Nudge (Shift = ×4) |
| `Esc` | Deselect / exit inline editing |

## Layout

```
TopBar          logo · undo/redo · export · open
InsertToolbar   text · shapes ▾ · image · align · z-order
┌──────────┬───────────────────────────┬────────────┐
│ SlideRail│         Canvas            │ Properties │
│ (thumbs) │  SVG + selection overlays │  (context) │
└──────────┴───────────────────────────┴────────────┘
```

## Structure

```
src/
├── App.tsx                  orchestrator: selection model, commit/refresh, insert/ops
├── constants.ts             EMU/RGB defaults (slide size, shape size, accent colour…)
├── hooks/
│   ├── useRenderer.ts       PptxRenderer lifecycle + thin wrappers over every 0.6.0 API
│   ├── useDrag.ts           move / resize / rotate / group-move on the canvas
│   ├── useInlineText.ts     double-click caret text editing
│   └── useKeyboard.ts       global shortcuts
├── components/
│   ├── TopBar.tsx           title, undo/redo, export, open
│   ├── InsertToolbar.tsx    insert + align + z-order
│   ├── SlideRail.tsx        vertical thumbnails (drag-reorder)
│   ├── PropertiesPanel.tsx  contextual editing (Arrange/Fill/Line/Text/Table/Image)
│   ├── TextPanel.tsx        per-paragraph / per-run text formatting
│   └── DropZone.tsx         file picker / drag-and-drop
└── utils/svg.ts             SVG DOM helpers + EMU↔px geometry + overlays
```

The slide SVG is inserted via `innerHTML`; selection, resize/rotate handles,
multi-select, and inline-text overlays are absolutely-positioned `<div>`s on top,
mapping between slide-absolute EMU and container CSS pixels.
