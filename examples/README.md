# pptx-svg Examples

## Vanilla JS

A single HTML file using pptx-svg via CDN (jsdelivr). No build step required.

**Run locally:**
```bash
# Any static file server works
npx serve examples/vanilla
# Open http://localhost:3000
```

Or open `vanilla/index.html` directly via a local server.

## React

A Vite + React + TypeScript example with simple editing features:
- Click a shape to select it
- Drag to move shapes
- Double-click text to edit
- Export saves all changes to PPTX

**Run:**
```bash
cd examples/react
npm install
npm run dev
# Open http://localhost:5173
```
