/**
 * pptx-svg React Editor — a PowerPoint / Google-Slides-style editor showcasing
 * every pptx-svg 0.6.0 editing API (undo/redo, inline text, z-order, multi-shape
 * transform, cross-slide copy/paste, table editing) on top of the existing
 * shape/text/fill/image/slide editing.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import { findShapeElement } from 'pptx-svg';
import { useRenderer } from './hooks/useRenderer';
import { useDrag } from './hooks/useDrag';
import { useKeyboard } from './hooks/useKeyboard';
import { useInlineText } from './hooks/useInlineText';
import { TopBar } from './components/TopBar';
import { InsertToolbar } from './components/InsertToolbar';
import { SlideRail } from './components/SlideRail';
import { PropertiesPanel } from './components/PropertiesPanel';
import { DropZone } from './components/DropZone';
import {
  insertSvgInto, showOverlay, showMultiOverlay, removeMultiOverlays,
  extractShapeInfo, okIndex, type ShapeInfo,
} from './utils/svg';
import {
  DEFAULT_SLIDE_CX, DEFAULT_SLIDE_CY, DEFAULT_SHAPE_SIZE,
  DEFAULT_IMAGE_CX, DEFAULT_IMAGE_CY, DEFAULT_STROKE_WIDTH, DEFAULT_FONT_SIZE,
  ACCENT_RGB, INK_RGB, PASTE_OFFSET,
} from './constants';

export default function App() {
  const renderer = useRenderer();
  const containerRef = useRef<HTMLDivElement>(null);
  const slide = renderer.slide;

  // ── State (+ refs so DOM event handlers read fresh values) ──
  const [selection, setSelectionState] = useState<ShapeInfo | null>(null);
  const selIdxRef = useRef(-1);
  const setSelection = useCallback((s: ShapeInfo | null) => { selIdxRef.current = s?.idx ?? -1; setSelectionState(s); }, []);

  const [multiSel, setMultiState] = useState<number[]>([]);
  const multiRef = useRef<number[]>([]);
  const setMulti = useCallback((m: number[]) => { multiRef.current = m; setMultiState(m); }, []);

  const [inlineShape, setInlineState] = useState<number | null>(null);
  const inlineRef = useRef<number | null>(null);
  const setInline = useCallback((v: number | null) => { inlineRef.current = v; setInlineState(v); }, []);

  const clipboardRef = useRef<string | null>(null);
  const [hasClipboard, setHasClipboard] = useState(false);

  const [thumbVer, setThumbVer] = useState<number[]>([]);
  const thumbSeed = useRef(1);
  const bumpThumb = useCallback((i: number) => setThumbVer(v => { const n = v.slice(); n[i] = (n[i] ?? 0) + 1; return n; }), []);
  /** Force every thumbnail (0..count-1) to re-render — after structural changes. */
  const reseedThumbs = useCallback((count: number) => {
    thumbSeed.current += 1;
    setThumbVer(Array.from({ length: count }, () => thumbSeed.current));
  }, []);

  // ── Helpers ──
  const svgIn = () => containerRef.current?.querySelector('svg') as SVGSVGElement | null;

  const drawSelection = useCallback((idx: number): ShapeInfo | null => {
    const c = containerRef.current; if (!c) return null;
    const svg = svgIn(); if (!svg) return null;
    const g = findShapeElement(svg, idx);
    if (!g) return null;
    showOverlay(c, g, true);
    return extractShapeInfo(g);
  }, []);

  const drawMulti = useCallback((indices: number[]) => {
    const c = containerRef.current, svg = svgIn();
    if (!c || !svg) return;
    const gs = indices.map(i => findShapeElement(svg, i)).filter(Boolean) as SVGGElement[];
    showMultiOverlay(c, gs);
  }, []);

  /** Size the canvas to fit the stage (contain, preserving slide aspect), centered. */
  const fitCanvas = useCallback(() => {
    const c = containerRef.current; if (!c) return;
    const svg = c.querySelector('svg'); const stage = c.parentElement;
    if (!svg || !stage) return;
    const vb = svg.getAttribute('viewBox')?.split(/\s+/);
    if (!vb || vb.length < 4) return;
    const aspect = parseFloat(vb[2]) / parseFloat(vb[3]);
    if (!isFinite(aspect) || aspect <= 0) return;
    const PAD = 24;
    const availW = stage.clientWidth - PAD * 2, availH = stage.clientHeight - PAD * 2;
    let w = availW, h = w / aspect;
    if (h > availH) { h = availH; w = h * aspect; }
    c.style.width = `${Math.max(0, Math.floor(w))}px`;
    c.style.height = `${Math.max(0, Math.floor(h))}px`;
  }, []);

  const renderCurrent = useCallback(() => {
    const c = containerRef.current; if (!c) return;
    insertSvgInto(c, renderer.renderSlide(slide));
    fitCanvas();
  }, [renderer, slide, fitCanvas]);

  /** Re-render the current slide + re-apply selection/multi overlays + sync history + bump thumb. */
  const refresh = useCallback((reselect?: number | null) => {
    renderCurrent();
    if (inlineRef.current !== null) { renderer.syncHistory(); bumpThumb(slide); return; }
    if (multiRef.current.length > 0) {
      drawMulti(multiRef.current);
      setSelectionState(null); selIdxRef.current = -1;
    } else {
      const idx = reselect === undefined ? selIdxRef.current : reselect;
      if (idx != null && idx >= 0) setSelection(drawSelection(idx));
      else setSelection(null);
    }
    renderer.syncHistory();
    bumpThumb(slide);
  }, [renderCurrent, renderer, drawMulti, drawSelection, setSelection, bumpThumb, slide]);

  const commit = useCallback((result: string, reselect?: number | null) => {
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    refresh(reselect);
  }, [renderer, refresh]);

  // ── Selection actions ──
  const selectIdx = useCallback((idx: number) => {
    setMulti([]); removeMultiOverlays(containerRef.current!);
    setSelection(drawSelection(idx));
  }, [drawSelection, setMulti, setSelection]);

  const deselect = useCallback(() => {
    setMulti([]); setSelection(null);
    const c = containerRef.current; if (c) { c.querySelector('.selection-overlay')?.remove(); removeMultiOverlays(c); }
  }, [setMulti, setSelection]);

  const toggleMulti = useCallback((idx: number) => {
    const base = multiRef.current.length ? multiRef.current.slice() : (selIdxRef.current >= 0 ? [selIdxRef.current] : []);
    const pos = base.indexOf(idx);
    if (pos >= 0) base.splice(pos, 1); else base.push(idx);
    const c = containerRef.current;
    if (base.length <= 1) {
      setMulti([]); if (c) removeMultiOverlays(c);
      if (base.length === 1) selectIdx(base[0]); else deselect();
      return;
    }
    setMulti(base); setSelection(null);
    if (c) c.querySelector('.selection-overlay')?.remove();
    drawMulti(base);
  }, [setMulti, setSelection, drawMulti, selectIdx, deselect]);

  // ── Inline text ──
  const enterInline = useCallback((idx: number) => {
    setInline(idx);
    // Populate the panel without the solid selection overlay; the inline hook
    // draws a dashed edit outline instead.
    const c = containerRef.current, svg = svgIn();
    const g = svg ? findShapeElement(svg, idx) : null;
    setSelection(g ? extractShapeInfo(g) : null);
    if (c) c.querySelector('.selection-overlay')?.remove();
    renderer.setStatus('Inline text — click to place caret, type to edit, Esc to finish');
  }, [setInline, setSelection, renderer]);

  const exitInline = useCallback(() => {
    const idx = inlineRef.current;
    setInline(null);
    if (idx != null && idx >= 0) setSelection(drawSelection(idx));
    renderer.setStatus('Done editing text');
  }, [setInline, setSelection, drawSelection, renderer]);

  useInlineText({
    containerRef, renderer, slide, shapeIdx: inlineShape,
    commit: (result) => { if (!result.startsWith('ERROR:')) { renderCurrent(); renderer.syncHistory(); bumpThumb(slide); } },
    onExit: exitInline,
  });

  // ── Canvas click / dblclick (select / multi / inline) ──
  useEffect(() => {
    const c = containerRef.current; if (!c) return;
    const onClick = (e: MouseEvent) => {
      const target = e.target as Element;
      if (target.classList.contains('resize-handle') || target.classList.contains('rotate-handle')) return;
      const g = target.closest('g[data-ooxml-shape-idx]') as SVGGElement | null;
      const idx = g ? parseInt(g.getAttribute('data-ooxml-shape-idx') ?? '-1', 10) : -1;
      if (inlineRef.current !== null) {
        if (idx === inlineRef.current) return;  // inside the editing shape → caret (inline hook)
        exitInline();                            // clicked elsewhere → leave edit mode, then select
      }
      if (e.shiftKey && idx >= 0) { toggleMulti(idx); return; }
      if (idx >= 0) selectIdx(idx); else deselect();
    };
    const onDbl = (e: MouseEvent) => {
      const g = (e.target as Element).closest('g[data-ooxml-shape-idx]') as SVGGElement | null;
      if (!g || !g.querySelector('tspan[data-ooxml-para-idx]')) return;
      enterInline(parseInt(g.getAttribute('data-ooxml-shape-idx') ?? '-1', 10));
    };
    c.addEventListener('click', onClick);
    c.addEventListener('dblclick', onDbl);
    return () => { c.removeEventListener('click', onClick); c.removeEventListener('dblclick', onDbl); };
  }, [toggleMulti, selectIdx, deselect, enterInline, exitInline]);

  // ── Drag / resize / rotate / group move ──
  useDrag({
    containerRef,
    selectedShapeIdx: inlineShape !== null ? -1 : (selection?.idx ?? -1),
    multiSel: inlineShape !== null ? [] : multiSel,
    slide,
    onTransformUpdate: (i, x, y, cx, cy, rot) => renderer.updateTransform(slide, i, x, y, cx, cy, rot),
    onGroupTransform: (items) => renderer.updateShapesTransform(slide, items),
    onDragEnd: useCallback((affected: number[], result: string) => {
      if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
      if (affected.length === 0) return; // pure click
      refresh(affected.length === 1 ? affected[0] : undefined);
    }, [renderer, refresh]),
  });

  // ── Navigation / load ──
  const pendingRef = useRef<number | null>(null);
  useEffect(() => {
    if (pendingRef.current !== null && containerRef.current) {
      insertSvgInto(containerRef.current, renderer.renderSlide(pendingRef.current));
      fitCanvas();
      pendingRef.current = null;
    }
  });

  const goTo = useCallback((idx: number) => {
    setInline(null); deselect();
    renderer.setSlide(idx);
    if (containerRef.current) { insertSvgInto(containerRef.current, renderer.renderSlide(idx)); fitCanvas(); }
    else pendingRef.current = idx;
  }, [renderer, deselect, setInline, fitCanvas]);

  const loadFile = useCallback(async (file: File) => {
    deselect(); setInline(null);
    await renderer.loadFile(file);
    pendingRef.current = 0;
  }, [renderer, deselect, setInline]);

  // Re-render every thumbnail whenever the slide count changes (load/add/delete).
  useEffect(() => { reseedThumbs(renderer.total); }, [renderer.total, reseedThumbs]);

  // Keep the canvas fitted + overlays positioned on window resize.
  useEffect(() => {
    const onResize = () => {
      fitCanvas();
      if (inlineRef.current !== null) return;
      if (multiRef.current.length > 0) drawMulti(multiRef.current);
      else if (selIdxRef.current >= 0) drawSelection(selIdxRef.current);
    };
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [fitCanvas, drawMulti, drawSelection]);

  // ── Insert ──
  const addShape = useCallback((geom: string) => {
    const isLine = geom === 'line';
    const cx = DEFAULT_SHAPE_SIZE, cy = isLine ? 0 : DEFAULT_SHAPE_SIZE;
    const x = Math.round((DEFAULT_SLIDE_CX - cx) / 2), y = Math.round((DEFAULT_SLIDE_CY - cy) / 2);
    const res = isLine
      ? renderer.addShape(slide, geom, x, y, cx, cy, -1, -1, -1)
      : renderer.addShape(slide, geom, x, y, cx, cy, ...ACCENT_RGB);
    if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
    const idx = okIndex(res);
    if (isLine) renderer.updateStroke(slide, idx, ...INK_RGB, DEFAULT_STROKE_WIDTH);
    commit(res, idx);
  }, [renderer, slide, commit]);

  const addTextBox = useCallback(() => {
    const cx = DEFAULT_SHAPE_SIZE, cy = Math.round(DEFAULT_SHAPE_SIZE / 3);
    const x = Math.round((DEFAULT_SLIDE_CX - cx) / 2), y = Math.round((DEFAULT_SLIDE_CY - cy) / 2);
    const res = renderer.addShape(slide, 'rect', x, y, cx, cy, -1, -1, -1);
    if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
    const idx = okIndex(res);
    renderer.addShapeText(slide, idx, 'Text', DEFAULT_FONT_SIZE, ...INK_RGB);
    commit('OK', idx);
  }, [renderer, slide, commit]);

  const addImage = useCallback(async (file: File) => {
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const cx = DEFAULT_IMAGE_CX, cy = DEFAULT_IMAGE_CY;
      const x = Math.round((DEFAULT_SLIDE_CX - cx) / 2), y = Math.round((DEFAULT_SLIDE_CY - cy) / 2);
      const res = await renderer.addImage(slide, data, file.type || 'image/png', x, y, cx, cy);
      if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
      commit(res, okIndex(res));
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, slide, commit]);

  const replaceImage = useCallback(async (file: File) => {
    if (!selection) return;
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      commit(await renderer.replaceImage(slide, selection.idx, data, file.type || 'image/png'), selection.idx);
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, slide, selection, commit]);

  // ── Shape ops ──
  const duplicate = useCallback(() => {
    if (!selection) return;
    const res = renderer.duplicateShape(slide, selection.idx);
    if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
    commit(res, okIndex(res));
  }, [renderer, slide, selection, commit]);

  const remove = useCallback(() => {
    if (selection) {
      const isPic = selection.shapeType === 'picture';
      const res = isPic ? renderer.deleteImage(slide, selection.idx) : renderer.deleteShape(slide, selection.idx);
      if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
      setSelection(null); commit('OK', null);
    } else if (multiRef.current.length) {
      // delete highest index first to keep indices stable
      const sorted = multiRef.current.slice().sort((a, b) => b - a);
      renderer.beginBatch();
      for (const i of sorted) renderer.deleteShape(slide, i);
      renderer.endBatch();
      setMulti([]); commit('OK', null);
    }
  }, [renderer, slide, selection, commit, setSelection, setMulti]);

  const zOrder = useCallback((how: 'front' | 'forward' | 'backward' | 'back') => {
    if (!selection) return;
    const fn = { front: renderer.bringToFront, forward: renderer.bringForward, backward: renderer.sendBackward, back: renderer.sendToBack }[how];
    const res = fn(slide, selection.idx);
    if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
    commit(res, okIndex(res, selection.idx));
  }, [renderer, slide, selection, commit]);

  const align = useCallback((how: 'l' | 'c' | 'r' | 't' | 'm' | 'b') => {
    if (!selection) return;
    const t = selection.t;
    let { x, y } = t;
    if (how === 'l') x = 0;
    else if (how === 'c') x = Math.round((DEFAULT_SLIDE_CX - t.cx) / 2);
    else if (how === 'r') x = DEFAULT_SLIDE_CX - t.cx;
    else if (how === 't') y = 0;
    else if (how === 'm') y = Math.round((DEFAULT_SLIDE_CY - t.cy) / 2);
    else if (how === 'b') y = DEFAULT_SLIDE_CY - t.cy;
    commit(renderer.updateTransform(slide, selection.idx, x, y, t.cx, t.cy, t.rot), selection.idx);
  }, [renderer, slide, selection, commit]);

  // ── Copy / paste ──
  const copy = useCallback(() => {
    if (!selection) return;
    const spec = renderer.getShapeSpec(slide, selection.idx);
    if (spec.startsWith('ERROR')) { renderer.setStatus(spec); return; }
    clipboardRef.current = spec; setHasClipboard(true);
    renderer.setStatus('Copied shape');
  }, [renderer, slide, selection]);

  const paste = useCallback(() => {
    if (!clipboardRef.current) return;
    const res = renderer.insertShapeSpec(slide, clipboardRef.current, PASTE_OFFSET, PASTE_OFFSET);
    if (res.startsWith('ERROR:')) { renderer.setStatus(res); return; }
    renderer.setStatus('Pasted shape');
    commit(res, okIndex(res));
  }, [renderer, slide, commit]);

  // ── Undo / redo ──
  const applyHistory = useCallback((res: string) => {
    if (res.startsWith('ERROR')) { renderer.setStatus(res); renderer.syncHistory(); return; }
    let info: { slides: number[]; slideCount: number } | null = null;
    try { info = JSON.parse(res); } catch { /* */ }
    setInline(null); deselect();
    const prevTotal = renderer.total;
    const count = info?.slideCount ?? prevTotal;
    renderer.setTotal(count);
    let cur = slide;
    if (cur >= count) cur = count - 1;
    if (cur < 0) cur = 0;
    renderer.setSlide(cur);
    if (containerRef.current) { insertSvgInto(containerRef.current, renderer.renderSlide(cur)); fitCanvas(); }
    if (count !== prevTotal) reseedThumbs(count);           // structure changed → all shift
    else { (info?.slides ?? []).forEach(i => bumpThumb(i)); bumpThumb(cur); }
    renderer.syncHistory();
  }, [renderer, slide, deselect, setInline, bumpThumb, reseedThumbs, fitCanvas]);

  const undo = useCallback(() => applyHistory(renderer.undo()), [renderer, applyHistory]);
  const redo = useCallback(() => applyHistory(renderer.redo()), [renderer, applyHistory]);

  // ── Nudge (single or group) ──
  const nudge = useCallback((dx: number, dy: number) => {
    if (multiRef.current.length > 0) {
      const svg = svgIn(); if (!svg) return;
      const items = multiRef.current.map(i => {
        const g = findShapeElement(svg, i); if (!g) return null;
        const m = extractShapeInfo(g);
        return { shapeIdx: i, x: m.t.x + dx, y: m.t.y + dy, cx: m.t.cx, cy: m.t.cy, rot: m.t.rot };
      }).filter(Boolean) as Array<{ shapeIdx: number; x: number; y: number; cx: number; cy: number; rot: number }>;
      commit(renderer.updateShapesTransform(slide, items));
    } else if (selIdxRef.current >= 0 && selection) {
      const t = selection.t;
      commit(renderer.updateTransform(slide, selection.idx, t.x + dx, t.y + dy, t.cx, t.cy, t.rot), selection.idx);
    }
  }, [renderer, slide, selection, commit]);

  useKeyboard({
    enabled: inlineShape === null,
    onUndo: undo, onRedo: redo,
    onCopy: copy, onPaste: paste,
    onDuplicate: duplicate, onDelete: remove,
    onNudge: nudge, onEscape: deselect,
  });

  // ── Slide management ──
  const addSlide = useCallback(async () => {
    try { const { insertedIdx } = await renderer.addSlide(slide); goTo(insertedIdx); }
    catch (e) { renderer.setStatus(`Error: ${(e as Error).message}`); }
  }, [renderer, slide, goTo]);
  const dupSlide = useCallback(async () => {
    try { const { insertedIdx } = await renderer.addSlide(slide, slide); goTo(insertedIdx); }
    catch (e) { renderer.setStatus(`Error: ${(e as Error).message}`); }
  }, [renderer, slide, goTo]);
  const delSlide = useCallback(async () => {
    if (renderer.total <= 1) return;
    try { await renderer.deleteSlide(slide); goTo(Math.min(slide, renderer.total - 2)); }
    catch (e) { renderer.setStatus(`Error: ${(e as Error).message}`); }
  }, [renderer, slide, goTo]);
  const reorder = useCallback(async (from: number, to: number) => {
    const order = Array.from({ length: renderer.total }, (_, i) => i);
    order.splice(to, 0, order.splice(from, 1)[0]);
    try { const { slideCount } = await renderer.reorderSlides(order); reseedThumbs(slideCount); goTo(to); }
    catch (e) { renderer.setStatus(`Error: ${(e as Error).message}`); }
  }, [renderer, goTo, reseedThumbs]);

  const loaded = renderer.total > 0;

  return (
    <div className="app">
      <TopBar status={renderer.status} canUndo={renderer.history.canUndo} canRedo={renderer.history.canRedo}
        onUndo={undo} onRedo={redo} onExport={renderer.exportPptx} onOpen={loadFile} />

      {!loaded ? (
        <div className="empty"><DropZone onFile={loadFile} /></div>
      ) : (
        <>
          <InsertToolbar onAddShape={addShape} onAddTextBox={addTextBox} onAddImage={addImage}
            hasSelection={!!selection} onAlign={align} onZ={zOrder} />
          <div className="body">
            <SlideRail total={renderer.total} current={slide} versions={thumbVer}
              renderSlide={renderer.renderSlide} onSelect={goTo}
              onAdd={addSlide} onDuplicate={dupSlide} onDelete={delSlide} onReorder={reorder} />

            <main className="stage">
              <div ref={containerRef} className="canvas" />
              {inlineShape !== null && <div className="inline-hint">Inline editing — type, Enter for new line, Esc to finish</div>}
            </main>

            <aside className="panel">
              {selection || multiSel.length > 0 ? (
                <PropertiesPanel renderer={renderer} slide={slide}
                  selection={selection ?? { idx: -1, label: '', shapeType: '', geom: '', fillHex: '', t: { x: 0, y: 0, cx: 0, cy: 0, rot: 0 }, paragraphs: [] }}
                  multiCount={multiSel.length} commit={commit}
                  onDuplicate={duplicate} onDelete={remove} onReplaceImage={replaceImage} />
              ) : (
                <div className="panel-empty">Select a shape to edit.<br />Double-click text to type. {hasClipboard ? '⌘V to paste.' : ''}</div>
              )}
            </aside>
          </div>
        </>
      )}
    </div>
  );
}
