/**
 * useRenderer — manages PptxRenderer lifecycle and all editing APIs (pptx-svg 0.6.0).
 *
 * Thin wrappers over every editing export so components deal only in
 * slide/shape indices, EMU values, and result strings. Also tracks undo/redo
 * availability so the UI can reflect it.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';
import type { ShapeSpec, TextLayout, TextHit, HistoryResult } from 'pptx-svg';

export type { ShapeSpec, TextLayout, TextHit, HistoryResult };

export interface HistoryState {
  canUndo: boolean;
  canRedo: boolean;
}

export function useRenderer() {
  const ref = useRef<PptxRenderer | null>(null);
  const [status, setStatus] = useState('Loading WebAssembly module…');
  const [slide, setSlide] = useState(0);
  const [total, setTotal] = useState(0);
  const [history, setHistory] = useState<HistoryState>({ canUndo: false, canRedo: false });

  const r = () => ref.current;

  // ── Init ──
  useEffect(() => {
    const renderer = new PptxRenderer({ logLevel: 'error', maxHistory: 100 });
    ref.current = renderer;
    renderer.init()
      .then(() => setStatus('Ready — open a .pptx file to start editing.'))
      .catch(err => setStatus(`Init failed: ${err.message}`));
  }, []);

  /** Refresh undo/redo availability into state (call after any mutation). */
  const syncHistory = useCallback(() => {
    const renderer = ref.current;
    if (!renderer) return;
    setHistory({ canUndo: renderer.canUndo(), canRedo: renderer.canRedo() });
  }, []);

  // ── Load ──
  const loadFile = useCallback(async (file: File) => {
    const renderer = ref.current;
    if (!renderer) return;
    setStatus(`Loading ${file.name}…`);
    try {
      const { slideCount } = await renderer.loadPptx(await file.arrayBuffer());
      setTotal(slideCount);
      setSlide(0);
      setHistory({ canUndo: false, canRedo: false });
      setStatus(`Loaded “${file.name}” — ${slideCount} slide(s)`);
    } catch (err) {
      setStatus(`Error: ${(err as Error).message}`);
    }
  }, []);

  const renderSlide = useCallback((idx: number): string =>
    r()?.renderSlideSvg(idx) ?? 'ERROR:renderer not initialized', []);

  const getSlideOoxml = useCallback((idx: number): string =>
    r()?.getSlideOoxml(idx) ?? 'ERROR:no renderer', []);

  // ── History (E6.1) ──
  const undo = useCallback((): string => r()?.undo() ?? 'ERROR:no renderer', []);
  const redo = useCallback((): string => r()?.redo() ?? 'ERROR:no renderer', []);
  const beginBatch = useCallback(() => r()?.beginBatch(), []);
  const endBatch = useCallback(() => r()?.endBatch(), []);

  // ── Shape transform ──
  const updateTransform = useCallback((s: number, i: number, x: number, y: number, cx: number, cy: number, rot: number): string =>
    r()?.updateShapeTransform(s, i, x, y, cx, cy, rot) ?? 'ERROR:no renderer', []);

  const updateShapesTransform = useCallback((s: number,
    items: Array<{ shapeIdx: number; x: number; y: number; cx: number; cy: number; rot: number }>): string =>
    r()?.updateShapesTransform(s, items) ?? 'ERROR:no renderer', []);

  // ── Fill / stroke ──
  const updateFill = useCallback((s: number, i: number, red: number, g: number, b: number): string =>
    r()?.updateShapeFill(s, i, red, g, b) ?? 'ERROR:no renderer', []);

  const updateGradientFill = useCallback((s: number, i: number, angle: number,
    stops: Array<{ pos: number; r: number; g: number; b: number }>): string =>
    r()?.updateShapeGradientFill(s, i, angle, stops) ?? 'ERROR:no renderer', []);

  const updateStroke = useCallback((s: number, i: number, red: number, g: number, b: number, w?: number, dash?: string): string =>
    r()?.updateShapeStroke(s, i, red, g, b, w, dash) ?? 'ERROR:no renderer', []);

  // ── Shape CRUD ──
  const deleteShape = useCallback((s: number, i: number): string => r()?.deleteShape(s, i) ?? 'ERROR:no renderer', []);
  const addShape = useCallback((s: number, geom: string, x: number, y: number, cx: number, cy: number, red?: number, g?: number, b?: number): string =>
    r()?.addShape(s, geom, x, y, cx, cy, red, g, b) ?? 'ERROR:no renderer', []);
  const duplicateShape = useCallback((s: number, i: number, dx?: number, dy?: number): string =>
    r()?.duplicateShape(s, i, dx, dy) ?? 'ERROR:no renderer', []);

  // ── Z-order (E6.3) ──
  const bringToFront = useCallback((s: number, i: number): string => r()?.bringToFront(s, i) ?? 'ERROR:no renderer', []);
  const sendToBack = useCallback((s: number, i: number): string => r()?.sendToBack(s, i) ?? 'ERROR:no renderer', []);
  const bringForward = useCallback((s: number, i: number): string => r()?.bringForward(s, i) ?? 'ERROR:no renderer', []);
  const sendBackward = useCallback((s: number, i: number): string => r()?.sendBackward(s, i) ?? 'ERROR:no renderer', []);

  // ── Copy / paste (E6.5) ──
  const getShapeSpec = useCallback((s: number, i: number): string => r()?.getShapeSpec(s, i) ?? 'ERROR:no renderer', []);
  const insertShapeSpec = useCallback((s: number, spec: string, dx?: number, dy?: number): string =>
    r()?.insertShapeSpec(s, spec, dx, dy) ?? 'ERROR:no renderer', []);

  // ── Inline text geometry (E6.2) ──
  const getTextLayout = useCallback((s: number, i: number): TextLayout | null => {
    const raw = r()?.getTextLayout(s, i);
    if (!raw || raw.startsWith('ERROR')) return null;
    try { return JSON.parse(raw) as TextLayout; } catch { return null; }
  }, []);
  const hitTestText = useCallback((s: number, i: number, x: number, y: number): TextHit | null => {
    const raw = r()?.hitTestText(s, i, x, y);
    if (!raw || raw.startsWith('ERROR')) return null;
    try { return JSON.parse(raw) as TextHit; } catch { return null; }
  }, []);
  const replaceTextRange = useCallback((s: number, i: number, sp: number, sc: number, ep: number, ec: number, text: string): string =>
    r()?.replaceTextRange(s, i, sp, sc, ep, ec, text) ?? 'ERROR:no renderer', []);

  // ── Text run / paragraph formatting ──
  const addShapeText = useCallback((s: number, i: number, text: string, size?: number, red?: number, g?: number, b?: number): string =>
    r()?.addShapeText(s, i, text, size, red, g, b) ?? 'ERROR:no renderer', []);
  const updateText = useCallback((s: number, i: number, p: number, ri: number, text: string): string =>
    r()?.updateShapeText(s, i, p, ri, text) ?? 'ERROR:no renderer', []);
  const updateTextRunStyle = useCallback((s: number, i: number, p: number, ri: number, bold: number, italic: number): string =>
    r()?.updateTextRunStyle(s, i, p, ri, bold, italic) ?? 'ERROR:no renderer', []);
  const updateTextRunFontSize = useCallback((s: number, i: number, p: number, ri: number, size: number): string =>
    r()?.updateTextRunFontSize(s, i, p, ri, size) ?? 'ERROR:no renderer', []);
  const updateTextRunColor = useCallback((s: number, i: number, p: number, ri: number, red: number, g: number, b: number): string =>
    r()?.updateTextRunColor(s, i, p, ri, red, g, b) ?? 'ERROR:no renderer', []);
  const updateTextRunFont = useCallback((s: number, i: number, p: number, ri: number, font: string, ea?: string, cs?: string): string =>
    r()?.updateTextRunFont(s, i, p, ri, font, ea, cs) ?? 'ERROR:no renderer', []);
  const updateTextRunDecoration = useCallback((s: number, i: number, p: number, ri: number, ul: string, st: string, bl: number): string =>
    r()?.updateTextRunDecoration(s, i, p, ri, ul, st, bl) ?? 'ERROR:no renderer', []);
  const updateParagraphAlign = useCallback((s: number, i: number, p: number, align: string): string =>
    r()?.updateParagraphAlign(s, i, p, align) ?? 'ERROR:no renderer', []);
  const addParagraph = useCallback((s: number, i: number, text: string, align?: string): string =>
    r()?.addParagraph(s, i, text, align) ?? 'ERROR:no renderer', []);
  const deleteParagraph = useCallback((s: number, i: number, p: number): string =>
    r()?.deleteParagraph(s, i, p) ?? 'ERROR:no renderer', []);
  const addRun = useCallback((s: number, i: number, p: number, text: string): string =>
    r()?.addRun(s, i, p, text) ?? 'ERROR:no renderer', []);
  const deleteRun = useCallback((s: number, i: number, p: number, ri: number): string =>
    r()?.deleteRun(s, i, p, ri) ?? 'ERROR:no renderer', []);

  // ── Table editing (E6.6) ──
  const updateTableCellText = useCallback((s: number, i: number, row: number, col: number, text: string): string =>
    r()?.updateTableCellText(s, i, row, col, text) ?? 'ERROR:no renderer', []);
  const addTableRow = useCallback((s: number, i: number, after?: number): string => r()?.addTableRow(s, i, after) ?? 'ERROR:no renderer', []);
  const deleteTableRow = useCallback((s: number, i: number, row: number): string => r()?.deleteTableRow(s, i, row) ?? 'ERROR:no renderer', []);
  const addTableColumn = useCallback((s: number, i: number, after?: number, w?: number): string => r()?.addTableColumn(s, i, after, w) ?? 'ERROR:no renderer', []);
  const deleteTableColumn = useCallback((s: number, i: number, col: number): string => r()?.deleteTableColumn(s, i, col) ?? 'ERROR:no renderer', []);

  // ── Image ──
  const addImage = useCallback(async (s: number, data: Uint8Array, mime: string, x: number, y: number, cx: number, cy: number): Promise<string> =>
    r()?.addImage(s, data, mime, x, y, cx, cy) ?? 'ERROR:no renderer', []);
  const replaceImage = useCallback(async (s: number, i: number, data: Uint8Array, mime: string): Promise<string> =>
    r()?.replaceImage(s, i, data, mime) ?? 'ERROR:no renderer', []);
  const deleteImage = useCallback((s: number, i: number): string => r()?.deleteImage(s, i) ?? 'ERROR:no renderer', []);

  // ── Slide management ──
  const addSlide = useCallback(async (after?: number, source?: number) => {
    const renderer = ref.current; if (!renderer) throw new Error('no renderer');
    const res = await renderer.addSlide(after, source);
    setTotal(res.slideCount); syncHistory();
    return res;
  }, [syncHistory]);
  const deleteSlide = useCallback(async (idx: number) => {
    const renderer = ref.current; if (!renderer) throw new Error('no renderer');
    const res = await renderer.deleteSlide(idx);
    setTotal(res.slideCount); syncHistory();
    return res;
  }, [syncHistory]);
  const reorderSlides = useCallback(async (order: number[]) => {
    const renderer = ref.current; if (!renderer) throw new Error('no renderer');
    const res = await renderer.reorderSlides(order); syncHistory();
    return res;
  }, [syncHistory]);

  // ── Export ──
  const exportPptx = useCallback(async () => {
    const renderer = ref.current; if (!renderer) throw new Error('no renderer');
    setStatus('Exporting…');
    try {
      const buffer = await renderer.exportPptx();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'edited.pptx';
      a.click();
      URL.revokeObjectURL(a.href);
      setStatus('Exported edited.pptx');
    } catch (err) {
      setStatus(`Export error: ${(err as Error).message}`);
    }
  }, []);

  return {
    status, setStatus, slide, setSlide, total, setTotal, history, syncHistory,
    loadFile, renderSlide, getSlideOoxml,
    undo, redo, beginBatch, endBatch,
    updateTransform, updateShapesTransform,
    updateFill, updateGradientFill, updateStroke,
    deleteShape, addShape, duplicateShape,
    bringToFront, sendToBack, bringForward, sendBackward,
    getShapeSpec, insertShapeSpec,
    getTextLayout, hitTestText, replaceTextRange,
    addShapeText, updateText, updateTextRunStyle, updateTextRunFontSize, updateTextRunColor,
    updateTextRunFont, updateTextRunDecoration, updateParagraphAlign,
    addParagraph, deleteParagraph, addRun, deleteRun,
    updateTableCellText, addTableRow, deleteTableRow, addTableColumn, deleteTableColumn,
    addImage, replaceImage, deleteImage,
    addSlide, deleteSlide, reorderSlides,
    exportPptx,
  };
}

export type RendererApi = ReturnType<typeof useRenderer>;
