/**
 * useRenderer — manages PptxRenderer lifecycle and all editing APIs.
 *
 * Encapsulates all Wasm interaction so components only deal with
 * slide index, SVG strings, and high-level edit operations.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';

export function useRenderer() {
  const rendererRef = useRef<PptxRenderer | null>(null);
  const [status, setStatus] = useState('Loading WebAssembly module...');
  const [slide, setSlide] = useState(0);
  const [total, setTotal] = useState(0);

  // ── Init ──
  useEffect(() => {
    const renderer = new PptxRenderer();
    rendererRef.current = renderer;
    renderer.init()
      .then(() => setStatus('Ready. Drop a .pptx file.'))
      .catch(err => setStatus(`Init failed: ${err.message}`));
  }, []);

  // ── Load PPTX ──
  const loadFile = useCallback(async (file: File) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    setStatus(`Loading ${file.name}...`);
    try {
      const { slideCount } = await renderer.loadPptx(await file.arrayBuffer());
      setTotal(slideCount);
      setSlide(0);
      setStatus(`Loaded "${file.name}" - ${slideCount} slide(s)`);
    } catch (err) {
      setStatus(`Error: ${(err as Error).message}`);
    }
  }, []);

  // ── Render SVG ──
  const renderSlide = useCallback((idx: number): string => {
    return rendererRef.current?.renderSlideSvg(idx) ?? 'ERROR: renderer not initialized';
  }, []);

  // ── Shape transform ──
  const updateTransform = useCallback((slideIdx: number, shapeIdx: number,
    x: number, y: number, cx: number, cy: number, rot: number): string => {
    return rendererRef.current?.updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot) ?? 'ERROR: no renderer';
  }, []);

  // ── Text content ──
  const updateText = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, text: string): string => {
    return rendererRef.current?.updateShapeText(slideIdx, shapeIdx, pi, ri, text) ?? 'ERROR: no renderer';
  }, []);

  // ── Fill ──
  const updateFill = useCallback((slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number): string => {
    return rendererRef.current?.updateShapeFill(slideIdx, shapeIdx, r, g, b) ?? 'ERROR: no renderer';
  }, []);

  // ── Stroke ──
  const updateStroke = useCallback((slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number, widthEmu?: number, dash?: string): string => {
    return rendererRef.current?.updateShapeStroke(slideIdx, shapeIdx, r, g, b, widthEmu, dash) ?? 'ERROR: no renderer';
  }, []);

  // ── Delete shape ──
  const deleteShape = useCallback((slideIdx: number, shapeIdx: number): string => {
    return rendererRef.current?.deleteShape(slideIdx, shapeIdx) ?? 'ERROR: no renderer';
  }, []);

  // ── Add shape ──
  const addShape = useCallback((slideIdx: number, geomType: string,
    x: number, y: number, cx: number, cy: number,
    fillR?: number, fillG?: number, fillB?: number): string => {
    return rendererRef.current?.addShape(slideIdx, geomType, x, y, cx, cy, fillR, fillG, fillB) ?? 'ERROR: no renderer';
  }, []);

  // ── Add text to shape ──
  const addShapeText = useCallback((slideIdx: number, shapeIdx: number, text: string,
    fontSize?: number, colorR?: number, colorG?: number, colorB?: number): string => {
    return rendererRef.current?.addShapeText(slideIdx, shapeIdx, text, fontSize, colorR, colorG, colorB) ?? 'ERROR: no renderer';
  }, []);

  // ── Duplicate shape ──
  const duplicateShape = useCallback((slideIdx: number, shapeIdx: number,
    dxEmu?: number, dyEmu?: number): string => {
    return rendererRef.current?.duplicateShape(slideIdx, shapeIdx, dxEmu, dyEmu) ?? 'ERROR: no renderer';
  }, []);

  // ── Text run formatting ──
  const updateTextRunStyle = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, bold: number, italic: number): string => {
    return rendererRef.current?.updateTextRunStyle(slideIdx, shapeIdx, pi, ri, bold, italic) ?? 'ERROR: no renderer';
  }, []);

  const updateTextRunFontSize = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, fontSize: number): string => {
    return rendererRef.current?.updateTextRunFontSize(slideIdx, shapeIdx, pi, ri, fontSize) ?? 'ERROR: no renderer';
  }, []);

  const updateTextRunColor = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, r: number, g: number, b: number): string => {
    return rendererRef.current?.updateTextRunColor(slideIdx, shapeIdx, pi, ri, r, g, b) ?? 'ERROR: no renderer';
  }, []);

  const updateTextRunFont = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, fontFace: string, eaFont?: string, csFont?: string): string => {
    return rendererRef.current?.updateTextRunFont(slideIdx, shapeIdx, pi, ri, fontFace, eaFont, csFont) ?? 'ERROR: no renderer';
  }, []);

  const updateTextRunDecoration = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, underline: string, strike: string, baseline: number): string => {
    return rendererRef.current?.updateTextRunDecoration(slideIdx, shapeIdx, pi, ri, underline, strike, baseline) ?? 'ERROR: no renderer';
  }, []);

  const updateParagraphAlign = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, align: string): string => {
    return rendererRef.current?.updateParagraphAlign(slideIdx, shapeIdx, pi, align) ?? 'ERROR: no renderer';
  }, []);

  // ── Paragraph/run CRUD ──
  const addParagraph = useCallback((slideIdx: number, shapeIdx: number,
    text: string, align?: string): string => {
    return rendererRef.current?.addParagraph(slideIdx, shapeIdx, text, align) ?? 'ERROR: no renderer';
  }, []);

  const deleteParagraph = useCallback((slideIdx: number, shapeIdx: number, pi: number): string => {
    return rendererRef.current?.deleteParagraph(slideIdx, shapeIdx, pi) ?? 'ERROR: no renderer';
  }, []);

  const addRun = useCallback((slideIdx: number, shapeIdx: number, pi: number, text: string): string => {
    return rendererRef.current?.addRun(slideIdx, shapeIdx, pi, text) ?? 'ERROR: no renderer';
  }, []);

  const deleteRun = useCallback((slideIdx: number, shapeIdx: number, pi: number, ri: number): string => {
    return rendererRef.current?.deleteRun(slideIdx, shapeIdx, pi, ri) ?? 'ERROR: no renderer';
  }, []);

  // ── Image operations ──
  const addImage = useCallback(async (slideIdx: number, data: Uint8Array, mime: string,
    x: number, y: number, cx: number, cy: number): Promise<string> => {
    const r = rendererRef.current;
    if (!r) return 'ERROR: no renderer';
    return r.addImage(slideIdx, data, mime, x, y, cx, cy);
  }, []);

  const replaceImage = useCallback(async (slideIdx: number, shapeIdx: number,
    data: Uint8Array, mime: string): Promise<string> => {
    const r = rendererRef.current;
    if (!r) return 'ERROR: no renderer';
    return r.replaceImage(slideIdx, shapeIdx, data, mime);
  }, []);

  const deleteImage = useCallback((slideIdx: number, shapeIdx: number): string => {
    return rendererRef.current?.deleteImage(slideIdx, shapeIdx) ?? 'ERROR: no renderer';
  }, []);

  // ── Slide management ──
  const addSlide = useCallback(async (afterIdx?: number, sourceSlideIdx?: number) => {
    const r = rendererRef.current;
    if (!r) throw new Error('no renderer');
    const result = await r.addSlide(afterIdx, sourceSlideIdx);
    setTotal(result.slideCount);
    return result;
  }, []);

  const deleteSlide = useCallback(async (slideIdx: number) => {
    const r = rendererRef.current;
    if (!r) throw new Error('no renderer');
    const result = await r.deleteSlide(slideIdx);
    setTotal(result.slideCount);
    return result;
  }, []);

  const reorderSlides = useCallback(async (newOrder: number[]) => {
    const r = rendererRef.current;
    if (!r) throw new Error('no renderer');
    const result = await r.reorderSlides(newOrder);
    return result;
  }, []);

  // ── Export ──
  const exportPptx = useCallback(async () => {
    const renderer = rendererRef.current;
    if (!renderer) throw new Error('renderer not initialized');
    setStatus('Exporting...');
    try {
      const buffer = await renderer.exportPptx();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'exported.pptx';
      a.click();
      URL.revokeObjectURL(a.href);
      setStatus('Exported!');
    } catch (err) {
      setStatus(`Export error: ${(err as Error).message}`);
    }
  }, []);

  return {
    status, setStatus,
    slide, setSlide, total,
    loadFile, renderSlide,
    // Shape editing
    updateTransform, updateText, updateFill, updateStroke,
    deleteShape, addShape, addShapeText, duplicateShape,
    // Text formatting
    updateTextRunStyle, updateTextRunFontSize, updateTextRunColor,
    updateTextRunFont, updateTextRunDecoration, updateParagraphAlign,
    addParagraph, deleteParagraph, addRun, deleteRun,
    // Image
    addImage, replaceImage, deleteImage,
    // Slide management
    addSlide, deleteSlide, reorderSlides,
    // Export
    exportPptx,
  };
}
