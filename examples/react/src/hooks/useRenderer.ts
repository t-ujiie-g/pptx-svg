/**
 * useRenderer — manages PptxRenderer lifecycle (init, load, render, export).
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
    const renderer = rendererRef.current;
    if (!renderer) return 'ERROR: renderer not initialized';
    return renderer.renderSlideSvg(idx);
  }, []);

  // ── Shape-level edits ──
  const updateTransform = useCallback((slideIdx: number, shapeIdx: number,
    x: number, y: number, cx: number, cy: number, rot: number): string => {
    return rendererRef.current?.updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot) ?? 'ERROR: no renderer';
  }, []);

  const updateText = useCallback((slideIdx: number, shapeIdx: number,
    pi: number, ri: number, text: string): string => {
    return rendererRef.current?.updateShapeText(slideIdx, shapeIdx, pi, ri, text) ?? 'ERROR: no renderer';
  }, []);

  const updateFill = useCallback((slideIdx: number, shapeIdx: number,
    r: number, g: number, b: number): string => {
    return rendererRef.current?.updateShapeFill(slideIdx, shapeIdx, r, g, b) ?? 'ERROR: no renderer';
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
    updateTransform, updateText, updateFill,
    exportPptx,
  };
}
