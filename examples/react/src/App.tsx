/**
 * pptx-svg React Example
 *
 * Demonstrates the v0.4.0 shape-level editing APIs:
 * - Click to select, drag to move, resize with handles
 * - Edit text per run, change fill color
 * - Export modified PPTX
 *
 * Architecture:
 *   App (state + composition)
 *   ├── DropZone         — file drop/browse UI
 *   ├── FillToolbar      — fill color picker (shown when shape selected)
 *   ├── TextRunsPanel    — per-run text editors (shown when shape has text)
 *   ├── SlideViewer      — SVG display + click-to-select
 *   └── hooks
 *       ├── useRenderer  — PptxRenderer lifecycle
 *       └── useDrag      — move/resize interaction
 */

import { useCallback, useRef, useState } from 'react';
import { useRenderer } from './hooks/useRenderer';
import { useDrag } from './hooks/useDrag';
import { DropZone } from './components/DropZone';
import { FillToolbar, TextRunsPanel } from './components/EditToolbar';
import { SlideViewer, insertSvg, reselectShape } from './components/SlideViewer';
import type { ShapeInfo } from './utils/svg';

export default function App() {
  const renderer = useRenderer();
  const containerRef = useRef<HTMLDivElement>(null);

  // ── Selection state ──
  const [selection, setSelection] = useState<ShapeInfo | null>(null);
  const selectedIdx = selection?.idx ?? -1;

  // ── Render current slide & re-select shape after edits ──
  const renderAndRefresh = useCallback((shapeIdx: number) => {
    insertSvg(containerRef.current, renderer.renderSlide(renderer.slide));
    if (shapeIdx >= 0) {
      setSelection(reselectShape(containerRef.current, shapeIdx));
    }
  }, [renderer]);

  // ── Drag/resize ──
  useDrag({
    containerRef,
    selectedShapeIdx: selectedIdx,
    slide: renderer.slide,
    onTransformUpdate: (shapeIdx, x, y, cx, cy, rot) =>
      renderer.updateTransform(renderer.slide, shapeIdx, x, y, cx, cy, rot),
    onDragEnd: useCallback((shapeIdx: number, result: string) => {
      if (result.startsWith('ERROR:')) {
        renderer.setStatus(result);
      } else {
        renderer.setStatus(`Updated #${shapeIdx}`);
        renderAndRefresh(shapeIdx);
      }
    }, [renderer, renderAndRefresh]),
  });

  // ── Handlers ──

  const handleSelect = useCallback((info: ShapeInfo | null) => {
    setSelection(info);
  }, []);

  const handleGoToSlide = useCallback((idx: number) => {
    renderer.setSlide(idx);
    setSelection(null);
    insertSvg(containerRef.current, renderer.renderSlide(idx));
  }, [renderer]);

  const handleLoadFile = useCallback(async (file: File) => {
    setSelection(null);
    await renderer.loadFile(file);
    insertSvg(containerRef.current, renderer.renderSlide(0));
  }, [renderer]);

  const handleApplyFill = useCallback((hex: string) => {
    if (selectedIdx < 0) return;
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    const result = renderer.updateFill(renderer.slide, selectedIdx, r, g, b);
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    renderer.setStatus(`Fill → ${hex}`);
    renderAndRefresh(selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleApplyText = useCallback((pi: number, ri: number, text: string) => {
    if (selectedIdx < 0) return;
    const result = renderer.updateText(renderer.slide, selectedIdx, pi, ri, text);
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    renderer.setStatus(`Text updated: P${pi}R${ri}`);
    renderAndRefresh(selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  // ── Render ──

  return (
    <div style={{ maxWidth: 1024, margin: '0 auto', padding: 24, fontFamily: 'system-ui, sans-serif' }}>
      <h1 style={{ marginBottom: 4 }}>pptx-svg React Example</h1>
      <p style={{ color: '#666', marginBottom: 20 }}>
        Click to select, drag to move, resize with handles. Edit text and fill from toolbar. Export saves changes.
      </p>

      <DropZone onFile={handleLoadFile} />

      {/* Navigation + Export */}
      {renderer.total > 0 && (
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 16, flexWrap: 'wrap' }}>
          <button onClick={() => handleGoToSlide(renderer.slide - 1)} disabled={renderer.slide === 0}>Prev</button>
          <span style={{ minWidth: 80, textAlign: 'center' }}>{renderer.slide + 1} / {renderer.total}</span>
          <button onClick={() => handleGoToSlide(renderer.slide + 1)} disabled={renderer.slide >= renderer.total - 1}>Next</button>
          <button onClick={renderer.exportPptx} style={{ marginLeft: 'auto' }}>Export PPTX</button>
        </div>
      )}

      {/* Editing toolbar */}
      {selection && (
        <>
          <FillToolbar
            fillColor={selection.fillHex ? '#' + selection.fillHex : '#4a90d9'}
            shapeLabel={selection.label}
            onApplyFill={handleApplyFill}
          />
          <TextRunsPanel runs={selection.textRuns} onApplyText={handleApplyText} />
        </>
      )}

      {/* SVG viewer */}
      <SlideViewer containerRef={containerRef} hasSelection={selectedIdx >= 0} onSelect={handleSelect} />

      {/* Shape info */}
      {selection && (
        <p style={{ marginTop: 8, fontFamily: 'monospace', fontSize: 13, color: '#888' }}>{selection.detail}</p>
      )}

      <p style={{ marginTop: 12, color: '#666', fontSize: 14 }}>{renderer.status}</p>

      <style>{`
        .selection-overlay { position: absolute; pointer-events: none; border: 2px solid #4a90d9; z-index: 100; }
        .resize-handle {
          position: absolute; width: 10px; height: 10px;
          background: #4a90d9; border: 1px solid #fff; border-radius: 2px;
          pointer-events: all; z-index: 101;
        }
        .resize-handle.nw { top: -5px; left: -5px; cursor: nw-resize; }
        .resize-handle.ne { top: -5px; right: -5px; cursor: ne-resize; }
        .resize-handle.sw { bottom: -5px; left: -5px; cursor: sw-resize; }
        .resize-handle.se { bottom: -5px; right: -5px; cursor: se-resize; }
        .resize-handle.n  { top: -5px; left: calc(50% - 5px); cursor: n-resize; }
        .resize-handle.s  { bottom: -5px; left: calc(50% - 5px); cursor: s-resize; }
        .resize-handle.w  { top: calc(50% - 5px); left: -5px; cursor: w-resize; }
        .resize-handle.e  { top: calc(50% - 5px); right: -5px; cursor: e-resize; }
        div svg { display: block; width: 100%; height: auto; }
        button {
          padding: 8px 16px; border: 1px solid #ccc; border-radius: 4px;
          cursor: pointer; background: white; font-size: 14px;
        }
        button:hover:not(:disabled) { background: #f5f5f5; }
        button:disabled { opacity: 0.5; cursor: not-allowed; }
      `}</style>
    </div>
  );
}
