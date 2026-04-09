/**
 * pptx-svg React Example — Full Editing Demo (v0.5.2)
 *
 * Layout: header bar, thumbnails, then 2-column (SVG left + sidebar right).
 * Sidebar is always visible (empty placeholder when nothing selected)
 * so SVG width never changes on selection.
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import { useRenderer } from './hooks/useRenderer';
import { useDrag } from './hooks/useDrag';
import { DropZone } from './components/DropZone';
import { ShapeToolbar, TextPanel } from './components/EditToolbar';
import { SlideViewer, insertSvg, reselectShape } from './components/SlideViewer';
import { SlideThumbnails } from './components/SlideThumbnails';
import { hexToRgb, type ShapeInfo } from './utils/svg';
import {
  DEFAULT_SLIDE_CX, DEFAULT_SLIDE_CY, DEFAULT_SHAPE_SIZE,
  DEFAULT_IMAGE_CX, DEFAULT_IMAGE_CY, DEFAULT_STROKE_WIDTH,
  DEFAULT_FONT_SIZE, SIDEBAR_WIDTH,
} from './constants';

export default function App() {
  const renderer = useRenderer();
  const containerRef = useRef<HTMLDivElement>(null);

  const [selection, setSelection] = useState<ShapeInfo | null>(null);
  const selectedIdx = selection?.idx ?? -1;

  const [thumbKey, setThumbKey] = useState(0);
  const refreshThumbs = () => setThumbKey(k => k + 1);

  // Fix: render first slide after React mounts the SlideViewer
  const pendingSlideRef = useRef<number | null>(null);
  useEffect(() => {
    if (pendingSlideRef.current !== null && containerRef.current) {
      insertSvg(containerRef.current, renderer.renderSlide(pendingSlideRef.current));
      pendingSlideRef.current = null;
    }
  });

  const renderAndRefresh = useCallback((shapeIdx?: number) => {
    insertSvg(containerRef.current, renderer.renderSlide(renderer.slide));
    if (shapeIdx !== undefined && shapeIdx >= 0) {
      setSelection(reselectShape(containerRef.current, shapeIdx));
    }
  }, [renderer]);

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

  // ── Navigation ──
  const handleGoToSlide = useCallback((idx: number) => {
    renderer.setSlide(idx);
    setSelection(null);
    if (containerRef.current) {
      insertSvg(containerRef.current, renderer.renderSlide(idx));
    } else {
      pendingSlideRef.current = idx;
    }
  }, [renderer]);

  const handleLoadFile = useCallback(async (file: File) => {
    setSelection(null);
    await renderer.loadFile(file);
    pendingSlideRef.current = 0;
    refreshThumbs();
  }, [renderer]);

  const callAndRefresh = (fn: () => string, msg: string, shapeIdx?: number) => {
    const result = fn();
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return result; }
    renderer.setStatus(msg);
    renderAndRefresh(shapeIdx);
    return result;
  };

  // ── Shape editing handlers ──
  const handleApplyFill = useCallback((hex: string) => {
    if (selectedIdx < 0) return;
    const [r, g, b] = hexToRgb(hex);
    callAndRefresh(() => renderer.updateFill(renderer.slide, selectedIdx, r, g, b), `Fill -> ${hex}`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleApplyStroke = useCallback((hex: string, dash: string) => {
    if (selectedIdx < 0) return;
    const [r, g, b] = hexToRgb(hex);
    callAndRefresh(() => renderer.updateStroke(renderer.slide, selectedIdx, r, g, b, DEFAULT_STROKE_WIDTH, dash),
      `Stroke -> ${hex} ${dash || 'solid'}`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleRemoveStroke = useCallback(() => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateStroke(renderer.slide, selectedIdx, -1, -1, -1, 0), 'Stroke removed', selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleDuplicate = useCallback(() => {
    if (selectedIdx < 0) return;
    const result = renderer.duplicateShape(renderer.slide, selectedIdx);
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    const newIdx = parseInt(result.split(':')[1]);
    renderer.setStatus(`Duplicated -> #${newIdx}`);
    renderAndRefresh(newIdx);
    refreshThumbs();
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleDelete = useCallback(() => {
    if (selectedIdx < 0) return;
    const isPic = selection?.shapeType === 'picture';
    const result = isPic
      ? renderer.deleteImage(renderer.slide, selectedIdx)
      : renderer.deleteShape(renderer.slide, selectedIdx);
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    renderer.setStatus(`Deleted #${selectedIdx}`);
    setSelection(null);
    renderAndRefresh();
    refreshThumbs();
  }, [selectedIdx, selection, renderer, renderAndRefresh]);

  // ── Text editing handlers ──
  const handleUpdateText = useCallback((pi: number, ri: number, text: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateText(renderer.slide, selectedIdx, pi, ri, text), `P${pi}R${ri} updated`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateStyle = useCallback((pi: number, ri: number, bold: number, italic: number) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateTextRunStyle(renderer.slide, selectedIdx, pi, ri, bold, italic), `Style updated`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateFontSize = useCallback((pi: number, ri: number, size: number) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateTextRunFontSize(renderer.slide, selectedIdx, pi, ri, size), `Size -> ${size / 100}pt`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateColor = useCallback((pi: number, ri: number, hex: string) => {
    if (selectedIdx < 0) return;
    const [r, g, b] = hexToRgb(hex);
    callAndRefresh(() => renderer.updateTextRunColor(renderer.slide, selectedIdx, pi, ri, r, g, b), `Color -> ${hex}`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateFont = useCallback((pi: number, ri: number, font: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateTextRunFont(renderer.slide, selectedIdx, pi, ri, font), `Font -> ${font}`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateDecoration = useCallback((pi: number, ri: number, underline: string, strike: string, baseline: number) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateTextRunDecoration(renderer.slide, selectedIdx, pi, ri, underline, strike, baseline), `Decoration updated`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleUpdateAlign = useCallback((pi: number, align: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.updateParagraphAlign(renderer.slide, selectedIdx, pi, align), `Align -> ${align}`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleAddParagraph = useCallback((text: string, align: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.addParagraph(renderer.slide, selectedIdx, text, align), 'Paragraph added', selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleDeleteParagraph = useCallback((pi: number) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.deleteParagraph(renderer.slide, selectedIdx, pi), `P${pi} deleted`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleAddRun = useCallback((pi: number, text: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.addRun(renderer.slide, selectedIdx, pi, text), `Run added`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleDeleteRun = useCallback((pi: number, ri: number) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.deleteRun(renderer.slide, selectedIdx, pi, ri), `R${ri} deleted`, selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  const handleAddShapeText = useCallback((text: string) => {
    if (selectedIdx < 0) return;
    callAndRefresh(() => renderer.addShapeText(renderer.slide, selectedIdx, text, DEFAULT_FONT_SIZE), 'Text added', selectedIdx);
  }, [selectedIdx, renderer, renderAndRefresh]);

  // ── Add shape ──
  const [addShapeType, setAddShapeType] = useState('rect');
  const [addShapeFill, setAddShapeFill] = useState('#4a90d9');

  const handleAddShape = useCallback(() => {
    const [r, g, b] = hexToRgb(addShapeFill);
    const isLine = addShapeType === 'line';
    const cx = DEFAULT_SHAPE_SIZE, cy = isLine ? 0 : DEFAULT_SHAPE_SIZE;
    const x = Math.round((DEFAULT_SLIDE_CX - cx) / 2);
    const y = Math.round((DEFAULT_SLIDE_CY - cy) / 2);
    const result = renderer.addShape(renderer.slide, addShapeType, x, y, cx, cy,
      isLine ? -1 : r, isLine ? -1 : g, isLine ? -1 : b);
    if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
    const newIdx = parseInt(result.split(':')[1]);
    if (isLine) renderer.updateStroke(renderer.slide, newIdx, r, g, b, DEFAULT_STROKE_WIDTH);
    renderer.setStatus(`Added ${addShapeType} -> #${newIdx}`);
    renderAndRefresh(newIdx);
    refreshThumbs();
  }, [addShapeType, addShapeFill, renderer, renderAndRefresh]);

  // ── Image operations ──
  const imageInputRef = useRef<HTMLInputElement>(null);
  const replaceInputRef = useRef<HTMLInputElement>(null);

  const handleAddImage = useCallback(async (file: File) => {
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const mime = file.type || 'image/png';
      const cx = DEFAULT_IMAGE_CX, cy = DEFAULT_IMAGE_CY;
      const x = Math.round((DEFAULT_SLIDE_CX - cx) / 2), y = Math.round((DEFAULT_SLIDE_CY - cy) / 2);
      const result = await renderer.addImage(renderer.slide, data, mime, x, y, cx, cy);
      if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
      const newIdx = parseInt(result.split(':')[1]);
      renderer.setStatus(`Added image -> #${newIdx}`);
      renderAndRefresh(newIdx);
      refreshThumbs();
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, renderAndRefresh]);

  const handleReplaceImage = useCallback(async (file: File) => {
    if (selectedIdx < 0) return;
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      const result = await renderer.replaceImage(renderer.slide, selectedIdx, data, file.type || 'image/png');
      if (result.startsWith('ERROR:')) { renderer.setStatus(result); return; }
      renderer.setStatus(`Replaced image`);
      renderAndRefresh(selectedIdx);
      refreshThumbs();
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [selectedIdx, renderer, renderAndRefresh]);

  // ── Slide management ──
  const handleAddSlide = useCallback(async () => {
    try {
      const { insertedIdx } = await renderer.addSlide(renderer.slide);
      renderer.setStatus(`Added slide ${insertedIdx + 1}`);
      refreshThumbs();
      handleGoToSlide(insertedIdx);
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, handleGoToSlide]);

  const handleDuplicateSlide = useCallback(async () => {
    try {
      const { insertedIdx } = await renderer.addSlide(renderer.slide, renderer.slide);
      renderer.setStatus(`Duplicated -> ${insertedIdx + 1}`);
      refreshThumbs();
      handleGoToSlide(insertedIdx);
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, handleGoToSlide]);

  const handleDeleteSlide = useCallback(async () => {
    if (renderer.total <= 1) return;
    const idx = renderer.slide;
    try {
      await renderer.deleteSlide(idx);
      renderer.setStatus(`Deleted slide ${idx + 1}`);
      refreshThumbs();
      handleGoToSlide(Math.min(idx, renderer.total - 2));
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, handleGoToSlide]);

  const handleMoveSlide = useCallback(async (dir: -1 | 1) => {
    const cur = renderer.slide;
    if (cur + dir < 0 || cur + dir >= renderer.total) return;
    const order = Array.from({ length: renderer.total }, (_, i) => i);
    [order[cur], order[cur + dir]] = [order[cur + dir], order[cur]];
    try {
      await renderer.reorderSlides(order);
      const newIdx = cur + dir;
      renderer.setSlide(newIdx);
      renderer.setStatus(`Moved slide ${dir < 0 ? 'left' : 'right'}`);
      refreshThumbs();
      handleGoToSlide(newIdx);
    } catch (err) { renderer.setStatus(`Error: ${(err as Error).message}`); }
  }, [renderer, handleGoToSlide]);

  // ── Render ──
  const loaded = renderer.total > 0;
  const isPicture = selection?.shapeType === 'picture';

  return (
    <div style={{ margin: '0 auto', padding: '12px 16px', fontFamily: 'system-ui, sans-serif' }}>
      {/* ── Header ── */}
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 12, marginBottom: 2 }}>
        <h1 style={{ margin: 0, fontSize: 18 }}>pptx-svg React Editor</h1>
        <span style={{ color: '#999', fontSize: 11 }}>Click shapes to select, drag to move. All changes export to PPTX.</span>
      </div>

      {/* ── Action bar (Export, Open, Load) ── */}
      {loaded && (
        <div style={{ display: 'flex', gap: 6, alignItems: 'center', marginBottom: 6, flexWrap: 'wrap', fontSize: 12 }}>
          <button onClick={renderer.exportPptx}>Export PPTX</button>
          <DropZone onFile={handleLoadFile} compact />
          <span style={{ width: 1, height: 18, background: '#ddd' }} />
          <button onClick={() => handleGoToSlide(renderer.slide - 1)} disabled={renderer.slide === 0}>&#x25C0;</button>
          <span style={{ minWidth: 50, textAlign: 'center', fontSize: 12 }}>{renderer.slide + 1} / {renderer.total}</span>
          <button onClick={() => handleGoToSlide(renderer.slide + 1)} disabled={renderer.slide >= renderer.total - 1}>&#x25B6;</button>
          <span style={{ width: 1, height: 18, background: '#ddd' }} />
          <button onClick={handleAddSlide}>+ Slide</button>
          <button onClick={handleDuplicateSlide}>Dup</button>
          <button onClick={() => handleMoveSlide(-1)} disabled={renderer.slide === 0}>&larr;</button>
          <button onClick={() => handleMoveSlide(1)} disabled={renderer.slide >= renderer.total - 1}>&rarr;</button>
          <button onClick={handleDeleteSlide} disabled={renderer.total <= 1}
            style={{ borderColor: '#e74c3c', color: '#e74c3c' }}>Del</button>
          <span style={{ width: 1, height: 18, background: '#ddd' }} />
          <select value={addShapeType} onChange={e => setAddShapeType(e.target.value)}
            style={{ padding: '2px 4px', border: '1px solid #ccc', borderRadius: 3, fontSize: 11 }}>
            <option value="rect">Rect</option>
            <option value="ellipse">Ellipse</option>
            <option value="roundRect">RoundRect</option>
            <option value="line">Line</option>
          </select>
          <input type="color" value={addShapeFill} onChange={e => setAddShapeFill(e.target.value)}
            style={{ width: 22, height: 18, border: '1px solid #ccc', borderRadius: 3, padding: 0, cursor: 'pointer' }} />
          <button onClick={handleAddShape}>+ Shape</button>
          <input ref={imageInputRef} type="file" accept="image/*" style={{ display: 'none' }}
            onChange={e => { const f = e.target.files?.[0]; if (f) handleAddImage(f); e.target.value = ''; }} />
          <button onClick={() => imageInputRef.current?.click()}>+ Image</button>
          {isPicture && (
            <>
              <input ref={replaceInputRef} type="file" accept="image/*" style={{ display: 'none' }}
                onChange={e => { const f = e.target.files?.[0]; if (f) handleReplaceImage(f); e.target.value = ''; }} />
              <button onClick={() => replaceInputRef.current?.click()}>Replace</button>
            </>
          )}
        </div>
      )}

      {!loaded && <DropZone onFile={handleLoadFile} />}

      {loaded && (
        <>
          {/* ── Thumbnails ── */}
          <SlideThumbnails
            total={renderer.total}
            current={renderer.slide}
            renderSlide={renderer.renderSlide}
            onSelect={handleGoToSlide}
            refreshKey={thumbKey}
          />

          {/* ── 2-column: SVG (left) + sidebar (right, always visible) ── */}
          <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
            {/* Left: SVG viewer */}
            <div style={{ flex: '1 1 0', minWidth: 0 }}>
              <SlideViewer containerRef={containerRef} hasSelection={selectedIdx >= 0} onSelect={setSelection} />
              <div style={{ display: 'flex', gap: 12, marginTop: 4, fontSize: 11, color: '#999' }}>
                {selection && <span style={{ fontFamily: 'monospace' }}>{selection.detail}</span>}
                <span>{renderer.status}</span>
              </div>
            </div>

            {/* Right: editing sidebar (always rendered, fixed width) */}
            <div style={{
              flex: `0 0 ${SIDEBAR_WIDTH}px`, width: SIDEBAR_WIDTH,
              maxHeight: 'calc(100vh - 160px)', overflowY: 'auto',
            }}>
              {selection ? (
                <>
                  <ShapeToolbar
                    shapeLabel={selection.label}
                    fillHex={selection.fillHex}
                    onApplyFill={handleApplyFill}
                    onApplyStroke={handleApplyStroke}
                    onRemoveStroke={handleRemoveStroke}
                    onDuplicate={handleDuplicate}
                    onDelete={handleDelete}
                  />
                  <div style={{ marginTop: 6 }}>
                    <TextPanel
                      paragraphs={selection.paragraphs}
                      onUpdateText={handleUpdateText}
                      onUpdateStyle={handleUpdateStyle}
                      onUpdateFontSize={handleUpdateFontSize}
                      onUpdateColor={handleUpdateColor}
                      onUpdateFont={handleUpdateFont}
                      onUpdateDecoration={handleUpdateDecoration}
                      onUpdateAlign={handleUpdateAlign}
                      onAddParagraph={handleAddParagraph}
                      onDeleteParagraph={handleDeleteParagraph}
                      onAddRun={handleAddRun}
                      onDeleteRun={handleDeleteRun}
                      onAddShapeText={handleAddShapeText}
                    />
                  </div>
                </>
              ) : (
                <div style={{
                  padding: 16, background: '#fafbfc', border: '1px solid #eee',
                  borderRadius: 6, color: '#bbb', fontSize: 12, textAlign: 'center',
                }}>
                  Click a shape to edit
                </div>
              )}
            </div>
          </div>
        </>
      )}

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
          padding: 4px 10px; border: 1px solid #ccc; border-radius: 4px;
          cursor: pointer; background: white; font-size: 12px;
        }
        button:hover:not(:disabled) { background: #f5f5f5; }
        button:disabled { opacity: 0.4; cursor: not-allowed; }
      `}</style>
    </div>
  );
}
