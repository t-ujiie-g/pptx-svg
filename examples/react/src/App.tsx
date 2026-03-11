import { useCallback, useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';

// Vite resolves this to a URL pointing to the actual .wasm asset.
// Used as fallback if the library's auto-resolution fails in dev mode.
import wasmUrl from '/node_modules/pptx-svg/dist/main.wasm?url';

export default function App() {
  const rendererRef = useRef<PptxRenderer | null>(null);
  const [slide, setSlide] = useState(0);
  const [total, setTotal] = useState(0);
  const [status, setStatus] = useState('Loading WebAssembly module...');
  const [editingText, setEditingText] = useState<{ el: SVGTSpanElement; text: string } | null>(null);
  const [selectedShape, setSelectedShape] = useState<SVGGElement | null>(null);

  // This ref holds the SVG viewer div — React never touches its children.
  const svgContainerRef = useRef<HTMLDivElement>(null);

  // ── Helpers ──

  /** Insert SVG into the viewer (outside React's control). */
  const insertSvg = useCallback((svgString: string) => {
    const container = svgContainerRef.current;
    if (!container) return;
    if (svgString.startsWith('ERROR:')) {
      container.innerHTML = `<span style="color:red">${svgString}</span>`;
      return;
    }
    container.innerHTML = svgString;
    // Remove fixed width/height so CSS can make SVG responsive
    const svgEl = container.querySelector('svg');
    if (svgEl) {
      const w = svgEl.getAttribute('width');
      const h = svgEl.getAttribute('height');
      if (w && h && !svgEl.getAttribute('viewBox')) {
        svgEl.setAttribute('viewBox', `0 0 ${w} ${h}`);
      }
      svgEl.removeAttribute('width');
      svgEl.removeAttribute('height');
    }
  }, []);

  /** Push current SVG state back to the renderer. */
  const syncSvgToRenderer = useCallback(() => {
    const renderer = rendererRef.current;
    const container = svgContainerRef.current;
    if (!renderer || !container) return;
    const svgEl = container.querySelector('svg');
    if (svgEl) {
      renderer.updateSlideFromSvg(slide, svgEl.outerHTML);
    }
  }, [slide]);

  // ── Init ──
  useEffect(() => {
    const renderer = new PptxRenderer();
    rendererRef.current = renderer;
    renderer.init()
      .catch(() => renderer.init(wasmUrl))
      .then(() => setStatus('Ready. Drop a .pptx file.'))
      .catch(err => setStatus(`Init failed: ${err.message}`));
  }, []);

  // ── Navigation ──
  const goToSlide = useCallback((idx: number) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    setSlide(idx);
    setSelectedShape(null);
    setEditingText(null);
    insertSvg(renderer.renderSlideSvg(idx));
  }, [insertSvg]);

  // ── File load ──
  const handleFile = useCallback(async (file: File) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    setStatus(`Loading ${file.name}...`);
    try {
      const { slideCount } = await renderer.loadPptx(await file.arrayBuffer());
      setTotal(slideCount);
      setStatus(`Loaded "${file.name}" - ${slideCount} slide(s)`);
      goToSlide(0);
    } catch (err) {
      setStatus(`Error: ${(err as Error).message}`);
    }
  }, [goToSlide]);

  // ── Export ──
  const handleExport = useCallback(async () => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    syncSvgToRenderer();
    setStatus('Exporting...');
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
  }, [syncSvgToRenderer]);

  // ── Click: select shape ──
  const handleClick = useCallback((e: MouseEvent) => {
    setEditingText(null);
    const target = e.target as Element;
    const shapeG = target.closest('g[data-ooxml-shape-type]') as SVGGElement | null;

    // Clear old selection
    svgContainerRef.current?.querySelectorAll('.pptx-selected').forEach(el =>
      el.classList.remove('pptx-selected')
    );

    if (shapeG) {
      shapeG.classList.add('pptx-selected');
      setSelectedShape(shapeG);
    } else {
      setSelectedShape(null);
    }
  }, []);

  // ── Double-click: edit text ──
  const handleDblClick = useCallback((e: MouseEvent) => {
    const target = e.target as Element;
    const tspan = target.closest('tspan[data-ooxml-run-idx]') as SVGTSpanElement | null;
    if (!tspan) return;
    e.preventDefault();
    setEditingText({ el: tspan, text: tspan.textContent || '' });
  }, []);

  // ── Commit text edit ──
  const commitTextEdit = useCallback(() => {
    if (!editingText) return;
    editingText.el.textContent = editingText.text;
    syncSvgToRenderer();
    setEditingText(null);
  }, [editingText, syncSvgToRenderer]);

  // ── Drag to move shape ──
  const dragState = useRef<{
    shape: SVGGElement;
    startX: number; startY: number;
    origTransform: string;
    origEmuX: number; origEmuY: number;
    emuPerCssPx: number; // EMU per 1 CSS pixel on screen
  } | null>(null);

  const handleMouseDown = useCallback((e: MouseEvent) => {
    if (!selectedShape) return;
    const svgEl = svgContainerRef.current?.querySelector('svg') as SVGSVGElement | null;
    if (!svgEl) return;

    const target = e.target as Element;
    const clickedShape = target.closest('g[data-ooxml-shape-type]');
    if (clickedShape !== selectedShape) return;

    const origEmuX = parseInt(selectedShape.getAttribute('data-ooxml-x') || '0', 10);
    const origEmuY = parseInt(selectedShape.getAttribute('data-ooxml-y') || '0', 10);

    // Calculate how many EMU correspond to 1 CSS pixel on screen.
    // The SVG viewBox is in "renderer pixels" (EMU / scale), and it's
    // displayed at the bounding-client-rect size by CSS.
    const rect = svgEl.getBoundingClientRect();
    const slideCx = parseInt(svgEl.getAttribute('data-ooxml-slide-cx') || '9144000', 10);
    const emuPerCssPx = slideCx / rect.width;

    dragState.current = {
      shape: selectedShape,
      startX: e.clientX,
      startY: e.clientY,
      origTransform: selectedShape.getAttribute('transform') || '',
      origEmuX, origEmuY,
      emuPerCssPx,
    };
    e.preventDefault();
  }, [selectedShape]);

  // Attach native event listeners to the container (avoids React DOM conflicts)
  useEffect(() => {
    const container = svgContainerRef.current;
    if (!container) return;
    container.addEventListener('click', handleClick);
    container.addEventListener('dblclick', handleDblClick);
    container.addEventListener('mousedown', handleMouseDown);
    return () => {
      container.removeEventListener('click', handleClick);
      container.removeEventListener('dblclick', handleDblClick);
      container.removeEventListener('mousedown', handleMouseDown);
    };
  }, [handleClick, handleDblClick, handleMouseDown]);

  // Global mouse move/up for dragging
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      const ds = dragState.current;
      if (!ds) return;
      const dx = e.clientX - ds.startX;
      const dy = e.clientY - ds.startY;

      // Update EMU position
      const newEmuX = ds.origEmuX + Math.round(dx * ds.emuPerCssPx);
      const newEmuY = ds.origEmuY + Math.round(dy * ds.emuPerCssPx);
      ds.shape.setAttribute('data-ooxml-x', String(newEmuX));
      ds.shape.setAttribute('data-ooxml-y', String(newEmuY));

      // Visual feedback: get the SVG's viewBox-to-screen ratio to convert
      // CSS pixel delta into SVG coordinate delta for the translate offset.
      const svgEl = svgContainerRef.current?.querySelector('svg') as SVGSVGElement | null;
      if (!svgEl) return;
      const rect = svgEl.getBoundingClientRect();
      const vb = svgEl.viewBox.baseVal;
      const svgPxPerCssPx = vb.width / rect.width;
      const svgDx = dx * svgPxPerCssPx;
      const svgDy = dy * svgPxPerCssPx;

      // Apply pixel delta as an additional translate on top of original transform
      ds.shape.setAttribute('transform', `translate(${svgDx},${svgDy}) ${ds.origTransform}`);
    };

    const handleMouseUp = () => {
      if (!dragState.current) return;
      dragState.current = null;
      // Sync updated data-ooxml-x/y to renderer for export
      syncSvgToRenderer();
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [syncSvgToRenderer]);

  // ── Drop handler ──
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  }, [handleFile]);

  return (
    <div style={{ maxWidth: 1024, margin: '0 auto', padding: 24, fontFamily: 'system-ui, sans-serif' }}>
      <h1 style={{ marginBottom: 4 }}>pptx-svg React Example</h1>
      <p style={{ color: '#666', marginBottom: 20 }}>
        Click a shape to select, drag to move. Double-click text to edit. Export saves changes.
      </p>

      {/* Drop zone */}
      <div
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        onClick={() => document.getElementById('file-input')?.click()}
        style={{
          border: '2px dashed #ccc', borderRadius: 8, padding: 32,
          textAlign: 'center', cursor: 'pointer', marginBottom: 16,
        }}
      >
        <p><strong>Drop a .pptx file here</strong> or click to browse</p>
        <input
          id="file-input" type="file" accept=".pptx" style={{ display: 'none' }}
          onChange={(e) => { if (e.target.files?.[0]) handleFile(e.target.files[0]); }}
        />
      </div>

      {/* Controls */}
      {total > 0 && (
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 16, flexWrap: 'wrap' }}>
          <button onClick={() => goToSlide(slide - 1)} disabled={slide === 0}>Prev</button>
          <span style={{ minWidth: 80, textAlign: 'center' }}>{slide + 1} / {total}</span>
          <button onClick={() => goToSlide(slide + 1)} disabled={slide >= total - 1}>Next</button>
          <button onClick={handleExport} style={{ marginLeft: 'auto' }}>Export PPTX</button>
        </div>
      )}

      {/* Text editor */}
      {editingText && (
        <div style={{
          display: 'flex', gap: 8, alignItems: 'center',
          marginBottom: 12, padding: 12, background: '#f0f7ff',
          border: '1px solid #4a90d9', borderRadius: 6,
        }}>
          <label style={{ fontWeight: 'bold', whiteSpace: 'nowrap' }}>Edit text:</label>
          <input
            type="text"
            value={editingText.text}
            onChange={(e) => setEditingText({ ...editingText, text: e.target.value })}
            onKeyDown={(e) => { if (e.key === 'Enter') commitTextEdit(); if (e.key === 'Escape') setEditingText(null); }}
            autoFocus
            style={{ flex: 1, padding: '6px 10px', border: '1px solid #ccc', borderRadius: 4, fontSize: 14 }}
          />
          <button onClick={commitTextEdit}>Apply</button>
          <button onClick={() => setEditingText(null)} style={{ background: '#eee', color: '#333' }}>Cancel</button>
        </div>
      )}

      {/* SVG viewer — React does NOT manage this div's children */}
      <div
        ref={svgContainerRef}
        style={{
          background: '#fff', border: '1px solid #ddd', borderRadius: 8,
          padding: 16, minHeight: 300, overflow: 'hidden',
          cursor: selectedShape ? 'move' : 'default',
        }}
      />

      <p style={{ marginTop: 12, color: '#666', fontSize: 14 }}>{status}</p>

      <style>{`
        .pptx-selected > :first-child {
          outline: 2px solid #4a90d9;
          outline-offset: 2px;
        }
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
