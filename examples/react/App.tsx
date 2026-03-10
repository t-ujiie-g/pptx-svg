/**
 * pptx-svg React Example
 *
 * Usage:
 *   npm create vite@latest my-app -- --template react-ts
 *   cd my-app
 *   npm install pptx-svg
 *   # Copy this file to src/App.tsx
 *   npm run dev
 */
import { useCallback, useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';

export default function App() {
  const rendererRef = useRef<PptxRenderer | null>(null);
  const [ready, setReady] = useState(false);
  const [slide, setSlide] = useState(0);
  const [total, setTotal] = useState(0);
  const [svg, setSvg] = useState('');
  const [status, setStatus] = useState('Loading WebAssembly module...');

  useEffect(() => {
    const renderer = new PptxRenderer();
    rendererRef.current = renderer;
    renderer.init()
      .then(() => { setReady(true); setStatus('Ready. Drop a .pptx file.'); })
      .catch(err => setStatus(`Init failed: ${err.message}`));
  }, []);

  const showSlide = useCallback((idx: number) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    setSlide(idx);
    setSvg(renderer.renderSlideSvg(idx));
  }, []);

  const handleFile = useCallback(async (file: File) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    setStatus(`Loading ${file.name}...`);
    const { slideCount } = await renderer.loadPptx(await file.arrayBuffer());
    setTotal(slideCount);
    setStatus(`Loaded "${file.name}" - ${slideCount} slide(s)`);
    setSlide(0);
    setSvg(renderer.renderSlideSvg(0));
  }, []);

  const handleExport = useCallback(async () => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    const buffer = await renderer.exportPptx();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'exported.pptx';
    a.click();
    URL.revokeObjectURL(a.href);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  }, [handleFile]);

  return (
    <div style={{ maxWidth: 960, margin: '0 auto', padding: 24, fontFamily: 'system-ui' }}>
      <h1>pptx-svg React Example</h1>

      <div
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        onClick={() => document.getElementById('file-input')?.click()}
        style={{
          border: '2px dashed #ccc', borderRadius: 8, padding: 40,
          textAlign: 'center', cursor: 'pointer', marginBottom: 16,
        }}
      >
        <p><strong>Drop a .pptx file here</strong> or click to browse</p>
        <input
          id="file-input" type="file" accept=".pptx" style={{ display: 'none' }}
          onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
        />
      </div>

      {total > 0 && (
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 16 }}>
          <button onClick={() => showSlide(slide - 1)} disabled={slide === 0}>Prev</button>
          <span>{slide + 1} / {total}</span>
          <button onClick={() => showSlide(slide + 1)} disabled={slide >= total - 1}>Next</button>
          <button onClick={handleExport}>Export PPTX</button>
        </div>
      )}

      {svg && (
        <div
          style={{ background: '#fff', border: '1px solid #ddd', borderRadius: 8, padding: 16 }}
          dangerouslySetInnerHTML={{ __html: svg }}
        />
      )}

      <p style={{ marginTop: 12, color: '#666' }}>{status}</p>
    </div>
  );
}
