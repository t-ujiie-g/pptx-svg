/**
 * DropZone — file drop area for loading .pptx files.
 * When used without children, shows a full drop area.
 * When compact=true (after initial load), shows only a small button.
 */

import { useCallback, useRef } from 'react';

interface DropZoneProps {
  onFile: (file: File) => void;
  compact?: boolean;
}

export function DropZone({ onFile, compact }: DropZoneProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
  }, [onFile]);

  if (compact) {
    return (
      <>
        <input
          ref={inputRef}
          type="file" accept=".pptx" style={{ display: 'none' }}
          onChange={(e) => { if (e.target.files?.[0]) onFile(e.target.files[0]); }}
        />
        <button onClick={() => inputRef.current?.click()} title="Load another .pptx file"
          style={{ padding: '4px 10px', border: '1px solid #ccc', borderRadius: 4, cursor: 'pointer', background: 'white', fontSize: 12 }}>
          Open
        </button>
      </>
    );
  }

  return (
    <div
      onDragOver={(e) => e.preventDefault()}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      style={{
        border: '2px dashed #ccc', borderRadius: 8, padding: 32,
        textAlign: 'center', cursor: 'pointer', marginBottom: 16,
      }}
    >
      <p><strong>Drop a .pptx file here</strong> or click to browse</p>
      <input
        ref={inputRef}
        type="file" accept=".pptx" style={{ display: 'none' }}
        onChange={(e) => { if (e.target.files?.[0]) onFile(e.target.files[0]); }}
      />
    </div>
  );
}
