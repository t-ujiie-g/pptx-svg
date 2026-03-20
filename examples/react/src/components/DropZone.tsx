/**
 * DropZone — file drop area for loading .pptx files.
 */

import { useCallback, useRef } from 'react';

interface DropZoneProps {
  onFile: (file: File) => void;
}

export function DropZone({ onFile }: DropZoneProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
  }, [onFile]);

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
