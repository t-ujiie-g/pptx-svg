/**
 * EditToolbar — fill color picker + text run editors.
 *
 * Shown when a shape is selected. Provides:
 * - Color picker + "Apply Fill" button
 * - Per-run text input fields with "Apply" buttons
 */

import { useState, useEffect } from 'react';
import type { TextRun } from '../utils/svg';

// ── Fill toolbar ──

interface FillToolbarProps {
  fillColor: string;
  shapeLabel: string;
  onApplyFill: (hex: string) => void;
}

export function FillToolbar({ fillColor, shapeLabel, onApplyFill }: FillToolbarProps) {
  const [color, setColor] = useState(fillColor);

  useEffect(() => { setColor(fillColor); }, [fillColor]);

  return (
    <div style={{
      display: 'flex', gap: 8, alignItems: 'center', marginBottom: 12,
      padding: 12, background: '#f0f7ff', border: '1px solid #4a90d9',
      borderRadius: 6, flexWrap: 'wrap',
    }}>
      <label style={{ fontWeight: 'bold', whiteSpace: 'nowrap', fontSize: 14 }}>Fill:</label>
      <input
        type="color" value={color}
        onChange={(e) => setColor(e.target.value)}
        style={{ width: 36, height: 28, border: '1px solid #ccc', borderRadius: 4, padding: 0, cursor: 'pointer' }}
      />
      <button onClick={() => onApplyFill(color)}>Apply Fill</button>
      <span style={{ width: 1, height: 24, background: '#ccc', margin: '0 4px' }} />
      <span style={{ color: '#666', fontSize: 13 }}>{shapeLabel}</span>
    </div>
  );
}

// ── Text runs panel ──

interface TextRunsPanelProps {
  runs: TextRun[];
  onApplyText: (pi: number, ri: number, text: string) => void;
}

export function TextRunsPanel({ runs, onApplyText }: TextRunsPanelProps) {
  if (runs.length === 0) return null;

  return (
    <div style={{
      display: 'flex', flexDirection: 'column', gap: 6, padding: 12,
      background: '#f8f9fa', border: '1px solid #ddd', borderRadius: 6,
      marginBottom: 12, maxHeight: 200, overflowY: 'auto',
    }}>
      {runs.map((run) => (
        <TextRunRow key={`${run.pi}:${run.ri}`} run={run} onApply={onApplyText} />
      ))}
    </div>
  );
}

// ── Single text run row ──

function TextRunRow({ run, onApply }: {
  run: TextRun;
  onApply: (pi: number, ri: number, text: string) => void;
}) {
  const [text, setText] = useState(run.text);

  useEffect(() => { setText(run.text); }, [run.text]);

  return (
    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
      <span style={{ color: '#999', fontSize: 12, fontFamily: 'monospace', minWidth: 50, whiteSpace: 'nowrap' }}>
        P{run.pi}R{run.ri}
      </span>
      <input
        type="text" value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter') onApply(run.pi, run.ri, text); }}
        style={{ flex: 1, padding: '4px 8px', border: '1px solid #ccc', borderRadius: 4, fontSize: 13 }}
      />
      <button onClick={() => onApply(run.pi, run.ri, text)} style={{ padding: '4px 10px', fontSize: 13 }}>
        Apply
      </button>
    </div>
  );
}
