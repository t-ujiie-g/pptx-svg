/**
 * EditToolbar — Shape editing controls (fill, stroke, actions)
 * TextPanel — Full text editing with paragraph management and run formatting
 */

import { useState, useEffect, useRef } from 'react';
import type { ParagraphInfo, TextRun } from '../utils/svg';

// ── Styles ──

const toolbarStyle: React.CSSProperties = {
  display: 'flex', gap: 6, alignItems: 'center',
  padding: '8px 10px', background: '#f8f9fb', border: '1px solid #dde3ea',
  borderRadius: 6, flexWrap: 'wrap', fontSize: 12,
};

const sepStyle: React.CSSProperties = {
  width: 1, height: 20, background: '#dde3ea', margin: '0 2px',
};

const labelStyle: React.CSSProperties = {
  fontWeight: 600, color: '#555', whiteSpace: 'nowrap', fontSize: 11,
};

const btnStyle: React.CSSProperties = {
  padding: '3px 8px', border: '1px solid #ccc', borderRadius: 4,
  background: '#fff', cursor: 'pointer', fontSize: 11,
};

const dangerBtnStyle: React.CSSProperties = {
  ...btnStyle, borderColor: '#e74c3c', color: '#e74c3c',
};

const colorInputStyle: React.CSSProperties = {
  width: 28, height: 22, border: '1px solid #ccc', borderRadius: 3,
  padding: 0, cursor: 'pointer',
};

const selectStyle: React.CSSProperties = {
  padding: '2px 4px', border: '1px solid #ccc', borderRadius: 3,
  background: '#fff', fontSize: 11, cursor: 'pointer',
};

// ── ShapeToolbar (fill, stroke, duplicate, delete) ──

interface ShapeToolbarProps {
  shapeLabel: string;
  fillHex: string;
  onApplyFill: (hex: string) => void;
  onApplyStroke: (hex: string, dash: string) => void;
  onRemoveStroke: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
}

export function ShapeToolbar({
  shapeLabel, fillHex,
  onApplyFill, onApplyStroke, onRemoveStroke,
  onDuplicate, onDelete,
}: ShapeToolbarProps) {
  const [fill, setFill] = useState(fillHex ? '#' + fillHex : '#4a90d9');
  const [strokeColor, setStrokeColor] = useState('#000000');
  const [strokeDash, setStrokeDash] = useState('');

  useEffect(() => { if (fillHex) setFill('#' + fillHex); }, [fillHex]);

  return (
    <div style={toolbarStyle}>
      <span style={labelStyle}>Fill</span>
      <input type="color" value={fill} onChange={e => setFill(e.target.value)} style={colorInputStyle} />
      <button style={btnStyle} onClick={() => onApplyFill(fill)}>Apply</button>

      <span style={sepStyle} />

      <span style={labelStyle}>Stroke</span>
      <input type="color" value={strokeColor} onChange={e => setStrokeColor(e.target.value)} style={colorInputStyle} />
      <select value={strokeDash} onChange={e => setStrokeDash(e.target.value)} style={selectStyle}>
        <option value="">Solid</option>
        <option value="dash">Dash</option>
        <option value="dot">Dot</option>
        <option value="dashDot">Dash-Dot</option>
        <option value="lgDash">Long Dash</option>
      </select>
      <button style={btnStyle} onClick={() => onApplyStroke(strokeColor, strokeDash)}>Apply</button>
      <button style={btnStyle} onClick={onRemoveStroke}>None</button>

      <span style={sepStyle} />

      <button style={btnStyle} onClick={onDuplicate}>Duplicate</button>
      <button style={dangerBtnStyle} onClick={onDelete}>Delete</button>

      <span style={sepStyle} />
      <span style={{ color: '#888', fontSize: 11 }}>{shapeLabel}</span>
    </div>
  );
}

// ── TextPanel ──

interface TextPanelProps {
  paragraphs: ParagraphInfo[];
  onUpdateText: (pi: number, ri: number, text: string) => void;
  onUpdateStyle: (pi: number, ri: number, bold: number, italic: number) => void;
  onUpdateFontSize: (pi: number, ri: number, size: number) => void;
  onUpdateColor: (pi: number, ri: number, hex: string) => void;
  onUpdateFont: (pi: number, ri: number, font: string) => void;
  onUpdateDecoration: (pi: number, ri: number, underline: string, strike: string, baseline: number) => void;
  onUpdateAlign: (pi: number, align: string) => void;
  onAddParagraph: (text: string, align: string) => void;
  onDeleteParagraph: (pi: number) => void;
  onAddRun: (pi: number, text: string) => void;
  onDeleteRun: (pi: number, ri: number) => void;
  onAddShapeText: (text: string) => void;
}

export function TextPanel({
  paragraphs,
  onUpdateText, onUpdateStyle, onUpdateFontSize, onUpdateColor,
  onUpdateFont, onUpdateDecoration, onUpdateAlign,
  onAddParagraph, onDeleteParagraph, onAddRun, onDeleteRun,
  onAddShapeText,
}: TextPanelProps) {
  return (
    <div style={{
      display: 'flex', flexDirection: 'column', gap: 4, padding: 10,
      background: '#fafbfc', border: '1px solid #dde3ea', borderRadius: 6,
      maxHeight: 350, overflowY: 'auto', fontSize: 12,
    }}>
      <span style={{ fontWeight: 600, color: '#555', fontSize: 11, marginBottom: 2 }}>Text</span>

      {paragraphs.length === 0 && (
        <AddShapeTextRow onAdd={onAddShapeText} />
      )}

      {paragraphs.map(para => (
        <ParagraphRow
          key={para.pi}
          para={para}
          onUpdateText={onUpdateText}
          onUpdateStyle={onUpdateStyle}
          onUpdateFontSize={onUpdateFontSize}
          onUpdateColor={onUpdateColor}
          onUpdateFont={onUpdateFont}
          onUpdateDecoration={onUpdateDecoration}
          onUpdateAlign={onUpdateAlign}
          onDeleteParagraph={onDeleteParagraph}
          onAddRun={onAddRun}
          onDeleteRun={onDeleteRun}
        />
      ))}

      {paragraphs.length > 0 && <AddParagraphRow onAdd={onAddParagraph} />}
    </div>
  );
}

// ── AddShapeTextRow (for shapes with no text body) ──

function AddShapeTextRow({ onAdd }: { onAdd: (text: string) => void }) {
  const [text, setText] = useState('');

  const doAdd = () => {
    onAdd(text || 'Text');
    setText('');
  };

  return (
    <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
      <input
        type="text" value={text}
        onChange={e => setText(e.target.value)}
        onKeyDown={e => { if (e.key === 'Enter') doAdd(); }}
        placeholder="Enter text..."
        style={{ flex: 1, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, fontSize: 12 }}
      />
      <button style={{ ...btnStyle, background: '#eef4fb', borderColor: '#4a90d9' }} onClick={doAdd}>
        + Add Text
      </button>
    </div>
  );
}

// ── ParagraphRow ──

function ParagraphRow({ para, onUpdateText, onUpdateStyle, onUpdateFontSize, onUpdateColor,
  onUpdateFont, onUpdateDecoration, onUpdateAlign, onDeleteParagraph, onAddRun, onDeleteRun,
}: {
  para: ParagraphInfo;
} & Pick<TextPanelProps, 'onUpdateText' | 'onUpdateStyle' | 'onUpdateFontSize' | 'onUpdateColor'
  | 'onUpdateFont' | 'onUpdateDecoration' | 'onUpdateAlign' | 'onDeleteParagraph' | 'onAddRun' | 'onDeleteRun'>) {

  return (
    <div style={{ borderTop: '1px solid #e8ecf0', paddingTop: 4, marginTop: 2 }}>
      <div style={{ display: 'flex', gap: 4, alignItems: 'center', marginBottom: 3 }}>
        <span style={{ fontWeight: 700, color: '#4a90d9', fontSize: 10, fontFamily: 'monospace' }}>
          P{para.pi}
        </span>
        <select
          value={para.align}
          onChange={e => onUpdateAlign(para.pi, e.target.value)}
          style={selectStyle}
          title="Paragraph alignment"
        >
          <option value="l">Left</option>
          <option value="ctr">Center</option>
          <option value="r">Right</option>
          <option value="just">Justify</option>
        </select>
        <button
          style={{ ...btnStyle, fontSize: 10 }}
          onClick={() => {
            const text = prompt('Text for new run:', 'New text');
            if (text !== null) onAddRun(para.pi, text);
          }}
        >+ Run</button>
        <button
          style={{ ...dangerBtnStyle, fontSize: 10, padding: '2px 6px' }}
          onClick={() => onDeleteParagraph(para.pi)}
          title="Delete paragraph"
        >Del</button>
      </div>

      {para.runs.map(run => (
        <RunRow
          key={`${run.pi}:${run.ri}`}
          run={run}
          onUpdateText={onUpdateText}
          onUpdateStyle={onUpdateStyle}
          onUpdateFontSize={onUpdateFontSize}
          onUpdateColor={onUpdateColor}
          onUpdateFont={onUpdateFont}
          onUpdateDecoration={onUpdateDecoration}
          onDeleteRun={onDeleteRun}
        />
      ))}
    </div>
  );
}

// ── RunRow ──

const fmtBtnBase: React.CSSProperties = {
  padding: '1px 5px', minWidth: 22, textAlign: 'center',
  border: '1px solid #ccc', borderRadius: 3, background: '#fff',
  cursor: 'pointer', fontFamily: 'serif', fontSize: 12, lineHeight: '18px',
};

const activeFmt: React.CSSProperties = { background: '#e3edf7', borderColor: '#4a90d9' };

function RunRow({ run, onUpdateText, onUpdateStyle, onUpdateFontSize, onUpdateColor,
  onUpdateFont, onUpdateDecoration, onDeleteRun,
}: {
  run: TextRun;
  onUpdateText: (pi: number, ri: number, text: string) => void;
  onUpdateStyle: (pi: number, ri: number, bold: number, italic: number) => void;
  onUpdateFontSize: (pi: number, ri: number, size: number) => void;
  onUpdateColor: (pi: number, ri: number, hex: string) => void;
  onUpdateFont: (pi: number, ri: number, font: string) => void;
  onUpdateDecoration: (pi: number, ri: number, underline: string, strike: string, baseline: number) => void;
  onDeleteRun: (pi: number, ri: number) => void;
}) {
  const [text, setText] = useState(run.text);
  const [font, setFont] = useState(run.font);
  const textRef = useRef(run.text);
  const fontRef = useRef(run.font);

  useEffect(() => { setText(run.text); textRef.current = run.text; }, [run.text]);
  useEffect(() => { setFont(run.font); fontRef.current = run.font; }, [run.font]);

  const applyText = () => {
    if (text !== textRef.current) onUpdateText(run.pi, run.ri, text);
  };
  const applyFont = () => {
    if (font !== fontRef.current) onUpdateFont(run.pi, run.ri, font);
  };

  return (
    <div style={{ display: 'flex', gap: 3, alignItems: 'center', marginBottom: 2, flexWrap: 'wrap' }}>
      <span style={{ color: '#999', fontSize: 10, fontFamily: 'monospace', minWidth: 24 }}>
        R{run.ri}
      </span>

      <input
        type="text" value={text}
        onChange={e => setText(e.target.value)}
        onBlur={applyText}
        onKeyDown={e => { if (e.key === 'Enter') applyText(); }}
        style={{ flex: 1, minWidth: 60, padding: '2px 5px', border: '1px solid #ccc', borderRadius: 3, fontSize: 11 }}
      />

      <button
        style={{ ...fmtBtnBase, fontWeight: 'bold', ...(run.bold ? activeFmt : {}) }}
        onClick={() => onUpdateStyle(run.pi, run.ri, run.bold ? 0 : 1, -1)}
        title="Bold"
      >B</button>

      <button
        style={{ ...fmtBtnBase, fontStyle: 'italic', ...(run.italic ? activeFmt : {}) }}
        onClick={() => onUpdateStyle(run.pi, run.ri, -1, run.italic ? 0 : 1)}
        title="Italic"
      >I</button>

      <button
        style={{ ...fmtBtnBase, textDecoration: 'underline', ...(run.underline ? activeFmt : {}) }}
        onClick={() => onUpdateDecoration(run.pi, run.ri, run.underline ? 'none' : 'sng', '', -1)}
        title="Underline"
      >U</button>

      <button
        style={{ ...fmtBtnBase, textDecoration: 'line-through', ...(run.strike ? activeFmt : {}) }}
        onClick={() => onUpdateDecoration(run.pi, run.ri, '', run.strike ? 'none' : 'sngStrike', -1)}
        title="Strikethrough"
      >S</button>

      <button
        style={{ ...fmtBtnBase, fontSize: 10, ...(run.baseline > 0 ? activeFmt : {}) }}
        onClick={() => onUpdateDecoration(run.pi, run.ri, '', '', run.baseline > 0 ? 0 : 30000)}
        title="Superscript"
      >A&sup2;</button>

      <button
        style={{ ...fmtBtnBase, fontSize: 10, ...(run.baseline < 0 ? activeFmt : {}) }}
        onClick={() => onUpdateDecoration(run.pi, run.ri, '', '', run.baseline < 0 ? 0 : -25000)}
        title="Subscript"
      >A&sub2;</button>

      <input
        type="number"
        defaultValue={run.fontSize > 0 ? Math.round(run.fontSize / 100) : ''}
        placeholder="pt"
        min={1} max={400}
        style={{ width: 38, padding: '1px 3px', border: '1px solid #ccc', borderRadius: 3, fontSize: 10, textAlign: 'center' }}
        title="Font size (pt)"
        onChange={e => {
          const pt = parseInt(e.target.value) || 0;
          if (pt > 0) onUpdateFontSize(run.pi, run.ri, pt * 100);
        }}
      />

      <input
        type="color"
        defaultValue={run.color.length === 6 ? '#' + run.color : '#000000'}
        style={{ ...colorInputStyle, width: 22, height: 18 }}
        title="Text color"
        onChange={e => onUpdateColor(run.pi, run.ri, e.target.value)}
      />

      <input
        type="text" value={font}
        onChange={e => setFont(e.target.value)}
        onBlur={applyFont}
        onKeyDown={e => { if (e.key === 'Enter') applyFont(); }}
        placeholder="Font"
        style={{ width: 60, padding: '1px 4px', border: '1px solid #ccc', borderRadius: 3, fontSize: 10 }}
        title="Font family"
      />

      <button
        style={{ ...dangerBtnStyle, padding: '1px 5px', fontSize: 11 }}
        onClick={() => onDeleteRun(run.pi, run.ri)}
        title="Delete run"
      >&times;</button>
    </div>
  );
}

// ── AddParagraphRow ──

function AddParagraphRow({ onAdd }: { onAdd: (text: string, align: string) => void }) {
  const [text, setText] = useState('');
  const [align, setAlign] = useState('');

  const doAdd = () => {
    onAdd(text, align);
    setText('');
  };

  return (
    <div style={{ display: 'flex', gap: 4, alignItems: 'center', borderTop: '1px solid #d0d8e0', paddingTop: 6, marginTop: 4 }}>
      <input
        type="text" value={text}
        onChange={e => setText(e.target.value)}
        onKeyDown={e => { if (e.key === 'Enter') doAdd(); }}
        placeholder="New paragraph..."
        style={{ flex: 1, padding: '3px 6px', border: '1px solid #ccc', borderRadius: 3, fontSize: 11 }}
      />
      <select value={align} onChange={e => setAlign(e.target.value)} style={selectStyle} title="Alignment">
        <option value="">Inherit</option>
        <option value="l">Left</option>
        <option value="ctr">Center</option>
        <option value="r">Right</option>
        <option value="just">Justify</option>
      </select>
      <button style={{ ...btnStyle, background: '#eef4fb', borderColor: '#4a90d9' }} onClick={doAdd}>
        + Para
      </button>
    </div>
  );
}
