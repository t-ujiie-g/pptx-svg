/**
 * PropertiesPanel — contextual editing panel for the selected shape.
 *
 * Sections (shown when relevant): Arrange (position/size/rotation/z-order),
 * Fill (solid + gradient), Line, Text (reuses TextPanel), Table, Image.
 *
 * It calls renderer APIs directly and reports the result string to `commit`,
 * which centralises re-render/reselect/history-sync in App.
 */

import { useEffect, useState } from 'react';
import { emuToPx, pxToEmu, degreesToOoxml, ooxmlToDegrees } from 'pptx-svg';
import { hexToRgb, okIndex, type ShapeInfo } from '../utils/svg';
import { TextPanel } from './TextPanel';
import type { RendererApi } from '../hooks/useRenderer';

const STROKE_PT = 12700;      // 1pt in EMU
const GRADIENT_END_POS = 100000; // gradient stop position scale (0–100000)

interface Props {
  renderer: RendererApi;
  slide: number;
  selection: ShapeInfo;
  multiCount: number;
  commit: (result: string, reselect?: number | null) => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onReplaceImage: (file: File) => void;
}

function Section({ title, children, defaultOpen = true }: { title: string; children: React.ReactNode; defaultOpen?: boolean }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="prop-section">
      <button className="prop-head" onClick={() => setOpen(o => !o)}>
        <span>{title}</span><span className="chev">{open ? '▾' : '▸'}</span>
      </button>
      {open && <div className="prop-body">{children}</div>}
    </div>
  );
}

export function PropertiesPanel({ renderer, slide, selection, multiCount, commit, onDuplicate, onDelete, onReplaceImage }: Props) {
  const idx = selection.idx;
  const t = selection.t;

  // Arrange (numeric, in px for display)
  const [xy, setXY] = useState({ x: 0, y: 0, w: 0, h: 0, r: 0 });
  useEffect(() => {
    setXY({ x: emuToPx(t.x), y: emuToPx(t.y), w: emuToPx(t.cx), h: emuToPx(t.cy), r: Math.round(ooxmlToDegrees(t.rot)) });
  }, [t.x, t.y, t.cx, t.cy, t.rot]);

  const norm360 = (d: number) => ((d % 360) + 360) % 360;
  const applyTransform = (next: Partial<typeof xy>) => {
    const v = { ...xy, ...next };
    setXY(v);
    commit(renderer.updateTransform(slide, idx, pxToEmu(v.x), pxToEmu(v.y), pxToEmu(Math.max(1, v.w)), pxToEmu(Math.max(0, v.h)), degreesToOoxml(norm360(v.r))));
  };

  // Fill
  const [fill, setFill] = useState('#4a90d9');
  const [g1, setG1] = useState('#ff5577');
  const [g2, setG2] = useState('#3366ff');
  const [gAngle, setGAngle] = useState(45);
  useEffect(() => { if (selection.fillHex) setFill('#' + selection.fillHex); }, [selection.fillHex]);

  // Line
  const [line, setLine] = useState('#222222');
  const [lineW, setLineW] = useState(1);
  const [dash, setDash] = useState('');

  // Table
  const [tr, setTR] = useState(0);
  const [tc, setTC] = useState(0);
  const [cellText, setCellText] = useState('');

  const isTable = selection.shapeType === 'table';
  const isPicture = selection.shapeType === 'picture';
  const hasText = selection.paragraphs.length > 0;

  return (
    <div className="props">
      <div className="props-title">
        {multiCount > 1 ? `${multiCount} shapes selected` : selection.label}
      </div>

      {multiCount > 1 ? (
        <div className="hint">Drag to move all together, or use arrow keys. Shift-click toggles selection.</div>
      ) : (
        <>
          <Section title="Arrange">
            <div className="grid2">
              <label>X<input type="number" value={xy.x} onChange={e => applyTransform({ x: +e.target.value })} /></label>
              <label>Y<input type="number" value={xy.y} onChange={e => applyTransform({ y: +e.target.value })} /></label>
              <label>W<input type="number" value={xy.w} onChange={e => applyTransform({ w: +e.target.value })} /></label>
              <label>H<input type="number" value={xy.h} onChange={e => applyTransform({ h: +e.target.value })} /></label>
              <label>Rotation<input type="number" value={xy.r} onChange={e => applyTransform({ r: +e.target.value })} /></label>
            </div>
            <div className="row">
              <button onClick={() => { const r = renderer.bringToFront(slide, idx); commit(r, okIndex(r, idx)); }}>To front</button>
              <button onClick={() => { const r = renderer.bringForward(slide, idx); commit(r, okIndex(r, idx)); }}>Forward</button>
              <button onClick={() => { const r = renderer.sendBackward(slide, idx); commit(r, okIndex(r, idx)); }}>Backward</button>
              <button onClick={() => { const r = renderer.sendToBack(slide, idx); commit(r, okIndex(r, idx)); }}>To back</button>
            </div>
            <div className="row">
              <button onClick={onDuplicate}>Duplicate</button>
              <button className="danger" onClick={onDelete}>Delete</button>
            </div>
          </Section>

          <Section title="Fill">
            <div className="row">
              <input type="color" value={fill} onChange={e => setFill(e.target.value)} />
              <button onClick={() => { const [r, g, b] = hexToRgb(fill); commit(renderer.updateFill(slide, idx, r, g, b)); }}>Solid</button>
            </div>
            <div className="row gradient">
              <input type="color" value={g1} onChange={e => setG1(e.target.value)} />
              <span>→</span>
              <input type="color" value={g2} onChange={e => setG2(e.target.value)} />
              <input type="number" className="angle" value={gAngle} onChange={e => setGAngle(+e.target.value)} title="angle°" />
              <button onClick={() => {
                const [r1, g1r, b1] = hexToRgb(g1), [r2, g2r, b2] = hexToRgb(g2);
                commit(renderer.updateGradientFill(slide, idx, degreesToOoxml(norm360(gAngle)),
                  [{ pos: 0, r: r1, g: g1r, b: b1 }, { pos: GRADIENT_END_POS, r: r2, g: g2r, b: b2 }]));
              }}>Gradient</button>
            </div>
          </Section>

          <Section title="Line">
            <div className="row">
              <input type="color" value={line} onChange={e => setLine(e.target.value)} />
              <input type="number" className="angle" value={lineW} min={0} step={0.5} onChange={e => setLineW(+e.target.value)} title="width (pt)" />
              <select value={dash} onChange={e => setDash(e.target.value)}>
                <option value="">solid</option><option value="dash">dash</option><option value="dot">dot</option>
                <option value="dashDot">dash-dot</option><option value="lgDash">long-dash</option>
              </select>
            </div>
            <div className="row">
              <button onClick={() => { const [r, g, b] = hexToRgb(line); commit(renderer.updateStroke(slide, idx, r, g, b, Math.round(lineW * STROKE_PT), dash)); }}>Apply</button>
              <button onClick={() => commit(renderer.updateStroke(slide, idx, -1, -1, -1, 0, ''))}>None</button>
            </div>
          </Section>

          {isPicture && (
            <Section title="Image">
              <div className="row">
                <label className="filebtn">Replace…
                  <input type="file" accept="image/*" hidden onChange={e => { const f = e.target.files?.[0]; if (f) onReplaceImage(f); e.target.value = ''; }} />
                </label>
                <button className="danger" onClick={onDelete}>Delete</button>
              </div>
            </Section>
          )}

          {isTable && (
            <Section title="Table">
              <div className="row">
                <label>R<input type="number" className="angle" value={tr} min={0} onChange={e => setTR(+e.target.value)} /></label>
                <label>C<input type="number" className="angle" value={tc} min={0} onChange={e => setTC(+e.target.value)} /></label>
              </div>
              <div className="row">
                <input type="text" placeholder="cell text" value={cellText} onChange={e => setCellText(e.target.value)} style={{ flex: 1 }} />
                <button onClick={() => commit(renderer.updateTableCellText(slide, idx, tr, tc, cellText))}>Set</button>
              </div>
              <div className="row">
                <button onClick={() => commit(renderer.addTableRow(slide, idx, tr))}>+Row</button>
                <button onClick={() => commit(renderer.deleteTableRow(slide, idx, tr))}>−Row</button>
                <button onClick={() => commit(renderer.addTableColumn(slide, idx, tc, 0))}>+Col</button>
                <button onClick={() => commit(renderer.deleteTableColumn(slide, idx, tc))}>−Col</button>
              </div>
            </Section>
          )}

          {hasText && (
            <Section title="Text">
              <TextPanel
                paragraphs={selection.paragraphs}
                onUpdateText={(pi, ri, text) => commit(renderer.updateText(slide, idx, pi, ri, text))}
                onUpdateStyle={(pi, ri, b, i) => commit(renderer.updateTextRunStyle(slide, idx, pi, ri, b, i))}
                onUpdateFontSize={(pi, ri, sz) => commit(renderer.updateTextRunFontSize(slide, idx, pi, ri, sz))}
                onUpdateColor={(pi, ri, hex) => { const [r, g, b] = hexToRgb(hex); commit(renderer.updateTextRunColor(slide, idx, pi, ri, r, g, b)); }}
                onUpdateFont={(pi, ri, font) => commit(renderer.updateTextRunFont(slide, idx, pi, ri, font))}
                onUpdateDecoration={(pi, ri, ul, st, bl) => commit(renderer.updateTextRunDecoration(slide, idx, pi, ri, ul, st, bl))}
                onUpdateAlign={(pi, align) => commit(renderer.updateParagraphAlign(slide, idx, pi, align))}
                onAddParagraph={(text, align) => commit(renderer.addParagraph(slide, idx, text, align))}
                onDeleteParagraph={(pi) => commit(renderer.deleteParagraph(slide, idx, pi))}
                onAddRun={(pi, text) => commit(renderer.addRun(slide, idx, pi, text))}
                onDeleteRun={(pi, ri) => commit(renderer.deleteRun(slide, idx, pi, ri))}
                onAddShapeText={(text) => commit(renderer.addShapeText(slide, idx, text, 1800))}
              />
            </Section>
          )}
        </>
      )}
    </div>
  );
}
