/** SlideRail — vertical thumbnail list with add/duplicate/delete + drag-reorder. */
import { memo, useEffect, useRef, useState } from 'react';
import { insertSvgInto } from '../utils/svg';

interface Props {
  total: number;
  current: number;
  versions: number[];                 // per-slide bump to refresh only changed thumbs
  renderSlide: (idx: number) => string;
  onSelect: (idx: number) => void;
  onAdd: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onReorder: (from: number, to: number) => void;
}

const Thumb = memo(function Thumb({ idx, active, ver, renderSlide, onSelect, onDragStart, onDrop }: {
  idx: number; active: boolean; ver: number;
  renderSlide: (i: number) => string;
  onSelect: (i: number) => void;
  onDragStart: (i: number) => void;
  onDrop: (i: number) => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [over, setOver] = useState(false);
  useEffect(() => {
    const el = ref.current; if (!el) return;
    const svg = renderSlide(idx);
    if (svg.startsWith('<svg')) insertSvgInto(el, svg);
  }, [idx, ver, renderSlide]);
  return (
    <div className={`rail-item ${active ? 'active' : ''} ${over ? 'drop' : ''}`}
      draggable
      onClick={() => onSelect(idx)}
      onDragStart={e => { e.dataTransfer.effectAllowed = 'move'; onDragStart(idx); }}
      onDragOver={e => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); onDrop(idx); }}>
      <span className="rail-num">{idx + 1}</span>
      <div className="rail-svg" ref={ref} />
    </div>
  );
});

export function SlideRail({ total, current, versions, renderSlide, onSelect, onAdd, onDuplicate, onDelete, onReorder }: Props) {
  const dragFrom = useRef<number | null>(null);
  return (
    <aside className="rail">
      <div className="rail-actions">
        <button title="Add slide" onClick={onAdd}>＋</button>
        <button title="Duplicate slide" onClick={onDuplicate}>⧉</button>
        <button title="Delete slide" className="danger" disabled={total <= 1} onClick={onDelete}>🗑</button>
      </div>
      <div className="rail-list">
        {Array.from({ length: total }, (_, i) => (
          <Thumb key={i} idx={i} active={i === current} ver={versions[i] ?? 0}
            renderSlide={renderSlide} onSelect={onSelect}
            onDragStart={i2 => { dragFrom.current = i2; }}
            onDrop={to => { const from = dragFrom.current; dragFrom.current = null; if (from !== null && from !== to) onReorder(from, to); }} />
        ))}
      </div>
    </aside>
  );
}
