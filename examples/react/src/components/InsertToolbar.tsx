/** InsertToolbar — insert shapes/text/image/table, plus align & z-order quick actions. */
import { useRef, useState } from 'react';

interface Props {
  onAddShape: (geom: string) => void;
  onAddTextBox: () => void;
  onAddImage: (file: File) => void;
  hasSelection: boolean;
  onAlign: (how: 'l' | 'c' | 'r' | 't' | 'm' | 'b') => void;
  onZ: (how: 'front' | 'forward' | 'backward' | 'back') => void;
}

export function InsertToolbar({ onAddShape, onAddTextBox, onAddImage, hasSelection, onAlign, onZ }: Props) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [shapeMenu, setShapeMenu] = useState(false);

  return (
    <div className="insertbar">
      <button onClick={onAddTextBox}>＋ Text</button>

      <div className="menu-wrap">
        <button onClick={() => setShapeMenu(v => !v)}>＋ Shape ▾</button>
        {shapeMenu && (
          <div className="menu" onMouseLeave={() => setShapeMenu(false)}>
            {(['rect', 'roundRect', 'ellipse', 'line'] as const).map(g => (
              <button key={g} onClick={() => { onAddShape(g); setShapeMenu(false); }}>{g}</button>
            ))}
          </div>
        )}
      </div>

      <button onClick={() => fileRef.current?.click()}>＋ Image</button>
      <input ref={fileRef} type="file" accept="image/*" hidden
        onChange={e => { const f = e.target.files?.[0]; if (f) onAddImage(f); e.target.value = ''; }} />

      <span className="sep" />

      <div className={`grp ${hasSelection ? '' : 'disabled'}`}>
        <span className="lbl">Align</span>
        <button title="Left" onClick={() => onAlign('l')}>⌳</button>
        <button title="Center" onClick={() => onAlign('c')}>↔</button>
        <button title="Right" onClick={() => onAlign('r')}>⌲</button>
        <button title="Top" onClick={() => onAlign('t')}>⌶</button>
        <button title="Middle" onClick={() => onAlign('m')}>↕</button>
        <button title="Bottom" onClick={() => onAlign('b')}>⌷</button>
      </div>

      <span className="sep" />

      <div className={`grp ${hasSelection ? '' : 'disabled'}`}>
        <span className="lbl">Order</span>
        <button title="To front" onClick={() => onZ('front')}>⤒</button>
        <button title="Forward" onClick={() => onZ('forward')}>▲</button>
        <button title="Backward" onClick={() => onZ('backward')}>▼</button>
        <button title="To back" onClick={() => onZ('back')}>⤓</button>
      </div>
    </div>
  );
}
