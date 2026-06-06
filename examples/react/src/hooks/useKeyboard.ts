/** useKeyboard — global editor shortcuts (disabled while typing in a field or inline). */
import { useEffect } from 'react';

interface Options {
  enabled: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onCopy: () => void;
  onPaste: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onNudge: (dx: number, dy: number) => void;
  onEscape: () => void;
}

const NUDGE = 45720; // 0.05in per arrow press

export function useKeyboard(o: Options) {
  useEffect(() => {
    if (!o.enabled) return;
    const onKey = (e: KeyboardEvent) => {
      const tag = (e.target as HTMLElement)?.tagName?.toUpperCase();
      const inField = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
      const mod = e.ctrlKey || e.metaKey;

      if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? o.onRedo() : o.onUndo(); return; }
      if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); o.onRedo(); return; }
      if (inField) return;
      if (mod && e.key.toLowerCase() === 'c') { e.preventDefault(); o.onCopy(); return; }
      if (mod && e.key.toLowerCase() === 'v') { e.preventDefault(); o.onPaste(); return; }
      if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); o.onDuplicate(); return; }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); o.onDelete(); return; }
      if (e.key === 'Escape') { o.onEscape(); return; }
      if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        const d = e.shiftKey ? NUDGE * 4 : NUDGE;
        if (e.key === 'ArrowLeft') o.onNudge(-d, 0);
        else if (e.key === 'ArrowRight') o.onNudge(d, 0);
        else if (e.key === 'ArrowUp') o.onNudge(0, -d);
        else if (e.key === 'ArrowDown') o.onNudge(0, d);
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [o]);
}
