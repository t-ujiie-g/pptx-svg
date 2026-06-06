/**
 * useInlineText — double-click-to-type inline text editing (E6.2).
 *
 * While a text shape is being edited it draws a dashed edit outline + a blinking
 * caret over the SVG, maps clicks to caret positions via hitTestText, moves the
 * caret with the arrow keys (getTextLayout geometry), and applies typing /
 * Backspace / Enter through replaceTextRange. Re-rendering the slide is delegated
 * to `commit`, which the parent uses to refresh the SVG; this hook then re-derives
 * the geometry.
 */

import { useCallback, useEffect, useRef } from 'react';
import { findShapeElement } from 'pptx-svg';
import { emuToContainerPx, clientToEmu, showEditOverlay, type ShapeTransformInfo } from '../utils/svg';
import type { RendererApi } from './useRenderer';

interface Caret { paraIdx: number; paraOffset: number; }

interface Options {
  containerRef: React.RefObject<HTMLDivElement | null>;
  renderer: RendererApi;
  slide: number;
  shapeIdx: number | null;       // active text shape, or null
  commit: (result: string) => void; // re-render slide SVG after an edit
  onExit: () => void;
}

export function useInlineText({ containerRef, renderer, slide, shapeIdx, commit, onExit }: Options) {
  const caretRef = useRef<Caret | null>(null);

  const svgEl = useCallback(() => containerRef.current?.querySelector('svg') as SVGSVGElement | null, [containerRef]);

  const clearOverlays = useCallback(() => {
    containerRef.current?.querySelectorAll('.edit-overlay, .text-caret').forEach(el => el.remove());
  }, [containerRef]);

  const drawCaret = useCallback(() => {
    const container = containerRef.current, svg = svgEl();
    if (!container || !svg || shapeIdx === null) return;
    container.querySelectorAll('.text-caret').forEach(el => el.remove());
    const layout = renderer.getTextLayout(slide, shapeIdx);
    const caret = caretRef.current;
    if (!layout || !caret) return;
    let cx: number | null = null, cy = 0, ch = 0;
    for (const line of layout.lines) {
      for (const run of line.runs) {
        if (run.paraIdx !== caret.paraIdx) continue;
        for (let i = 0; i < run.chars.length; i++) {
          if (run.paraCharStart + i === caret.paraOffset) { cx = run.chars[i].x; cy = line.y; ch = line.h; }
        }
        if (run.paraCharStart + run.chars.length === caret.paraOffset) { cx = run.x + run.w; cy = line.y; ch = line.h; }
      }
      if (cx === null && line.runs.length === 0 && line.paraIdx === caret.paraIdx) { cx = layout.box.x; cy = line.y; ch = line.h; }
    }
    if (cx === null) { const f = layout.lines[0]; if (!f) return; cx = layout.box.x; cy = f.y; ch = f.h; }
    const p = emuToContainerPx(svg, container, cx, cy);
    const el = document.createElement('div');
    el.className = 'text-caret';
    el.style.left = `${p.left}px`; el.style.top = `${p.top}px`; el.style.height = `${ch / p.epp}px`;
    container.appendChild(el);
  }, [containerRef, svgEl, renderer, slide, shapeIdx]);

  // Draw the dashed edit outline around the shape + the caret (no per-char boxes).
  const drawLayout = useCallback(() => {
    const container = containerRef.current, svg = svgEl();
    if (!container || !svg || shapeIdx === null) return;
    clearOverlays();
    const g = findShapeElement(svg, shapeIdx);
    if (g) showEditOverlay(container, g);
    drawCaret();
  }, [containerRef, svgEl, shapeIdx, clearOverlays, drawCaret]);

  // Redraw geometry whenever inline editing turns on / the slide re-renders.
  useEffect(() => {
    if (shapeIdx === null) { clearOverlays(); caretRef.current = null; return; }
    caretRef.current = { paraIdx: 0, paraOffset: 0 };
    // Defer so the SVG is in the DOM.
    const id = requestAnimationFrame(drawLayout);
    return () => cancelAnimationFrame(id);
  }, [shapeIdx, slide, drawLayout, clearOverlays]);

  // Click inside the editing shape → caret position. (Clicks elsewhere are
  // handled by the parent, which exits inline mode.)
  useEffect(() => {
    const container = containerRef.current;
    if (!container || shapeIdx === null) return;
    const onClick = (e: MouseEvent) => {
      const g = (e.target as Element).closest('g[data-ooxml-shape-idx]');
      if (g?.getAttribute('data-ooxml-shape-idx') !== String(shapeIdx)) return;
      const svg = svgEl();
      if (!svg) return;
      const { x, y } = clientToEmu(svg, e.clientX, e.clientY);
      const hit = renderer.hitTestText(slide, shapeIdx, x, y);
      if (!hit) return;
      caretRef.current = { paraIdx: hit.paraIdx, paraOffset: hit.paraOffset };
      drawCaret();
    };
    container.addEventListener('click', onClick);
    return () => container.removeEventListener('click', onClick);
  }, [containerRef, svgEl, renderer, slide, shapeIdx, drawCaret]);

  // Move the caret with arrow keys (no edit, just reposition).
  const moveCaret = useCallback((dir: 'left' | 'right' | 'up' | 'down') => {
    if (shapeIdx === null) return;
    const layout = renderer.getTextLayout(slide, shapeIdx);
    const caret = caretRef.current;
    if (!layout || !caret) return;
    // paragraph lengths + max paragraph index
    const paraLen = new Map<number, number>();
    let maxPara = 0;
    for (const line of layout.lines) {
      maxPara = Math.max(maxPara, line.paraIdx);
      if (line.runs.length === 0) paraLen.set(line.paraIdx, paraLen.get(line.paraIdx) ?? 0);
      for (const run of line.runs) {
        paraLen.set(run.paraIdx, Math.max(paraLen.get(run.paraIdx) ?? 0, run.paraCharStart + run.chars.length));
      }
    }
    const lenOf = (p: number) => paraLen.get(p) ?? 0;

    if (dir === 'left') {
      if (caret.paraOffset > 0) caretRef.current = { paraIdx: caret.paraIdx, paraOffset: caret.paraOffset - 1 };
      else if (caret.paraIdx > 0) caretRef.current = { paraIdx: caret.paraIdx - 1, paraOffset: lenOf(caret.paraIdx - 1) };
    } else if (dir === 'right') {
      if (caret.paraOffset < lenOf(caret.paraIdx)) caretRef.current = { paraIdx: caret.paraIdx, paraOffset: caret.paraOffset + 1 };
      else if (caret.paraIdx < maxPara) caretRef.current = { paraIdx: caret.paraIdx + 1, paraOffset: 0 };
    } else {
      // up / down: find the current line, then the nearest char on the line above/below.
      const lines = layout.lines.map((line, li) => {
        // collect caret-able x positions on this line
        const spots: Array<{ x: number; paraIdx: number; paraOffset: number }> = [];
        for (const run of line.runs) {
          for (let i = 0; i < run.chars.length; i++) spots.push({ x: run.chars[i].x, paraIdx: run.paraIdx, paraOffset: run.paraCharStart + i });
          spots.push({ x: run.x + run.w, paraIdx: run.paraIdx, paraOffset: run.paraCharStart + run.chars.length });
        }
        if (spots.length === 0) spots.push({ x: layout.box.x, paraIdx: line.paraIdx, paraOffset: 0 });
        return { li, y: line.y, spots };
      });
      // current caret x
      let curX = layout.box.x, curLine = 0;
      for (const L of lines) {
        const s = L.spots.find(sp => sp.paraIdx === caret.paraIdx && sp.paraOffset === caret.paraOffset);
        if (s) { curX = s.x; curLine = L.li; break; }
      }
      const target = lines.find(L => dir === 'up' ? L.li === curLine - 1 : L.li === curLine + 1);
      if (target) {
        let best = target.spots[0];
        for (const sp of target.spots) if (Math.abs(sp.x - curX) < Math.abs(best.x - curX)) best = sp;
        caretRef.current = { paraIdx: best.paraIdx, paraOffset: best.paraOffset };
      }
    }
    drawCaret();
  }, [renderer, slide, shapeIdx, drawCaret]);

  // Keyboard typing.
  useEffect(() => {
    if (shapeIdx === null) return;
    const replace = (s: number, e: number, text: string): boolean => {
      const c = caretRef.current; if (!c) return false;
      const res = renderer.replaceTextRange(slide, shapeIdx, c.paraIdx, s, c.paraIdx, e, text);
      if (res.startsWith('ERROR')) return false;
      commit(res);
      return true;
    };
    const onKey = (e: KeyboardEvent) => {
      const tag = (e.target as HTMLElement)?.tagName?.toUpperCase();
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
      const c = caretRef.current; if (!c) return;
      if (e.key === 'Escape') { e.preventDefault(); onExit(); return; }
      if (e.key === 'ArrowLeft') { e.preventDefault(); moveCaret('left'); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); moveCaret('right'); return; }
      if (e.key === 'ArrowUp') { e.preventDefault(); moveCaret('up'); return; }
      if (e.key === 'ArrowDown') { e.preventDefault(); moveCaret('down'); return; }
      if (e.key === 'Backspace') {
        e.preventDefault();
        if (c.paraOffset > 0 && replace(c.paraOffset - 1, c.paraOffset, '')) { caretRef.current = { paraIdx: c.paraIdx, paraOffset: c.paraOffset - 1 }; requestAnimationFrame(drawLayout); }
        return;
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        if (replace(c.paraOffset, c.paraOffset, '\n')) { caretRef.current = { paraIdx: c.paraIdx, paraOffset: c.paraOffset + 1 }; requestAnimationFrame(drawLayout); }
        return;
      }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        if (replace(c.paraOffset, c.paraOffset, e.key)) { caretRef.current = { paraIdx: c.paraIdx, paraOffset: c.paraOffset + 1 }; requestAnimationFrame(drawLayout); }
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [renderer, slide, shapeIdx, commit, onExit, drawLayout, moveCaret]);
}

// Re-export so App can reference the transform type without importing svg.ts twice.
export type { ShapeTransformInfo };
