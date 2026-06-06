/**
 * useDrag — pointer interactions on the canvas:
 *   • move a single selected shape
 *   • move a multi-selection together (one undo step via onGroupTransform)
 *   • resize via the 8 handles
 *   • rotate via the top handle (Shift = 15° snap)
 *
 * Visual feedback is applied live to the SVG group / overlay; the actual
 * transform is committed on mouseup.
 */

import { useEffect, useRef } from 'react';
import { findShapeElement, getShapeTransform, degreesToOoxml, ooxmlToDegrees } from 'pptx-svg';
import { getEmuPerCssPx, type HandlePos } from '../utils/svg';

const MIN_SHAPE_EMU = 50000;
const MIN_OVERLAY_PX = 10;

interface GroupItem { g: SVGGElement; idx: number; origTransform: string; t: { x: number; y: number; cx: number; cy: number; rot: number }; }

type DragState =
  | { mode: 'move'; shape: SVGGElement; idx: number; svgEl: SVGSVGElement; startX: number; startY: number; origTransform: string; t: { x: number; y: number; cx: number; cy: number; rot: number }; epp: number }
  | { mode: 'group'; items: GroupItem[]; svgEl: SVGSVGElement; startX: number; startY: number; epp: number }
  | { mode: 'resize'; idx: number; handle: HandlePos; startX: number; startY: number; t: { x: number; y: number; cx: number; cy: number; rot: number }; epp: number; ov: { l: number; t: number; w: number; h: number }; nx?: number; ny?: number; ncx?: number; ncy?: number }
  | { mode: 'rotate'; idx: number; cx: number; cy: number; startAngle: number; t: { x: number; y: number; cx: number; cy: number; rot: number }; overlay: HTMLElement | null; curRot: number };

interface UseDragOptions {
  containerRef: React.RefObject<HTMLDivElement | null>;
  selectedShapeIdx: number;
  multiSel: number[];
  slide: number;
  onTransformUpdate: (idx: number, x: number, y: number, cx: number, cy: number, rot: number) => string;
  onGroupTransform: (items: Array<{ shapeIdx: number; x: number; y: number; cx: number; cy: number; rot: number }>) => string;
  onDragEnd: (affected: number[], result: string) => void;
}

function angleDeg(cx: number, cy: number, px: number, py: number): number {
  return Math.atan2(py - cy, px - cx) * 180 / Math.PI;
}

export function useDrag({
  containerRef, selectedShapeIdx, multiSel, slide,
  onTransformUpdate, onGroupTransform, onDragEnd,
}: UseDragOptions) {
  const dragRef = useRef<DragState | null>(null);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const onMouseDown = (e: MouseEvent) => {
      const target = e.target as Element;
      const svgEl = container.querySelector('svg') as SVGSVGElement | null;
      if (!svgEl) return;
      const epp = getEmuPerCssPx(svgEl);

      // Rotate handle
      if (target.classList.contains('rotate-handle')) {
        if (selectedShapeIdx < 0) return;
        const g = findShapeElement(svgEl, selectedShapeIdx);
        if (!g) return;
        e.preventDefault(); e.stopPropagation();
        const r = g.getBoundingClientRect();
        const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
        const t = getShapeTransform(g);
        const overlay = container.querySelector('.selection-overlay') as HTMLElement | null;
        if (overlay) overlay.style.transformOrigin = 'center center';
        dragRef.current = { mode: 'rotate', idx: selectedShapeIdx, cx, cy, startAngle: angleDeg(cx, cy, e.clientX, e.clientY), t, overlay, curRot: t.rot };
        return;
      }

      // Resize handle
      if (target.classList.contains('resize-handle')) {
        if (selectedShapeIdx < 0) return;
        const g = findShapeElement(svgEl, selectedShapeIdx);
        if (!g) return;
        e.preventDefault(); e.stopPropagation();
        const t = getShapeTransform(g);
        const overlay = container.querySelector('.selection-overlay') as HTMLElement | null;
        dragRef.current = {
          mode: 'resize', idx: selectedShapeIdx, handle: (target as HTMLElement).dataset.handle as HandlePos,
          startX: e.clientX, startY: e.clientY, t, epp,
          ov: { l: overlay ? parseFloat(overlay.style.left) : 0, t: overlay ? parseFloat(overlay.style.top) : 0, w: overlay ? parseFloat(overlay.style.width) : 0, h: overlay ? parseFloat(overlay.style.height) : 0 },
        };
        return;
      }

      // Move: clicking the selected shape (or any shape within the multi-selection)
      const clicked = target.closest('g[data-ooxml-shape-idx]') as SVGGElement | null;
      if (!clicked) return;
      const idx = parseInt(clicked.getAttribute('data-ooxml-shape-idx') ?? '-1', 10);

      if (multiSel.length > 1 && multiSel.includes(idx)) {
        // Group move
        const items: GroupItem[] = [];
        for (const si of multiSel) {
          const g = findShapeElement(svgEl, si);
          if (!g) continue;
          items.push({ g, idx: si, origTransform: g.getAttribute('transform') || '', t: getShapeTransform(g) });
        }
        if (items.length === 0) return;
        e.preventDefault();
        dragRef.current = { mode: 'group', items, svgEl, startX: e.clientX, startY: e.clientY, epp };
        container.style.cursor = 'grabbing';
        return;
      }

      if (idx !== selectedShapeIdx) return;
      e.preventDefault();
      dragRef.current = {
        mode: 'move', shape: clicked, idx, svgEl, startX: e.clientX, startY: e.clientY,
        origTransform: clicked.getAttribute('transform') || '', t: getShapeTransform(clicked), epp,
      };
      container.style.cursor = 'grabbing';
    };

    container.addEventListener('mousedown', onMouseDown);
    return () => container.removeEventListener('mousedown', onMouseDown);
  }, [containerRef, selectedShapeIdx, multiSel]);

  useEffect(() => {
    const container = containerRef.current;

    const onMouseMove = (e: MouseEvent) => {
      const ds = dragRef.current;
      if (!ds || !container) return;

      if (ds.mode === 'move' || ds.mode === 'group') {
        const dx = e.clientX - ds.startX, dy = e.clientY - ds.startY;
        const vb = ds.svgEl.viewBox.baseVal;
        const ratio = vb.width / ds.svgEl.getBoundingClientRect().width;
        if (ds.mode === 'move') {
          ds.shape.setAttribute('transform', `translate(${dx * ratio},${dy * ratio}) ${ds.origTransform}`);
          const ov = container.querySelector('.selection-overlay') as HTMLElement | null;
          if (ov) ov.style.transform = `translate(${dx}px,${dy}px)`;
        } else {
          for (const it of ds.items) it.g.setAttribute('transform', `translate(${dx * ratio},${dy * ratio}) ${it.origTransform}`);
          container.querySelectorAll('.multi-overlay').forEach(o => ((o as HTMLElement).style.transform = `translate(${dx}px,${dy}px)`));
        }
      } else if (ds.mode === 'resize') {
        const h = ds.handle, de = Math.round((e.clientX - ds.startX) * ds.epp), df = Math.round((e.clientY - ds.startY) * ds.epp);
        let nx = ds.t.x, ny = ds.t.y, ncx = ds.t.cx, ncy = ds.t.cy;
        if (h.includes('e')) ncx += de;
        if (h.includes('w')) { nx += de; ncx -= de; }
        if (h.includes('s')) ncy += df;
        if (h.includes('n')) { ny += df; ncy -= df; }
        ncx = Math.max(ncx, MIN_SHAPE_EMU); ncy = Math.max(ncy, MIN_SHAPE_EMU);
        ds.nx = nx; ds.ny = ny; ds.ncx = ncx; ds.ncy = ncy;
        const ov = container.querySelector('.selection-overlay') as HTMLElement | null;
        if (ov) {
          const dxp = e.clientX - ds.startX, dyp = e.clientY - ds.startY;
          let l = ds.ov.l, t = ds.ov.t, w = ds.ov.w, hh = ds.ov.h;
          if (h.includes('e')) w += dxp;
          if (h.includes('w')) { l += dxp; w -= dxp; }
          if (h.includes('s')) hh += dyp;
          if (h.includes('n')) { t += dyp; hh -= dyp; }
          ov.style.left = `${l}px`; ov.style.top = `${t}px`;
          ov.style.width = `${Math.max(w, MIN_OVERLAY_PX)}px`; ov.style.height = `${Math.max(hh, MIN_OVERLAY_PX)}px`;
          ov.style.transform = '';
        }
      } else if (ds.mode === 'rotate') {
        const cur = angleDeg(ds.cx, ds.cy, e.clientX, e.clientY);
        const deltaDeg = cur - ds.startAngle;
        const origDeg = ooxmlToDegrees(ds.t.rot);
        let deg = origDeg + deltaDeg;
        if (e.shiftKey) deg = Math.round(deg / 15) * 15;
        deg = ((deg % 360) + 360) % 360;
        ds.curRot = degreesToOoxml(deg);
        if (ds.overlay) ds.overlay.style.transform = `rotate(${deg - origDeg}deg)`;
      }
    };

    const onMouseUp = (e: MouseEvent) => {
      const ds = dragRef.current;
      if (!ds) return;
      dragRef.current = null;
      if (container) container.style.cursor = '';
      let result = 'OK', affected: number[] = [];

      if (ds.mode === 'move') {
        const dx = e.clientX - ds.startX, dy = e.clientY - ds.startY;
        if (Math.abs(dx) < 2 && Math.abs(dy) < 2) { onDragEnd([], 'OK'); return; } // click, not drag
        const nx = ds.t.x + Math.round(dx * ds.epp), ny = ds.t.y + Math.round(dy * ds.epp);
        result = onTransformUpdate(ds.idx, nx, ny, ds.t.cx, ds.t.cy, ds.t.rot);
        affected = [ds.idx];
        if (result.startsWith('ERROR:')) ds.shape.setAttribute('transform', ds.origTransform);
      } else if (ds.mode === 'group') {
        const dx = e.clientX - ds.startX, dy = e.clientY - ds.startY;
        if (Math.abs(dx) < 2 && Math.abs(dy) < 2) { onDragEnd([], 'OK'); return; }
        const items = ds.items.map(it => ({ shapeIdx: it.idx, x: it.t.x + Math.round(dx * ds.epp), y: it.t.y + Math.round(dy * ds.epp), cx: it.t.cx, cy: it.t.cy, rot: it.t.rot }));
        result = onGroupTransform(items);
        affected = ds.items.map(it => it.idx);
      } else if (ds.mode === 'resize') {
        result = onTransformUpdate(ds.idx, ds.nx ?? ds.t.x, ds.ny ?? ds.t.y, ds.ncx ?? ds.t.cx, ds.ncy ?? ds.t.cy, ds.t.rot);
        affected = [ds.idx];
      } else if (ds.mode === 'rotate') {
        result = onTransformUpdate(ds.idx, ds.t.x, ds.t.y, ds.t.cx, ds.t.cy, ds.curRot);
        affected = [ds.idx];
      }
      onDragEnd(affected, result);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
    return () => { window.removeEventListener('mousemove', onMouseMove); window.removeEventListener('mouseup', onMouseUp); };
  }, [containerRef, slide, onTransformUpdate, onGroupTransform, onDragEnd]);
}
