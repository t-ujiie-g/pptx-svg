/**
 * useDrag — drag-to-move and resize-handles interaction logic.
 *
 * Attaches mousedown to the container, and mousemove/mouseup to window.
 * On mouseup, calls the renderer's updateShapeTransform() and refreshes.
 */

import { useEffect, useRef } from 'react';
import { findShapeElement, getShapeTransform, emuToPx } from 'pptx-svg';
import { getEmuPerCssPx, type HandlePos } from '../utils/svg';

const MIN_SHAPE_EMU = 50000;
const MIN_OVERLAY_PX = 10;

// ── Drag state types ──

interface DragBase {
  shapeIdx: number;
  startX: number; startY: number;
  origEmuX: number; origEmuY: number;
  origEmuCx: number; origEmuCy: number;
  origRot: number;
  emuPerCssPx: number;
}

interface DragMove extends DragBase {
  mode: 'move';
  shape: SVGGElement;
  svgEl: SVGSVGElement;
  origTransform: string;
}

interface DragResize extends DragBase {
  mode: 'resize';
  handle: HandlePos;
  origOverlayLeft: number; origOverlayTop: number;
  origOverlayWidth: number; origOverlayHeight: number;
  newEmuX?: number; newEmuY?: number;
  newEmuCx?: number; newEmuCy?: number;
}

type DragState = DragMove | DragResize;

// ── Hook ──

interface UseDragOptions {
  containerRef: React.RefObject<HTMLDivElement | null>;
  selectedShapeIdx: number;
  slide: number;
  onTransformUpdate: (shapeIdx: number, x: number, y: number, cx: number, cy: number, rot: number) => string;
  onDragEnd: (shapeIdx: number, result: string) => void;
}

export function useDrag({
  containerRef, selectedShapeIdx, slide,
  onTransformUpdate, onDragEnd,
}: UseDragOptions) {
  const dragRef = useRef<DragState | null>(null);

  // ── Mousedown: start move or resize ──
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleMouseDown = (e: MouseEvent) => {
      const target = e.target as Element;
      const svgEl = container.querySelector('svg') as SVGSVGElement | null;

      // Resize handle click
      if (target.classList.contains('resize-handle')) {
        if (selectedShapeIdx < 0 || !svgEl) return;
        const g = findShapeElement(svgEl, selectedShapeIdx);
        if (!g) return;
        const t = getShapeTransform(g);
        e.preventDefault(); e.stopPropagation();
        const overlay = container.querySelector('.selection-overlay') as HTMLElement | null;
        dragRef.current = {
          mode: 'resize',
          handle: (target as HTMLElement).dataset.handle as HandlePos,
          shapeIdx: selectedShapeIdx,
          startX: e.clientX, startY: e.clientY,
          origEmuX: t.x, origEmuY: t.y, origEmuCx: t.cx, origEmuCy: t.cy, origRot: t.rot,
          emuPerCssPx: getEmuPerCssPx(svgEl),
          origOverlayLeft: overlay ? parseFloat(overlay.style.left) : 0,
          origOverlayTop: overlay ? parseFloat(overlay.style.top) : 0,
          origOverlayWidth: overlay ? parseFloat(overlay.style.width) : 0,
          origOverlayHeight: overlay ? parseFloat(overlay.style.height) : 0,
        };
        return;
      }

      // Move: only if clicking the selected shape
      if (selectedShapeIdx < 0 || !svgEl) return;
      const clicked = target.closest('g[data-ooxml-shape-idx]') as SVGGElement | null;
      if (!clicked) return;
      const idx = parseInt(clicked.getAttribute('data-ooxml-shape-idx') ?? '-1', 10);
      if (idx !== selectedShapeIdx) return;

      const t = getShapeTransform(clicked);
      e.preventDefault();
      dragRef.current = {
        mode: 'move', shape: clicked, shapeIdx: idx, svgEl,
        startX: e.clientX, startY: e.clientY,
        origTransform: clicked.getAttribute('transform') || '',
        origEmuX: t.x, origEmuY: t.y, origEmuCx: t.cx, origEmuCy: t.cy, origRot: t.rot,
        emuPerCssPx: getEmuPerCssPx(svgEl),
      };
      container.style.cursor = 'grabbing';
    };

    container.addEventListener('mousedown', handleMouseDown);
    return () => container.removeEventListener('mousedown', handleMouseDown);
  }, [containerRef, selectedShapeIdx]);

  // ── Global mousemove/mouseup ──
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      const ds = dragRef.current;
      if (!ds) return;
      const dx = e.clientX - ds.startX;
      const dy = e.clientY - ds.startY;

      if (ds.mode === 'move') {
        // Visual feedback: translate the SVG group + overlay
        const rect = ds.svgEl.getBoundingClientRect();
        const vb = ds.svgEl.viewBox.baseVal;
        const r = vb.width / rect.width;
        ds.shape.setAttribute('transform', `translate(${dx * r},${dy * r}) ${ds.origTransform}`);

        const overlay = containerRef.current?.querySelector('.selection-overlay') as HTMLElement | null;
        if (overlay) overlay.style.transform = `translate(${dx}px, ${dy}px)`;
      } else {
        // Resize: compute new EMU dimensions
        const epp = ds.emuPerCssPx;
        const h = ds.handle;
        let nx = ds.origEmuX, ny = ds.origEmuY, ncx = ds.origEmuCx, ncy = ds.origEmuCy;
        const de = Math.round(dx * epp), df = Math.round(dy * epp);
        if (h.includes('e')) ncx += de;
        if (h.includes('w')) { nx += de; ncx -= de; }
        if (h.includes('s')) ncy += df;
        if (h.includes('n')) { ny += df; ncy -= df; }
        if (ncx < MIN_SHAPE_EMU) ncx = MIN_SHAPE_EMU;
        if (ncy < MIN_SHAPE_EMU) ncy = MIN_SHAPE_EMU;
        ds.newEmuX = nx; ds.newEmuY = ny; ds.newEmuCx = ncx; ds.newEmuCy = ncy;

        // Visual feedback: resize overlay
        const overlay = containerRef.current?.querySelector('.selection-overlay') as HTMLElement | null;
        if (overlay) {
          let ol = ds.origOverlayLeft, ot = ds.origOverlayTop;
          let ow = ds.origOverlayWidth, oh = ds.origOverlayHeight;
          if (h.includes('e')) ow += dx;
          if (h.includes('w')) { ol += dx; ow -= dx; }
          if (h.includes('s')) oh += dy;
          if (h.includes('n')) { ot += dy; oh -= dy; }
          if (ow < MIN_OVERLAY_PX) ow = MIN_OVERLAY_PX;
          if (oh < MIN_OVERLAY_PX) oh = MIN_OVERLAY_PX;
          overlay.style.left = `${ol}px`;
          overlay.style.top = `${ot}px`;
          overlay.style.width = `${ow}px`;
          overlay.style.height = `${oh}px`;
          overlay.style.transform = '';
        }
      }
    };

    const handleMouseUp = (e: MouseEvent) => {
      const ds = dragRef.current;
      if (!ds) return;
      const dx = e.clientX - ds.startX;
      const dy = e.clientY - ds.startY;
      let result: string;

      if (ds.mode === 'move') {
        const epp = ds.emuPerCssPx;
        const nx = ds.origEmuX + Math.round(dx * epp);
        const ny = ds.origEmuY + Math.round(dy * epp);
        result = onTransformUpdate(ds.shapeIdx, nx, ny, ds.origEmuCx, ds.origEmuCy, ds.origRot);
        if (result.startsWith('ERROR:')) {
          ds.shape.setAttribute('transform', ds.origTransform);
        }
      } else {
        const nx = ds.newEmuX ?? ds.origEmuX;
        const ny = ds.newEmuY ?? ds.origEmuY;
        const ncx = ds.newEmuCx ?? ds.origEmuCx;
        const ncy = ds.newEmuCy ?? ds.origEmuCy;
        result = onTransformUpdate(ds.shapeIdx, nx, ny, ncx, ncy, ds.origRot);
      }

      dragRef.current = null;
      if (containerRef.current) containerRef.current.style.cursor = '';
      onDragEnd(ds.shapeIdx, result);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [containerRef, slide, onTransformUpdate, onDragEnd]);
}
