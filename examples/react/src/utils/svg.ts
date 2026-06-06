/**
 * SVG DOM utilities for the pptx-svg React editor.
 *
 * The slide SVG is inserted into a container via innerHTML; selection,
 * resize/rotate handles, multi-select, marquee, and inline-text overlays are
 * absolutely-positioned <div>s on top of it. All geometry maps between
 * slide-absolute EMU and container CSS pixels.
 */

import { getShapeTransform, emuToPx, ooxmlToDegrees } from 'pptx-svg';

// ── Color ──
export function hexToRgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16)];
}
export function rgbToHex(r: number, g: number, b: number): string {
  const h = (n: number) => n.toString(16).padStart(2, '0');
  return `#${h(r)}${h(g)}${h(b)}`;
}

/** Parse the shape index from an `"OK:<idx>"` result string (else `fallback`). */
export function okIndex(result: string, fallback = -1): number {
  const n = parseInt(result.split(':')[1] ?? '', 10);
  return Number.isNaN(n) ? fallback : n;
}

// ── Types ──
export interface TextRun {
  pi: number; ri: number; text: string;
  bold: boolean; italic: boolean; fontSize: number; color: string;
  font: string; eaFont: string; csFont: string;
  underline: string; strike: string; baseline: number;
}
export interface ParagraphInfo { pi: number; align: string; runs: TextRun[]; }
export interface ShapeTransformInfo { x: number; y: number; cx: number; cy: number; rot: number; }
export interface ShapeInfo {
  idx: number;
  label: string;
  shapeType: string;   // autoshape | picture | table | group | chart | …
  geom: string;
  fillHex: string;
  t: ShapeTransformInfo;
  paragraphs: ParagraphInfo[];
}

const HANDLE_POSITIONS = ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se'] as const;
export type HandlePos = typeof HANDLE_POSITIONS[number];

// ── SVG insertion ──
export function insertSvgInto(container: HTMLElement, svgString: string) {
  clearOverlays(container);
  if (svgString.startsWith('ERROR:')) {
    container.innerHTML = `<span style="color:#d33;font-family:monospace;font-size:12px">${svgString}</span>`;
    return;
  }
  container.innerHTML = svgString;
  const svgEl = container.querySelector('svg');
  if (svgEl) {
    const w = svgEl.getAttribute('width'), h = svgEl.getAttribute('height');
    if (w && h && !svgEl.getAttribute('viewBox')) svgEl.setAttribute('viewBox', `0 0 ${w} ${h}`);
    svgEl.removeAttribute('width');
    svgEl.removeAttribute('height');
  }
}

// ── Overlays ──
export function clearOverlays(container: HTMLElement) {
  container.querySelectorAll('.selection-overlay, .multi-overlay, .edit-overlay, .text-caret').forEach(el => el.remove());
}

/** Dashed outline around a shape while its text is being edited inline (no handles). */
export function showEditOverlay(container: HTMLElement, g: SVGGElement) {
  container.querySelectorAll('.edit-overlay').forEach(el => el.remove());
  const b = shapeBox(container, g);
  const o = document.createElement('div');
  o.className = 'edit-overlay';
  o.style.left = `${b.left}px`; o.style.top = `${b.top}px`;
  o.style.width = `${b.width}px`; o.style.height = `${b.height}px`;
  container.appendChild(o);
}
export function removeOverlay(container: HTMLElement) {
  container.querySelector('.selection-overlay')?.remove();
}
export function removeMultiOverlays(container: HTMLElement) {
  container.querySelectorAll('.multi-overlay').forEach(el => el.remove());
}

/** Box of a shape <g> in container-local CSS px. */
function shapeBox(container: HTMLElement, g: SVGGElement) {
  const r = g.getBoundingClientRect();
  const cr = container.getBoundingClientRect();
  return {
    left: r.left - cr.left + container.scrollLeft,
    top: r.top - cr.top + container.scrollTop,
    width: r.width, height: r.height,
  };
}

export function showOverlay(container: HTMLElement, g: SVGGElement, withRotate = true) {
  removeOverlay(container);
  const b = shapeBox(container, g);
  const overlay = document.createElement('div');
  overlay.className = 'selection-overlay';
  overlay.style.left = `${b.left}px`;
  overlay.style.top = `${b.top}px`;
  overlay.style.width = `${b.width}px`;
  overlay.style.height = `${b.height}px`;
  for (const pos of HANDLE_POSITIONS) {
    const h = document.createElement('div');
    h.className = `resize-handle ${pos}`;
    h.dataset.handle = pos;
    overlay.appendChild(h);
  }
  if (withRotate) {
    const rot = document.createElement('div');
    rot.className = 'rotate-handle';
    rot.dataset.handle = 'rot';
    overlay.appendChild(rot);
  }
  container.appendChild(overlay);
}

export function showMultiOverlay(container: HTMLElement, gs: SVGGElement[]) {
  removeMultiOverlays(container);
  for (const g of gs) {
    const b = shapeBox(container, g);
    const o = document.createElement('div');
    o.className = 'multi-overlay';
    o.style.left = `${b.left}px`; o.style.top = `${b.top}px`;
    o.style.width = `${b.width}px`; o.style.height = `${b.height}px`;
    container.appendChild(o);
  }
}

// ── EMU ↔ container px ──
export function getEmuPerCssPx(svgEl: SVGSVGElement | null): number {
  if (!svgEl) return 9525;
  const rect = svgEl.getBoundingClientRect();
  const slideCx = parseInt(svgEl.getAttribute('data-ooxml-slide-cx') || '9144000', 10);
  return slideCx / rect.width;
}

/** Client (mouse) point → slide-absolute EMU. */
export function clientToEmu(svgEl: SVGSVGElement, clientX: number, clientY: number) {
  const sr = svgEl.getBoundingClientRect();
  const epp = getEmuPerCssPx(svgEl);
  return { x: Math.round((clientX - sr.left) * epp), y: Math.round((clientY - sr.top) * epp) };
}

/** Slide-absolute EMU → container-local CSS px (for drawing overlays). */
export function emuToContainerPx(svgEl: SVGSVGElement, container: HTMLElement, ex: number, ey: number) {
  const sr = svgEl.getBoundingClientRect();
  const cr = container.getBoundingClientRect();
  const epp = getEmuPerCssPx(svgEl);
  return {
    left: sr.left - cr.left + container.scrollLeft + ex / epp,
    top: sr.top - cr.top + container.scrollTop + ey / epp,
    epp,
  };
}

// ── Shape info extraction ──
export function extractShapeInfo(g: SVGGElement): ShapeInfo {
  const idx = parseInt(g.getAttribute('data-ooxml-shape-idx') ?? '-1', 10);
  const fillHex = g.getAttribute('data-ooxml-fill') || '';
  const shapeType = g.getAttribute('data-ooxml-shape-type') || 'shape';
  const geom = g.getAttribute('data-ooxml-geom') || '';
  const t = getShapeTransform(g);
  const typeLabel = geom ? `${shapeType}/${geom}` : shapeType;
  return {
    idx,
    label: `#${idx} · ${typeLabel}`,
    shapeType,
    geom,
    fillHex: fillHex.length === 6 ? fillHex : '',
    t: { x: t.x, y: t.y, cx: t.cx, cy: t.cy, rot: t.rot },
    paragraphs: extractParagraphs(g),
  };
}

/** Number of distinct shapes on the current SVG (for hit-testing helpers). */
export function describeTransform(t: ShapeTransformInfo): string {
  return `x ${emuToPx(t.x)} · y ${emuToPx(t.y)} · w ${emuToPx(t.cx)} · h ${emuToPx(t.cy)} · ${Math.round(ooxmlToDegrees(t.rot))}°`;
}

function extractParagraphs(g: SVGGElement): ParagraphInfo[] {
  const paraTspans = g.querySelectorAll('tspan[data-ooxml-para-idx]');
  const map = new Map<number, ParagraphInfo>();
  for (const pts of paraTspans) {
    const pi = parseInt(pts.getAttribute('data-ooxml-para-idx')!);
    if (!map.has(pi)) map.set(pi, { pi, align: pts.getAttribute('data-ooxml-para-align') || 'l', runs: [] });
    const para = map.get(pi)!;
    for (const rts of pts.querySelectorAll('tspan[data-ooxml-run-idx]')) {
      const ri = parseInt(rts.getAttribute('data-ooxml-run-idx')!);
      const existing = para.runs.find(r => r.ri === ri);
      if (existing) { existing.text += rts.textContent || ''; continue; }
      para.runs.push({
        pi, ri, text: rts.textContent || '',
        bold: rts.getAttribute('data-ooxml-bold') === 'true',
        italic: rts.getAttribute('font-style') === 'italic',
        fontSize: parseInt(rts.getAttribute('data-ooxml-font-size') || '0'),
        color: rts.getAttribute('data-ooxml-color') || '',
        font: rts.getAttribute('data-ooxml-run-font') || '',
        eaFont: rts.getAttribute('data-ooxml-ea-font') || '',
        csFont: rts.getAttribute('data-ooxml-cs-font') || '',
        underline: rts.getAttribute('data-ooxml-underline') || '',
        strike: rts.getAttribute('data-ooxml-strike') || '',
        baseline: parseInt(rts.getAttribute('data-ooxml-baseline') || '0'),
      });
    }
  }
  return Array.from(map.values()).sort((a, b) => a.pi - b.pi);
}
