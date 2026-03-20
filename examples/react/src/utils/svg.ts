/**
 * SVG DOM utilities for the pptx-svg React example.
 *
 * Handles SVG insertion, selection overlay, and text run extraction.
 */

import { getShapeTransform, emuToPx } from 'pptx-svg';

// ── Types ──

export interface TextRun {
  pi: number;
  ri: number;
  text: string;
}

export interface ShapeInfo {
  idx: number;
  label: string;
  detail: string;
  fillHex: string;
  textRuns: TextRun[];
}

// ── Constants ──

const HANDLE_POSITIONS = ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se'] as const;
export type HandlePos = typeof HANDLE_POSITIONS[number];

// ── SVG insertion ──

/**
 * Insert an SVG string into a container element.
 * Sets up viewBox for responsive sizing and removes fixed width/height.
 */
export function insertSvgInto(container: HTMLElement, svgString: string) {
  container.querySelector('.selection-overlay')?.remove();

  if (svgString.startsWith('ERROR:')) {
    container.innerHTML = `<span style="color:red;font-family:monospace">${svgString}</span>`;
    return;
  }

  container.innerHTML = svgString;
  const svgEl = container.querySelector('svg');
  if (svgEl) {
    const w = svgEl.getAttribute('width');
    const h = svgEl.getAttribute('height');
    if (w && h && !svgEl.getAttribute('viewBox')) {
      svgEl.setAttribute('viewBox', `0 0 ${w} ${h}`);
    }
    svgEl.removeAttribute('width');
    svgEl.removeAttribute('height');
  }
}

// ── Selection overlay ──

/** Remove the selection overlay from a container. */
export function removeOverlay(container: HTMLElement) {
  container.querySelector('.selection-overlay')?.remove();
}

/** Show a selection overlay with resize handles over a shape element. */
export function showOverlay(container: HTMLElement, shapeG: SVGGElement) {
  removeOverlay(container);

  const shapeRect = shapeG.getBoundingClientRect();
  const cr = container.getBoundingClientRect();
  const ox = shapeRect.left - cr.left - container.clientLeft + container.scrollLeft;
  const oy = shapeRect.top - cr.top - container.clientTop + container.scrollTop;

  const overlay = document.createElement('div');
  overlay.className = 'selection-overlay';
  overlay.style.cssText =
    `position:absolute;pointer-events:none;border:2px solid #4a90d9;z-index:100;` +
    `left:${ox}px;top:${oy}px;width:${shapeRect.width}px;height:${shapeRect.height}px`;

  for (const pos of HANDLE_POSITIONS) {
    const h = document.createElement('div');
    h.className = `resize-handle ${pos}`;
    h.dataset.handle = pos;
    overlay.appendChild(h);
  }

  container.appendChild(overlay);
}

// ── EMU-per-CSS-pixel ──

/** Calculate the EMU-per-CSS-pixel ratio from the SVG element. */
export function getEmuPerCssPx(svgEl: SVGSVGElement | null): number {
  if (!svgEl) return 9525;
  const rect = svgEl.getBoundingClientRect();
  const slideCx = parseInt(svgEl.getAttribute('data-ooxml-slide-cx') || '9144000', 10);
  return slideCx / rect.width;
}

// ── Shape info extraction ──

/** Extract shape info (label, detail, fill, text runs) from a shape `<g>` element. */
export function extractShapeInfo(shapeG: SVGGElement): ShapeInfo {
  const idx = parseInt(shapeG.getAttribute('data-ooxml-shape-idx') ?? '-1', 10);
  const fillHex = shapeG.getAttribute('data-ooxml-fill') || '';

  const t = getShapeTransform(shapeG);
  const shapeType = shapeG.getAttribute('data-ooxml-shape-type') || '?';
  const geom = shapeG.getAttribute('data-ooxml-geom') || '';

  return {
    idx,
    label: `Shape #${idx} (${shapeType}${geom ? '/' + geom : ''})`,
    detail: `x=${emuToPx(t.x)}px y=${emuToPx(t.y)}px w=${emuToPx(t.cx)}px h=${emuToPx(t.cy)}px rot=${t.rot}`,
    fillHex: fillHex.length === 6 ? fillHex : '',
    textRuns: extractTextRuns(shapeG),
  };
}

/** Extract text runs from a shape, merging wrapped tspan fragments. */
function extractTextRuns(shapeG: SVGGElement): TextRun[] {
  const runTspans = shapeG.querySelectorAll('tspan[data-ooxml-run-idx]');
  if (runTspans.length === 0) return [];

  const seen = new Map<string, number>();
  const runs: TextRun[] = [];

  for (const ts of runTspans) {
    const ri = ts.getAttribute('data-ooxml-run-idx');
    const paraTspan = ts.closest('tspan[data-ooxml-para-idx]');
    const pi = paraTspan ? paraTspan.getAttribute('data-ooxml-para-idx') : null;
    if (pi === null || ri === null) continue;

    const key = `${pi}:${ri}`;
    if (seen.has(key)) {
      runs[seen.get(key)!].text += ts.textContent || '';
    } else {
      seen.set(key, runs.length);
      runs.push({ pi: parseInt(pi), ri: parseInt(ri), text: ts.textContent || '' });
    }
  }

  return runs;
}
