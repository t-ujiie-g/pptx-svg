/**
 * SVG DOM utilities for the pptx-svg React example.
 *
 * Handles SVG insertion, selection overlay, and text run extraction
 * including formatting attributes (bold, italic, font size, color, etc).
 */

import { getShapeTransform, emuToPx } from 'pptx-svg';

// ── Color ──

/** Parse '#rrggbb' hex string to [r, g, b] tuple. */
export function hexToRgb(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ];
}

// ── Types ──

export interface TextRun {
  pi: number;
  ri: number;
  text: string;
  bold: boolean;
  italic: boolean;
  fontSize: number;    // hundredths of a point (0 = inherit)
  color: string;       // hex 6-char (no #), '' = inherit
  font: string;        // Latin font name
  eaFont: string;      // East Asian font name
  csFont: string;      // Complex Script font name
  underline: string;   // 'sng', 'dbl', '' = none
  strike: string;      // 'sngStrike', 'dblStrike', '' = none
  baseline: number;    // 30000 = super, -25000 = sub, 0 = normal
}

export interface ParagraphInfo {
  pi: number;
  align: string;
  runs: TextRun[];
}

export interface ShapeInfo {
  idx: number;
  label: string;
  detail: string;
  fillHex: string;
  shapeType: string;
  paragraphs: ParagraphInfo[];
}

// ── Constants ──

const HANDLE_POSITIONS = ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se'] as const;
export type HandlePos = typeof HANDLE_POSITIONS[number];

// ── SVG insertion ──

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

export function removeOverlay(container: HTMLElement) {
  container.querySelector('.selection-overlay')?.remove();
}

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

export function getEmuPerCssPx(svgEl: SVGSVGElement | null): number {
  if (!svgEl) return 9525;
  const rect = svgEl.getBoundingClientRect();
  const slideCx = parseInt(svgEl.getAttribute('data-ooxml-slide-cx') || '9144000', 10);
  return slideCx / rect.width;
}

// ── Shape info extraction ──

export function extractShapeInfo(shapeG: SVGGElement): ShapeInfo {
  const idx = parseInt(shapeG.getAttribute('data-ooxml-shape-idx') ?? '-1', 10);
  const fillHex = shapeG.getAttribute('data-ooxml-fill') || '';
  const shapeType = shapeG.getAttribute('data-ooxml-shape-type') || '?';
  const geom = shapeG.getAttribute('data-ooxml-geom') || '';

  const t = getShapeTransform(shapeG);

  return {
    idx,
    label: `Shape #${idx} (${shapeType}${geom ? '/' + geom : ''})`,
    detail: `x=${emuToPx(t.x)}px y=${emuToPx(t.y)}px w=${emuToPx(t.cx)}px h=${emuToPx(t.cy)}px rot=${t.rot}`,
    fillHex: fillHex.length === 6 ? fillHex : '',
    shapeType,
    paragraphs: extractParagraphs(shapeG),
  };
}

/** Extract paragraphs with full run formatting from SVG data attributes. */
function extractParagraphs(shapeG: SVGGElement): ParagraphInfo[] {
  const paraTspans = shapeG.querySelectorAll('tspan[data-ooxml-para-idx]');
  const parasMap = new Map<number, ParagraphInfo>();

  for (const pts of paraTspans) {
    const pi = parseInt(pts.getAttribute('data-ooxml-para-idx')!);
    if (!parasMap.has(pi)) {
      parasMap.set(pi, {
        pi,
        align: pts.getAttribute('data-ooxml-para-align') || 'l',
        runs: [],
      });
    }
    const para = parasMap.get(pi)!;
    const runTspans = pts.querySelectorAll('tspan[data-ooxml-run-idx]');

    for (const rts of runTspans) {
      const ri = parseInt(rts.getAttribute('data-ooxml-run-idx')!);
      const existing = para.runs.find(r => r.ri === ri);
      if (existing) {
        existing.text += rts.textContent || '';
      } else {
        para.runs.push({
          pi,
          ri,
          text: rts.textContent || '',
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
  }

  return Array.from(parasMap.values());
}
