/**
 * SlideViewer — SVG slide display with click-to-select.
 *
 * This component renders SVG outside of React's virtual DOM (via innerHTML),
 * since pptx-svg outputs raw SVG strings. React only manages the container div.
 */

import { useEffect } from 'react';
import { insertSvgInto, showOverlay, removeOverlay, extractShapeInfo, type ShapeInfo } from '../utils/svg';

interface SlideViewerProps {
  containerRef: React.RefObject<HTMLDivElement | null>;
  hasSelection: boolean;
  onSelect: (info: ShapeInfo | null) => void;
}

export function SlideViewer({ containerRef, hasSelection, onSelect }: SlideViewerProps) {
  // Click to select/deselect shapes
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const handleClick = (e: MouseEvent) => {
      const target = e.target as Element;
      if (target.classList.contains('resize-handle')) return;

      const shapeG = target.closest('g[data-ooxml-shape-idx]') as SVGGElement | null;
      if (shapeG) {
        showOverlay(container, shapeG);
        onSelect(extractShapeInfo(shapeG));
      } else {
        removeOverlay(container);
        onSelect(null);
      }
    };

    container.addEventListener('click', handleClick);
    return () => container.removeEventListener('click', handleClick);
  }, [containerRef, onSelect]);

  return (
    <div
      ref={containerRef}
      style={{
        background: '#fff', border: '1px solid #ddd', borderRadius: 8,
        padding: 16, minHeight: 300, overflow: 'hidden',
        position: 'relative',
        cursor: hasSelection ? 'move' : 'default',
      }}
    />
  );
}

// ── Imperative helpers (called from App via containerRef) ──

/** Insert SVG string into container. */
export function insertSvg(container: HTMLDivElement | null, svgString: string) {
  if (container) insertSvgInto(container, svgString);
}

/** Re-show selection overlay for a shape by index. Returns shape info or null. */
export function reselectShape(container: HTMLDivElement | null, shapeIdx: number): ShapeInfo | null {
  if (!container) return null;
  const svgEl = container.querySelector('svg');
  if (!svgEl) return null;
  const g = svgEl.querySelector<SVGGElement>(`g[data-ooxml-shape-idx="${shapeIdx}"]`);
  if (!g) return null;
  showOverlay(container, g);
  return extractShapeInfo(g);
}
