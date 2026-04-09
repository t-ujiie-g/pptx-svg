/**
 * SlideThumbnails — horizontal thumbnail strip for slide navigation.
 * Renders thumbnails asynchronously so the main slide appears first.
 */

import { useEffect, useRef } from 'react';
import { insertSvgInto } from '../utils/svg';

interface SlideThumbnailsProps {
  total: number;
  current: number;
  renderSlide: (idx: number) => string;
  onSelect: (idx: number) => void;
  /** Incremented to trigger re-render of thumbnails */
  refreshKey: number;
}

export function SlideThumbnails({ total, current, renderSlide, onSelect, refreshKey }: SlideThumbnailsProps) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const container = containerRef.current;
    if (!container || total === 0) return;

    // Create placeholder thumbnails immediately (no SVG yet)
    container.innerHTML = '';
    const thumbEls: HTMLDivElement[] = [];
    for (let i = 0; i < total; i++) {
      const thumb = document.createElement('div');
      thumb.style.cssText = `
        flex: 0 0 auto; width: 120px; height: 68px;
        border: 2px solid ${i === current ? '#4a90d9' : '#ddd'};
        border-radius: 6px; overflow: hidden; cursor: pointer;
        position: relative; background: #f5f5f5;
      `;

      const label = document.createElement('span');
      label.textContent = `${i + 1}`;
      label.style.cssText = `
        position: absolute; bottom: 0; right: 0;
        background: rgba(0,0,0,0.6); color: #fff;
        font-size: 10px; padding: 1px 5px;
        border-top-left-radius: 4px; z-index: 1;
      `;
      thumb.appendChild(label);

      const idx = i;
      thumb.addEventListener('click', () => onSelect(idx));
      container.appendChild(thumb);
      thumbEls.push(thumb);
    }

    // Render SVG into each thumbnail asynchronously, one at a time
    let cancelled = false;
    let nextIdx = 0;

    function renderNext() {
      if (cancelled || nextIdx >= total) return;
      const i = nextIdx++;
      const svg = renderSlide(i);
      if (svg.startsWith('<svg') && thumbEls[i]) {
        const wrapper = document.createElement('div');
        wrapper.style.cssText = 'width:100%;height:100%;pointer-events:none;position:absolute;top:0;left:0;';
        insertSvgInto(wrapper, svg);
        thumbEls[i].appendChild(wrapper);
      }
      requestAnimationFrame(renderNext);
    }

    // Start after a microtask so the main slide renders first
    requestAnimationFrame(renderNext);

    return () => { cancelled = true; };
  }, [total, current, renderSlide, onSelect, refreshKey]);

  if (total === 0) return null;

  return (
    <div
      ref={containerRef}
      style={{
        display: 'flex', gap: 6, padding: '6px 0',
        overflowX: 'auto', marginBottom: 6, minHeight: 72,
      }}
    />
  );
}
