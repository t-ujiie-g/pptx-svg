/**
 * Default font fallback mappings.
 *
 * Maps fonts commonly used in PPTX (Windows/Office) to cross-platform alternatives
 * available on macOS, Linux, and the web. Entries are checked at runtime via FFI
 * and appended to the CSS `font-family` list in rendered SVG.
 *
 * To customize, pass `fontFallbacks` option to PptxRenderer constructor.
 * User entries are merged with (and override) these defaults.
 */

/** Font fallback mapping: source font name → ordered list of fallback font names. */
export type FontFallbackMap = Record<string, string[]>;

export const DEFAULT_FONT_FALLBACKS: FontFallbackMap = {
  // ── Japanese Gothic (ゴシック系) ──────────────────────────────────────────
  'ＭＳ Ｐゴシック': ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'ＭＳ ゴシック':   ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'メイリオ':       ['Hiragino Sans', 'Noto Sans JP', 'Yu Gothic'],
  '游ゴシック':     ['Yu Gothic', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],
  '游ゴシック体':   ['Yu Gothic', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],

  // ── Japanese Mincho (明朝系) ──────────────────────────────────────────────
  'ＭＳ Ｐ明朝': ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'ＭＳ 明朝':   ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  '游明朝':     ['Yu Mincho', 'Hiragino Mincho ProN', 'Noto Serif JP'],
  '游明朝体':   ['Yu Mincho', 'Hiragino Mincho ProN', 'Noto Serif JP'],

  // ── Japanese decorative / HG fonts ────────────────────────────────────────
  'HGPゴシックE':     ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'HGPゴシックM':     ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'HGP明朝E':         ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'HGP明朝B':         ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'HG丸ゴシックM-PRO': ['Hiragino Maru Gothic ProN', 'Noto Sans JP', 'Yu Gothic'],

  // ── Western fonts (Windows → macOS/Web) ───────────────────────────────────
  'Calibri':    ['Helvetica Neue', 'Helvetica', 'Arial'],
  'Cambria':    ['Georgia', 'Times New Roman'],
  'Segoe UI':   ['SF Pro Text', 'Helvetica Neue', 'Arial'],

  // ── Korean ────────────────────────────────────────────────────────────────
  'Malgun Gothic': ['Apple SD Gothic Neo', 'Noto Sans KR'],
};
