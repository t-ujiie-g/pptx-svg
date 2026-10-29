/** Font fallback mapping: source font name → ordered list of fallback font names. */
export type FontFallbackMap = Record<string, string[]>;

/** Built-in fallbacks for common Office fonts; the `fontFallbacks` option overrides entries. */
export const DEFAULT_FONT_FALLBACKS: FontFallbackMap = {
  'ＭＳ Ｐゴシック': ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'ＭＳ ゴシック':   ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'メイリオ':       ['Hiragino Sans', 'Noto Sans JP', 'Yu Gothic'],
  '游ゴシック':     ['Yu Gothic', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],
  '游ゴシック体':   ['Yu Gothic', 'Hiragino Kaku Gothic ProN', 'Noto Sans JP'],

  'ＭＳ Ｐ明朝': ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'ＭＳ 明朝':   ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  '游明朝':     ['Yu Mincho', 'Hiragino Mincho ProN', 'Noto Serif JP'],
  '游明朝体':   ['Yu Mincho', 'Hiragino Mincho ProN', 'Noto Serif JP'],

  'HGPゴシックE':     ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'HGPゴシックM':     ['Hiragino Kaku Gothic ProN', 'Noto Sans JP', 'Yu Gothic', 'Meiryo'],
  'HGP明朝E':         ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'HGP明朝B':         ['Hiragino Mincho ProN', 'Noto Serif JP', 'Yu Mincho'],
  'HG丸ゴシックM-PRO': ['Hiragino Maru Gothic ProN', 'Noto Sans JP', 'Yu Gothic'],

  'Calibri':    ['Helvetica Neue', 'Helvetica', 'Arial'],
  'Cambria':    ['Georgia', 'Times New Roman'],
  'Segoe UI':   ['SF Pro Text', 'Helvetica Neue', 'Arial'],

  'Malgun Gothic': ['Apple SD Gothic Neo', 'Noto Sans KR'],
};
