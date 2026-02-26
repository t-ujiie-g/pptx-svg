# PPTX Viewer/Editor — 進捗管理

## 技術スタック
- **言語**: MoonBit (wasm-gc ターゲット)
- **レンダリング**: SVG (data-ooxml-* 属性による往復変換対応)
- **編集機能**: SVG 編集 → SlideData 逆変換 → PPTX エクスポート

---

## Phase 1: Foundation ✅

**目標**: PPTX バイト列から slide1.xml を抽出してブラウザに表示する

- [x] プロジェクト初期化（moon.mod.json, ディレクトリ構成）
- [x] `src/ffi/ffi.mbt` — JS host 関数の FFI 宣言（外部依存なし）
- [x] `src/main/main.mbt` — initialize_pptx, get_slide_count, get_slide_xml_raw, get_entry_list
- [x] `web/host.js` — JS 側 ZIP 解凍（DecompressionStream）+ Wasm 初期化
- [x] `web/index.html` — ファイルドロップ UI
- [x] `test_fixtures/minimal.pptx` — テスト用最小 PPTX
- [x] `test_fixtures/test_node.mjs` — Node.js テスト（全 5 テスト PASS）
- [x] ブラウザ確認用 HTTP サーバー（localhost:8765）

---

## Phase 2: 基本 SVG レンダリング ✅

**目標**: 矩形・楕円・テキスト・画像を含むスライドを SVG で表示

- [x] `src/xml/xml.mbt` — 汎用 XML パーサー（DOM ツリー）
- [x] `src/ooxml/ooxml.mbt` — AutoShape / Picture / TextBody パース
- [x] `src/renderer/renderer.mbt` — rect / ellipse / roundRect / text / image → SVG
- [x] `render_slide_svg(idx)` エクスポート
- [x] 動作確認: タイトルスライドが表示される

---

## ライブラリ化: Round-trip PPTX ↔ SVG ✅

**目標**: SVG を編集して PPTX にエクスポートできる双方向変換ライブラリ

### Phase A: data-ooxml-* 属性付き SVG ✅
- [x] `Color::to_hex()`, `ShapeGeom::to_prst()` メソッド追加
- [x] SVG root に `data-ooxml-slide-cx/cy`, `data-ooxml-bg`, `data-ooxml-scale`
- [x] 各シェイプを `<g data-ooxml-*>` でラップ
- [x] テキスト tspan に `data-ooxml-para-idx/align`, `data-ooxml-run-idx/bold/font-size/color`

### Phase B: OOXML シリアライザ ✅
- [x] `src/serializer/serializer.mbt` — `serialize_slide(SlideData) -> String`
- [x] AutoShape, Picture, テキスト対応
- [x] `get_slide_ooxml(idx)` エクスポート

### Phase C: SVG パーサー ✅
- [x] `src/svg_parser/svg_parser.mbt` — `parse_svg_to_slide(svg_string) -> SlideData`
- [x] data-ooxml-* 属性から SlideData を再構築
- [x] `update_slide_from_svg(idx, svg)` エクスポート
- [x] スライドキャッシュ（`g_slides`, `g_modified`）

### Phase D: JS エクスポートパイプライン ✅
- [x] `host.js` — `#buildZip()`, `#deflate()`, CRC-32 実装
- [x] `PptxRenderer` に `updateSlideFromSvg()`, `getSlideOoxml()`, `exportPptx()` 追加
- [x] `get_modified_entries()` Wasm エクスポート
- [x] `web/index.html` — Export PPTX ボタン、OOXML デバッグ表示

---

## Phase 3: テーマ・継承解決 🔲

**目標**: 色・フォントが Theme/Master/Layout から正しく継承される

- [ ] theme1.xml パース
- [ ] slideMaster / slideLayout パース
- [ ] HLS 変換 + lumMod/lumOff/shade/tint
- [ ] テキストスタイル継承（lstStyle の親子解決）

---

## Phase 4: プリセット図形ライブラリ 🔲

**目標**: 200種のプリセット図形を正しい SVG パスで描画

- [ ] 上位 20 種ハードコード
- [ ] DrawingML ガイド式エバリュエータ
- [ ] グラデーション（linearGradient / radialGradient）
- [ ] 線スタイル（破線・矢印ヘッド）

---

## Phase E: 編集対応 🔲

**目標**: SVG 上での編集を検知してPPTXに反映

- [ ] 位置/サイズ変更検知（pixel→EMU 逆変換）
- [ ] テキスト内容変更の反映
- [ ] ビジュアルスタイル変更（data 属性なしの場合 SVG 属性からフォールバック）

---

## 既知の制限

- SmartArt / チャートはグレーフォールバック
- EMF/WMF 画像は非対応
- アニメーション・トランジションは無視
- Node.js では動作しない（wasm-gc はブラウザのみ）
