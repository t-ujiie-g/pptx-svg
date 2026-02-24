# PPTX Viewer/Editor — 進捗管理

## 技術スタック
- **言語**: MoonBit (wasm-gc ターゲット)
- **レンダリング**: SVG
- **編集機能**: テキスト編集・図形移動リサイズ・図形追加削除・スライド並び替え

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
- [x] wasm-gc リリースビルド成功（7.1KB）
- [x] ブラウザ確認用 HTTP サーバー（localhost:8765）

**学習事項:**
- 外部パッケージ（bobzhang/zip, ruifeng/XMLParser）は現在のコンパイラと非互換
- ZIP 解凍は JS 側（DecompressionStream）で行う設計が最適
- MoonBit は標準ライブラリのみで動作、7.1KB の小さな Wasm バイナリ生成
- wasm-gc の js-string builtins により String の FFI コストはゼロ

---

## Phase 2: 基本 SVG レンダリング 🔲

**目標**: 矩形・楕円・テキスト・画像を含むスライドを SVG で表示

- [ ] `src/ooxml/color.mbt` — RGB 直接指定のカラー解決
- [ ] `src/ooxml/text.mbt` — txBody / paragraph / run パース
- [ ] `src/ooxml/shape.mbt` — AutoShape / TextBox / Picture パース
- [ ] `src/renderer/svg_builder.mbt` — SVG 文字列生成ヘルパー
- [ ] `src/renderer/shape_renderer.mbt` — rect / ellipse → SVG
- [ ] `src/renderer/text_renderer.mbt` — テキスト → `<foreignObject>` + CSS
- [ ] `src/renderer/image_renderer.mbt` — 画像 → base64 data URI
- [ ] `src/renderer/slide_renderer.mbt` — スライド全体の SVG 合成
- [ ] `render_slide_svg(idx)` エクスポート
- [ ] 動作確認: タイトルスライドが近似的に正しく表示される

**→ Phase 2 完了後にユーザー確認**

---

## Phase 3: テーマ・継承解決 🔲

**目標**: 色・フォントが Theme/Master/Layout から正しく継承される

- [ ] `src/ooxml/theme.mbt` — theme1.xml パース
- [ ] `src/ooxml/master.mbt` — slideMaster パース
- [ ] `src/ooxml/layout.mbt` — slideLayout パース
- [ ] `src/ooxml/color.mbt` 完全実装 — HLS 変換 + lumMod/lumOff/shade/tint
- [ ] テキストスタイル継承（lstStyle の親子解決）
- [ ] 動作確認: テーマカラー・継承フォントが正しく表示される

**→ Phase 3 完了後にユーザー確認**

---

## Phase 4: プリセット図形ライブラリ 🔲

**目標**: 200種のプリセット図形を正しい SVG パスで描画

- [ ] `src/renderer/preset_shapes.mbt` — 上位 20 種ハードコード
- [ ] DrawingML ガイド式エバリュエータ（残りのプリセット用）
- [ ] グラデーション（linearGradient / radialGradient）
- [ ] 線スタイル（破線・矢印ヘッド）
- [ ] 図形の回転・反転（SVG transform）
- [ ] 動作確認: 複雑な図形を含む .pptx が正しく表示される

**→ Phase 4 完了後にユーザー確認**

---

## Phase 5: 編集操作 🔲

**目標**: テキスト編集・移動・リサイズ・追加削除・スライド並び替え

- [ ] `src/editor/operations.mbt` — EditOp enum + apply 関数
- [ ] `src/editor/history.mbt` — undo/redo スタック
- [ ] `apply_edit` / `undo` / `redo` / `get_shape_info` エクスポート
- [ ] `web/editor_ui.js` — 選択オーバーレイ + 8 方向リサイズハンドル
- [ ] `web/editor_ui.js` — ドラッグで図形移動
- [ ] `web/editor_ui.js` — テキスト編集（textarea オーバーレイ）
- [ ] `web/editor_ui.js` — スライドパネル + 並び替え
- [ ] 動作確認: 各編集操作 + undo/redo が正しく動く

**→ Phase 5 完了後にユーザー確認**

---

## Phase 6: 高度機能（継続） 🔲

- [ ] テーブル完全対応
- [ ] PPTX エクスポート（JSZip 連携）
- [ ] 埋め込みフォント対応
- [ ] スライドノート表示
- [ ] サムネイル生成

---

## 既知の制限・注意事項

- SmartArt / チャートはグレーフォールバックで表示（初期フェーズ）
- EMF/WMF 画像は非対応（プレースホルダー表示）
- アニメーション・トランジションは無視
- Safari の WasmGC サポートは比較的最近のため要確認
