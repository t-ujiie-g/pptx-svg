# pptx-svg — PPTX 完全互換レンダリングライブラリ

**目標**: OOXML PresentationML (ECMA-376 / ISO 29500) 準拠の PPTX → SVG レンダリング + PPTX エクスポート
**スコープ外**: アニメーション (`p:timing`)、トランジション (`p:transition`)、マクロ/VBA

---

## 完了済み

- [x] **基盤**: ZIP 解凍 (JS DecompressionStream) + Wasm FFI + 3-tier js-string 互換 + 汎用 XML パーサー
- [x] **シェイプ基本**: AutoShape (rect/ellipse/roundRect/line) + 回転/フリップ + Picture (data URI base64)
- [x] **テーブル完全対応**: `p:graphicFrame` + `a:tbl` — セル背景/テキスト/グラデーション + セル結合 (gridSpan/rowSpan/vMerge) + ボーダー (lnL/R/T/B + 対角線 lnTlToBr/lnBlToTr + noFill) + マージン (marL/R/T/B) + アンカー + テーブルスタイル (tblStyleId + tableStyles.xml パース) + 条件書式 (firstRow/lastRow/firstCol/lastCol/bandRow/bandCol)
- [x] **テーマ解決**: theme1.xml カラー 12 色 + フォントスキーム + lumMod/lumOff/shade/tint/satMod
- [x] **Round-trip**: data-ooxml-* SVG ↔ SlideData ↔ OOXML ↔ ZIP
- [x] **マスター/レイアウト継承**: slideMaster/slideLayout パース + placeholder transform/bodyProps/背景/テキストスタイル継承 + `p:clrMapOvr` + マスターシェイプ描画
- [x] **テキスト完全対応**: 段落書式 (spacing/indent/tab/RTL) + バレット全種 (文字/自動番号/画像/フォント/サイズ/色) + ラン装飾 (下線/取消線/上下付き/間隔/kern/cap) + フォント (EA/CS/Sym/テーマ参照) + bodyPr (anchor/insets/autofit/wrap/vert/rot/columns) + ハイパーリンク (click/hover/色)
- [x] **塗りつぶし完全対応**: グラデーション (linear/path/tileFlip) + パターン (全 48 種) + 画像フィル (stretch/tile/crop) + アルファ/透過
- [x] **線/ストローク完全対応**: 破線 11 種 + カスタム破線 + 矢印 5 種 + 線結合/線端/複合線/noFill + Line ジオメトリ `<line>` SVG
- [x] **グループシェイプ**: `p:grpSp` 再帰パース + 座標変換 + ネスト + SVG `<g>`
- [x] **コネクタ**: `p:cxnSp` 直線/折れ線/曲線 + 矢印 + 調整値 round-trip
- [x] **プリセットジオメトリ**: `a:prstGeom` ガイド式エバリュエータ + ~154 種 SVG path 生成
- [x] **カスタムジオメトリ**: `a:custGeom` パース + ガイド式 + pathLst → SVG path + round-trip
- [x] **テキスト矩形**: `a:rect` プリセット/カスタムジオメトリのテキスト配置領域計算
- [x] **接続ポイント**: `a:cxnLst` パース + `stCxnId/endCxnId` コネクタ接続 round-trip
- [x] **ギア歯**: gear6/gear9 正確な歯型パス (6/9 歯 + 中心穴)
- [x] **ライブラリ化**: TypeScript 分割 (lib/ → dist/) + PptxRenderer クラス + Wasm 3-tier フォールバック
- [x] **画像クロップ**: `srcRect` → SVG `<clipPath>` (Picture + AutoShape blipFill 両対応)
- [x] **画像アルファ**: `a:alphaModFix` パース + SVG `opacity` 属性
- [x] **外部画像参照**: `TargetMode="External"` — Relationship に target_mode 追加 + 外部 URL 直接参照
- [x] **SVG 画像**: `a:extLst` 内 `asvg:svgBlip` / `a16:svgBlip` パース + SVG 画像優先表示
- [x] **画像エフェクト**: `a:lum` brightness/contrast → SVG `<filter>` feComponentTransfer
- [x] **Duotone**: `a:duotone` パース + SVG filter (grayscale→2色マッピング)
- [x] **Color change**: `a:clrChange` パース + round-trip (SVG filter での精密再現は不可 — データ保持のみ)
- [x] **カラー修飾子拡張**: comp/inv/gamma/invGamma/hueMod/hueOff/satOff
- [x] **シェイプリンク**: `a:hlinkClick`/`a:hlinkHover` on `p:cNvPr` (パース/SVG属性/round-trip/シリアライズ)
- [x] **背景 blip/パターン**: `p:bg/p:bgPr` の `a:blipFill`/`a:pattFill` パース/レンダリング/シリアライズ
- [x] **線グラデーション/パターン**: `a:ln` 内 `a:gradFill`/`a:pattFill` パース/SVGレンダリング/シリアライズ
- [x] **プレースホルダ自動内容**: スライド番号(`sldNum`)/日付(`dt`)/フッター(`ftr`) 空プレースホルダへの自動テキスト注入
- [x] **エフェクト基本**: `a:effectLst` パース + 外側シャドウ(`a:outerShdw`→feDropShadow) + 内側シャドウ(`a:innerShdw`) + グロー(`a:glow`) + ソフトエッジ(`a:softEdge`→feGaussianBlur) — データモデル/パース/SVGフィルタ/round-trip/シリアライズ

---

## 8. ライブラリ公開 [P1]

- [ ] `dist/pptx-svg.wasm` ビルド成果物コピー (ビルドスクリプト)
- [ ] npm publish ワークフロー
- [ ] エラーハンドリング設計
- [ ] テスト (Vitest / ビジュアルリグレッション / ブラウザ互換)
- [ ] ドキュメント (README / CHANGELOG / JSDoc)

---

## 9. エフェクト [P2]

- [x] 外側シャドウ (`a:outerShdw`) → SVG `<filter>` feDropShadow
- [x] 内側シャドウ (`a:innerShdw`)
- [x] グロー (`a:glow rad`)
- [x] ソフトエッジ (`a:softEdge rad`) → SVG feGaussianBlur
- [ ] リフレクション (`a:reflection`) — 構造的アプローチ (フィルタではなくシェイプ複製+反転+フェード)
- [ ] 3D — ベベル / 押し出し / 照明 (ベストエフォート)
- [ ] テキスト単位の影・エフェクト (`a:rPr` 内 `a:effectLst`)

---

## 10. チャート [P2]

ChartML (ECMA-376 Part 1 Chapter 21) パーサー + SVG レンダラーが必要。

- [ ] 基盤: `c:chartSpace` / 軸 / 凡例 / データラベル / タイトル / プロットエリア
- [ ] 棒/折れ線/円/ドーナツ/散布/面/レーダー/バブル/株価/等高線
- [ ] 複合チャート、3D チャート (2D 投影で近似)
- [ ] 系列書式 / データポイント個別書式 / トレンドライン / エラーバー

---

## 11. テキスト — 高度機能 [P2]

- [ ] テキストワープ (`a:prstTxWarp`) — WordArt 的な曲線パス上テキスト
- [ ] テキストアウトライン (`a:rPr/a:ln`) — 文字の輪郭線
- [ ] テキストグラデーション塗り (`a:rPr/a:gradFill`) — 文字へのグラデーション

---

## 保留 — 対応予定なし

以下は技術的制約により対応しない項目。データとしては保持するが、レンダリングは行わない。

- [ ] **EMF / WMF** — GDI 描画命令の逐次解釈が必要 (仕様 ~500 ページ)。純粋 Wasm での再実装は工数対効果が極めて低い
- [ ] **TIFF** — ブラウザ `<img>` 非対応。複数圧縮方式 (LZW/CCITT/ZIP) のデコーダ自作が必要。遭遇頻度が低い
- [ ] **SmartArt** (`dgm:*` DiagramML) — レイアウトアルゴリズムの再実装が必要。フォールバック画像利用は検討余地あり
- [ ] **OLE / 埋め込み** (`p:oleObj`) — フォールバック画像表示のみ検討余地あり
- [ ] **メディア** — ビデオポスターフレーム / オーディオアイコン
- [ ] **数式** (OMML `m:oMath`) — 数式レンダラーの自作が必要
- [ ] **スピーカーノート / コメント** — レンダリング対象外 (API 経由取得は検討余地あり)
- [ ] **埋め込みフォント** (`a:fontScheme` + fontdata)

---

## 優先度サマリー

| 優先度 | 内容 | 状態 |
|--------|------|------|
| **P0** | 基盤/テーマ/マスター継承/テキスト完全対応 | **完了** |
| **P1** | 塗り/線/グループ/コネクタ/ジオメトリ/テーブル完全対応 | **完了** |
| **P1** | スライド・シェイプ残機能 (カラー修飾子/背景blip・パターン/シェイプリンク/線グラデ・パターン/PH自動内容) | **完了** |
| **P1** | ライブラリ公開 (ビルド/npm/テスト/ドキュメント) | 未着手 |
| **P2** | 画像 (クロップ/アルファ/外部参照/SVG画像/エフェクト/Duotone/clrChange) | **完了** |
| **P2** | エフェクト基本 (outerShdw/innerShdw/glow/softEdge) | **完了** |
| **P2** | エフェクト残 (リフレクション/3D/テキストエフェクト) | 未着手 |
| **P2** | チャート/テキスト高度 | 未着手 |
| **—** | EMF/WMF/TIFF/SmartArt/OLE/メディア/数式/ノート/フォント | **保留** |
