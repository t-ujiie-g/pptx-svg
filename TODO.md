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

---

## 7. 画像 — 残機能 [P2]

- [ ] SVG 画像 (`a:blip` + SVG extension `a:extLst`)
- [ ] 画像エフェクト (`a:clrChange`, brightness/contrast)
- [ ] Duotone (`a:duotone`)
- [ ] EMF / WMF / TIFF — ラスタライズ

---

## 8. エフェクト [P2]

- [ ] 外側シャドウ (`a:outerShdw`) → SVG `<filter>` feDropShadow
- [ ] 内側シャドウ (`a:innerShdw`)
- [ ] グロー (`a:glow rad`)
- [ ] ソフトエッジ (`a:softEdge rad`) → SVG feGaussianBlur
- [ ] リフレクション (`a:reflection`)
- [ ] 3D — ベベル / 押し出し / 照明 (ベストエフォート)
- [ ] テキスト単位の影・エフェクト (`a:rPr` 内 `a:effectLst`)

---

## 9. チャート [P2]

ChartML (ECMA-376 Part 1 Chapter 21) パーサー + SVG レンダラーが必要。

- [ ] 基盤: `c:chartSpace` / 軸 / 凡例 / データラベル / タイトル / プロットエリア
- [ ] 棒/折れ線/円/ドーナツ/散布/面/レーダー/バブル/株価/等高線
- [ ] 複合チャート、3D チャート (2D 投影で近似)
- [ ] 系列書式 / データポイント個別書式 / トレンドライン / エラーバー

---

## 10. テキスト — 高度機能 [P2]

- [ ] テキストワープ (`a:prstTxWarp`) — WordArt 的な曲線パス上テキスト
- [ ] テキストアウトライン (`a:rPr/a:ln`) — 文字の輪郭線
- [ ] テキストグラデーション塗り (`a:rPr/a:gradFill`) — 文字へのグラデーション

---

## 11. スライド・シェイプ — 残機能 [P1]

- [ ] プレースホルダ自動内容 — 日付 (`ph type="dt"`) / スライド番号 (`sldNum`) / フッター (`ftr`)
- [ ] 背景 blip/パターン — `p:bg/p:bgPr` の `a:blipFill` / `a:pattFill` (現状 solid+gradient のみ)
- [ ] シェイプリンク — 図形自体の `a:hlinkClick` / `a:hlinkHover` (テキストリンクとは別)
- [ ] 線のグラデーション/パターン塗り (`a:ln/a:gradFill`, `a:ln/a:pattFill`)
- [ ] 複数スライドマスター対応 (現状1つ目のみ読み込み)
- [ ] カラー修飾子追加 — `comp` / `inv` / `gamma` / `hueMod` (現状 lumMod/lumOff/shade/tint/satMod/alpha)

---

## 12. SmartArt [P3]

- [ ] `dgm:*` (DiagramML) パース + レイアウトアルゴリズム
- [ ] フォールバック: SmartArt の画像キャッシュ利用

---

## 13. その他のオブジェクト [P3]

- [ ] OLE / 埋め込み (`p:oleObj` — フォールバック画像表示)
- [ ] メディア (ビデオポスターフレーム / オーディオアイコン)
- [ ] 数式 (OMML `m:oMath`)
- [ ] スピーカーノート / コメント (API 経由取得)
- [ ] 埋め込みフォント (`a:fontScheme` + fontdata)

---

## 14. ライブラリ公開 — 残タスク [P1]

- [ ] `dist/pptx-svg.wasm` ビルド成果物コピー (ビルドスクリプト)
- [ ] npm publish ワークフロー
- [ ] エラーハンドリング設計
- [ ] テスト (Vitest / ビジュアルリグレッション / ブラウザ互換)
- [ ] ドキュメント (README / CHANGELOG / JSDoc)

---

## 優先度サマリー

| 優先度 | 内容 | 状態 |
|--------|------|------|
| **P0** | 基盤/テーマ/マスター継承/テキスト完全対応 | **完了** |
| **P1** | 塗り/線/グループ/コネクタ/ジオメトリ/テーブル完全対応 | **完了** |
| **P1** | スライド・シェイプ残機能, ライブラリ公開 | 未着手 |
| **P2** | 画像 (クロップ/アルファ/外部参照 完了, SVG画像/エフェクト/EMF 未着手) | **一部完了** |
| **P2** | エフェクト/チャート/テキスト高度 | 未着手 |
| **P3** | SmartArt/OLE/メディア/数式/ノート/埋め込みフォント | 未着手 |
