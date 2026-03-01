# pptx-render — PPTX 完全互換レンダリングライブラリ

**目標**: OOXML PresentationML (ECMA-376 / ISO 29500) 準拠の PPTX → SVG レンダリング + PPTX エクスポート
**スコープ外**: アニメーション (`p:timing`)、トランジション (`p:transition`)、マクロ/VBA

---

## 完了済み ✅

- [x] ZIP 解凍 (JS DecompressionStream) + Wasm FFI + 3-tier js-string 互換
- [x] 汎用 XML パーサー (DOM ツリー)
- [x] AutoShape (rect / ellipse / roundRect / line) + 回転 / フリップ
- [x] Picture (PNG / JPEG / GIF / SVG → data URI base64)
- [x] テキスト (段落 / ラン / bold / italic / fontSize / color / fontFace / level)
- [x] `a:defRPr` デフォルトランプロパティ (bold / italic / fontSize / color / fontFace のフォールバック)
- [x] XML エンティティデコード (`&amp;` → `&` テキストノード)
- [x] テーブル (`p:graphicFrame` + `a:tbl` — セル背景 / テキスト)
- [x] スライド背景色 (`p:bg` → `p:bgPr` → solidFill)
- [x] テーマ解決 (theme1.xml カラー 12 色 + フォントスキーム + lumMod/lumOff/shade/tint/satMod)
- [x] Round-trip パイプライン (data-ooxml-* SVG ↔ SlideData ↔ OOXML ↔ ZIP)
- [x] スライドマスター / レイアウト継承 (背景 / テキストスタイル / placeholder transform / master shapes)
- [x] フォントテーマ参照 (`+mj-lt`/`+mn-lt`/`+mj-ea`/`+mn-ea`)
- [x] 東アジアフォント (`a:ea typeface`) + round-trip
- [x] 行間 (`a:lnSpc`) EMU / 百分率
- [x] 文字間隔 (`a:rPr spc`) + letter-spacing
- [x] `a:lstStyle` シェイプ固有リストスタイル継承

---

## 1. スライドマスター / レイアウト継承 [P0 — 最優先]

ほぼ全ての PPTX で必要。マスター/レイアウトがないと背景・デフォルト書式・プレースホルダが再現できない。

- [x] slideMaster*.xml パース — デフォルトスタイル / 背景 / シェイプ
- [x] slideLayout*.xml パース — レイアウトテンプレート
- [x] スタイル継承チェーン: slide → layout → master → theme
- [x] プレースホルダタイプ解決 (`<p:ph type="title/body/..." idx="N">`)
- [x] マスター/レイアウトからの背景継承 (slide に `p:bg` がなければ親を参照)
- [ ] `p:clrMapOvr` カラーマップオーバーライド
- [x] デフォルトテキストスタイル (`p:txStyles` titleStyle/bodyStyle/otherStyle レベル 0-8)
- [x] `a:lstStyle` (シェイプ/プレースホルダ固有のリストスタイル) レベル継承
- [x] プレースホルダの transform 継承 (slide → layout → master)
- [x] プレースホルダの bodyProps 継承 (anchor 等)
- [x] マスター/レイアウト上のシェイプ描画 (ロゴ・フッター等 — 非PHのみ)

---

## 2. テキスト — 完全対応 [P0]

### 2.1 段落プロパティ
- [x] スペーシング: `a:spcBef` (前) / `a:spcAft` (後) — パース + レンダリング
- [x] 行間 (`a:lnSpc`) — パース + レンダリング (EMU / 百分率)
- [x] インデント: `a:pPr marL/indent` — 左マージン、ぶら下げインデント
- [ ] タブストップ (`a:tabLst`)
- [ ] RTL / BiDi (`a:pPr rtl`)

### 2.2 箇条書き / 番号付きリスト
- [x] 文字バレット (`a:buChar char="●"`)
- [x] 自動番号 (`a:buAutoNum type="arabicPeriod/alphaLcParenR/..."`)
- [ ] バレットフォント (`a:buFont typeface`)
- [ ] バレットサイズ (`a:buSzPct` / `a:buSzPts`)
- [ ] バレット色 (`a:buClr`)
- [x] バレット非表示 (`a:buNone`)
- [ ] 画像バレット (`a:buBlip`)

### 2.3 ラン装飾
- [x] 下線 (`a:rPr u="sng/dbl/heavy/dotted/dash/wavy/..."` — 18 種)
- [x] 取り消し線 (`a:rPr strike="sngStrike/dblStrike"`)
- [x] 上付き / 下付き (`a:rPr baseline="30000/-25000"`)
- [x] 文字間隔 (`a:rPr spc`) — パース + letter-spacing レンダリング
- [ ] カーニング (`a:rPr kern`)
- [ ] キャピタライズ (`a:rPr cap="all/small"`)

### 2.4 フォント
- [x] 東アジアフォント (`a:ea typeface`) — パース + font-family レンダリング + round-trip
- [ ] 複合スクリプト (`a:cs typeface`)
- [ ] シンボルフォント (`a:sym typeface`)
- [x] フォントテーマ参照 (`+mj-lt`/`+mn-lt`/`+mj-ea`/`+mn-ea` → theme font 解決)

### 2.5 テキストボディ
- [x] `a:bodyPr` — アンカー (`anchor="ctr/b/t"`) 垂直位置合わせ
- [x] 内部マージン (`lIns/tIns/rIns/bIns`)
- [x] シェイプ自動フィット (`a:spAutoFit`) — テキスト溢れ時にシェイプ拡大
- [ ] テキスト自動フィット (`a:normAutofit fontScale/lnSpcReduction`)
- [ ] テキスト折り返し (`wrap="square/none"`)
- [ ] カラム数 (`numCol`) / カラム間隔 (`spcCol`)
- [ ] 縦書き (`vert="eaVert/vert/vert270/wordArtVert/..."`)
- [ ] テキスト回転 (`rot`)

### 2.6 ハイパーリンク
- [ ] `a:hlinkClick r:id` → Relationship 経由で URL 解決
- [ ] `a:hlinkMouseOver` — ホバーリンク
- [ ] リンク色 (hlink / folHlink テーマカラー)

---

## 3. 塗りつぶし — 完全対応 [P1]

### 3.1 グラデーション (`a:gradFill`)
- [ ] リニア (`a:lin ang="..." scaled="1"`) → SVG `<linearGradient>`
- [ ] パス型 (`a:path path="circle/rect/shape"`) → SVG `<radialGradient>`
- [ ] グラデーションストップ (`a:gs pos`) の色解決 (schemeClr / srgbClr + モディファイア)
- [ ] タイルフリップ (`tileFlip`)

### 3.2 パターンフィル (`a:pattFill`)
- [ ] 48 種のプリセットパターン (`prst="pct5/pct10/ltDnDiag/..."`)
- [ ] 前景色 / 背景色
- [ ] SVG `<pattern>` へマッピング

### 3.3 画像フィル (`a:blipFill` in shapes)
- [ ] ストレッチ (`a:stretch` + `a:fillRect`)
- [ ] タイル (`a:tile tx/ty/sx/sy/flip/algn`)
- [ ] ソースRect (`a:srcRect l/t/r/b`) — クロッピング

### 3.4 透過 / アルファ
- [ ] Color に alpha フィールド追加
- [ ] `a:alpha val` → SVG `opacity` / `fill-opacity`
- [ ] `a:alphaModFix amt` — 固定アルファ変更
- [ ] グラデーションストップ個別アルファ

---

## 4. 線 / ストローク — 完全対応 [P1]

- [ ] 破線スタイル (`a:prstDash val="dash/dot/dashDot/lgDash/sysDot/..."`)
- [ ] カスタム破線 (`a:custDash`)
- [ ] 矢印ヘッド (`a:headEnd type="triangle/stealth/diamond/oval/arrow" w/len`)
- [ ] 矢印テイル (`a:tailEnd` — 同上)
- [ ] 線結合 (`a:round` / `a:bevel` / `a:miter lim`)
- [ ] 線端 (`a:ln cap="flat/rnd/sq"`)
- [ ] 複合線 (`a:ln cmpd="sng/dbl/thickThin/thinThick/tri"`)
- [ ] 線なし (`a:noFill` inside `a:ln`)

---

## 5. シェイプ — 完全対応 [P1]

### 5.1 グループシェイプ (`p:grpSp`)
- [ ] 再帰的子シェイプパース
- [ ] グループ変換 (`a:xfrm` + `a:chOff` / `a:chExt`) の座標変換
- [ ] ネストグループ
- [ ] SVG `<g transform="...">` レンダリング

### 5.2 コネクタ (`p:cxnSp`)
- [ ] 始点 / 終点座標
- [ ] 直線コネクタ (`straightConnector1`)
- [ ] 曲線コネクタ (`curvedConnector2-5`)
- [ ] 折れ線コネクタ (`bentConnector2-5`)
- [ ] 矢印ヘッドとの組み合わせ

### 5.3 プリセットジオメトリ — 全 220+ 種 (`a:prstGeom`)

ECMA-376 Part 1 §20.1.10.56 `ST_ShapeType` で定義される全図形を実装する。
描画定義は `presetShapeDefinitions.xml` (ECMA-376 Appendix D) にガイド式で記載。

**実装方針**: DrawingML ガイド式エバリュエータを作り、`presetShapeDefinitions.xml` の定義から SVG path を生成する。個別ハードコードはしない。

#### カテゴリ別 (全て対応必須)

| カテゴリ | 種数 | 例 |
|---------|------|-----|
| 基本図形 | ~40 | rect, ellipse, triangle, diamond, pentagon, hexagon, octagon, trapezoid, parallelogram, plus, cross, donut, blockArc, ... |
| ブロック矢印 | ~25 | rightArrow, leftArrow, upArrow, downArrow, leftRightArrow, quadArrow, curvedRightArrow, circularArrow, uturnArrow, ... |
| フローチャート | ~27 | flowChartProcess, flowChartDecision, flowChartTerminator, flowChartDocument, flowChartMultidocument, flowChartPredefinedProcess, ... |
| 吹き出し | ~20 | wedgeRoundRectCallout, wedgeEllipseCallout, cloudCallout, borderCallout1-3, accentCallout1-3, accentBorderCallout1-3, ... |
| 星・リボン | ~15 | star4, star5, star6, star8, star10, star12, star16, star24, star32, ribbon, ribbon2, wave, doubleWave, ... |
| 数式記号 | ~5 | mathPlus, mathMinus, mathMultiply, mathDivide, mathEqual |
| アクションボタン | ~14 | actionButtonBlank, actionButtonHome, actionButtonHelp, actionButtonForwardNext, actionButtonBackPrevious, ... |
| 装飾 | ~20 | heart, lightningBolt, sun, moon, smileyFace, foldedCorner, irregularSeal1-2, gear6, gear9, funnel, ... |
| その他 | ~50+ | frame, bevel, plaque, can, cube, corner, diagStripe, pie, chord, teardrop, ... |

#### ガイド式エバリュエータ
- [ ] `a:gd` (ガイド定義) パース — `fmla` 式: `+- */ ?: val pin cos sin tan at2 cat2 sat2 sqrt mod abs`
- [ ] `a:avLst` (調整値) パース — パラメトリック図形
- [ ] `a:gdLst` (ガイドリスト) 評価
- [ ] `a:pathLst` → SVG `<path d="...">` 変換
  - [ ] moveTo / lineTo / arcTo / cubicBezTo / quadBezTo / close
- [ ] `a:rect` (テキスト矩形) 計算
- [ ] `a:cxnLst` (接続ポイント) — コネクタ接続用

### 5.4 カスタムジオメトリ (`a:custGeom`)
- [ ] `a:pathLst` → SVG path (プリセットと同じエンジン)
- [ ] フリーフォーム図形対応

---

## 6. テーブル — 完全対応 [P1]

- [ ] セル結合 — 水平 (`gridSpan`) / 垂直 (`vMerge` / `rowSpan`)
- [ ] セルボーダー (`a:tcBorders` — `lnL/lnR/lnT/lnB/lnTlToBr/lnBlToTr`)
- [ ] セルマージン (`a:tcPr marL/marR/marT/marB`)
- [ ] セルアンカー (`a:tcPr anchor="ctr/b/t"`)
- [ ] テーブルスタイル (`a:tblStyleId` → theme のテーブルスタイル定義)
- [ ] バンド行/列条件書式 (`firstRow/lastRow/firstCol/lastCol/bandRow/bandCol`)
- [ ] セルグラデーション塗りつぶし

---

## 7. 画像 — 完全対応 [P2]

- [ ] クロッピング (`a:srcRect l/t/r/b` — %)
- [ ] SVG 画像 (`a:blip` + SVG extension `a:extLst`)
- [ ] 画像エフェクト (`a:clrChange`, brightness/contrast)
- [ ] Duotone (`a:duotone`)
- [ ] 画像のアルファ (`a:alphaModFix`)
- [ ] EMF (Enhanced Metafile) — ラスタライズ
- [ ] WMF (Windows Metafile) — ラスタライズ
- [ ] TIFF 画像

---

## 8. エフェクト [P2]

- [ ] 外側シャドウ (`a:outerShdw blurRad/dist/dir/algn/...`) → SVG `<filter>` feDropShadow
- [ ] 内側シャドウ (`a:innerShdw`)
- [ ] グロー (`a:glow rad`)
- [ ] ソフトエッジ (`a:softEdge rad`) → SVG feGaussianBlur
- [ ] リフレクション (`a:reflection blurRad/stA/endA/endPos/dir/...`)
- [ ] 3D — ベベル / 押し出し / 照明 (ベストエフォート、完全再現は困難)

---

## 9. チャート — 全タイプ対応 [P2]

ChartML (ECMA-376 Part 1 Chapter 21) パーサー + SVG レンダラーが必要。

### 9.1 基盤
- [ ] `c:chartSpace` / `c:chart` パース
- [ ] 軸 (`c:valAx`, `c:catAx`, `c:dateAx`, `c:serAx`) — ラベル / スケール / グリッド線
- [ ] 凡例 (`c:legend`)
- [ ] データラベル (`c:dLbls`)
- [ ] タイトル (`c:title`)
- [ ] プロットエリア (`c:plotArea`)

### 9.2 チャートタイプ (全て対応)

| タイプ | 要素 | 備考 |
|--------|------|------|
| 棒グラフ | `c:barChart` | 縦/横 (`barDir`), 積み上げ (`grouping`) |
| 折れ線グラフ | `c:lineChart` | マーカー, 平滑化 |
| 円グラフ | `c:pieChart` | 分離 (`explosion`) |
| ドーナツグラフ | `c:doughnutChart` | 穴サイズ (`holeSize`) |
| 散布図 | `c:scatterChart` | マーカー, 線スタイル |
| 面グラフ | `c:areaChart` | 積み上げ対応 |
| レーダーチャート | `c:radarChart` | 塗りつぶし / 線 |
| バブルチャート | `c:bubbleChart` | サイズ軸 |
| 株価チャート | `c:stockChart` | HLC / OHLC |
| 等高線 | `c:surfaceChart` | 3D / ワイヤーフレーム |
| 円グラフ付き | `c:ofPieChart` | 円 of 円 / 棒 of 円 |
| 3D チャート | `c:bar3DChart`, `c:line3DChart`, `c:pie3DChart`, `c:area3DChart`, `c:surface3DChart` | 2D 投影で近似 |
| 複合チャート | 複数系列 | 異なるタイプの系列を重ね合わせ |

### 9.3 チャート装飾
- [ ] 系列の塗りつぶし / 線スタイル
- [ ] データポイント個別書式
- [ ] トレンドライン (`c:trendline`)
- [ ] エラーバー (`c:errBars`)
- [ ] 近似曲線

---

## 10. SmartArt [P3]

- [ ] `dgm:*` (DiagramML) パース
- [ ] レイアウトアルゴリズム (list / cycle / hierarchy / process / relationship / matrix / pyramid)
- [ ] フォールバック: SmartArt の画像キャッシュ (`drs/` 内の EMF/PNG) 利用

---

## 11. その他のオブジェクト [P3]

### 11.1 OLE / 埋め込み
- [ ] `p:oleObj` — フォールバック画像表示
- [ ] 埋め込み Excel / Word のプレビュー

### 11.2 メディア
- [ ] ビデオ (`p:vid`) — ポスターフレーム表示
- [ ] オーディオ (`p:audio`) — アイコン表示

### 11.3 数式
- [ ] OMML (`m:oMath`) — Office Math Markup Language
- [ ] フォールバック画像

### 11.4 コメント / ノート
- [ ] スピーカーノート (`notesSlide*.xml`) — API 経由で取得可能に
- [ ] コメント (`comments*.xml`) — 同上

---

## 12. ライブラリ公開 [P1 — インフラ]

### 12.1 TypeScript ライブラリ化
- [x] `host.js` → TypeScript 分離 (`lib/`)
- [x] `lib/pptx-renderer.ts` — PptxRenderer クラス (コア API)
- [x] `lib/wasm-compat.ts` — 3-tier Wasm インスタンス化
- [x] `lib/zip.ts` — ZIP 解凍 / 構築
- [x] `lib/utils.ts` — CRC-32, base64
- [x] `lib/index.ts` — 公開 API re-export
- [x] `web/` — デモ UI (ライブラリをインポートして使う)

### 12.2 ビルド / パッケージング
- [x] `tsconfig.json` — ESM 出力 (`dist/`)
- [x] `package.json` — name, version, exports, types, files
- [x] `dist/` — JS + `.d.ts` + source maps 出力
- [ ] `dist/pptx-render.wasm` — ビルド成果物コピー (ビルドスクリプト)
- [ ] npm publish ワークフロー

### 12.3 API 設計
- [x] `PptxRenderer.init(wasmUrl | wasmBytes)` — Wasm 初期化
- [x] `PptxRenderer.loadPptx(ArrayBuffer)` — PPTX ロード
- [x] `PptxRenderer.renderSlideSvg(idx) → string (SVG)`
- [x] `PptxRenderer.getSlideCount() → number`
- [x] `PptxRenderer.exportPptx() → Promise<ArrayBuffer>`
- [x] `PptxRenderer.getSlideXmlRaw(idx) → string` — デバッグ用
- [ ] エラーハンドリング (例外 vs Result 型)

### 12.4 テスト
- [ ] ユニットテスト (Vitest or similar)
- [ ] ビジュアルリグレッションテスト (参照 SVG との diff)
- [ ] 複数 PPTX ファイルでの互換性テスト
- [ ] ブラウザ互換テスト (Chrome 111+, Firefox 120+, Safari 17+)

### 12.5 ドキュメント
- [ ] README.md — 使い方、API リファレンス、ブラウザ互換性
- [ ] CHANGELOG.md
- [ ] JSDoc / TSDoc

---

## 優先度サマリー

| 優先度 | 内容 | 理由 |
|--------|------|------|
| **P0** | 1. マスター/レイアウト継承, 2. テキスト完全対応 | ほぼ全 PPTX で必要 |
| **P1** | 3-6. 塗り/線/シェイプ/テーブル, 12. ライブラリ公開 | ビジネス文書の 90% をカバー |
| **P2** | 7-9. 画像高度/エフェクト/チャート | 完全互換に必要 |
| **P3** | 10-11. SmartArt/OLE/メディア/数式 | 特殊ケース |
