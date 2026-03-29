# pptx-svg — PPTX 完全互換レンダリングライブラリ

**目標**: OOXML PresentationML (ECMA-376 / ISO 29500) 準拠の PPTX → SVG レンダリング + PPTX エクスポート
**スコープ外**: マクロ/VBA

**互換性方針**: 取り込んだ要素はすべてデータモデルに保持し、エクスポート時に再出力する。
レンダリング不可の要素も round-trip で欠落させない。表示は可能な範囲で 2D 近似する。

---

## 完了済み

### 基盤・コア
- [x] ZIP 解凍 (JS DecompressionStream, Central Directory ベース) + Wasm FFI + 3-tier js-string 互換 + 汎用 XML パーサー
- [x] Google スライド互換 (ZIP Data Descriptor フラグ対応)
- [x] テーマ解決: theme1.xml カラー 12 色 + フォントスキーム + 全カラー修飾子
- [x] マスター/レイアウト継承: slideMaster/slideLayout パース + placeholder 継承 + `p:clrMapOvr`
- [x] Round-trip: data-ooxml-* SVG ↔ SlideData ↔ OOXML ↔ ZIP
- [x] ライブラリ化: TypeScript 分割 + PptxRenderer クラス + npm + CI/CD

### シェイプ・描画
- [x] AutoShape (~154 プリセット) + カスタムジオメトリ (`a:custGeom`) + コネクタ (直線/折れ線/曲線)
- [x] グループシェイプ: `p:grpSp` 再帰 + 座標変換 + ネスト
- [x] 塗りつぶし: 単色 + グラデーション (線形/放射) + パターン (48 種) + 画像フィル (stretch/tile/crop) + アルファ
- [x] 線/ストローク: 破線 11 種 + 矢印 5 種 + 線結合/線端/複合線 + グラデーション/パターンストローク
- [x] エフェクト: outerShdw / innerShdw / glow / softEdge / reflection → SVG フィルタ
- [x] 画像: PNG/JPEG/GIF/SVG/WebP + クロップ + アルファ + 明度/コントラスト + デュオトーン + カラーチェンジ
- [x] EMF→SVG 変換 (内蔵コンバータ)
- [x] テーブル: セル結合 + ボーダー (対角線) + マージン + テーブルスタイル + 条件書式
- [x] 背景: 単色 + グラデーション + 画像 + パターン
- [x] シェイプリンク / プレースホルダ自動内容 (スライド番号/日付/フッター)
- [x] 3D データ保持 (bevel/extrusion/contour/material/camera/lighting)

### テキスト
- [x] 段落書式 + バレット (文字/自動番号/画像) + ラン装飾 + フォント (Latin/EA/CS/Symbol) + bodyPr + ハイパーリンク
- [x] 折り返し + autoFit (normAutofit フォント縮小) + 縦書き + 多段組 + RTL + タブ
- [x] テキストワープ (prstTxWarp) データ保持 + テキストアウトライン + テキストグラデーション/パターン塗り

### チャート
- [x] 13 種: 棒/折れ線/円/ドーナツ/散布/面/レーダー/バブル/株価/等高線/OfPie (+ 3D variants パース)
- [x] データラベル + dPt + トレンドライン + エラーバー + 複合チャート + view3D/serAx 保持

### データ保持・フォールバック
- [x] SmartArt: mc:Fallback シェイプ描画 + mc:Choice (DiagramML) 原 XML 保持
- [x] OLE / 埋め込み: フォールバック画像描画 + 原 XML 保持
- [x] メディア (動画/音声): ポスターフレーム描画 + 原 XML 保持
- [x] 数式 (OMML): プレーンテキストフォールバック + 原 XML 保持
- [x] WMF / TIFF / 埋め込みフォント: ZIP 内バイナリ round-trip 保持

### メタデータ
- [x] スピーカーノート: `getSlideNotes()` API + round-trip 保持
- [x] コメント: `getSlideComments()` / `getCommentAuthors()` API + round-trip 保持

---

## 未対応 — レンダリング品質向上

### P1: Round-trip データ欠損の修正

現在、編集済みスライドのシリアライズ時に `p:transition` と `p:timing` が出力されない。
未編集スライドは元 XML をそのまま使うため問題ないが、編集したスライドではデータが失われる。

- [x] **アニメーション/トランジション round-trip 保持** — SlideData に `transition_xml` / `timing_xml` フィールドを追加し、パース→保持→シリアライズで欠損を防ぐ。レンダリングは不要（静的表示のみ）
- [x] **隠しスライド検出** — `<p:sld show="0">` をパースし、SlideData に `hidden: Bool` フィールドを追加。API で取得可能にする

### P2: 数式レンダリング (OMML → SVG)

現在はプレーンテキストフォールバックのみ。OMML の構造を SVG の配置に変換することで視覚的に正しい数式表示を実現する。

- [x] **分数** (`m:f`) — 分子/分母を上下に配置 + 分数線を描画
- [x] **上付き/下付き** (`m:sSup` / `m:sSub` / `m:sSubSup`) — baseline シフトで配置
- [x] **根号** (`m:rad`) — √ 記号 + 上線を SVG path で描画
- [x] **大型演算子** (`m:nary`) — ∫ / Σ / Π + 上下限の配置
- [x] **括弧** (`m:d`) — 伸縮する括弧/中括弧の描画
- [x] **アクセント** (`m:acc`) — 上付き記号 (ˆ, ˜, ¯ 等)
- [x] **行列** (`m:m`) — グリッド配置
- [x] **オーバー/アンダーバー** (`m:bar`) — 上線/下線の描画

### P3: テキストワープ視覚レンダリング

現在はデータ保持のみ。prstTxWarp のプリセット (arch, wave, chevron 等) に沿ったパス上テキスト配置を実装する。

- [x] **テキストワープレンダリング** — ワープパス計算 + SVG `<textPath>` によるカーブ沿い配置

### P3: WMF → SVG 変換

EMF コンバータと同様の手法で WMF (16-bit GDI) をパースし SVG に変換する。

- [x] **WMF コンバータ** — WMF レコードパース + GDI 状態管理 + SVG パス/テキスト出力

---

## 未対応 — 新機能

### P3: Office 2016+ チャート (cx namespace)

Office 2016+ で追加されたチャート型。`c:chart` ではなく `cx:chart` XML スキーマを使用する。

- [ ] **ウォーターフォールチャート** (`cx:waterfall`)
- [ ] **ツリーマップ** (`cx:treemap`)
- [ ] **サンバースト** (`cx:sunburst`)
- [ ] **ヒストグラム** (`cx:histogram`)
- [ ] **箱ひげ図** (`cx:boxWhisker`)
- [ ] **ファネル** (`cx:funnel`)

### P4: 追加エフェクト

- [x] **blur エフェクト** (`a:blur`) — SVG `filter: blur()` で容易に対応可能
- [x] **プリセットシャドウ** (`a:prstShdw`) — 現在は `a:outerShdw` のみ対応
- [x] **fillOverlay** (`a:fillOverlay`) — 塗りつぶしのオーバーレイ合成

### P4: その他

- [x] **均等割り付け** (align="just") — SVG にはネイティブの justify がないため、ワード間スペーシングで近似

---

## スコープ外

以下は対応しない。

- **マクロ / VBA** — セキュリティ上の理由
- **3D シェイプ視覚レンダリング** (bevel/extrusion/lighting) — SVG では困難。データは round-trip 保持済み
- **アーティスティックエフェクト** (`a:artisticEffect`) — ビットマップ処理が必要
- **背景除去** (`a14:backgroundRemoval`) — 画像処理が必要
- **Zoom スライド / セクション Zoom** — 対話的機能
- **インクアノテーション** (`p:inkGrp`) — 稀な機能
- **3D モデル** (GLB/GLTF) — mc:Fallback の静止画で表示済み

---

## 優先度サマリー

| 優先度 | 内容 | 表示 | 保持 | 状態 |
|--------|------|------|------|------|
| **P0** | 基盤/テーマ/マスター継承/テキスト/Round-trip | 完全 | 完全 | **完了** |
| **P1** | 塗り/線/グループ/ジオメトリ/テーブル/画像/エフェクト/背景 | 完全 | 完全 | **完了** |
| **P1** | チャート 13 種 + データラベル/トレンドライン/エラーバー/複合 | 完全 | 完全 | **完了** |
| **P1** | ライブラリ公開 (npm/CI/ドキュメント) | — | — | **完了** |
| **P2** | SmartArt / OLE / メディア / 数式 / WMF・TIFF・フォント保持 | FB | 完全 | **完了** |
| **P3** | ノート / コメント (API) | — | 完全 | **完了** |
| **P1** | アニメーション/トランジション round-trip 保持 | — | 保持 | **完了** |
| **P1** | 隠しスライド検出 | — | — | **完了** |
| **P2** | 数式レンダリング (OMML → SVG) | 数式 | 完全 | **完了** |
| **P3** | テキストワープ視覚レンダリング | ワープ | 完全 | **完了** |
| **P3** | WMF → SVG 変換 | 変換 | 完全 | **完了** |
| **P3** | Office 2016+ チャート (cx namespace) | 完全 | 完全 | 未着手 |
| **P4** | 追加エフェクト (blur/prstShdw/fillOverlay) | 完全 | 完全 | **完了** |
| **P4** | 均等割り付け / TIFF デコーダ | 近似 | 完全 | 均等割り付け完了 |
