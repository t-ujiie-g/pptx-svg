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

## 🔒 セキュリティ対応（2026-04-27 レビュー）

PPTX を信頼できない入力として扱うブラウザライブラリ。`renderSlideSvg()` の出力を `innerHTML` で挿入する利用パターン（`web/index.html:118`、`web/editing.html:369`、`examples/vanilla/index.html:70`、`examples/react/src/utils/svg.ts:69`）を前提に、細工した PPTX による攻撃ベクタを洗い出した。

### 🔴 高（XSS が成立 — 細工した PPTX で任意コード実行）

- [x] **H1: ハイパーリンク URL のスキーム未検証** ✅ 2026-04-27
  - 調査: `src/renderer/renderer_text.mbt:851-854` で `<a href="${svg_escape(url)}">` を出力。`svg_escape` は XML メタ文字のみ。`.rels` の `Target="javascript:..."` がそのまま出力され、`innerHTML` で DOM に入った後にユーザがクリックすると親ページのオリジンで JS が走る。`<image href>` 経路（外部画像、`renderer.mbt:549-551`）も同様。
  - 対応: `renderer.mbt` に `sanitize_url()` を追加（http/https/mailto/tel/ftp(s)/data:image/* と相対 URL のみ許可、先頭の制御文字も剥がす）。`renderer_text.mbt` のハイパーリンク出力時に通し、空ならリンクラップ自体を省略。

- [x] **H2: SVG 属性ヘルパー `a()` / `da()` が値を XML エスケープしない** ✅ 2026-04-27
  - 調査: `src/renderer/renderer.mbt:160-173` で `" name=\"" + value + "\""` と直結。ユーザー制御文字列（`font_face`、`href`、`ph_type`、`warp_prst`、`3d-cam` 等多数）が値に流れる。`Target` 属性は XML パース時に `decode_entities` を経るため、`Target="x&quot; onload=&quot;alert(1)"` で literal な `"` を混入させ属性ブレイク → 任意属性注入が可能。
  - 対応: `a()` と `da()` の中で `@xml.xml_escape(value)` を呼ぶよう変更。`ai()`/`dai()` は数値専用なのでエスケープ不要パスを保持（`a()` 経由をやめてインライン化）。二重エスケープ回避のため `renderer_text.mbt:843` と `renderer_math.mbt:1275,1294` の `da("math-xml", @xml.xml_escape(...))` から手動エスケープを除去。

- [x] **H3: 外部画像 URL のスキーム未検証** ✅ 2026-04-27
  - 調査: `renderer.mbt:549-551` で `r.target_mode == "External"` のとき Target をそのまま `a("href", href)` に渡す。H2 と組み合わせ属性ブレイクの起点になる。
  - 対応: `resolve_image_href` の External 分岐で `sanitize_url(target)` を経由するよう変更。

### 🟠 中（DoS / ロジック起因）

- [x] **M1: ZIP 解凍サイズ無制限（解凍爆弾）** ✅ 2026-04-27
  - 調査: `lib/zip.ts:130-174` の `extractZip` は `cd.compressedSize` / `uncompressedSize` を CD から読むのみで、`DecompressionStream` の出力に上限がない。1KB の DEFLATE ブロックで GB 級まで膨張可能 → タブハング DoS。
  - 対応: `MAX_INFLATE_BYTES = 256 MiB`（単一エントリ）と `MAX_ARCHIVE_INFLATE_BYTES = 1 GiB`（ZIP 全体）の二重キャップを導入。`inflate()` のストリーミングループで超過時に `reader.cancel()` + throw、`extractZip` で `cd.uncompressedSize` の事前チェック + 総量カウントを追加。

- [x] **M2: EMF/WMF のレコード内点数チェック欠如** ✅ 2026-04-27
  - 調査: `lib/emf-converter.ts:577-590` `readPoints` の `count = dv.getUint32(offset+24)` は無制限で、悪意ある EMF が `count = 0xFFFFFFF0` を渡すと 40 億回ループ→ OOM。`wmf-converter.ts:367` `nPolys` も同様。
  - 対応: EMF `readPoints` に `recordEnd` 引数を追加し `min(declared, capacity, MAX_POINTS_PER_RECORD = 100,000)` で clamp（4 つの呼び出し元で `offset + size` を渡す）。WMF `META_POLYGON` / `META_POLYLINE` / `META_POLYPOLYGON` も同様にレコード末尾でキャパシティ計算 → 超過時はレコードスキップ。

- [x] **M3: EMF `STRETCHDIBITS` のサイズ未検証** ✅ 2026-04-27
  - 調査: `lib/emf-converter.ts:533-548` の `bmpSize = 14 + cbBmi + cbBits`。`cbBmi` / `cbBits` は Uint32 で攻撃者制御。`new Uint8Array(4e9)` で `RangeError` は出るが、その前に巨大バッファ確保でメモリ圧迫。
  - 対応: `offBmi >= 88`（ヘッダ末尾以降）と `bmiEnd <= recEnd` / `bitsEnd <= recEnd`、加算オーバーフローガード（`bmiEnd >= offset + offBmi`）を追加。検証失敗時はレコードスキップ。

- [x] **M4: XML パーサが深いネストでスタックオーバーフロー** ✅ 2026-04-28
  - 調査: `src/xml/xml.mbt:304-393` の `Parser::parse_children` が再帰でネストを処理、深さ上限なし。極端にネストした OOXML で Wasm スタック枯渇。
  - 対応: `parse_children(depth)` に深さカウンタを追加、`max_xml_depth = 1024` 超過で `skip_to_close()`（非再帰の lexical skipper）に切り替えて当該要素を読み飛ばす。`parse_xml` のエントリは `depth = 0` から開始。

- [x] **M5: 文字列連結が O(N²)（性能 DoS）** ✅ 2026-04-28
  - 調査: `src/xml/xml.mbt` の `read_until_char`、`collect_chars`、`decode_entities`、`xml_escape` 等で `result = result + char_to_str(c)` を 1 文字ずつ実行。MoonBit の文字列はイミュータブルなので二次オーダー。`StringBuilder` 禁止という制約下、数 MB の slide.xml で UI が固まる。
  - 対応: `concat_balanced(parts: Array[String])` ヘルパーを追加（バランス型ボトムアップ merge で O(N log N)、`+` のみ使用なので Tier-2/3 ブラウザ互換を維持）。ホットパス 7 関数（`collect_chars` / `read_until_char` / `read_until_str` / `decode_entities` / `parse_name` / `str_suffix` / `str_replace_all` / `str_substring` / `xml_escape`）を Array push + concat_balanced 形式に書き換え。

### 🟡 低（限定的、または利用者次第）

- [x] **L1: ノート/コメント API が XML エンティティを未デコード** ✅ 2026-04-27
  - 調査: `lib/pptx-renderer.ts:1355-1402` の `extractNotesText` / `parseComments` が `<a:t>` / `<p:text>` の inner XML を `match[1]` のまま返す。`textContent` 用途では問題ないが、`innerHTML` に渡されると `<img onerror=...>` 等が再解釈される。
  - 対応: `decodeXmlEntities()` プライベートメソッドを追加（`&amp; &lt; &gt; &quot; &apos; &#xNN; &#NN;` を一括処理）。`extractNotesText` のテキスト抽出、`parseComments` の `text`、`parseCommentAuthors` の `name` / `initials` に適用。

- [x] **L2: 動的 RegExp に rid / rels 値を素で埋め込み** ✅ 2026-04-28
  - 調査: `lib/pptx-renderer.ts:1196-1198` 等で `new RegExp(\`...Id="${rid}"...\`)` のように rid を直接埋める。悪意 PPTX が rid に正規表現メタ文字を仕込むと `SyntaxError`。例外は `ERROR:` 文字列として返り、その文字列が H2 経由で innerHTML に入るとさらに悪用可。
  - 対応: モジュールレベルに `escapeRegex(s)` ヘルパーを追加し、`extractRelTarget` / `updatePresentationXmlFor{Add,Delete,Reorder}` / `resolveRidTarget` / `resolveRelTarget` の `new RegExp` 13 箇所で rid・target・typeSuffix・relType をすべてエスケープ。

### 🟢 良かった点（現状維持）
- XML パーサが DOCTYPE / 外部エンティティを完全スキップする実装で、**XXE は構造的に発生不可能**（`xml.mbt:344-350`）。
- ZIP 内パスの `..` traversal は in-memory `Map<string, ...>` 上にとどまりファイルシステム影響なし。
- EMF/WMF 由来 SVG は `data:image/svg+xml,...` URI として `<image href>` に渡すため、その内部の `<script>` は画像コンテキストで非実行。
- `emfToSvg` / `wmfToSvg` の文字列出力は `escapeXml(...)` 済み。

### 推奨対応順
1. **H2** を最初に（差分最小、H1/H3 を含む属性注入起点が一掃される）
2. **H1**（URL スキーム allowlist）
3. **M1**（解凍サイズ cap）
4. **M2/M3**（EMF/WMF のサイズ・点数 clamp）
5. **L1**（ノート/コメント entity デコード）
6. **M4/M5/L2** は影響範囲を見て段階対応

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

## 未対応 — レンダリング精度向上（新規）

### R1: テキストレンダリング

- [ ] **OpenType フォント機能** — `a:latin` の `numSpacing`/`numForm` 属性未パース。CSS `font-feature-settings` (リガチャ等) 未出力
- [ ] **RTL + 縦書きの組み合わせ** — 右から左テキストの縦書きレイアウトが未検証
- [ ] **wordArtVert（レガシー WordArt 縦書き）** — Office との表示差異の可能性あり

### R2: 図形・エフェクト

- [ ] **コネクタの自動ルーティング** — `st_cxn_id`/`end_cxn_id` はパース済みだが、接続先シェイプの接続ポイントへの自動位置合わせ未実装。表示上はユーザ指定の座標で正しく描画される
- [ ] **SmartArt のネイティブレンダリング** — `dgm:*`（DiagramML）の解釈が未実装。`mc:Fallback` の静的シェイプで表示。色変更・レイアウト変更には非対応
- [ ] **サーフェスチャートの 3D 塗りつぶし** — ワイヤフレーム表示のみ。塗りつぶしサーフェスは未実装

### R3: 数式（OMML）マイナー要素

- [ ] **m:box（ファントムボックス）** — 不可視ボックス要素未対応
- [ ] **m:phant（ファントム要素）** — ファントム要素未対応
- [ ] **m:intLim（積分の極限位置）** — 積分記号の上下極限の配置が一部不完全
- [ ] **行列セルのアラインメント指定** — 行列セル内の揃え (mcJc) が未対応

---

## 未対応 — OOXML 機能（新規）

### F1: アクセシビリティ・メタデータ

- [ ] **代替テキスト（p:cNvPr descr）** — シェイプの alt text をパースし、SVG `<title>` / `aria-label` 属性に出力。round-trip 保持
- [ ] **セクション（p:sectionLst）** — プレゼンテーションのセクション情報をパース。API で取得可能にする

### F2: テーマ・スタイル

- [ ] **テーマバリアント（複数カラースキーム）** — 単一テーマのみ対応。同一テーマの色バリアント切り替え未実装
- [ ] **埋め込みフォント** — フォントファイル（fontFile）の抽出と CSS `@font-face` 生成が未対応。システムフォントにフォールバック

### F3: インタラクション

- [ ] **アクションボタン（p:actionLst）** — ハイパーリンクはサポート済みだが、スライド遷移アクション・OLE アクション等が未対応
- [ ] **トランジション視覚化** — XML は round-trip 保持済み。CSS/SVG アニメーションによるトランジション再生は未実装
- [ ] **アニメーション再生** — XML は round-trip 保持済み。タイムライン・キーフレーム・イベントエンジンが未実装
- [ ] **動画/音声の再生** — メディアファイルは ZIP 内保持済み。`<video>`/`<audio>` 要素生成やプレーヤ UI なし

---

## 未対応 — 編集 API 拡充（新規）

### E1: スライド操作

- [x] **スライドの追加** — `addSlide(afterIdx?, sourceSlideIdx?)` API。presentation.xml sldIdLst + .rels + [Content_Types].xml 自動更新。レイアウト継承対応
- [x] **スライドの削除** — `deleteSlide(slideIdx)` API。presentation.xml + .rels クリーンアップ + スライド再ナンバリング。最後の1枚は削除不可
- [x] **スライドの並べ替え** — `reorderSlides(newOrder)` API。sldIdLst 順序変更 + スライドファイル再ナンバリング。順列バリデーション付き

### E2: シェイプ操作

- [x] **シェイプの追加** — `addShape(slideIdx, geomType, x, y, cx, cy, fillR, fillG, fillB)` API。rect/ellipse/roundRect/line 対応
- [x] **シェイプの削除 API** — `deleteShape(slideIdx, shapeIdx)` API。グループ内シェイプ対応（composite index）
- [x] **シェイプの複製** — `duplicateShape(slideIdx, shapeIdx, dxEmu, dyEmu)` API。オフセット付きコピー
- [x] **グラデーション塗りつぶし編集 API** — `updateShapeGradientFill(slideIdx, shapeIdx, angle, stops)` API。線形グラデーション対応
- [x] **ストローク編集 API** — `updateShapeStroke(slideIdx, shapeIdx, r, g, b, widthEmu, dash)` API。色/幅/破線パターン + 削除対応

#### E2 既知の問題

- [x] **ライン追加時にストロークが設定されない** — `addShape(line)` でデフォルト黒 1pt ストロークを自動設定
- [x] **シェイプへのテキスト追加** — `addShapeText(slideIdx, shapeIdx, text, fontSize, colorR, colorG, colorB)` API を追加。新規段落+ランを作成

### E2.5: テキスト編集

現在 `updateShapeText(slideIdx, shapeIdx, paraIdx, runIdx, text)` で既存ランの文字列のみ変更可能。以下が未対応:

- [x] **段落の追加・削除** — `addParagraph(slideIdx, shapeIdx, text, align)` / `deleteParagraph(slideIdx, shapeIdx, paraIdx)` API
- [x] **ランの追加・削除** — `addRun(slideIdx, shapeIdx, paraIdx, text)` / `deleteRun(slideIdx, shapeIdx, paraIdx, runIdx)` API
- [x] **テキスト書式変更 API（太字/イタリック）** — `updateTextRunStyle(slideIdx, shapeIdx, paraIdx, runIdx, bold, italic)` API。1=set, 0=unset, -1=no change
- [x] **フォントサイズ変更 API** — `updateTextRunFontSize(slideIdx, shapeIdx, paraIdx, runIdx, fontSize)` API。hundredths of a point (1800=18pt)
- [x] **テキスト色変更 API** — `updateTextRunColor(slideIdx, shapeIdx, paraIdx, runIdx, r, g, b)` API。r=-1 で inherit
- [x] **フォントファミリ変更 API** — `updateTextRunFont(slideIdx, shapeIdx, paraIdx, runIdx, fontFace, eaFont, csFont)` API
- [x] **段落配置変更 API** — `updateParagraphAlign(slideIdx, shapeIdx, paraIdx, align)` API。"l"/"ctr"/"r"/"just"/""
- [x] **下線/取り消し線/上付き/下付き変更 API** — `updateTextRunDecoration(slideIdx, shapeIdx, paraIdx, runIdx, underline, strike, baseline)` API

### E3: テーブル操作

- [ ] **行/列の追加・削除**
- [ ] **セルの結合・分割**
- [ ] **セルテキスト編集 API**

### E4: チャート操作

- [ ] **チャートデータの編集** — シリーズ値・カテゴリの変更。現在 ChartShape のシリアライザは空文字列を返す (`serializer.mbt:1390`)。チャート XML は外部ファイル参照 (`chartX.xml`) で保持されるが、編集→再シリアライズのパスがない
- [ ] **チャート種別の変更**
- [ ] **データシリーズの追加・削除**

### E5: 画像操作

- [x] **画像の追加** — メディアファイル追加 + `.rels` + `[Content_Types].xml` 更新が必要
- [x] **画像の差し替え** — メディアファイル置換
- [x] **画像の削除** — シェイプ削除 + 孤立メディアクリーンアップ

---

## 未対応 — エクスポート基盤強化（新規）

### X1: パッケージング

- [x] **`[Content_Types].xml` の動的再生成** — メディア追加/削除時にコンテントタイプを自動更新（E5 で実装済み）
- [x] **`.rels` ファイルの動的再生成** — シェイプ/画像/チャートの変更時にリレーションシップを自動更新（E5 で実装済み）
- [x] **孤立メディアファイルのクリーンアップ** — 削除されたシェイプが参照していたメディアを ZIP から除去（E5 で実装済み）
- [ ] **メディアファイルの重複排除** — 同一画像の複数参照を単一ファイルに統合

### X2: 新規プレゼンテーション作成

- [ ] **空のプレゼンテーション作成** — `presentation.xml` + 必須 `.rels` + テーマの最小構成生成
- [ ] **テンプレートからの作成** — 既存 PPTX をベースに新規スライドを追加

---

## 完了済み — 新機能

### P3: Office 2016+ チャート (cx namespace)

- [x] **ウォーターフォールチャート** (`cx:waterfall`) — 累積バー + コネクタライン + subtotal 対応
- [x] **ツリーマップ** (`cx:treemap`) — slice-and-dice 矩形分割レイアウト
- [x] **サンバースト** (`cx:sunburst`) — 同心円ドーナツ (cos1000/sin1000 再利用)
- [x] **ヒストグラム** (`cx:histogram`) — 棒グラフレンダラー再利用
- [x] **箱ひげ図** (`cx:boxWhisker`) — ボックスプロット (min/Q1/median/Q3/max)
- [x] **ファネル** (`cx:funnel`) — 中央揃え幅比例バー

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
- **カスタム XML データバインディング** — customXmlPart 未対応（利用頻度極めて低い）
- **印刷設定** (`p:sldPr`) — 印刷関連プロパティ（Web ビューアでは不要）

---

## 優先度サマリー

| 優先度 | 内容 | 表示 | 保持 | 状態 |
|--------|------|------|------|------|
| **P0** | 基盤/テーマ/マスター継承/テキスト/Round-trip | 完全 | 完全 | **完了** |
| **P1** | 塗り/線/グループ/ジオメトリ/テーブル/画像/エフェクト/背景 | 完全 | 完全 | **完了** |
| **P1** | チャート 17 種 + データラベル/トレンドライン/エラーバー/複合 | 完全 | 完全 | **完了** |
| **P1** | ライブラリ公開 (npm/CI/ドキュメント) | — | — | **完了** |
| **P2** | SmartArt / OLE / メディア / 数式 / WMF・TIFF・フォント保持 | FB | 完全 | **完了** |
| **P3** | ノート / コメント (API) | — | 完全 | **完了** |
| **P1** | アニメーション/トランジション round-trip 保持 | — | 保持 | **完了** |
| **P1** | 隠しスライド検出 | — | — | **完了** |
| **P2** | 数式レンダリング (OMML → SVG) | 数式 | 完全 | **完了** |
| **P3** | テキストワープ視覚レンダリング | ワープ | 完全 | **完了** |
| **P3** | WMF → SVG 変換 | 変換 | 完全 | **完了** |
| **P3** | Office 2016+ チャート (cx namespace) | 完全 | 完全 | **完了** |
| **P4** | 追加エフェクト (blur/prstShdw/fillOverlay) | 完全 | 完全 | **完了** |
| **P4** | 均等割り付け | 近似 | 完全 | **完了** |
| — | — | — | — | — |
| **R1** | テキスト: OpenType 機能 / RTL+縦書き / wordArtVert | 部分 | 完全 | 未着手 |
| **R2** | コネクタ自動接続 / SmartArt ネイティブ / サーフェス 3D | 近似 | 完全 | 未着手 |
| **R3** | OMML マイナー要素 (box/phant/intLim/行列揃え) | 近似 | 保持 | 未着手 |
| **F1** | 代替テキスト / セクション | — | — | 未着手 |
| **F2** | テーマバリアント / 埋め込みフォント | FB | — | 未着手 |
| **F3** | アクション / トランジション / アニメーション / メディア再生 | — | 保持 | 未着手 |
| **E1** | スライド追加/削除/並べ替え | — | — | **完了** |
| **E2** | シェイプ追加/削除/複製 / グラデーション編集 / ストローク編集 | — | — | **一部残** |
| **E2** | (既知問題) line 追加時ストローク未設定 / シェイプテキスト追加 | — | — | 未着手 |
| **E2.5** | テキスト編集: 段落/ラン追加削除 / 書式変更 (太字/サイズ/色/配置) | — | — | **完了** |
| **E3** | テーブル行列操作 / セル編集 | — | — | 未着手 |
| **E4** | チャートデータ編集 / シリアライズ | — | — | 未着手 |
| **E5** | 画像追加/差替/削除 | — | — | 未着手 |
| **X1** | Content_Types / .rels 動的再生成 / メディアクリーンアップ | — | — | 未着手 |
| **X2** | 新規プレゼンテーション作成 | — | — | 未着手 |
