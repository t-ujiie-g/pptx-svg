# pptx-svg

PPTX と SVG の双方向変換ライブラリ。外部依存なし、ブラウザのみで動作します。
- [デモサイト(GitHub Pages)](https://t-ujiie-g.github.io/pptx-svg/)

[English](README.md)

## 特徴

- **PPTX → SVG**: PowerPoint スライドを高品質な SVG に変換
- **SVG → PPTX**: SVG を編集して有効な .pptx ファイルにエクスポート（ロスレス往復）
- **サーバー不要**: ZIP 解凍・OOXML パース・SVG 生成すべてブラウザ内で完結
- **外部依存なし**: Wasm バイナリ約 230KB、npm 依存パッケージなし
- **フレームワーク非依存**: React、Vue、Svelte、バニラ JS など何でも利用可能

## インストール

```bash
npm install pptx-svg
```

## クイックスタート

```ts
import { PptxRenderer } from 'pptx-svg';

const renderer = new PptxRenderer();
await renderer.init();                        // Wasm は自動で読み込まれます

const file = await fetch('presentation.pptx');
await renderer.loadPptx(await file.arrayBuffer());

const svgString = renderer.renderSlideSvg(0); // スライド1を SVG に変換
document.getElementById('viewer').innerHTML = svgString;
```

### React

```tsx
import { useEffect, useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-svg';

function SlideViewer({ pptxBuffer }: { pptxBuffer: ArrayBuffer }) {
  const [svg, setSvg] = useState('');
  const rendererRef = useRef<PptxRenderer | null>(null);

  useEffect(() => {
    const renderer = new PptxRenderer();
    rendererRef.current = renderer;
    renderer.init()
      .then(() => renderer.loadPptx(pptxBuffer))
      .then(() => setSvg(renderer.renderSlideSvg(0)));
  }, [pptxBuffer]);

  return <div dangerouslySetInnerHTML={{ __html: svg }} />;
}
```

### バニラ JS（バンドラなし）

```html
<script type="importmap">
{ "imports": { "pptx-svg": "https://cdn.jsdelivr.net/npm/pptx-svg/dist/index.js" } }
</script>
<script type="module">
  import { PptxRenderer } from 'pptx-svg';
  const renderer = new PptxRenderer();
  await renderer.init();
  // ...
</script>
```

完全なサンプルは [`examples/`](examples/) を参照してください。  
- [デモサイト(GitHub Pages)](https://t-ujiie-g.github.io/pptx-svg/)

## API リファレンス

### `PptxRenderer`

```ts
import { PptxRenderer } from 'pptx-svg';

const renderer = new PptxRenderer(options?);
```

**オプション:**

| オプション | 型 | デフォルト | 説明 |
|-----------|------|-----------|------|
| `measureText` | `(text, fontFace, fontSizePx) => number` | Canvas 2D | テキスト幅計測のカスタム関数。 |
| `fontFallbacks` | `Record<string, string[]>` | (組み込み) | フォントフォールバックマッピング。組み込みデフォルトとマージされます。 |
| `logLevel` | `'silent' \| 'error' \| 'warn' \| 'info' \| 'debug'` | `'error'` | コンソール出力の詳細度。 |

**メソッド:**

| メソッド | 戻り値 | 説明 |
|--------|---------|------|
| `init(wasmSource?)` | `Promise<void>` | Wasm モジュールを読み込み。引数なしで自動解決。URL や ArrayBuffer で上書き可能。 |
| `loadPptx(buffer)` | `Promise<{ slideCount }>` | ArrayBuffer から PPTX を読み込み。 |
| `getSlideCount()` | `number` | スライド数。 |
| `isSlideHidden(idx)` | `boolean` | スライドが非表示（`show="0"`）かどうか。 |
| `renderSlideSvg(idx)` | `string` | スライドを SVG 文字列として描画（0始まり）。 |
| `updateSlideFromSvg(idx, svg)` | `string` | 編集済み SVG からスライドデータを更新。`"OK"` または `"ERROR:..."` を返す。 |
| `getSlideOoxml(idx)` | `string` | スライドの OOXML XML を取得。 |
| `exportPptx()` | `Promise<ArrayBuffer>` | 変更を反映した .pptx ファイルとしてエクスポート。 |
| `getSlideXmlRaw(idx)` | `string` | 生のスライド XML（デバッグ用）。 |
| `getEntryList()` | `string[]` | 全 ZIP エントリパス（デバッグ用）。 |

**シェイプ単位の編集メソッド:**

| メソッド | 戻り値 | 説明 |
|--------|---------|------|
| `renderShapeSvg(slideIdx, shapeIdx)` | `string` | 単一シェイプを SVG フラグメントとして描画。 |
| `updateShapeTransform(slideIdx, shapeIdx, x, y, cx, cy, rot)` | `string` | 位置/サイズ/回転を更新（EMU単位）。再描画SVGを返す。 |
| `updateShapeText(slideIdx, shapeIdx, paraIdx, runIdx, text)` | `string` | テキスト内容を更新。再描画SVGを返す。 |
| `updateShapeFill(slideIdx, shapeIdx, r, g, b)` | `string` | 塗りつぶし色を更新（0-255）。再描画SVGを返す。 |

`update*` メソッドはキャッシュされた SlideData を直接更新し、エクスポート用にスライドを変更済みとしてマークし、再描画されたシェイプSVGを返します。使用パターンは [`docs/editing-guide.md`](docs/editing-guide.md) を参照。

**ノート・コメント:**

| メソッド | 戻り値 | 説明 |
|--------|---------|------|
| `getSlideNotes(idx)` | `string[]` | スピーカーノートを段落文字列の配列で取得。 |
| `getSlideComments(idx)` | `SlideComment[]` | コメント一覧（テキスト・著者ID・日時・位置）。 |
| `getCommentAuthors()` | `CommentAuthor[]` | コメント著者一覧（ID・名前・イニシャル）。 |

ノートとコメントは round-trip エクスポート時に自動的に保持されます。

**単位変換ヘルパー:**

```ts
import { pxToEmu, emuToPx, ptToHundredths, hundredthsToPt, degreesToOoxml, ooxmlToDegrees } from 'pptx-svg';

pxToEmu(100)          // 952500 EMU
emuToPx(914400)       // 96 px
ptToHundredths(18)    // 1800
hundredthsToPt(1800)  // 18
degreesToOoxml(90)    // 5400000
ooxmlToDegrees(5400000) // 90
```

**SVG DOM ヘルパー:**

```ts
import { findShapeElement, getShapeTransform, getAllShapes, getSlideScale } from 'pptx-svg';

const shapes = getAllShapes(svgElement);           // 全シェイプ <g> 要素
const g = findShapeElement(svgElement, 0);         // インデックスでシェイプ取得
const transform = getShapeTransform(g);            // { x, y, cx, cy, rot } EMU単位
const scale = getSlideScale(svgElement);           // SVGピクセルあたりのEMU
```

## 対応機能

### 完全対応

- **シェイプ**: AutoShape (rect/ellipse/roundRect/line/プリセット約154種), カスタムジオメトリ (`a:custGeom`), コネクタ (直線/折れ線/曲線)
- **テキスト**: 段落, ラン, バレット (文字/自動/画像), フォント (Latin/EA/CS/Symbol), 太字/斜体/下線/取消線, 上付き/下付き, 文字間隔, カーニング, 大文字化, ハイパーリンク, タブ, RTL, 均等割り付け (word-spacing 分配)
- **テキストボディ**: 縦方向整列, 余白, 自動調整, フォントスケール, 回転, 縦書き, 多段組, テキストワープ (prstTxWarp)
- **塗りつぶし**: 単色, グラデーション (線形/放射 + ストップ), パターン (48プリセット), 画像フィル (stretch/tile/crop)
- **線/ストローク**: 破線11種, 矢印5種, 線端/結合, 複合線, グラデーション/パターンストローク
- **エフェクト**: 外部シャドウ, 内部シャドウ, プリセットシャドウ, グロー, ソフトエッジ, リフレクション, ブラー, フィルオーバーレイ (全て SVG フィルタ)
- **画像**: PNG/JPEG/GIF/SVG, クロップ, アルファ, 明るさ/コントラスト, デュオトーン, カラーチェンジ
- **テーブル**: セル結合, ボーダー (対角線含む), マージン, アンカー, テーブルスタイル, 条件書式
- **チャート**: 棒 (集合/積み上げ/100%積み上げ)/折れ線/円/ドーナツ/散布/面/レーダー/バブル/株価/等高線/ofPie (13種), データラベル, トレンドライン, エラーバー, 複合チャート
- **グループシェイプ**: 再帰ネスト + 座標変換
- **テーマ**: 12テーマカラー, フォントスキーム, 全カラー修飾子
- **マスター/レイアウト継承**: プレースホルダ継承, `p:clrMapOvr`
- **背景**: 単色, グラデーション, 画像, パターン
- **3D**: Round-trip 用データ保持 (ベベル, 押し出し, 輪郭, マテリアル, カメラ, ライティング)
- **プレースホルダ自動内容**: スライド番号, 日付, フッター
- **スピーカーノート**: `getSlideNotes()` で取得可能、round-trip エクスポートで保持
- **コメント**: `getSlideComments()` / `getCommentAuthors()` で取得可能、round-trip エクスポートで保持
- **SmartArt**: `mc:AlternateContent` のフォールバックシェイプで描画、`mc:Choice` (DiagramML) は round-trip 保持
- **OLE / 埋め込みオブジェクト**: `p:oleObj` のフォールバック画像で描画、原 XML は round-trip 保持
- **メディア** (動画/音声): ポスターフレーム画像で描画、原 XML は round-trip 保持
- **EMF / WMF 画像**: 内蔵コンバータで SVG に変換
- **数式** (OMML `m:oMath`): 分数・根号・積分・行列・アクセント・大型演算子の SVG レンダリング、原 XML は round-trip 保持

### 制限付き対応
- **TIFF 画像** — バイナリ round-trip 保持、一部ブラウザで `<img>` 非対応
- **埋め込みフォント** — バイナリ round-trip 保持、描画はシステムフォントフォールバック

### データ保持（描画なし）

- **アニメーション** (`p:timing`) — round-trip エクスポートで保持、静的レンダリングのみ
- **トランジション** (`p:transition`) — round-trip エクスポートで保持、静的レンダリングのみ
- **非表示スライド** — `isSlideHidden()` API で検出可能、`show="0"` は round-trip エクスポートで保持

### スコープ外

- **マクロ / VBA** — セキュリティ上の理由で非対応

## SVG 出力フォーマット

生成される SVG には `data-ooxml-*` 属性が埋め込まれ、OOXML メタデータを保持します。属性の完全なリファレンスは [`docs/svg-specification.md`](docs/svg-specification.md) を参照してください。

## アーキテクチャ

```
[ブラウザ]
  PptxRenderer (TypeScript)
    ├── ZIP 解凍 (DecompressionStream)
    ├── ZIP 構築 (CompressionStream + CRC-32)
    └── FFI ─── WebAssembly GC (MoonBit)
                  ├── XML パーサー
                  ├── OOXML パーサー (型定義, テーマ, テキスト, シェイプ, チャート)
                  ├── SVG レンダラー (シェイプ, テキスト, フィル, ジオメトリ, チャート)
                  ├── SVG パーサー (data-ooxml-* → SlideData)
                  └── OOXML シリアライザー (SlideData → XML)
```

## 開発

### 前提条件

- [MoonBit ツールチェイン](https://moonbitlang.com/download/)
- Node.js 18+

### ビルド

```bash
npm run build          # Wasm + TypeScript + wasm を dist/ にコピー
```

### テスト

```bash
npm test               # 全テスト (MoonBit ユニット + Node.js 統合)
npm run test:moon      # MoonBit ユニットテストのみ
npm run test:node      # Node.js 統合テストのみ
```

### ブラウザテスト

```bash
python3 -m http.server 8765 --directory .
# http://localhost:8765/web/index.html を開く
```

## リリース

npm へのリリースは GitHub Actions でバージョンタグ push 時に自動実行されます:

```bash
# package.json のバージョンを更新後:
git tag v0.1.0
git push origin v0.1.0
```

GitHub リポジトリ設定で `NPM_TOKEN` シークレットの設定が必要です。

## コントリビュート

1. リポジトリをフォーク
2. フィーチャーブランチを作成
3. 既存のコードスタイルに従って変更
4. `src/*/..._test.mbt` に MoonBit ユニットテスト、または `test_fixtures/` に統合テストを追加
5. `npm run build && npm test` で検証
6. プルリクエストを提出

## ライセンス

[MIT](LICENSE)
