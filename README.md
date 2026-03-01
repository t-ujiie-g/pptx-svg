# pptx-svg

MoonBit (wasm-gc) で構築した、ブラウザで動作する PPTX ビューア／エディタ。
ZIP 解凍・OOXML 処理・SVG レンダリング・PPTX エクスポートをすべてクライアントサイドで完結させます。

## 特徴

- **双方向変換**: PPTX → SVG → 編集 → PPTX エクスポート
- **サーバー不要**: ZIP 解凍・OOXML パース・SVG 生成・ZIP 再構築すべてブラウザ内で完結
- **ロスレス往復**: SVG に `data-ooxml-*` 属性を埋め込み、OOXML メタデータを保持
- **軽量**: Wasm バイナリ約 24KB、外部依存なし

## 技術スタック

| レイヤー | 技術 |
|---------|------|
| ロジック | [MoonBit](https://moonbitlang.com/) → WebAssembly GC (wasm-gc) |
| レンダリング | SVG (data-ooxml-* 属性付き) |
| ZIP 解凍/生成 | ブラウザ標準 `DecompressionStream` / `CompressionStream` API (JS 側) |
| 文字列 FFI | `use-js-builtin-string: true` (MoonBit String = JS String, ゼロコスト) |
| ホスト層 | 素の JavaScript ES Modules |

## アーキテクチャ

```
[ブラウザ]
  ┌─────────────────────────────────────────────────┐
  │  web/index.html                                 │
  │  lib/ → dist/  ← PptxRenderer クラス           │
  │    │                                            │
  │    ├─ ZIP 解凍 (DecompressionStream)            │
  │    ├─ ZIP 生成 (CompressionStream + CRC-32)     │
  │    │                                            │
  │    └─ FFI ──────────────────────────────┐      │
  │                                          │      │
  │  [WebAssembly GC]                        │      │
  │  _build/.../main.wasm                    │      │
  │    src/ffi/         ← FFI 宣言          │      │
  │    src/xml/         ← 汎用 XML パーサー  │      │
  │    src/ooxml/       ← OOXML 型+パーサー  │      │
  │    src/renderer/    ← SlideData → SVG    │      │
  │    src/svg_parser/  ← SVG → SlideData    │      │
  │    src/serializer/  ← SlideData → XML    │      │
  │    src/main/        ← 公開 API ──────────┘      │
  └─────────────────────────────────────────────────┘
```

**データフロー (Round-trip):**
1. ユーザーが .pptx ファイルをドロップ
2. JS が ZIP をパース・解凍し、エントリを Map に保存
3. `render_slide_svg(idx)` → data-ooxml-* 属性付き SVG
4. (ブラウザ側で SVG を編集)
5. `update_slide_from_svg(idx, svg)` → SlideData をキャッシュに更新
6. `exportPptx()` → 変更された slide XML で ZIP を再構築 → .pptx ダウンロード

## ブラウザ互換性

| ブラウザ | wasm-gc | js-string builtins | 動作 |
|---------|---------|-------------------|------|
| Chrome 117+ | ✓ | Tier-1 (builtins) | ✓ |
| Edge 117+ | ✓ | Tier-1 (builtins) | ✓ |
| Chrome 115–116 | ✓ | Tier-2 (importedStringConstants) | ✓ |
| Chrome 111–114 | ✓ | Tier-3 (full manual) | ✓ |
| Firefox 120+ | ✓ | Tier-1 (builtins) | ✓ |
| Safari 17+ | ✓ | Tier-1 (builtins) | ✓ |
| Chrome < 111 | ✗ | — | ✗ (wasm-gc 未対応) |

> `lib/wasm-compat.ts` は 3 段階フォールバックで Wasm を初期化します。

## クイックスタート

### 前提条件

- [MoonBit toolchain](https://moonbitlang.com/download/) (`moon` コマンド)
- Python 3.9+ (HTTP サーバー用)
- Chrome 111+ / Firefox 120+ / Safari 17+

### ビルド

```bash
moon build --target wasm-gc --release
# → _build/wasm-gc/release/build/main/main.wasm (~24KB)
```

### 開発サーバー起動

```bash
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html
```

### テスト (JS ホスト層)

```bash
node test_fixtures/test_node.mjs
```

## プロジェクト構造

```
pptx-svg/
├── moon.mod.json              # MoonBit プロジェクト設定（外部依存なし）
├── package.json               # npm パッケージ定義
├── src/                       # MoonBit (Wasm-GC)
│   ├── ffi/ffi.mbt            # JS host 関数の FFI 宣言
│   ├── xml/xml.mbt            # 汎用 XML パーサー（DOM ツリー）
│   ├── ooxml/ooxml.mbt        # OOXML 型定義 + PPTX slide XML パーサー
│   ├── renderer/renderer.mbt  # SlideData → SVG (data-ooxml-* 属性付き)
│   ├── svg_parser/svg_parser.mbt  # SVG → SlideData (逆変換)
│   ├── serializer/serializer.mbt  # SlideData → OOXML slide XML
│   └── main/main.mbt          # Wasm エクスポート API + スライドキャッシュ
├── lib/                       # TypeScript ライブラリソース
│   ├── index.ts               # 公開 API re-exports
│   ├── pptx-renderer.ts       # PptxRenderer クラス (コア API)
│   ├── wasm-compat.ts         # 3-tier Wasm js-string フォールバック
│   ├── zip.ts                 # ZIP 解凍 / 構築
│   └── utils.ts               # bytesToBase64, crc32
├── dist/                      # コンパイル済み JS + .d.ts (tsc 出力)
├── web/
│   ├── host.js                # レガシー JS ホスト (参考用)
│   └── index.html             # デモ UI (dist/ をインポート)
└── test_fixtures/
    ├── minimal.pptx           # 2 スライドのテスト用最小 PPTX
    └── test_node.mjs          # Node.js テスト (JS レイヤーのみ)
```

**モジュール依存関係 (サイクルなし):**
```
main → renderer   → ooxml → xml
     → svg_parser → ooxml → xml
     → serializer → ooxml
     → ffi
```

## API リファレンス

### Wasm エクスポート関数

| 関数 | 戻り値 | 説明 |
|------|--------|------|
| `initialize_pptx()` | `"OK:<count>"` or `"ERROR:..."` | PPTX を初期化しスライド数を取得 |
| `get_slide_count()` | `Int` | スライド数 |
| `get_slide_xml_raw(idx)` | `String` | 生の slide XML |
| `get_entry_list()` | `String` | ZIP エントリ一覧 (改行区切り) |
| `render_slide_svg(idx)` | `String` | data-ooxml-* 属性付き SVG |
| `update_slide_from_svg(idx, svg)` | `"OK"` or `"ERROR:..."` | SVG から SlideData を更新 |
| `get_slide_ooxml(idx)` | `String` | OOXML slide XML (変更済みなら再生成) |
| `get_modified_entries()` | `String` | 変更エントリ (path\tcontent\n 形式) |

### JS API (PptxRenderer クラス)

```javascript
await renderer.init(wasmUrl)              // Wasm モジュール初期化
await renderer.loadPptx(arrayBuffer)      // PPTX ロード → { slideCount }
renderer.renderSlideSvg(slideIdx)         // SVG 文字列取得
renderer.updateSlideFromSvg(idx, svg)     // SVG → 内部データ更新
renderer.getSlideOoxml(idx)               // OOXML XML 取得
await renderer.exportPptx()               // PPTX ArrayBuffer エクスポート
```

## 開発ノート

### MoonBit 制約事項

- **整数補間禁止**: `"\{n}"` は `fromCharCodeArray` を使うため Chrome 117+ 未満で動かない → `int_to_str(n)` ヘルパーで代替
- **StringBuilder 禁止**: `to_string()` が `fromCharCodeArray` を呼ぶ → `+` (concat) で文字列を組み立てる
- **Char→String**: `@ffi.ffi_char_code_to_str(Char::to_int(c))` を使用
- **外部パッケージ不使用**: 現在のコンパイラ (2026年版) と非互換のため全て自前実装
- **pub(all)**: 外部パッケージから構造体を構築するには `pub(all) struct` が必要

## 既知の制限

- SmartArt / チャートはグレーフォールバック
- EMF/WMF 画像は非対応
- アニメーション・トランジションは無視
- Node.js では動作しない（wasm-gc はブラウザのみ）
