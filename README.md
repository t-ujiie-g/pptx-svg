# pptx-render

MoonBit (wasm-gc) で構築した、ブラウザで動作する PPTX ビューア／エディタ。
ZIP 解凍・OOXML 処理・SVG レンダリングをすべてクライアントサイドで完結させ、
サーバーへのアップロードなしにスライドを表示・編集します。

## 技術スタック

| レイヤー | 技術 |
|---------|------|
| ロジック | [MoonBit](https://moonbitlang.com/) → WebAssembly GC (wasm-gc) |
| レンダリング | SVG |
| ZIP 解凍 | ブラウザ標準 `DecompressionStream` API (JS 側) |
| 文字列 FFI | `use-js-builtin-string: true` (MoonBit String = JS String, ゼロコスト) |
| ホスト層 | 素の JavaScript ES Modules |

## アーキテクチャ

```
[ブラウザ]
  ┌────────────────────────────────────────────┐
  │  web/index.html                            │
  │  web/host.js   ← PptxRenderer クラス      │
  │    │                                       │
  │    ├─ ZIP 解凍 (DecompressionStream)       │
  │    │   テキスト → Map<path, string>         │
  │    │   バイナリ → Map<path, Uint8Array>     │
  │    │                                       │
  │    └─ FFI ─────────────────────────────┐  │
  │                                         │  │
  │  [WebAssembly GC]                       │  │
  │  _build/.../main.wasm                   │  │
  │    src/ffi/ffi.mbt  ← FFI 宣言         │  │
  │    src/main/main.mbt ← OOXML ロジック  │  │
  │      initialize_pptx()                  │  │
  │      get_slide_count()     ─────────────┘  │
  │      get_slide_xml_raw(idx)                │
  │      get_entry_list()                      │
  └────────────────────────────────────────────┘
```

**データフロー:**
1. ユーザーが .pptx ファイルをドロップ
2. JS が ZIP をパース・解凍し、エントリを Map に保存
3. MoonBit `initialize_pptx()` が FFI 経由で `presentation.xml` を取得してスライド数をカウント
4. MoonBit `get_slide_xml_raw(idx)` が各スライドの XML を返す
5. (Phase 2 以降) MoonBit `render_slide_svg(idx)` が SVG 文字列を生成

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

> **Note:** `host.js` は 3 段階フォールバックで初期化を試みます。
> コンソールで `[pptx] Wasm init: tier-X` を確認できます。

## クイックスタート

### 前提条件

- [MoonBit toolchain](https://moonbitlang.com/download/) (`moon` コマンド)
  - 動作確認バージョン: `0.1.20260209`
- Python 3.9+ (HTTP サーバー・スクリプト実行用)
- Chrome 111+ / Firefox 120+ / Safari 17+

### ビルド

```bash
moon build --target wasm-gc --release
# → _build/wasm-gc/release/build/main/main.wasm (約 1.6KB)
```

### 開発サーバー起動

```bash
python3 -m http.server 8765 --directory .
# → http://localhost:8765/web/index.html
```

ブラウザで上記 URL を開き、.pptx ファイルをドロップして動作確認。

### テスト (JS ホスト層)

```bash
node test_fixtures/test_node.mjs
```

Node.js 側の ZIP 解凍・スライドカウントロジックをテストします (Wasm 不使用)。

## プロジェクト構造

```
pptx-render/
├── moon.mod.json              # MoonBit プロジェクト設定（外部依存なし）
├── README.md
├── TODO.md                    # フェーズ別進捗管理
│
├── src/
│   ├── ffi/
│   │   ├── ffi.mbt            # JS host 関数の FFI 宣言
│   │   └── moon.pkg.json
│   ├── main/
│   │   ├── main.mbt           # エクスポート関数・OOXML ロジック
│   │   └── moon.pkg.json      # link 設定・エクスポート名・use-js-builtin-string
│   ├── ooxml/                 # (Phase 2+) OOXML パーサー
│   ├── renderer/              # (Phase 2+) SVG レンダラー
│   ├── editor/                # (Phase 5+) 編集操作
│   ├── xml/                   # (Phase 2+) XML パーサー
│   └── zip/                   # (将来) ZIP ライブラリ (現在は JS 側で処理)
│
├── web/
│   ├── index.html             # デモ UI (ファイルドロップ・XML ビューア)
│   └── host.js                # JS ホスト層 (ZIP 解凍・Wasm 初期化・FFI)
│
├── scripts/
│   └── gen_string_constants.py  # STRING_CONSTANTS 自動生成ツール
│
└── test_fixtures/
    ├── minimal.pptx           # 2 スライドのテスト用最小 PPTX
    └── test_node.mjs          # Node.js テスト (JS レイヤーのみ)
```

## 開発ノート

### MoonBit API (コンパイラ 0.1.20260209)

| 廃止 API | 現行 API |
|---------|---------|
| `unsafe_char_at(i)` | `get_char(i).unwrap()` |
| `String::substring(a, b)` | `s[a:b]` (ただし CreatingViewError を raise) |
| `"\{int_val}"` (整数補間) | `int_to_str(n)` (main.mbt 内のヘルパー) |

> **整数補間の注意:** MoonBit の `"\{n}"` は内部で `fromCharCodeArray` を呼び出します。
> この関数は Chrome 117+ の `{ builtins: ['js-string'] }` がなければ提供できないため、
> 整数→文字列変換は `int_to_str()` ヘルパーで代替しています。

### FFI 設計 (wasm-gc)

```moonbit
// MoonBit 側: JS 関数をインポート
pub fn ffi_get_file(path : String) -> String = "pptx_ffi" "get_file"
pub fn ffi_log(msg : String) -> Unit = "pptx_ffi" "log"

// JS 側: importObject で提供
'pptx_ffi': {
  get_file: (path) => this.#files.get(path) ?? '',
  log:      (msg)  => console.log('[pptx]', msg),
}
```

### STRING_CONSTANTS の更新

MoonBit のソースコードを変更して文字列リテラルを追加・削除した場合、
`host.js` の `STRING_CONSTANTS` 配列を更新する必要があります。

```bash
moon build --target wasm-gc --release
python3 scripts/gen_string_constants.py --update
```

`--update` オプションを省略すると、更新すべき配列の内容がターミナルに出力されます。

### 外部パッケージについて

`bobzhang/zip` と `ruifeng/XMLParser` は現在のコンパイラ (2026年版) と非互換のため不使用。
ZIP 解凍は JS の `DecompressionStream` API で代替。XML パーサーは Phase 2 で自前実装予定。

## フェーズ進捗

| フェーズ | 内容 | 状態 |
|---------|------|------|
| Phase 1 | Foundation — ZIP 解凍・FFI・スライドカウント | ✅ 完了 |
| Phase 2 | 基本 SVG レンダリング (rect / ellipse / text / image) | 🔲 |
| Phase 3 | テーマ・カラー継承解決 | 🔲 |
| Phase 4 | プリセット図形ライブラリ (200 種) | 🔲 |
| Phase 5 | 編集操作 (移動・リサイズ・テキスト編集・undo/redo) | 🔲 |
| Phase 6 | 高度機能 (テーブル・PPTX エクスポート・埋め込みフォント) | 🔲 |

## 既知の制限

- SmartArt / チャートはグレーフォールバックで表示（Phase 2 初期）
- EMF/WMF 画像は非対応
- アニメーション・トランジションは無視
- Node.js では動作しない（wasm-gc は Node.js 非対応、ブラウザのみ）
