# VBA Specification Driven Development Template (XVBA Edition)

AI（Antigravity等）を活用し、VSCodeで仕様からVBAコードを生成・管理するためのプロジェクトテンプレートです。

## 🚀 はじめに (Setup)

このテンプレートを使用するには、以下のセットアップが必要です。

### 1. 必要なツールのインストール
- [Visual Studio Code](https://code.visualstudio.com/)
- [XVBA 拡張機能](https://marketplace.visualstudio.com/items?itemName=AndrewButwin.xvba)
- [VBA (serkonda7)](https://marketplace.visualstudio.com/items?itemName=serkonda7.vscode-vba)

### 2. Excel 側の設定
XVBAがExcelを操作できるように許可する必要があります。
1. Excelを開く -> `ファイル` -> `オプション` -> `トラスト センター` (セキュリティ センター)
2. `トラスト センターの設定...` をクリック
3. `マクロの設定` -> **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」** にチェックを入れる

### 3. 初期化
`bin/` フォルダの中に、マクロ有効ブック (`.xlsm`) を配置してください。
（デフォルト設定では `bin/project.xlsm` を読みに行くように `package.json` で設定されています）

---

## 🛠 ワークフロー (Workflow)

### 1. 仕様の作成 (docs/specs)
`docs/specs/` 内にMarkdown形式で仕様を記述します。AI（Antigravity等）にこのファイルを読み込ませ、実装イメージを共有します。

### 2. AIによる実装 (AI Implementation)
AIに仕様を元にしたコード生成を依頼します。生成されたコードは `src/` フォルダ内の odpowiedni なファイル（`.bas`, `.cls`など）に保存させます。

### 3. Excel への同期 (Sync with XVBA)
1. VSCodeで `F1` キーを押し、`XVBA: Run Live Server` または同期コマンドを実行します。
2. `src/` の変更がリアルタイム（または手動同期時）にExcelへ反映されます。

### 4. Git コミット (Version Control)
`src/` フォルダ内のテキストファイルのみをコミット対象とします（バイナリであるExcel本体はコミットしないことを推奨します）。これにより、クリーンな差分管理が可能になります。

---

## 📂 ディレクトリ構成
- `bin/`: Excelバイナリを格納
- `src/`: VBAソースコード本体（Git管理対象）
- `docs/`: 
  - `specs/`: 機能仕様書
  - `designs/`: UI設計図など
- `.vscode/`: VSCode用設定
