# CellFocus — Excel 集中編集アドイン

選択したセル範囲を独立したタブに分離し、リボンなしで集中編集できる Excel Office.js アドイン。

## 主な機能

- **範囲タブ** — 複数のセル範囲を同時に8タブまで管理
- **完全な書式再現** — フォント・背景色・罫線・数値フォーマット・セル結合をそのまま表示
- **双方向同期** — アドイン側の編集は即座に元セルへ反映、Excel側の変更もアドインに反映
- **キーボードショートカット** — `Ctrl+Shift+F` で選択範囲を即座に開く
- **右クリックメニュー** — セル右クリック →「CellFocusで開く」
- **ダークテーマ** — 目に優しいミニマルUI

## インストール（サイドロード）

1. `manifest.xml` をダウンロード
2. Excel Desktop を開く
3. **ファイル** → **オプション** → **トラストセンター** → **トラストセンターの設定** → **信頼されているアドイン カタログ**
4. または: **挿入** → **アドイン** → **個人用アドイン** → **マイアドインの管理** → **カタログの共有フォルダー** を使用

> **注意**: Excel Desktop (Windows) 専用です。Excel Online は非対応。

## ローカル開発

**前提条件**: Node.js 18+, Excel Desktop (Windows)

```bash
# 依存パッケージのインストール
npm install

# HTTPS開発証明書の生成（初回のみ）
npx office-addin-dev-certs install

# 開発サーバー起動 + Excelへ自動サイドロード
npm start

# プロダクションビルド
npm run build

# アドインの停止・アンロード
npm stop
```

開発時は `manifest.dev.xml`（localhost:3000）が使用されます。

## デプロイ

CellFocus は 2 つのデプロイ経路をサポートします。

### Microsoft Marketplace / AppSource 公開配布

公開 HTTPS ホスティングに `dist/` を配置し、Partner Center に `dist/manifest.xml` を提出します。

GitHub Pages を使う場合:

```bash
npm run build:github-pages
```

このリポジトリの標準公開 URL:

```text
https://younnieCutler.github.io/excel-new-tab-cell-range/
```

`main` に push すると GitHub Actions が `dist/` を `gh-pages` ブランチへ公開します。

任意の公開 HTTPS ホストを使う場合:

```bash
npm run build:marketplace -- \
  --base-url https://cellfocus.example.com \
  --support-url https://cellfocus.example.com/support.html

npm run validate:marketplace
```

詳細: [docs/marketplace-release-ko.md](docs/marketplace-release-ko.md)

### 顧客テナント配布

顧客環境の HTTPS 静的ホスティング + Microsoft 365 Admin 中央配布で使います。

```bash
npm run build:customer -- --base-url https://cellfocus.customer.example
```

生成された `dist/` を顧客環境の HTTPS 静的ホスティングに配置し、`dist/manifest.xml` を Microsoft 365 Admin Center で組織配布します。

詳細: [docs/customer-deployment-ko.md](docs/customer-deployment-ko.md)

## アーキテクチャ

```
src/
├── taskpane/
│   ├── taskpane.js          # エントリーポイント
│   ├── taskpane.html / .css # UI（ダークテーマ）
│   └── modules/
│       ├── i18n.js          # 日本語デフォルト・英語フォールバック
│       ├── utils.js         # アドレス解析・書式変換
│       ├── tabManager.js    # タブ管理（最大8タブ）
│       ├── syncEngine.js    # Excel ↔ アドイン双方向同期
│       └── gridRenderer.js  # HTMLテーブル書式レンダラー
└── commands/
    └── commands.js          # 右クリック・ショートカット処理
```

**技術スタック**: Office.js (Shared Runtime) / Vanilla JS / Webpack / HTTPS 静的ホスティング

## ライセンス

MIT
