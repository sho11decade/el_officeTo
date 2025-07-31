# Office Image Extractor

Microsoft Officeファイル（Word、Excel、PowerPoint）に含まれる画像を抽出・閲覧・保存するElectronアプリケーションです。

## 特徴

- **対応ファイル形式**
  - Word: .docx, .doc
  - Excel: .xlsx, .xls
  - PowerPoint: .pptx, .ppt

- **主要機能**
  - ドラッグ&ドロップによるファイル読み込み
  - 埋め込み画像の自動抽出
  - グリッド/リスト表示の切り替え
  - 画像詳細表示（拡大表示）
  - 個別画像保存
  - 一括画像保存
  - 画像メタデータ表示

## 使用方法

### 1. アプリケーションの起動

```bash
npm start
```

### 2. ファイルの読み込み

以下のいずれかの方法でOfficeファイルを読み込みます：
- 「ファイルを開く」ボタンをクリック
- ファイルをアプリケーションウィンドウにドラッグ&ドロップ

### 3. 画像の閲覧

- 抽出された画像は自動的に表示されます
- グリッドビューまたはリストビューで表示切り替え可能
- 画像をクリックすると詳細表示モーダルが開きます

### 4. 画像の保存

- **個別保存**: 画像詳細モーダルの「保存」ボタン
- **一括保存**: ヘッダーの「すべて保存」ボタン

## 開発環境のセットアップ

### 必要要件

- Node.js (v14 以上)
- npm

### インストール

```bash
# 依存関係のインストール
npm install

# 開発モードで起動
npm run dev

# アプリケーションのビルド
npm run build

# 配布用パッケージの作成
npm run dist
```

## プロジェクト構造

```
office-image-extractor/
├── src/
│   ├── main.js              # メインプロセス
│   ├── preload.js           # プリロードスクリプト
│   ├── renderer/            # レンダラープロセス
│   │   ├── index.html       # メインUI
│   │   ├── styles.css       # スタイルシート
│   │   └── renderer.js      # UI制御JavaScript
│   └── utils/
│       └── imageExtractor.js # 画像抽出ユーティリティ
├── assets/                  # アプリケーションアセット
├── package.json
└── README.md
```

## 技術スタック

- **フレームワーク**: Electron
- **ランタイム**: Node.js
- **UI**: HTML5, CSS3, JavaScript (ES6+)

### 主要ライブラリ

- `electron` - デスクトップアプリケーションフレームワーク
- `mammoth` - Word文書処理
- `xlsx` - Excel文書処理
- `yauzl` - ZIP（Office文書）解凍
- `sharp` - 画像処理
- `electron-builder` - アプリケーションビルド

## 制限事項

- .doc、.xls、.ppt形式は部分的なサポート
- 一部の埋め込み画像形式は抽出できない場合があります
- 大容量ファイルの処理には時間がかかる場合があります

## ライセンス

MIT License

## 貢献

プルリクエストやイシューの報告を歓迎します。

## サポート

問題や質問がある場合は、GitHubのIssuesページでお知らせください。
