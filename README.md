# Office Image Extractor

Microsoft Officeファイル（Word、Excel、PowerPoint）に含まれる画像を抽出・閲覧・保存するElectronアプリケーションです。

[![Release](https://img.shields.io/github/v/release/sho11decade/el_officeTo)](https://github.com/sho11decade/el_officeTo/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Electron](https://img.shields.io/badge/Electron-37.2.4-blue.svg)](https://electronjs.org/)

## 📥 ダウンロード

[**最新リリースをダウンロード**](https://github.com/sho11decade/el_officeTo/releases/latest)

### インストール方法

#### Windows
- **インストーラー版**: `Office Image Extractor Setup 1.0.0.exe` - 通常のソフトウェアとしてインストール
- **ポータブル版**: `Office Image Extractor 1.0.0.exe` - インストール不要、USBメモリでも利用可能

## ✨ 特徴

- **対応ファイル形式**
  - Word: .docx, .doc
  - Excel: .xlsx, .xls  
  - PowerPoint: .pptx, .ppt

- **主要機能**
  - 🖱️ **ドラッグ&ドロップ対応** - 日本語ファイル名も完全サポート
  - 🖼️ **画像一括抽出** - 埋め込み画像の自動検出・抽出
  - 👁️ **表示モード切替** - グリッド/リスト表示
  - 🔍 **画像プレビュー** - 拡大表示とメタデータ確認
  - 💾 **柔軟な保存** - 個別保存・一括保存・ZIP出力
  - ⚡ **高性能** - 仮想スクロール・遅延読み込み対応
  - 🔒 **セキュア** - Content Security Policy適用

## 🚀 使用方法

### 1. ファイルの読み込み
- **ドラッグ&ドロップ**: ファイルを画面にドロップ
- **ファイル選択**: 「ファイルを開く」ボタンをクリック

### 2. 画像の閲覧
- 抽出された画像が自動表示されます
- グリッド/リスト表示の切り替え可能
- 画像クリックで詳細表示

### 3. 画像の保存
- **個別保存**: 詳細表示の「保存」ボタン
- **一括保存**: 「すべて保存」ボタン
- **ZIP出力**: 「ZIPエクスポート」ボタン

## 🛠️ 開発者向け情報

### 必要要件
- Node.js 16.0.0 以上
- npm 8.0.0 以上

### セットアップ

```bash
# リポジトリのクローン
git clone https://github.com/sho11decade/el_officeTo.git
cd el_officeTo

# 依存関係のインストール
npm install

# 開発モードで起動
npm start

# 配布用ビルド（Windows）
npm run build:win
```

### 利用可能なスクリプト

```bash
npm start              # アプリケーション起動
npm run dev            # 開発モード起動
npm run build:win      # Windows用ビルド
npm run build:mac      # macOS用ビルド
npm run build:linux    # Linux用ビルド
npm run dist           # 全プラットフォーム用ビルド
```
- npm

### インストール

```bash
# 依存関係のインストール
## 📁 プロジェクト構造

```
office-image-extractor/
├── src/
│   ├── main.js              # メインプロセス（セキュリティ強化）
│   ├── preload.js           # プリロードスクリプト（IPC通信）
│   ├── renderer/            # レンダラープロセス
│   │   ├── index.html       # メインUI（CSP対応）
│   │   ├── styles.css       # モダンスタイルシート
│   │   └── renderer.js      # UI制御（パフォーマンス最適化）
│   └── utils/
│       └── imageExtractor.js # 画像抽出エンジン（メモリ最適化）
├── assets/                  # アプリケーションアセット
│   └── icon.png            # アプリケーションアイコン
├── dist/                   # ビルド出力
├── .github/               # GitHub設定
│   └── copilot-instructions.md
├── package.json
├── OPTIMIZATION_REPORT.md  # 最適化レポート
└── README.md
```

## 🔧 技術スタック

- **フレームワーク**: Electron 37.2.4
- **ランタイム**: Node.js
- **UI**: HTML5, CSS3, JavaScript (ES6+)
- **セキュリティ**: Content Security Policy, Context Isolation

### 主要ライブラリ

| ライブラリ | バージョン | 用途 |
|-----------|----------|------|
| `electron` | ^37.2.4 | デスクトップアプリケーションフレームワーク |
| `mammoth` | ^1.9.1 | Word文書（.docx）処理 |
| `xlsx` | ^0.18.5 | Excel文書（.xlsx）処理 |
| `yauzl` | ^3.2.0 | ZIP（Office文書）解凍 |
| `yazl` | ^3.3.1 | ZIP圧縮（エクスポート用） |
| `sharp` | ^0.34.3 | 高性能画像処理 |
| `electron-builder` | ^26.0.12 | アプリケーションビルド |

## ⚡ パフォーマンス最適化

- **仮想スクロール**: 大量画像の効率的表示
- **遅延読み込み**: 必要時のみ画像読み込み
- **メモリ管理**: 自動ガベージコレクション
- **バッチ処理**: 効率的なファイル処理
- **Intersection Observer**: スムーズなスクロール

## 🔒 セキュリティ機能

- Content Security Policy (CSP) 適用
- Context Isolation 有効化
- Node Integration 無効化
- 安全なIPC通信
- ファイルパス検証

## 🌐 国際化対応

- 日本語ファイル名完全サポート
- UTF-8/Shift_JIS エンコーディング対応
- クロスプラットフォーム文字処理

## ⚠️ 制限事項・既知の問題

### ファイル形式制限
- **.doc、.xls、.ppt**: 部分的なサポート（.docx、.xlsx、.pptx推奨）
- **暗号化ファイル**: パスワード保護されたファイルは非対応
- **SVG画像**: 一部表示に制限がある場合があります

### パフォーマンス
- **大容量ファイル**: 500MB以上のファイルは処理時間が長くなる場合があります
- **同時処理**: 一度に多数のファイルを処理する際はメモリ使用量にご注意ください

### セキュリティ脆弱性
- **xlsx パッケージ**: 既知の脆弱性（プロトタイプ汚染）が存在します
  - 対策: 信頼できるファイルのみ処理してください
  - 将来のバージョンで exceljs への移行を予定

## 🤝 貢献

プロジェクトへの貢献を歓迎します！

### 貢献方法
1. このリポジトリをフォーク
2. 機能ブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. プルリクエストを作成

### 開発ガイドライン
- ES6+ の機能を積極的に使用
- セキュリティを最優先に考慮
- パフォーマンスを意識したコード作成
- 適切なエラーハンドリングを実装

## 📄 ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 🆘 サポート

### 問題を報告する
- バグ報告: [Issues](https://github.com/sho11decade/el_officeTo/issues/new?template=bug_report.md)
- 機能要求: [Issues](https://github.com/sho11decade/el_officeTo/issues/new?template=feature_request.md)

### コミュニティ
- [GitHub Discussions](https://github.com/sho11decade/el_officeTo/discussions) - 質問や提案
- [リリースノート](https://github.com/sho11decade/el_officeTo/releases) - 更新情報

## 🙏 謝辞

このプロジェクトは以下のオープンソースプロジェクトに依存しています：
- [Electron](https://electronjs.org/) - クロスプラットフォームデスクトップアプリ
- [mammoth.js](https://github.com/mwilliamson/mammoth.js) - Word文書処理
- [SheetJS](https://sheetjs.com/) - Excel文書処理
- [Sharp](https://sharp.pixelplumbing.com/) - 高性能画像処理

---

**Office Image Extractor** - Microsoft Officeファイルから画像を簡単抽出 🖼️
