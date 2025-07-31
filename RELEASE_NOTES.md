# Office Image Extractor v1.0.0 🎉

Microsoft Officeファイルから画像を簡単に抽出・閲覧・保存するElectronアプリケーションの初回リリースです。

## ✨ 主要機能

### 📄 対応ファイル形式
- **Word**: .docx, .doc
- **Excel**: .xlsx, .xls  
- **PowerPoint**: .pptx, .ppt

### 🖱️ 簡単操作
- **ドラッグ&ドロップ**: ファイルを画面にドロップするだけ
- **日本語ファイル名**: 完全対応
- **ファイル選択**: 従来のファイル選択ダイアログも利用可能

### 🖼️ 画像管理
- **自動抽出**: 埋め込み画像を瞬時に検出・抽出
- **表示切替**: グリッド/リスト表示
- **プレビュー**: 拡大表示とメタデータ確認
- **保存オプション**: 個別・一括・ZIP出力

### ⚡ 高性能
- **仮想スクロール**: 大量画像の快適表示
- **遅延読み込み**: 必要時のみ読み込み
- **メモリ最適化**: 効率的なリソース管理

### 🔒 セキュリティ
- **Content Security Policy**: XSS攻撃防止
- **Context Isolation**: 安全なプロセス分離
- **ファイル検証**: 安全なファイルアクセス

## 📥 ダウンロード

### Windows (推奨)

#### インストーラー版 (92.38 MB)
- `Office-Image-Extractor-Setup-1.0.0.exe`
- 通常のソフトウェアとしてインストール
- デスクトップショートカット作成
- スタートメニューに登録
- アンインストーラー付き

#### ポータブル版 (92.18 MB)
- `Office-Image-Extractor-1.0.0.exe`
- インストール不要
- USBメモリで持ち運び可能
- どこでも実行可能

## 🚀 使い方

1. **アプリを起動**
2. **ファイルをドロップ** または 「ファイルを開く」をクリック
3. **画像を確認** - 自動的に抽出・表示されます
4. **保存** - 個別またはまとめて保存

## 🔧 システム要件

- **OS**: Windows 10/11 (64-bit)
- **メモリ**: 4GB RAM (推奨: 8GB)
- **ストレージ**: 200MB の空き容量
- **その他**: Microsoft Office ファイル形式に対応

## ⚠️ 制限事項

- **.doc/.xls/.ppt**: 制限的サポート（.docx/.xlsx/.pptx 推奨）
- **暗号化ファイル**: パスワード保護ファイルは非対応
- **大容量ファイル**: 500MB以上は処理時間が長くなる場合があります

## 🛠️ 技術仕様

- **フレームワーク**: Electron 37.2.4
- **画像処理**: Sharp (高性能画像処理ライブラリ)
- **Office処理**: mammoth.js (Word), xlsx (Excel), yauzl (PowerPoint)
- **セキュリティ**: CSP適用、Context Isolation有効

## 📝 リリース情報

- **リリース日**: 2025年7月31日
- **バージョン**: 1.0.0
- **ビルド**: Electron 37.2.4
- **ライセンス**: MIT License

## 🤝 サポート

- **バグ報告**: [Issues](https://github.com/sho11decade/el_officeTo/issues)
- **機能要求**: [Feature Requests](https://github.com/sho11decade/el_officeTo/issues/new?template=feature_request.md)
- **ドキュメント**: [README](https://github.com/sho11decade/el_officeTo/blob/main/README.md)

## 🙏 謝辞

このプロジェクトの開発にご協力いただいた全ての方々に感謝いたします。

---

**Happy extracting! 🖼️✨**
