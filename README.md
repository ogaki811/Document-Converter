# Document Converter

EXCELやWordファイルを、ブラウザ上で瞬時に **画像 (PNG)** ファイルへ変換するモダンなWebアプリケーションです。

![License](https://img.shields.io/badge/license-MIT-blue.svg)

## 特徴

*   **完全プライベート**: ファイルはサーバーにアップロードされず、すべての処理がブラウザ内で完結します。
*   **簡単操作**: ドラッグ＆ドロップでファイルをアップロードできます。
*   **対応フォーマット**:
    *   Excel: `.xlsx`, `.xls`
    *   Word: `.docx` のみ（旧形式 `.doc` は非対応）
*   **出力**: 高品質なPNG画像としてダウンロードされます。
*   **モダンなUI**: Tailwind CSSを使用した洗練されたデザイン。

## 使用技術

*   **React 19**: フロントエンドフレームワーク
*   **Vite**: 高速なビルドツール
*   **Tailwind CSS v4**: デザインシステム
*   **SheetJS (xlsx)**: Excelファイルの解析
*   **Mammoth.js**: Wordファイルの解析
*   **html2canvas**: HTML要素の画像化

## セットアップと実行

このプロジェクトは Node.js 環境が必要です。

1.  依存関係のインストール:
    ```bash
    npm install
    ```

2.  開発サーバーの起動:
    ```bash
    npm run dev
    ```

3.  ブラウザでアクセス:
    `http://localhost:5173` (ポート番号は異なる場合があります)

## ビルド

本番用にビルドする場合:

```bash
npm run build
```

## 注意事項

*   複雑なレイアウトのWord文書やExcelシートは、レイアウトが崩れる場合があります（ブラウザベースの簡易レンダリングのため）。
*   非常に大きなファイルの変換には時間がかかる場合があります。

## ライセンス

MIT
