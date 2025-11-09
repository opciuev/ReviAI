# 🤖 ReviAI - AI評審作業システム

ExcelからPDFを生成し、Gemini 2.5 ProでAI評審を実行し、結果をフォーマット済みExcelファイルとして保存する自動化システムです。

---

## ✨ 主な機能

- 📄 **Excel→PDF変換**: 選択したシートをA4横向きPDFとして出力
- 🤖 **AI評審**: Gemini 2.5 Proによる自動文書評価
- 📊 **結果保存**: 8列の構造化表をフォーマット済みExcelで出力
- 🎨 **モダンなGUI**: PySide6による3ステップウィザード形式
- 🌏 **日本語完全対応**: すべてのテキスト処理で日本語をサポート

---

## 🚀 クイックスタート

### 1. 前提条件

- ✅ Python 3.8以降
- ✅ Microsoft Excel（xlwings依存）
- ✅ Gemini API key（[こちらから取得](https://aistudio.google.com/apikey)）

### 2. インストール

```bash
cd d:\ReviAI
pip install -r requirements.txt
```

### 3. 設定

`config.ini`を開いて、Gemini API keyを設定:

```ini
[API]
gemini_api_key = あなたのAPIキーをここに貼り付け
```

### 4. 起動

```bash
cd src
python main.py
```

---

## 📖 使用方法

### 第1段階: Excel→PDF生成
1. Excelファイルを選択
2. 出力したいシートをチェック
3. バージョン番号を入力（例: 6）
4. 「PDF生成」をクリック

### 第2段階: AI評審
1. 生成されたPDFが自動表示されます
2. 「AI評審開始」をクリック
3. 完了まで待機（30秒〜2分）

### 第3段階: 結果保存
1. 回数を入力（例: 6 → 「第六回.xlsx」）
2. 「Excelに保存」をクリック

---

## 📁 プロジェクト構造

```
ReviAI/
├── config.ini                # 設定ファイル
├── prompt_template.txt       # 評審プロンプト（編集可能）
├── requirements.txt          # Python依存関係
├── README.md                 # このファイル
├── 開発計画.md               # 詳細な開発計画書
├── logs/                     # ログファイル
├── output/
│   ├── pdfs/                 # 生成PDF
│   └── results/              # 評審結果Excel
└── src/                      # ソースコード
```

---

## 🔧 主要技術

| 技術 | 用途 |
|-----|------|
| **PySide6** | GUI |
| **xlwings** | Excel→PDF変換 |
| **Gemini 2.5 Pro** | AI評審 |
| **Pydantic** | データ検証 |
| **openpyxl** | Excel出力 |

---

## 💰 コスト

- **1文書あたり**: 約5.3円（10ページPDF + 100行出力）
- **Gemini API**: $1.25/1M入力tokens、$10/1M出力tokens

---

## 📝 ログ

すべての操作は`logs/app.log`に記録されます。問題が発生した場合は、このファイルを確認してください。

---

## 🐛 トラブルシューティング

### API keyエラー
→ `config.ini`の`gemini_api_key`を確認

### Excelが見つからない
→ Microsoft Excelがインストールされているか確認

### AI評審が失敗
→ PDFが読み取り可能か、`logs/app.log`でエラー詳細を確認

---

## 📚 詳細ドキュメント

詳細な技術情報は[開発計画.md](開発計画.md)を参照してください。

---

## 🎯 次のステップ

1. `config.ini`でAPI keyを設定
2. テストExcelファイルで動作確認
3. `prompt_template.txt`を必要に応じてカスタマイズ
4. 本番運用開始

---

**ReviAI - 文書評審を自動化し、効率を最大化**
