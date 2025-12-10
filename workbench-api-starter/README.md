# KPMG Workbench API Starter

KPMG Workbench API を使用するための開発環境セットアップツールです。

## クイックスタート

### 1. セットアップ

`setup.bat` をダブルクリックして実行します。

```
setup.bat
```

このスクリプトは以下を自動で行います:
- Python のインストール確認
- 仮想環境の作成
- 依存関係のインストール
- `.env` ファイルの作成

### 2. API Key の設定

`.env` ファイルを開いて、API Key を設定します:

```
KPMG_WORKBENCH_API_KEY=your_actual_api_key
KPMG_CHARGE_CODE=your_charge_code
```

API Key は [KPMG Workbench Developer Portal](https://developer.workbench.kpmg/) から取得できます。

### 3. テストの実行

`run.bat` をダブルクリックして API テストを実行します。

```
run.bat
```

## 環境要件

- **OS**: Windows 10 / 11
- **Python**: 3.8 以上
- **ネットワーク**: KPMG 社内ネットワークまたは VPN 接続

## ファイル構成

```
workbench-api-starter/
├── setup.bat          # セットアップスクリプト
├── run.bat            # テスト実行スクリプト
├── requirements.txt   # Python 依存関係
├── .env.example       # 環境変数テンプレート
├── .env               # 環境変数 (setup.bat で作成)
├── test_api.py        # API テストスクリプト
└── README.md          # このファイル
```

## テスト対象 API

| API | 説明 | エンドポイント |
|-----|------|---------------|
| Chat Completion | チャット生成 | `/genai/azure/inference/chat/completions` |
| Embeddings | テキスト埋め込み | `/genai/azure/inference/embeddings` |
| Text Translation | テキスト翻訳 | `/translator/azure/text/translate` |
| Text Chunking | テキスト分割 | `/text-chunking/microsoft/simple/simple` |

## よくある質問

### Q: Python がインストールされていないと表示される

以下のリンクから Python をダウンロードしてインストールしてください:
- https://www.python.org/downloads/

インストール時に **「Add Python to PATH」** にチェックを入れてください。

### Q: API Key はどこで取得できますか？

[KPMG Workbench Developer Portal](https://developer.workbench.kpmg/) にログインし、
サブスクリプションページから API Key を取得してください。

### Q: チャージコードとは何ですか？

プロジェクトに紐づいた費用計上コードです。
プロジェクトマネージャーに確認してください。

### Q: API テストが失敗する

以下を確認してください:
1. `.env` ファイルに正しい API Key が設定されているか
2. KPMG 社内ネットワークまたは VPN に接続されているか
3. サブスクリプションが有効か (Developer Portal で確認)

## 参考リンク

- [KPMG Workbench Developer Portal](https://developer.workbench.kpmg/)
- [API ドキュメント](https://developer.workbench.kpmg/docs)
