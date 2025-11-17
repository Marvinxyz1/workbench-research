# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

このリポジトリは **KPMG Workbench 戦略評価レポート** の自動生成システムです。Python + python-docx を使用して、日本語と英語の混在した Word 文書(.docx)を生成します。

主な用途:
- KPMG Workbench の評価フレームワークに基づいた詳細レポートの生成
- 学習パス、認証要件、技術評価など複数の「維度(dimension)」セクションを含む
- ハイパーリンク、テーブル、KPMGブランドカラーを含む書式設定済みドキュメント

## 重要なコマンド

### メイン実行方法

**完全版レポートを生成** (推奨):
```bash
cd scripts
python main.py
```
出力先: `generated_docs/KPMG_Workbench戦略評価レポート_完成版.docx`

**スタンドアロン版（日本語版のみ）**:
```bash
python generate_assessment_doc_jp.py
```
出力先: プロジェクトルートに直接生成

### 依存関係インストール

```bash
pip install python-docx
```

## プロジェクト構造

### 主要ファイル

- `scripts/main.py` - メイン実行エントリーポイント。すべてのモジュールを統合
- `scripts/utils.py` - 共通ユーティリティ関数（ハイパーリンク、見出しスタイル設定など）
- `scripts/modules/` - 各セクションを生成するモジュール群
- `generate_assessment_doc_jp.py` - スタンドアロン版ジェネレータ（維度0のみ）

### モジュール構成

`scripts/modules/` 内の各モジュールは独立したセクションを生成:

- `cover.py` - 表紙、目次、エグゼクティブサマリー
- `dimension_0.py` - 事前準備（学習パス、認証要件）- **最も詳細で重要**
- `dimension_1.py` - 技術能力と効率評価
- `dimension_2.py` - Agentic AI核心能力評価
- `dimension_3.py` - 商業価値と顧客応用
- `dimension_4.py` - 学習資源とコミュニティサポート
- `dimension_5.py` - 戦略価値と組織影響
- `dimension_6.py` - リスクとコンプライアンス評価
- `dimension_7.py` - チーム推進と運営考慮
- `appendix.py` - 付録（参考資料、タイムライン）

各モジュールは `add_dimension_X(doc)` という関数をエクスポートし、python-docx の `Document` オブジェクトを受け取って直接編集します。

## 開発パターン

### 新しいセクション/維度の追加

1. `scripts/modules/dimension_X.py` を作成
2. `add_dimension_X(doc)` 関数を実装
3. `scripts/modules/__init__.py` に import を追加
4. `scripts/main.py` で関数を呼び出し

### 既存セクションの編集

1. 該当する `scripts/modules/dimension_X.py` を直接編集
2. `python scripts/main.py` で再生成
3. Word で開いて確認

### スタイリング規約

- **KPMGブランドカラー**を使用:
  - `KPMG_BLUE = RGBColor(0, 94, 184)`
  - `KPMG_DARK_BLUE = RGBColor(0, 51, 141)`
- **日本語フォント**: `Meiryo UI` (東アジア文字用)
- **英語フォント**: `Arial`
- 両方のフォントを設定するには `set_run_font()` (dimension_0.py で定義) を使用

### ハイパーリンク

`utils.add_hyperlink(paragraph, text, url)` を使用:
```python
p = doc.add_paragraph('詳細は ')
add_hyperlink(p, 'こちら', 'https://example.com')
p.add_run(' をご覧ください。')
```

### テーブルスタイル

`dimension_0.py` の `set_table_style()` を使用して KPMG ブランドのテーブルスタイルを適用:
- ヘッダー行に KPMG Blue 背景色
- ヘッダーテキストは白色・太字
- データ行は標準書式

## ファイルエンコーディング

- すべての Python ファイルは **UTF-8 エンコーディング**
- ファイル冒頭に `# -*- coding: utf-8 -*-` を含める
- Windows コンソールでの実行時は `sys.stdout` のエンコーディングを UTF-8 に変更 (main.py と generate_assessment_doc_jp.py で実装済み)

## 出力後の手順

Word 文書生成後、ユーザーが手動で実施する必要がある操作:

1. Word で文書を開く
2. 目次ページにカーソルを置く
3. 「参考資料」タブ → 「目次」 → 自動目次スタイルを選択
4. 内容記入完了後、目次を右クリック → 「フィールド更新」 → 「目次をすべて更新」

理由: python-docx は目次フィールドを自動生成できないため、Word の機能を使用する必要がある。

## Git 履歴に関する注意

- `generate_assessment_doc_jp.py` は現在未追跡ファイル (git status で確認済み)
- `scripts/` 内のモジュール化版が推奨アーキテクチャ
- スタンドアロン版 (`generate_assessment_doc_jp.py`) は維度0のみを生成する簡易版

## 言語とローカライゼーション

- ドキュメントコンテンツ: 主に**日本語**
- コード内コメント: 主に**日本語**
- 一部のハイパーリンクテキスト: 英語（KPMG 公式サイトのタイトル）
- フォーマット: 日本語と英語の混在に対応
