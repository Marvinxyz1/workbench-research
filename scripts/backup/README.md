# KPMG Workbench 戦略評価フレームワーク

モジュール化されたWord文書生成システム（日本語版）

## 📁 ディレクトリ構造

```
scripts/
├── main.py                          # メイン実行スクリプト（ここから実行）
├── utils.py                         # 共通ツール関数
├── modules/                         # モジュールディレクトリ
│   ├── __init__.py
│   ├── cover.py                     # 表紙・目次・エグゼクティブサマリー
│   ├── dimension_0.py               # 維度0: 事前準備
│   ├── dimension_1.py               # 維度1: 技術能力と効率評価
│   ├── dimension_2.py               # 維度2: Agentic AI核心能力評価
│   ├── dimension_3.py               # 維度3: 商業価値と顧客応用
│   ├── dimension_4.py               # 維度4: 学習資源とコミュニティサポート
│   ├── dimension_5.py               # 維度5: 戦略価値と組織影響
│   ├── dimension_6.py               # 維度6: リスクとコンプライアンス評価
│   ├── dimension_7.py               # 維度7: チーム推進と運営考慮
│   └── appendix.py                  # 付録（参考資料・タイムライン）
├── generate_assessment_doc.py       # 旧版（参考用）
└── generate_assessment_doc_jp.py    # 旧日文版（参考用）
```

## 🚀 使い方

### 完全版レポートを生成

すべてのセクションを含む完全なレポートを生成：

```bash
cd scripts
python main.py
```

生成されたファイル: `../generated_docs/KPMG_Workbench戦略評価レポート_完全版.docx`

### 個別セクションを修正

各維度は独立したモジュールなので、個別に編集可能：

1. `modules/dimension_X.py` を編集
2. `main.py` を再実行して文書を再生成

### カスタムレポートを生成

特定のセクションのみを含むレポートを作成する場合、`main.py`をコピーして不要なセクションをコメントアウト：

```python
# 例: 維度2-7を省略
# add_dimension_2(doc)
# add_dimension_3(doc)
# ...
```

## 📝 モジュール説明

### utils.py
- `add_hyperlink()` - ハイパーリンク追加関数
- `set_heading_style()` - 見出しスタイル設定関数

### modules/cover.py
- `create_cover_page()` - 表紙生成
- `create_toc()` - 目次生成
- `add_executive_summary()` - エグゼクティブサマリー追加

### modules/dimension_X.py
各維度の評価コンテンツを生成する関数を提供：
- `add_dimension_X(doc)` - Word文書にセクションを追加

### modules/appendix.py
- `add_appendix_references()` - 参考資料リスト
- `add_appendix_timeline()` - 評価タイムライン

## 🔧 依存関係

```bash
pip install python-docx
```

## 📄 出力ファイル

生成された文書は `../generated_docs/` ディレクトリに保存されます。

## ⚠️ 注意事項

1. Word文書を開いた後、目次を手動で生成する必要があります：
   - カーソルを目次ページに置く
   - 参考資料 → 目次 → 自動目次スタイルを選択

2. 内容記入完了後、目次を更新：
   - 目次を右クリック → フィールド更新 → 目次をすべて更新

## 🎯 今後の拡張

- 新しい維度を追加する場合：
  1. `modules/dimension_X.py` を作成
  2. `modules/__init__.py` にインポート追加
  3. `main.py` に関数呼び出し追加

---

**バージョン**: 1.0.0
**最終更新**: 2025-11-13
