#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - メイン実行スクリプト

このスクリプトは、すべてのモジュールをインポートして完全な評価レポートを生成します。
"""

import sys
import io
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

# モジュールのインポート
from utils import set_heading_style
from modules.cover import create_cover_page, create_toc, add_executive_summary
from modules.dimension_0 import add_dimension_0
from modules.dimension_1 import add_dimension_1
from modules.dimension_2 import add_dimension_2
from modules.dimension_3 import add_dimension_3
from modules.dimension_4 import add_dimension_4
from modules.dimension_5 import add_dimension_5
from modules.dimension_6 import add_dimension_6
from modules.dimension_7 import add_dimension_7
from modules.appendix import add_appendix_references, add_appendix_timeline


def main():
    """メイン関数 - 完全な評価レポートを生成"""

    # Windows コンソールエンコーディング問題を修正
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print('=' * 80)
    print('KPMG Workbench 戦略評価レポート（日本語版）生成開始')
    print('=' * 80)
    print()

    # ドキュメントを作成
    doc = Document()

    # デフォルトフォントを設定
    doc.styles['Normal'].font.name = 'Arial'  # 英語
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'Meiryo UI')  # 日本語
    doc.styles['Normal'].font.size = Pt(11)

    # 見出しスタイルを設定
    print('📋 フォーマット設定中...')
    set_heading_style(doc)
    print('✅ フォーマット設定完了\n')

    # 各セクションを作成
    print('📄 各セクション生成中:')
    print('  ├─ 表紙ページ...')
    create_cover_page(doc)

    print('  ├─ 目次ページ...')
    create_toc(doc)

    # print('  ├─ エグゼクティブサマリー...')
    # add_executive_summary(doc)

    print('  └─ 維度0: 事前準備（詳細版）...')
    add_dimension_0(doc)

    # print('  ├─ 維度1: 技術能力と効率評価...')
    # add_dimension_1(doc)

    # print('  ├─ 維度2: Agentic AI核心能力評価...')
    # add_dimension_2(doc)

    # print('  ├─ 維度3: 商業価値と顧客応用...')
    # add_dimension_3(doc)

    # print('  ├─ 維度4: 学習資源とコミュニティサポート...')
    # add_dimension_4(doc)

    # print('  ├─ 維度5: 戦略価値と組織影響...')
    # add_dimension_5(doc)

    # print('  ├─ 維度6: リスクとコンプライアンス評価...')
    # add_dimension_6(doc)

    # print('  ├─ 維度7: チーム推進と運営考慮...')
    # add_dimension_7(doc)

    # print('  ├─ 付録: 参考資料...')
    # add_appendix_references(doc)

    # print('  └─ 付録: 評価タイムライン...')
    # add_appendix_timeline(doc)

    print('\n✅ すべてのセクション生成完了\n')

    # 保存パスを設定（プロジェクトルート/generated_docs/に保存）
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_dir = os.path.join(project_root, 'generated_docs')
    output_path = os.path.join(output_dir, 'KPMG_Workbench戦略評価レポート_完成版.docx')

    # ディレクトリが存在しない場合は作成
    os.makedirs(output_dir, exist_ok=True)

    # 保存
    print('💾 ドキュメント保存中...')
    doc.save(output_path)

    print()
    print('=' * 80)
    print('✅ ドキュメント生成成功!')
    print('=' * 80)
    print(f'📁 ファイル場所: {output_path}')
    print()
    print('📝 次のステップ:')
    print('   1. Wordでドキュメントを開く')
    print('   2. カーソルを「目次」ページのプレースホルダーに置く')
    print('   3. 参考資料 → 目次 → 自動目次スタイルを選択')
    print('   4. 内容記入完了後、目次を右クリック → フィールド更新 → 目次をすべて更新')
    print('   5. WordのPDF保存機能で最終版を出力')
    print()
    print('=' * 80)
    print('🎉 完了!')
    print('=' * 80)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'\n❌ エラーが発生しました: {e}')
        import traceback
        traceback.print_exc()
        sys.exit(1)
