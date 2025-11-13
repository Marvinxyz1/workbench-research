# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度3: 商業価値と顧客応用
"""

def add_dimension_3(doc):
    """維度3: 商業価値と顧客応用

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('3. 商業価値と顧客応用', 1)

    doc.add_heading('3.1 Demo環境の価値', 2)
    doc.add_paragraph('【プリセールス加速】\n')
    doc.add_paragraph('Demo開発時間比較: [X週] → [Y日]\n')
    doc.add_paragraph('【提案競争力向上】\n')
    doc.add_paragraph('【カスタマイズ能力】\n')

    doc.add_heading('3.2 社内サービス開発', 2)
    doc.add_paragraph('【開発可能な社内ツールリスト】\n')
    doc.add_paragraph('【既存システムとの統合難易度】\n')
    doc.add_paragraph('【ナレッジ蓄積メカニズム】\n')

    doc.add_heading('3.3 顧客プロジェクト境界と移行コスト (リスク!)', 2)
    doc.add_paragraph('【「顧客サービスに使用不可」の境界】\n')
    doc.add_paragraph('Demo → 本番環境移行コスト: ___%\n')
    doc.add_paragraph('【データ分離リスク】\n')
    doc.add_paragraph('【コンプライアンス保証】\n')

    doc.add_page_break()
