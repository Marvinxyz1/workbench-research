# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度6: リスクとコンプライアンス評価
"""

def add_dimension_6(doc):
    """維度6: リスクとコンプライアンス評価

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('6. リスクとコンプライアンス評価 ⚠️', 1)

    doc.add_heading('6.1 データセキュリティと分離リスク', 2)
    doc.add_paragraph('【データ漏洩リスク】\n')
    doc.add_paragraph('【マルチテナント分離メカニズム】\n')
    doc.add_paragraph('【アクセス制御粒度】\n')

    doc.add_heading('6.2 技術依存とロックインリスク', 2)
    doc.add_paragraph('【プラットフォームロックインリスク】\n')
    doc.add_paragraph('移行コスト評価: ___\n')
    doc.add_paragraph('【スキル移転性】\n')

    doc.add_heading('6.3 コンプライアンスと監査', 2)
    doc.add_paragraph('【GDPR/データ主権コンプライアンス】\n')
    doc.add_paragraph('【監査ログ完全性】\n')

    doc.add_heading('6.4 コストリスク', 2)

    table = doc.add_table(rows=5, cols=3)
    table.style = 'Light Grid Accent 1'

    headers = table.rows[0].cells
    headers[0].text = 'コスト項目'
    headers[1].text = '予想金額'
    headers[2].text = '備考'

    items = [
        ('ライセンス費用', '', ''),
        ('コンピューティングコスト(GPU)', '', ''),
        ('トレーニングコスト', '', ''),
        ('合計', '', '')
    ]

    for i, (item, cost, note) in enumerate(items, 1):
        row = table.rows[i].cells
        row[0].text = item
        row[1].text = cost
        row[2].text = note

    doc.add_paragraph()
    doc.add_paragraph('【機会コスト】\n')

    doc.add_page_break()
