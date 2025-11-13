# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度7: チーム推進と運営考慮
"""

def add_dimension_7(doc):
    """維度7: チーム推進と運営考慮

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('7. チーム推進と運営考慮', 1)

    doc.add_heading('7.1 チーム適合性', 2)
    doc.add_paragraph('迅速に習得できるメンバーの割合: ___%\n')
    doc.add_paragraph('【トレーニングニーズとコスト】\n')
    doc.add_paragraph('【ワークロード影響】\n')

    doc.add_heading('7.2 コストと収益(ROI)', 2)

    table = doc.add_table(rows=5, cols=3)
    table.style = 'Light Grid Accent 1'

    headers = table.rows[0].cells
    headers[0].text = '指標'
    headers[1].text = '現状'
    headers[2].text = 'Workbench使用後'

    items = [
        ('Demo開発時間', '', ''),
        ('プロジェクト受注率', '', ''),
        ('チームモラール', '', ''),
        ('ナレッジ蓄積', '', '')
    ]

    for i, (metric, before, after) in enumerate(items, 1):
        row = table.rows[i].cells
        row[0].text = metric
        row[1].text = before
        row[2].text = after

    doc.add_paragraph()

    doc.add_heading('7.3 Go/No-Go決定基準 🎯', 2)

    doc.add_heading('全面展開条件 ✅', 3)
    doc.add_paragraph('☐ 個人認証完了時間 < 20時間', style='List Bullet')
    doc.add_paragraph('☐ Demo開発速度向上 > 30%', style='List Bullet')
    doc.add_paragraph('☐ 少なくとも1つのPoCが顧客から好評', style='List Bullet')
    doc.add_paragraph('☐ 技術サポート対応時間 < 24時間', style='List Bullet')
    doc.add_paragraph('☐ データ分離がセキュリティ審査を通過', style='List Bullet')
    doc.add_paragraph('☐ チーム > 70%が学習価値があると認識', style='List Bullet')

    doc.add_heading('展開延期条件 ❌', 3)
    doc.add_paragraph('☐ 学習コスト > 1ヶ月フルタイム投入', style='List Bullet')
    doc.add_paragraph('☐ 分離メカニズム不明確/コンプライアンスリスク有', style='List Bullet')
    doc.add_paragraph('☐ チーム < 50%が独立開発可能', style='List Bullet')
    doc.add_paragraph('☐ 移行コスト > 50%開発時間', style='List Bullet')
    doc.add_paragraph('☐ 公式サポート不足', style='List Bullet')
    doc.add_paragraph('☐ より良い代替案有', style='List Bullet')

    doc.add_heading('7.4 パイロット方案 (Phased Rollout)', 2)

    table = doc.add_table(rows=5, cols=4)
    table.style = 'Light Grid Accent 1'

    headers = table.rows[0].cells
    headers[0].text = '段階'
    headers[1].text = '期間'
    headers[2].text = '参加者'
    headers[3].text = '成果物'

    phases = [
        ('Phase 1: 個人探索', '1ヶ月', 'SCレベル1名', '技術評価レポート'),
        ('Phase 2: グループパイロット', '2ヶ月', 'SC + 1-2名', 'PoC + ROIデータ'),
        ('Phase 3: 決定ポイント', '1週', '管理層', 'Go/No-Go決定'),
        ('Phase 4: 全面展開', '3ヶ月', '全チーム', 'トレーニング完了+継続最適化')
    ]

    for i, (phase, duration, people, output) in enumerate(phases, 1):
        row = table.rows[i].cells
        row[0].text = phase
        row[1].text = duration
        row[2].text = people
        row[3].text = output

    doc.add_page_break()
