# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度1: 技術能力と効率評価
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import add_hyperlink


def add_dimension_1(doc):
    """維度1: 技術能力と効率評価

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('1. 技術能力と効率評価', 1)

    doc.add_heading('1.1 技術スタック互換性', 2)

    doc.add_heading('デザインシステム (Design Systems)', 3)
    p = doc.add_paragraph('参考: ')
    add_hyperlink(p, 'Design Systems for KPMG Workbench',
                  'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/Design-Systems-for-KPMG-Workbench.aspx')
    doc.add_paragraph('既存UIフレームワークとの互換性: ___\n')
    doc.add_paragraph('移行コスト評価: ___\n')

    doc.add_heading('開発プロセス (SDLC)', 3)
    p = doc.add_paragraph('参考: ')
    add_hyperlink(p, 'Software Development Lifecycle',
                  'https://docs.code.kpmg.com/GTK/Engineering-Ecosystem/Software-Development-Lifecycle-%28SDLC%29/sdlc/')
    doc.add_paragraph('既存ワークフローとの競合点: ___\n')
    doc.add_paragraph('CI/CDプロセス比較: ___\n')

    doc.add_heading('セキュリティ規範 (Secret Management)', 3)
    p = doc.add_paragraph('参考: ')
    add_hyperlink(p, 'Secret Management Best Practices',
                  'https://handbook.code.kpmg.com/digital-grc/secrets-management-best-practices/')
    doc.add_paragraph('既存ソリューションとの優劣: ___\n')
    doc.add_paragraph('GitHub EMU認証プロセスの複雑度: ___\n')

    doc.add_heading('1.2 使いやすさと学習曲線', 2)
    p = doc.add_paragraph('参考: ')
    add_hyperlink(p, 'KPMG Workbench User Guide',
                  'https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/')
    doc.add_paragraph('開発者体験(DX)スコア: ___/10\n')
    doc.add_paragraph('ドキュメント完全性スコア: ___/10\n')
    doc.add_paragraph('Hello Worldプロジェクト設定時間: ___時間\n')

    doc.add_heading('1.3 機能と制限', 2)
    doc.add_paragraph('【プリインストールツールリスト】\n')
    doc.add_paragraph('【既存環境と比較した新機能】\n')
    doc.add_paragraph('【リソースアクセスの利便性】\n')
    doc.add_paragraph('【クォータ制限とコスト】\n')
    doc.add_paragraph('【最大の弱点】\n')

    doc.add_heading('1.4 開発効率向上(定量化可能)', 2)

    table = doc.add_table(rows=4, cols=4)
    table.style = 'Light Grid Accent 1'

    headers = table.rows[0].cells
    headers[0].text = 'テスト項目'
    headers[1].text = '既存プロセス'
    headers[2].text = 'Workbench'
    headers[3].text = '向上率'

    items = [
        ('RAG Chatbot Demo', '', '', ''),
        ('コードから展開まで', '', '', ''),
        ('バグ修正効率', '', '', '')
    ]

    for i, (item, old, new, imp) in enumerate(items, 1):
        row = table.rows[i].cells
        row[0].text = item
        row[1].text = old
        row[2].text = new
        row[3].text = imp

    doc.add_page_break()
