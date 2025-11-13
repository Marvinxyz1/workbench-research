# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度5: 戦略価値と組織影響
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import add_hyperlink


def add_dimension_5(doc):
    """維度5: 戦略価値と組織影響

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('5. 戦略価値と組織影響', 1)

    doc.add_heading('5.1 KPMG AIビジョンとの整合性', 2)
    p = doc.add_paragraph('参考: ')
    add_hyperlink(p, 'KPMG Global AI戦略',
                  'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI')
    doc.add_paragraph()
    p = doc.add_paragraph('Trusted AI原則: ')
    add_hyperlink(p, 'Trusted AI Principles',
                  'https://spo-global.kpmg.com/sites/go-oi-bus-People/SitePages/Trusted-AI.aspx')
    doc.add_paragraph('【AI Backboneポジショニング価値】\n')
    doc.add_paragraph('【Championsケースになる可能性】\n')

    doc.add_heading('5.2 チーム横断協力機会', 2)
    doc.add_paragraph('【ナレッジ共有メカニズム】\n')
    doc.add_paragraph('【リソース再利用可能性】\n')

    doc.add_heading('5.3 チームブランドとキャリア開発', 2)
    doc.add_paragraph('【社内可視性向上】\n')
    doc.add_paragraph('【個人成長機会】\n')

    doc.add_page_break()
