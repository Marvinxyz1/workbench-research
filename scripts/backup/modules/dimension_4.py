# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度4: 学習資源とコミュニティサポート評価
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import add_hyperlink


def add_dimension_4(doc):
    """維度4: 学習資源とコミュニティサポート評価

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('4. 学習資源とコミュニティサポート評価', 1)

    doc.add_heading('4.1 公式学習資源品質', 2)
    doc.add_paragraph('【Tech Talksシリーズ (2025年4-6月)】\n')
    doc.add_paragraph('最も価値のあるエピソード: ___\n')
    doc.add_paragraph('【Developers Conference録画 (2024年11月)】\n')
    doc.add_paragraph('【ドキュメント完全性】\n')

    doc.add_heading('4.2 コミュニティとサポート', 2)
    doc.add_paragraph('技術サポート対応時間: ___時間\n')
    doc.add_paragraph('【Slack/Teamsチャンネルアクティビティ】\n')
    p = doc.add_paragraph('【Championsネットワーク】参考: ')
    add_hyperlink(p, 'Global AI Ninjas and Navigators',
                  'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI/SitePages/Global-AI-Ninjas-Navigators.aspx')

    doc.add_heading('4.3 外部リソース', 2)
    p = doc.add_paragraph('ホワイトペーパー価値: ')
    add_hyperlink(p, 'AI Adoption in the Workplace',
                  'https://kpmg.com/au/en/insights/artificial-intelligence-ai/workplace-ai-adoption-success-insights-stories.html')
    doc.add_paragraph()
    p = doc.add_paragraph('Podcast価値: ')
    add_hyperlink(p, 'You Can with AI Podcast',
                  'https://kpmg.com/us/en/podcasts/you-can-with-ai.html')

    doc.add_page_break()
