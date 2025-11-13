# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 付録
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import add_hyperlink


def add_appendix_references(doc):
    """参考資料付録を追加

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('付録: 参考資料', 1)

    doc.add_heading('A. 公式学習資源', 2)

    resources = [
        ('KPMG Workbench Learning & Development Hub',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-development.aspx'),
        ('Developer Learning Path',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-and-development-development-track.aspx'),
        ('Product Management Learning Path',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-learning-and-development-product-management-track.aspx'),
    ]

    for title, url in resources:
        p = doc.add_paragraph('• ', style='List Bullet')
        add_hyperlink(p, title, url)

    doc.add_heading('B. 技術ドキュメント', 2)

    tech_docs = [
        ('KPMG Workbench User Guide',
         'https://docs.code.kpmg.com/GTK/AI-Framework/KPMG-Workbench/'),
        ('Design Systems for KPMG Workbench',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/Design-Systems-for-KPMG-Workbench.aspx'),
        ('Software Development Lifecycle (SDLC)',
         'https://docs.code.kpmg.com/GTK/Engineering-Ecosystem/Software-Development-Lifecycle-%28SDLC%29/sdlc/'),
        ('Secret Management Best Practices',
         'https://handbook.code.kpmg.com/digital-grc/secrets-management-best-practices/'),
    ]

    for title, url in tech_docs:
        p = doc.add_paragraph('• ', style='List Bullet')
        add_hyperlink(p, title, url)

    doc.add_heading('C. 戦略資源', 2)

    strategic = [
        ('KPMG Global aIQ Hub',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI'),
        ('Global AI Ninjas and Navigators',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI/SitePages/Global-AI-Ninjas-Navigators.aspx'),
        ('Trusted AI Learning Path',
         'https://spo-global.kpmg.com/sites/go-oi-bus-People/SitePages/Trusted-AI.aspx'),
    ]

    for title, url in strategic:
        p = doc.add_paragraph('• ', style='List Bullet')
        add_hyperlink(p, title, url)

    doc.add_heading('D. ホワイトペーパーとインサイト', 2)

    insights = [
        ('The Agentic AI Advantage',
         'https://kpmg.com/us/en/articles/2025/the-agentic-ai-advantage.html'),
        ('AI Adoption in the Workplace',
         'https://kpmg.com/au/en/insights/artificial-intelligence-ai/workplace-ai-adoption-success-insights-stories.html'),
        ('KPMG Revolutionizes AI Delivery',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-AI/SitePages/KPMG-revolutionizes-AI-delivery-with-a-first-of-its-kind-global-AI-platform.aspx'),
    ]

    for title, url in insights:
        p = doc.add_paragraph('• ', style='List Bullet')
        add_hyperlink(p, title, url)

    doc.add_heading('E. 会議録画 (推奨)', 2)

    videos = [
        ('Microsoft Keynote - Agentic AI Thinking',
         'https://spo-global.kpmg.com/:v:/r/sites/GO-OI-BUS-GTK-WB/KPMGWorkbenchDevCon/Microsoft%20Keynote%20recording%20-%20Agentic%20AI%20Thinking.mp4'),
        ('KPMG Keynote - Workbench Champions',
         'https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/KPMGWorkbenchDevCon/KPMG%20Keynote%20recording%20-%20Workbench%20Champions.mp4'),
        ('Q&A with Product Owners',
         'https://spo-global.kpmg.com/:v:/r/sites/GO-OI-BUS-GTK-WB/KPMGWorkbenchDevCon/The%20Open%20Forum%20Live%20Q%26A%20with%20Workbench%20Product.mp4'),
    ]

    for title, url in videos:
        p = doc.add_paragraph('• ', style='List Bullet')
        add_hyperlink(p, title, url)

    doc.add_heading('F. 外部リソース', 2)

    p = doc.add_paragraph('• ', style='List Bullet')
    add_hyperlink(p, 'You Can with AI Podcast',
                  'https://kpmg.com/us/en/podcasts/you-can-with-ai.html')

    doc.add_page_break()


def add_appendix_timeline(doc):
    """評価タイムライン付録を追加

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('付録: 評価タイムライン (推奨)', 1)

    table = doc.add_table(rows=5, cols=3)
    table.style = 'Light Grid Accent 1'

    headers = table.rows[0].cells
    headers[0].text = '週次'
    headers[1].text = 'タスク'
    headers[2].text = '成果物'

    timeline = [
        ('W1', 'Prerequisites + Developer Learning Path完了', '認証バッジ'),
        ('W2', 'テストDemo 1開発 (例: RAG Chatbot)', '技術能力評価データ'),
        ('W3', 'テストDemo 2開発 (例: Audit Agent)', 'Agentic AI能力評価'),
        ('W4', 'リスク分析、ROI計算、レポート作成', '完全評価レポート + Executive Summary')
    ]

    for i, (week, task, output) in enumerate(timeline, 1):
        row = table.rows[i].cells
        row[0].text = week
        row[1].text = task
        row[2].text = output
