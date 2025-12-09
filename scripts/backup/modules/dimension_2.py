# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 維度2: Agentic AI核心能力評価
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils import add_hyperlink


def add_dimension_2(doc):
    """維度2: Agentic AI核心能力評価

    Args:
        doc: python-docx文書オブジェクト
    """
    doc.add_heading('2. Agentic AI核心能力評価 🔥', 1)

    doc.add_paragraph('【重要性】これはWorkbenchのコア競争力です!')

    doc.add_heading('2.1 戦略的背景', 2)
    p = doc.add_paragraph('KPMG公式ポジショニング参考: ')
    add_hyperlink(p, 'The Agentic AI Advantageホワイトペーパー',
                  'https://kpmg.com/us/en/articles/2025/the-agentic-ai-advantage.html')

    doc.add_heading('2.2 Agent開発能力テスト', 2)
    doc.add_paragraph('【Agentフレームワークのサポート有無】\n')
    doc.add_paragraph('【プリセットAgentテンプレート】\n')
    doc.add_paragraph('【LangChain/AutoGPT/CrewAIとの比較優位性】\n')

    doc.add_heading('2.3 実際のシナリオテスト (必須!)', 2)
    doc.add_paragraph('テストタスク: [例: 自動監査Agent]\n')
    doc.add_paragraph('開発時間: ___時間\n')
    doc.add_paragraph('Agent精度: ___%\n')
    doc.add_paragraph('統合難易度: ___\n')

    doc.add_heading('2.4 ナレッジベース統合', 2)
    doc.add_paragraph('【RAG能力】\n')
    doc.add_paragraph('【KPMGナレッジベースとの統合】\n')

    doc.add_heading('2.5 マルチAgent協調', 2)
    doc.add_paragraph('【マルチAgentサポート有無】\n')
    doc.add_paragraph('【Orchestrationメカニズム】\n')

    doc.add_page_break()
