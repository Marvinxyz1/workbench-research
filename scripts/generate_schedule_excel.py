# -*- coding: utf-8 -*-
"""
KPMG Workbench 開発スケジュール Excel 生成ツール v2
ガントチャート、タスク詳細、休暇カレンダー、マイルストーンを含む完全なスケジュール表を生成
- Phase列のセル結合と背景色
- バージョン番号自動インクリメント（既存ファイルを上書きしない）
"""

import sys
import os
import re
import glob
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment,
    NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

# Windows console UTF-8 support
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# ==================== KPMG ブランドカラー ====================
KPMG_BLUE = "005EB8"
KPMG_DARK_BLUE = "00338D"
KC_COLOR = "4472C4"      # 青色 - KC組
ATH_COLOR = "70AD47"     # 緑色 - ATH組
JOINT_COLOR = "7030A0"   # 紫色 - 連合タスク
HOLIDAY_COLOR = "D9D9D9" # グレー - 休暇
MAGENTA_COLOR = "E91E8C" # マゼンタ - Phase 4

# Phase カラー
PHASE_COLORS = {
    "Phase 1": KPMG_DARK_BLUE,  # 濃紺 - オンボーディング
    "Phase 2": KPMG_BLUE,       # 青 - 誰でも1週間で
    "Phase 3": ATH_COLOR,       # 緑 - 海外MFナレッジ
    "Phase 4": MAGENTA_COLOR,   # マゼンタ - スピーディな
}

# Phase display names (for merged cells)
PHASE_DISPLAY = {
    "Phase 1": "オンボーディング\nフェーズ\n(10-11月)",
    "Phase 2": "誰でも1週間で\n迷いなく開始\n(12-1月)",
    "Phase 3": "海外MF\nナレッジ探索\n(2-3月)",
    "Phase 4": "スピーディな\nプロダクトプロセス\n(4-6月)",
}

# ステータスカラー（タスクデータのstatus値と一致）
STATUS_COLORS = {
    "未開始": "BFBFBF",    # グレー
    "進行中": "FFC000",    # 黄色
    "完了": "92D050",      # 緑色
    "ブロック": "FF0000",  # 赤色
}

# ==================== タスクデータ (WBS) ====================
ALL_TASKS = [
    # Phase 1: オンボーディングフェーズ (10-11月)
    {"wbs": "1.1", "phase": "Phase 1", "subtask": "学習・認証", "name": "Prerequisites 認証完了", "team": "KC", "months": [10], "deliverable": "認証バッジ", "status": "未開始"},
    {"wbs": "1.1.1", "phase": "Phase 1", "subtask": "学習・認証", "name": "Developer Learning Path 完了", "team": "KC", "months": [10, 11], "deliverable": "学習修了証", "status": "未開始"},
    {"wbs": "1.1.2", "phase": "Phase 1", "subtask": "学習・認証", "name": "Knowledge Badge 取得", "team": "KC", "months": [11], "deliverable": "バッジ証明", "status": "未開始"},
    {"wbs": "1.1.3", "phase": "Phase 1", "subtask": "学習・認証", "name": "Tech Talks 重要回視聴", "team": "KC", "months": [10, 11], "deliverable": "視聴記録", "status": "未開始"},
    {"wbs": "1.2", "phase": "Phase 1", "subtask": "技術評価", "name": "Workbench 環境セットアップ", "team": "KC", "months": [10], "deliverable": "環境構築完了", "status": "未開始"},
    {"wbs": "1.2.1", "phase": "Phase 1", "subtask": "技術評価", "name": "API 機能調査（Document Translation等）", "team": "KC", "months": [10, 11], "deliverable": "API調査レポート", "status": "未開始"},
    {"wbs": "1.2.2", "phase": "Phase 1", "subtask": "技術評価", "name": "Agent 開発フレームワーク評価", "team": "KC", "months": [11], "deliverable": "評価レポート", "status": "未開始"},
    {"wbs": "1.3", "phase": "Phase 1", "subtask": "課題対応", "name": "手続上の障壁把握", "team": "KC", "months": [10], "deliverable": "課題リスト", "status": "未開始"},
    {"wbs": "1.3.1", "phase": "Phase 1", "subtask": "課題対応", "name": "APIアクセス課題解決（12/1完了）", "team": "KC", "months": [11], "deliverable": "解決策ドキュメント", "status": "未開始"},
    {"wbs": "1.3.2", "phase": "Phase 1", "subtask": "課題対応", "name": "Global WB Community ローンチ（11/13）", "team": "KC+ATH", "months": [11], "deliverable": "Community稼働", "status": "未開始"},
    {"wbs": "1.4", "phase": "Phase 1", "subtask": "報告・計画", "name": "技術評価レポート作成", "team": "KC", "months": [11], "deliverable": "評価レポート", "status": "未開始"},
    {"wbs": "1.4.1", "phase": "Phase 1", "subtask": "報告・計画", "name": "Phase 2 計画策定", "team": "KC", "months": [11], "deliverable": "計画書", "status": "未開始"},

    # Phase 2: "誰でも1週間で迷いなく開始" (12-1月)
    {"wbs": "2.1", "phase": "Phase 2", "subtask": "KC組タスク", "name": "残りAPIテスト完了（Document Translation等）", "team": "KC", "months": [12], "deliverable": "テストレポート", "status": "未開始"},
    {"wbs": "2.1.1", "phase": "Phase 2", "subtask": "KC組タスク", "name": "APIテスト報告総括", "team": "KC", "months": [12], "deliverable": "総括レポート", "status": "未開始"},
    {"wbs": "2.1.2", "phase": "Phase 2", "subtask": "KC組タスク", "name": "Cookbook基礎テンプレート整備", "team": "KC", "months": [12], "deliverable": "テンプレート", "status": "未開始"},
    {"wbs": "2.1.3", "phase": "Phase 2", "subtask": "KC組タスク", "name": "ATH組への技術サポート資料提供", "team": "KC", "months": [12], "deliverable": "技術資料", "status": "未開始"},
    {"wbs": "2.1.4", "phase": "Phase 2", "subtask": "KC組タスク", "name": "ATH組成果物Review", "team": "KC", "months": [1], "deliverable": "Reviewレポート", "status": "未開始"},
    {"wbs": "2.1.5", "phase": "Phase 2", "subtask": "KC組タスク", "name": "Cookbook上級用例補充（2シナリオ）", "team": "KC", "months": [1], "deliverable": "上級用例", "status": "未開始"},
    {"wbs": "2.1.6", "phase": "Phase 2", "subtask": "KC組タスク", "name": "Agentベストプラクティスドキュメント作成", "team": "KC", "months": [1], "deliverable": "ベストプラクティス", "status": "未開始"},
    {"wbs": "2.1.7", "phase": "Phase 2", "subtask": "KC組タスク", "name": "Phase 2 総括レポート", "team": "KC", "months": [1], "deliverable": "総括レポート", "status": "未開始"},
    {"wbs": "2.2", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "KJ開発者Onboardingサイト設計", "team": "ATH", "months": [12], "deliverable": "設計書", "status": "未開始"},
    {"wbs": "2.2.1", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "Onboardingサイトフロントエンド開発", "team": "ATH", "months": [12], "deliverable": "フロントエンド", "status": "未開始"},
    {"wbs": "2.2.2", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "Onboardingサイトコンテンツ作成", "team": "ATH", "months": [1], "deliverable": "コンテンツ", "status": "未開始"},
    {"wbs": "2.2.3", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "Agent検証環境構築", "team": "ATH", "months": [12], "deliverable": "環境構築", "status": "未開始"},
    {"wbs": "2.2.4", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "API/Agent基礎検証テスト", "team": "ATH", "months": [12], "deliverable": "テスト結果", "status": "未開始"},
    {"wbs": "2.2.5", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "Cookbookコア用例作成（5シナリオ）", "team": "ATH", "months": [1], "deliverable": "コア用例", "status": "未開始"},
    {"wbs": "2.2.6", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "開発者コミュニティプラットフォーム構築", "team": "ATH", "months": [1], "deliverable": "プラットフォーム", "status": "未開始"},
    {"wbs": "2.2.7", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "開発者コミュニティ初期コンテンツ", "team": "ATH", "months": [1], "deliverable": "初期コンテンツ", "status": "未開始"},
    {"wbs": "2.2.8", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "Onboardingサイトローンチ＆テスト", "team": "ATH", "months": [1], "deliverable": "サイト稼働", "status": "未開始"},
    {"wbs": "2.2.9", "phase": "Phase 2", "subtask": "ATH組タスク", "name": "開発者コミュニティ正式ローンチ", "team": "ATH", "months": [1], "deliverable": "コミュニティ稼働", "status": "未開始"},

    # Phase 3: "海外MFナレッジ探索" (2-3月)
    {"wbs": "3.1", "phase": "Phase 3", "subtask": "POC討論", "name": "POC項目テーマブレインストーミング", "team": "KC+ATH", "months": [1], "deliverable": "テーマリスト", "status": "未開始"},
    {"wbs": "3.1.1", "phase": "Phase 3", "subtask": "POC討論", "name": "関係者ニーズ・アイデア収集", "team": "ATH", "months": [1], "deliverable": "ニーズリスト", "status": "未開始"},
    {"wbs": "3.1.2", "phase": "Phase 3", "subtask": "POC討論", "name": "POC項目フィージビリティ初期評価", "team": "KC+ATH", "months": [1], "deliverable": "評価結果", "status": "未開始"},
    {"wbs": "3.1.3", "phase": "Phase 3", "subtask": "POC討論", "name": "POC項目リスト確定（2-3件）", "team": "KC+ATH", "months": [1], "deliverable": "確定リスト", "status": "未開始"},
    {"wbs": "3.2", "phase": "Phase 3", "subtask": "海外MF情報収集", "name": "US/UK/AU等MFのAgent開発事例収集", "team": "ATH", "months": [2], "deliverable": "事例リスト", "status": "未開始"},
    {"wbs": "3.2.1", "phase": "Phase 3", "subtask": "海外MF情報収集", "name": "海外MF技術ドキュメント翻訳整理", "team": "ATH", "months": [2], "deliverable": "翻訳ドキュメント", "status": "未開始"},
    {"wbs": "3.2.2", "phase": "Phase 3", "subtask": "海外MF情報収集", "name": "ナレッジ再利用フレームワーク構築", "team": "ATH", "months": [2], "deliverable": "フレームワーク", "status": "未開始"},
    {"wbs": "3.2.3", "phase": "Phase 3", "subtask": "海外MF情報収集", "name": "海外MF技術交流会議（2回）", "team": "KC", "months": [2], "deliverable": "会議記録", "status": "未開始"},
    {"wbs": "3.2.4", "phase": "Phase 3", "subtask": "海外MF情報収集", "name": "海外ソリューション実現可能性評価レポート", "team": "KC", "months": [2], "deliverable": "評価レポート", "status": "未開始"},
    {"wbs": "3.3", "phase": "Phase 3", "subtask": "POC開始", "name": "POC項目A開始", "team": "ATH", "months": [3], "deliverable": "POC A進捗", "status": "未開始"},
    {"wbs": "3.3.1", "phase": "Phase 3", "subtask": "POC開始", "name": "POC項目B開始", "team": "ATH", "months": [3], "deliverable": "POC B進捗", "status": "未開始"},
    {"wbs": "3.3.2", "phase": "Phase 3", "subtask": "POC開始", "name": "POC技術サポート", "team": "KC", "months": [3], "deliverable": "技術サポート", "status": "未開始"},
    {"wbs": "3.3.3", "phase": "Phase 3", "subtask": "POC開始", "name": "POC中期レビュー", "team": "KC", "months": [3], "deliverable": "中期レビュー", "status": "未開始"},

    # Phase 4: "スピーディなプロダクトプロセス" (4-6月)
    {"wbs": "4.1", "phase": "Phase 4", "subtask": "POC完成", "name": "POC項目A完成＆評価", "team": "ATH", "months": [3, 4], "deliverable": "POC A完成", "status": "未開始"},
    {"wbs": "4.1.1", "phase": "Phase 4", "subtask": "POC完成", "name": "POC項目B完成＆評価", "team": "ATH", "months": [3, 4], "deliverable": "POC B完成", "status": "未開始"},
    {"wbs": "4.2", "phase": "Phase 4", "subtask": "本番アプリ開発", "name": "本番アプリアーキテクチャ設計Review", "team": "KC", "months": [4], "deliverable": "設計Review", "status": "未開始"},
    {"wbs": "4.2.1", "phase": "Phase 4", "subtask": "本番アプリ開発", "name": "本番アプリ開発（項目A）", "team": "ATH", "months": [4, 5], "deliverable": "アプリA", "status": "未開始"},
    {"wbs": "4.2.2", "phase": "Phase 4", "subtask": "本番アプリ開発", "name": "本番アプリ開発（項目B）", "team": "ATH", "months": [4, 5], "deliverable": "アプリB", "status": "未開始"},
    {"wbs": "4.2.3", "phase": "Phase 4", "subtask": "本番アプリ開発", "name": "セキュリティ・コンプライアンス審査", "team": "KC", "months": [4], "deliverable": "審査結果", "status": "未開始"},
    {"wbs": "4.3", "phase": "Phase 4", "subtask": "リリース", "name": "UATテスト", "team": "ATH", "months": [5], "deliverable": "テスト結果", "status": "未開始"},
    {"wbs": "4.3.1", "phase": "Phase 4", "subtask": "リリース", "name": "本番リリース準備", "team": "ATH", "months": [5], "deliverable": "リリース準備", "status": "未開始"},
    {"wbs": "4.3.2", "phase": "Phase 4", "subtask": "リリース", "name": "本番リリース＆監視", "team": "ATH", "months": [5, 6], "deliverable": "リリース完了", "status": "未開始"},
    {"wbs": "4.3.3", "phase": "Phase 4", "subtask": "リリース", "name": "最終発表承認", "team": "KC", "months": [5], "deliverable": "承認完了", "status": "未開始"},
    {"wbs": "4.4", "phase": "Phase 4", "subtask": "プロセス確立", "name": "新プロセスドキュメント化", "team": "ATH", "months": [6], "deliverable": "プロセスドキュメント", "status": "未開始"},
    {"wbs": "4.4.1", "phase": "Phase 4", "subtask": "プロセス確立", "name": "プロジェクト総括レポート", "team": "KC", "months": [6], "deliverable": "総括レポート", "status": "未開始"},
]

# Milestones (including Phase 1)
MILESTONES = [
    # Phase 1 里程碑
    {"name": "Phase 1 開始", "target_date": "2024-10-01",
     "actual_date": "", "deliverable": "オンボーディング開始", "status": "未開始"},
    {"name": "Prerequisites認証完了", "target_date": "2024-10-31",
     "actual_date": "", "deliverable": "認証バッジ", "status": "未開始"},
    {"name": "技術評価レポート完成", "target_date": "2024-11-30",
     "actual_date": "", "deliverable": "評価レポート", "status": "未開始"},
    # Phase 2 里程碑
    {"name": "Phase 2 開始", "target_date": "2024-12-01",
     "actual_date": "", "deliverable": "チーム拡大", "status": "未開始"},
    {"name": "KC組休暇開始", "target_date": "2024-12-15",
     "actual_date": "", "deliverable": "引継ぎ完了", "status": "未開始"},
    {"name": "ATH組休暇開始", "target_date": "2024-12-27",
     "actual_date": "", "deliverable": "-", "status": "未開始"},
    {"name": "全員復帰", "target_date": "2025-01-06",
     "actual_date": "", "deliverable": "開発再開", "status": "未開始"},
    {"name": "Onboardingサイト公開", "target_date": "2025-01-31",
     "actual_date": "", "deliverable": "サイト稼働", "status": "未開始"},
    {"name": "開発者コミュニティ公開", "target_date": "2025-01-31",
     "actual_date": "", "deliverable": "コミュニティ稼働", "status": "未開始"},
    # Phase 3 里程碑
    {"name": "Phase 3 開始", "target_date": "2025-02-01",
     "actual_date": "", "deliverable": "海外MF探索開始", "status": "未開始"},
    {"name": "POC項目確定", "target_date": "2025-01-31",
     "actual_date": "", "deliverable": "確定リスト", "status": "未開始"},
    {"name": "POC開始", "target_date": "2025-03-01",
     "actual_date": "", "deliverable": "POC開発開始", "status": "未開始"},
    # Phase 4 里程碑
    {"name": "Phase 4 開始", "target_date": "2025-04-01",
     "actual_date": "", "deliverable": "本番開発開始", "status": "未開始"},
    {"name": "POC完成", "target_date": "2025-04-15",
     "actual_date": "", "deliverable": "POC評価完了", "status": "未開始"},
    {"name": "UAT完成", "target_date": "2025-05-31",
     "actual_date": "", "deliverable": "テスト合格", "status": "未開始"},
    {"name": "本番リリース", "target_date": "2025-05-31",
     "actual_date": "", "deliverable": "アプリ公開", "status": "未開始"},
    {"name": "プロジェクト完了", "target_date": "2025-06-30",
     "actual_date": "", "deliverable": "総括レポート", "status": "未開始"},
]

# Holiday periods
HOLIDAYS = [
    {"team": "KC", "start": "2024-12-15", "end": "2025-01-05", "note": "年末年始休暇"},
    {"team": "ATH", "start": "2024-12-27", "end": "2025-01-05", "note": "年末年始休暇"},
]


def get_next_version(output_dir, base_name):
    """ファイルの次のバージョン番号を取得"""
    pattern = os.path.join(output_dir, f"{base_name}_v*.xlsx")
    existing_files = glob.glob(pattern)

    if not existing_files:
        return 1

    versions = []
    for f in existing_files:
        match = re.search(r'_v(\d+)\.xlsx$', f)
        if match:
            versions.append(int(match.group(1)))

    return max(versions) + 1 if versions else 1


def get_month_year(month):
    """月に基づいて年を返す（10-12月は2024年、1-6月は2025年）"""
    if month >= 10:
        return 2024
    else:
        return 2025


def get_month_last_day(year, month):
    """指定月の最終日を取得"""
    if month == 12:
        return 31
    next_month = datetime(year, month + 1, 1) if month < 12 else datetime(year + 1, 1, 1)
    return (next_month - timedelta(days=1)).day


def months_to_dates(months):
    """月配列を開始/終了日付文字列に変換する"""
    if not months:
        return None, None

    start_month = min(months)
    end_month = max(months)

    start_year = get_month_year(start_month)
    end_year = get_month_year(end_month)

    start_date = f"{start_year}-{start_month:02d}-01"
    end_day = get_month_last_day(end_year, end_month)
    end_date = f"{end_year}-{end_month:02d}-{end_day:02d}"

    return start_date, end_date


def month_to_col(month, base_col=7):
    """月をガントチャートの列番号にマッピング
    月順序: 10月, 11月, 12月, 1月, 2月, 3月, 4月, 5月, 6月
    base_col: 最初の月（10月）に対応する列番号
    """
    if month >= 10:
        return base_col + (month - 10)  # 10月=7, 11月=8, 12月=9
    else:
        return base_col + 3 + (month - 1)  # 1月=10, 2月=11, ..., 6月=15


def create_header_style():
    """KPMGブランドのヘッダースタイルを作成"""
    return {
        'font': Font(name='Meiryo UI', bold=True, color="FFFFFF", size=11),
        'fill': PatternFill(start_color=KPMG_BLUE, end_color=KPMG_BLUE, fill_type="solid"),
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }


def apply_header_style(cell):
    """セルにヘッダースタイルを適用"""
    style = create_header_style()
    cell.font = style['font']
    cell.fill = style['fill']
    cell.alignment = style['alignment']
    cell.border = style['border']


def apply_cell_style(cell, team=None):
    """標準セルスタイルを適用"""
    cell.font = Font(name='Meiryo UI')  # デフォルトフォント
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # チーム列のチームベース色分け
    if team:
        if team == "KC":
            cell.fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
            cell.font = Font(name='Meiryo UI', color="FFFFFF", bold=True)
        elif team == "ATH":
            cell.fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
            cell.font = Font(name='Meiryo UI', color="FFFFFF", bold=True)
        elif "KC+ATH" in team:
            cell.fill = PatternFill(start_color=JOINT_COLOR, end_color=JOINT_COLOR, fill_type="solid")
            cell.font = Font(name='Meiryo UI', color="FFFFFF", bold=True)


def apply_phase_style(cell, phase):
    """Phase固有のスタイルを適用"""
    color = PHASE_COLORS.get(phase, KPMG_BLUE)
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    cell.font = Font(name='Meiryo UI', color="FFFFFF", bold=True, size=10)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True, text_rotation=0)
    cell.border = Border(
        left=Side(style='medium'),
        right=Side(style='medium'),
        top=Side(style='medium'),
        bottom=Side(style='medium')
    )


def get_week_start_dates(start_date_str, end_date_str):
    """週の開始日（月曜日）のリストを生成"""
    start = datetime.strptime(start_date_str, "%Y-%m-%d")
    end = datetime.strptime(end_date_str, "%Y-%m-%d")

    # 最初の月曜日を探す
    days_until_monday = (7 - start.weekday()) % 7
    if days_until_monday == 0 and start.weekday() != 0:
        days_until_monday = 7
    current = start + timedelta(days=days_until_monday) if start.weekday() != 0 else start

    weeks = []
    while current <= end:
        weeks.append(current)
        current += timedelta(days=7)

    return weeks


def is_in_holiday(date, team):
    """日付がチームの休暇期間内かどうかを確認"""
    for holiday in HOLIDAYS:
        if holiday['team'] == team or team == "KC+ATH":
            h_start = datetime.strptime(holiday['start'], "%Y-%m-%d")
            h_end = datetime.strptime(holiday['end'], "%Y-%m-%d")
            if h_start <= date <= h_end:
                return True
    return False


def create_gantt_sheet(wb):
    """ガントチャートシートを作成（月次ビュー、WBS形式）"""
    ws = wb.create_sheet("ガントチャート")

    # 月ヘッダー (10月-6月)
    month_labels = ["10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月"]
    month_values = [10, 11, 12, 1, 2, 3, 4, 5, 6]

    # ヘッダー: Phase, WBS, Sub-Task, Action, Owner, Deliverable + 月
    headers = ["Phase", "WBS", "Sub-Task", "Action", "Owner", "Deliverable"]
    base_month_col = len(headers) + 1

    # ヘッダーを書き込む
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # 月ヘッダーを書き込む
    for i, month_label in enumerate(month_labels):
        col = base_month_col + i
        cell = ws.cell(row=1, column=col, value=month_label)
        apply_header_style(cell)
        ws.column_dimensions[get_column_letter(col)].width = 5

    # 列幅を設定
    ws.column_dimensions['A'].width = 19  # Phase
    ws.column_dimensions['B'].width = 8   # WBS
    ws.column_dimensions['C'].width = 18  # Sub-Task
    ws.column_dimensions['D'].width = 40  # Action
    ws.column_dimensions['E'].width = 8   # Owner
    ws.column_dimensions['F'].width = 20  # Deliverable

    # セル結合のためにPhaseとSubtaskでタスクをグループ化
    phase_groups = {}
    subtask_groups = {}
    for i, task in enumerate(ALL_TASKS):
        phase = task['phase']
        subtask_key = f"{phase}|{task['subtask']}"

        if phase not in phase_groups:
            phase_groups[phase] = {'start': i, 'end': i}
        else:
            phase_groups[phase]['end'] = i

        if subtask_key not in subtask_groups:
            subtask_groups[subtask_key] = {'start': i, 'end': i}
        else:
            subtask_groups[subtask_key]['end'] = i

    # タスクを書き込む
    for row, task in enumerate(ALL_TASKS, 2):
        # Phase列
        phase_cell = ws.cell(row=row, column=1, value=PHASE_DISPLAY.get(task['phase'], task['phase']))
        apply_phase_style(phase_cell, task['phase'])

        # WBS
        ws.cell(row=row, column=2, value=task['wbs'])
        apply_cell_style(ws.cell(row=row, column=2))

        # Sub-Task
        ws.cell(row=row, column=3, value=task['subtask'])
        apply_cell_style(ws.cell(row=row, column=3))

        # Action（タスク名）
        ws.cell(row=row, column=4, value=task['name'])
        apply_cell_style(ws.cell(row=row, column=4))

        # Owner（チーム）
        team_cell = ws.cell(row=row, column=5, value=task['team'])
        apply_cell_style(team_cell, task['team'])

        # 成果物
        ws.cell(row=row, column=6, value=task['deliverable'])
        apply_cell_style(ws.cell(row=row, column=6))

        # 月別にガントバーを描画（Phaseカラーを使用）
        for i, month in enumerate(month_values):
            col = base_month_col + i
            cell = ws.cell(row=row, column=col)

            if month in task['months']:
                phase_color = PHASE_COLORS.get(task['phase'], KPMG_BLUE)
                cell.fill = PatternFill(start_color=phase_color, end_color=phase_color, fill_type="solid")

            cell.border = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )

    # Phaseセルを結合
    for phase, indices in phase_groups.items():
        start_row = indices['start'] + 2
        end_row = indices['end'] + 2
        if start_row != end_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)

    # Sub-Taskセルを結合
    for key, indices in subtask_groups.items():
        start_row = indices['start'] + 2
        end_row = indices['end'] + 2
        if start_row != end_row:
            ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)

    # 凡例を追加（Phaseカラー）
    legend_row = len(ALL_TASKS) + 4
    ws.cell(row=legend_row, column=1, value="凡例:")
    ws.cell(row=legend_row, column=2, value="Phase 1").fill = PatternFill(start_color=KPMG_DARK_BLUE, end_color=KPMG_DARK_BLUE, fill_type="solid")
    ws.cell(row=legend_row, column=2).font = Font(name='Meiryo UI', color="FFFFFF")
    ws.cell(row=legend_row, column=3, value="Phase 2").fill = PatternFill(start_color=KPMG_BLUE, end_color=KPMG_BLUE, fill_type="solid")
    ws.cell(row=legend_row, column=3).font = Font(name='Meiryo UI', color="FFFFFF")
    ws.cell(row=legend_row, column=4, value="Phase 3").fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=4).font = Font(name='Meiryo UI', color="FFFFFF")
    ws.cell(row=legend_row, column=5, value="Phase 4").fill = PatternFill(start_color=MAGENTA_COLOR, end_color=MAGENTA_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=5).font = Font(name='Meiryo UI', color="FFFFFF")

    # ペインを固定
    ws.freeze_panes = 'G2'

    return ws


def create_all_tasks_sheet(wb):
    """全タスク詳細シートを作成（WBS形式：subtask、月、成果物）"""
    ws = wb.create_sheet("タスク詳細")

    # ALL_TASKSの新構造に合わせたヘッダー: wbs, subtask, name, team, months, deliverable, status
    headers = ["Phase", "WBS", "Sub-Task", "Action", "Owner", "Months", "Deliverable", "Status"]

    # ヘッダーを書き込む
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # 列幅を設定
    col_widths = [18, 8, 18, 42, 10, 15, 22, 10]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # セル結合のためにPhaseとSubtaskでタスクをグループ化
    phase_groups = {}
    subtask_groups = {}
    for i, task in enumerate(ALL_TASKS):
        phase = task['phase']
        subtask_key = f"{phase}|{task['subtask']}"

        if phase not in phase_groups:
            phase_groups[phase] = {'start': i, 'end': i}
        else:
            phase_groups[phase]['end'] = i

        if subtask_key not in subtask_groups:
            subtask_groups[subtask_key] = {'start': i, 'end': i}
        else:
            subtask_groups[subtask_key]['end'] = i

    # 新フィールド名でタスクを書き込む
    for row, task in enumerate(ALL_TASKS, 2):
        # Phase列
        phase_cell = ws.cell(row=row, column=1, value=PHASE_DISPLAY.get(task['phase'], task['phase']))
        apply_phase_style(phase_cell, task['phase'])

        # WBS番号
        ws.cell(row=row, column=2, value=task['wbs'])
        apply_cell_style(ws.cell(row=row, column=2))

        # Sub-Taskカテゴリ
        ws.cell(row=row, column=3, value=task['subtask'])
        apply_cell_style(ws.cell(row=row, column=3))

        # Action（タスク名）
        ws.cell(row=row, column=4, value=task['name'])
        apply_cell_style(ws.cell(row=row, column=4))

        # Owner（チーム）
        team_cell = ws.cell(row=row, column=5, value=task['team'])
        apply_cell_style(team_cell, task['team'])

        # 月を "10月, 11月" 形式でフォーマット
        months_str = ", ".join([f"{m}月" for m in task['months']])
        ws.cell(row=row, column=6, value=months_str)
        apply_cell_style(ws.cell(row=row, column=6))

        # 成果物
        ws.cell(row=row, column=7, value=task['deliverable'])
        apply_cell_style(ws.cell(row=row, column=7))

        # ステータス
        ws.cell(row=row, column=8, value=task['status'])
        apply_cell_style(ws.cell(row=row, column=8))

    # Phaseセルを結合
    for phase, indices in phase_groups.items():
        start_row = indices['start'] + 2
        end_row = indices['end'] + 2
        if start_row != end_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)

    # Sub-Taskセルを結合
    for key, indices in subtask_groups.items():
        start_row = indices['start'] + 2
        end_row = indices['end'] + 2
        if start_row != end_row:
            ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)

    # ステータス列Hにデータ検証を追加
    status_validation = DataValidation(
        type="list",
        formula1='"未開始,進行中,完了,ブロック"',
        allow_blank=True
    )
    status_validation.add(f'H2:H{len(ALL_TASKS) + 1}')
    ws.add_data_validation(status_validation)

    # ステータスに条件付き書式を追加
    for status, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(
            f'H2:H{len(ALL_TASKS) + 1}',
            CellIsRule(
                operator='equal',
                formula=[f'"{status}"'],
                fill=PatternFill(start_color=color, end_color=color, fill_type="solid")
            )
        )

    # ペインを固定
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:I{len(ALL_TASKS) + 1}'

    return ws


def create_team_sheet(wb, team_name, team_filter):
    """チーム別タスクシートを作成（WBS形式）"""
    ws = wb.create_sheet(f"{team_name}組タスク")

    # WBS形式のヘッダー
    headers = ["Phase", "WBS", "Sub-Task", "Action", "Months", "Deliverable", "Status"]

    # ヘッダーを書き込む
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # 列幅を設定
    col_widths = [18, 8, 18, 42, 15, 22, 10]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # チームでタスクをフィルタ
    team_tasks = [t for t in ALL_TASKS if team_filter in t['team']]

    # Phaseでグループ化（セル結合用）
    phase_groups = {}
    for i, task in enumerate(team_tasks):
        phase = task['phase']
        if phase not in phase_groups:
            phase_groups[phase] = {'start': i, 'end': i}
        else:
            phase_groups[phase]['end'] = i

    # タスクを書き込む
    for row, task in enumerate(team_tasks, 2):
        # Phase列
        phase_cell = ws.cell(row=row, column=1, value=PHASE_DISPLAY.get(task['phase'], task['phase']))
        apply_phase_style(phase_cell, task['phase'])

        # WBS番号
        ws.cell(row=row, column=2, value=task['wbs'])
        apply_cell_style(ws.cell(row=row, column=2))

        # Sub-Taskカテゴリ
        ws.cell(row=row, column=3, value=task['subtask'])
        apply_cell_style(ws.cell(row=row, column=3))

        # Action（タスク名）
        ws.cell(row=row, column=4, value=task['name'])
        apply_cell_style(ws.cell(row=row, column=4))

        # 月
        months_str = ", ".join([f"{m}月" for m in task['months']])
        ws.cell(row=row, column=5, value=months_str)
        apply_cell_style(ws.cell(row=row, column=5))

        # 成果物
        ws.cell(row=row, column=6, value=task['deliverable'])
        apply_cell_style(ws.cell(row=row, column=6))

        # ステータス
        ws.cell(row=row, column=7, value=task['status'])
        apply_cell_style(ws.cell(row=row, column=7))

    # Phaseセルを結合
    for phase, indices in phase_groups.items():
        start_row = indices['start'] + 2
        end_row = indices['end'] + 2
        if start_row != end_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)

    # ステータス列Gにデータ検証を追加
    status_validation = DataValidation(
        type="list",
        formula1='"未開始,進行中,完了,ブロック"',
        allow_blank=True
    )
    if team_tasks:
        status_validation.add(f'G2:G{len(team_tasks) + 1}')
        ws.add_data_validation(status_validation)

    # ペインを固定
    ws.freeze_panes = 'B2'

    return ws


def create_holiday_sheet(wb):
    """休暇カレンダーシートを作成"""
    ws = wb.create_sheet("休暇カレンダー")

    # タイトル
    title_cell = ws.cell(row=1, column=1, value="2024-2025 年末年始休暇カレンダー")
    title_cell.font = Font(name='Meiryo UI', bold=True, size=14, color=KPMG_BLUE)
    ws.merge_cells('A1:G1')

    # 休暇詳細
    headers = ["組織", "開始日", "終了日", "休暇日数", "備考"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        apply_header_style(cell)

    for row, holiday in enumerate(HOLIDAYS, 4):
        start = datetime.strptime(holiday['start'], "%Y-%m-%d")
        end = datetime.strptime(holiday['end'], "%Y-%m-%d")
        days = (end - start).days + 1

        ws.cell(row=row, column=1, value=holiday['team'])
        ws.cell(row=row, column=2, value=holiday['start'])
        ws.cell(row=row, column=3, value=holiday['end'])
        ws.cell(row=row, column=4, value=days)
        ws.cell(row=row, column=5, value=holiday['note'])

        for col in range(1, 6):
            apply_cell_style(ws.cell(row=row, column=col))

    # カレンダービュー
    ws.cell(row=7, column=1, value="カレンダービュー:").font = Font(name='Meiryo UI', bold=True)

    # 2024年12月と2025年1月のカレンダーを生成
    months = [
        ("2024年12月", 2024, 12),
        ("2025年1月", 2025, 1)
    ]

    start_col = 1
    for month_name, year, month in months:
        ws.cell(row=8, column=start_col, value=month_name).font = Font(name='Meiryo UI', bold=True)

        # 曜日ヘッダー
        week_headers = ["月", "火", "水", "木", "金", "土", "日"]
        for i, header in enumerate(week_headers):
            ws.cell(row=9, column=start_col + i, value=header)

        # 月の最初の日を計算
        first_day = datetime(year, month, 1)
        # 月の日数を取得
        if month == 12:
            days_in_month = 31
        else:
            days_in_month = 31  # 1月

        # カレンダーを埋める
        current_row = 10
        current_col = start_col + first_day.weekday()

        for day in range(1, days_in_month + 1):
            date = datetime(year, month, day)
            cell = ws.cell(row=current_row, column=current_col, value=day)

            # 休暇をチェック
            is_kc_holiday = False
            is_ath_holiday = False

            for holiday in HOLIDAYS:
                h_start = datetime.strptime(holiday['start'], "%Y-%m-%d")
                h_end = datetime.strptime(holiday['end'], "%Y-%m-%d")
                if h_start <= date <= h_end:
                    if holiday['team'] == "KC":
                        is_kc_holiday = True
                    elif holiday['team'] == "ATH":
                        is_ath_holiday = True

            if is_kc_holiday and is_ath_holiday:
                cell.fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
            elif is_kc_holiday:
                cell.fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
                cell.font = Font(name='Meiryo UI', color="FFFFFF")
            elif is_ath_holiday:
                cell.fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
                cell.font = Font(name='Meiryo UI', color="FFFFFF")

            current_col += 1
            if current_col > start_col + 6:
                current_col = start_col
                current_row += 1

        start_col += 9  # 次の月へ移動

    # 列幅を設定
    for col in range(1, 20):
        ws.column_dimensions[get_column_letter(col)].width = 5

    # 凡例
    legend_row = 17
    ws.cell(row=legend_row, column=1, value="凡例:")
    ws.cell(row=legend_row, column=2, value="KC休暇").fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=2).font = Font(name='Meiryo UI', color="FFFFFF")
    ws.cell(row=legend_row, column=4, value="ATH休暇").fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=4).font = Font(name='Meiryo UI', color="FFFFFF")
    ws.cell(row=legend_row, column=6, value="両方休暇").fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    return ws


def create_milestone_sheet(wb):
    """マイルストーン追跡シートを作成"""
    ws = wb.create_sheet("マイルストーン")

    headers = ["マイルストーン", "予定日", "実績日", "成果物", "ステータス"]

    # ヘッダーを書き込む
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # 列幅を設定
    col_widths = [25, 12, 12, 25, 10]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # マイルストーンを書き込む
    for row, milestone in enumerate(MILESTONES, 2):
        ws.cell(row=row, column=1, value=milestone['name'])
        ws.cell(row=row, column=2, value=milestone['target_date'])
        ws.cell(row=row, column=3, value=milestone['actual_date'])
        ws.cell(row=row, column=4, value=milestone['deliverable'])
        ws.cell(row=row, column=5, value=milestone['status'])

        # 罫線を適用
        for col in range(1, 6):
            apply_cell_style(ws.cell(row=row, column=col))

    # ステータスにデータ検証を追加
    status_validation = DataValidation(
        type="list",
        formula1='"未開始,進行中,完了,遅延"',
        allow_blank=True
    )
    status_validation.add(f'E2:E{len(MILESTONES) + 1}')
    ws.add_data_validation(status_validation)

    # ペインを固定
    ws.freeze_panes = 'A2'

    return ws


def main():
    """Excelファイルを生成するメイン関数"""
    print("=" * 60)
    print("KPMG Workbench 開発スケジュール Excel 生成ツール v2")
    print("=" * 60)

    # 出力ディレクトリの存在を確認
    output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "generated_docs", "schedules")
    os.makedirs(output_dir, exist_ok=True)

    # 次のバージョン番号を取得
    base_name = "KPMG_Workbench_Schedule"
    version = get_next_version(output_dir, base_name)

    print(f"\nバージョン番号: v{version}")

    # ワークブックを作成
    wb = Workbook()

    # デフォルトシートを削除
    default_sheet = wb.active
    wb.remove(default_sheet)

    # シートを作成
    print("\n[1/6] ガントチャートを作成中（Phaseセル結合付き）...")
    create_gantt_sheet(wb)

    print("[2/6] タスク詳細を作成中（Phaseセル結合付き）...")
    create_all_tasks_sheet(wb)

    print("[3/6] KC組タスクビューを作成中...")
    create_team_sheet(wb, "KC", "KC")

    print("[4/6] ATH組タスクビューを作成中...")
    create_team_sheet(wb, "ATH", "ATH")

    print("[5/6] 休暇カレンダーを作成中...")
    create_holiday_sheet(wb)

    print("[6/6] マイルストーン追跡を作成中...")
    create_milestone_sheet(wb)

    # バージョン番号付きでワークブックを保存
    output_filename = f"{base_name}_v{version}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    wb.save(output_path)

    print("\n" + "=" * 60)
    print(f"Excelファイルを生成しました: {output_path}")
    print(f"バージョン: v{version}")
    print("=" * 60)

    # サマリー
    kc_tasks = len([t for t in ALL_TASKS if t['team'] == 'KC'])
    ath_tasks = len([t for t in ALL_TASKS if t['team'] == 'ATH'])
    joint_tasks = len([t for t in ALL_TASKS if 'KC+ATH' in t['team']])

    print(f"\nタスク統計:")
    print(f"  - KC組タスク: {kc_tasks} 件")
    print(f"  - ATH組タスク: {ath_tasks} 件")
    print(f"  - 連合タスク: {joint_tasks} 件")
    print(f"  - 合計: {len(ALL_TASKS)} 件")
    print(f"  - マイルストーン: {len(MILESTONES)} 件")

    # 自動的にファイルを開く
    import subprocess
    subprocess.run(['open', output_path])

    return output_path


if __name__ == "__main__":
    main()
