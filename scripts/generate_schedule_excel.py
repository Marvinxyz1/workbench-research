# -*- coding: utf-8 -*-
"""
KPMG Workbench 开发日程表 Excel 生成器
生成包含甘特图、任务详情、休假日历、里程碑的完整日程表
"""

import sys
import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment,
    NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# Windows console UTF-8 support
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

# ==================== KPMG Brand Colors ====================
KPMG_BLUE = "005EB8"
KPMG_DARK_BLUE = "00338D"
KC_COLOR = "4472C4"      # 蓝色 - KC组
ATH_COLOR = "70AD47"     # 绿色 - ATH组
JOINT_COLOR = "7030A0"   # 紫色 - 联合任务
HOLIDAY_COLOR = "D9D9D9" # 灰色 - 休假

# Status colors
STATUS_COLORS = {
    "未开始": "BFBFBF",    # 灰色
    "进行中": "FFC000",    # 黄色
    "已完成": "92D050",    # 绿色
    "阻塞": "FF0000",      # 红色
}

# ==================== Task Data ====================
ALL_TASKS = [
    # Phase 2: 12月～1月 KC组任务（休假前）
    {"id": "KC-01", "phase": "Phase 2", "name": "完成剩余API测试（Document Translation等）", "team": "KC", "days": 5, "start": "2024-12-09", "end": "2024-12-13", "status": "未开始", "note": ""},
    {"id": "KC-02", "phase": "Phase 2", "name": "编写API测试报告总结", "team": "KC", "days": 2, "start": "2024-12-14", "end": "2024-12-15", "status": "未开始", "note": ""},
    {"id": "KC-03", "phase": "Phase 2", "name": "整理Cookbook基础模板和目录结构", "team": "KC", "days": 2, "start": "2024-12-09", "end": "2024-12-10", "status": "未开始", "note": ""},
    {"id": "KC-04", "phase": "Phase 2", "name": "与ATH组交接工作，提供技术支持文档", "team": "KC", "days": 1, "start": "2024-12-15", "end": "2024-12-15", "status": "未开始", "note": "休假前最后任务"},

    # Phase 2: ATH组任务
    {"id": "ATH-01", "phase": "Phase 2", "name": "KJ开发者Onboarding网站设计", "team": "ATH", "days": 5, "start": "2024-12-09", "end": "2024-12-13", "status": "未开始", "note": ""},
    {"id": "ATH-02", "phase": "Phase 2", "name": "Onboarding网站前端开发", "team": "ATH", "days": 10, "start": "2024-12-16", "end": "2024-12-27", "status": "未开始", "note": ""},
    {"id": "ATH-03", "phase": "Phase 2", "name": "Onboarding网站内容填充", "team": "ATH", "days": 5, "start": "2025-01-06", "end": "2025-01-10", "status": "未开始", "note": ""},
    {"id": "ATH-04", "phase": "Phase 2", "name": "Agent验证环境搭建", "team": "ATH", "days": 5, "start": "2024-12-09", "end": "2024-12-13", "status": "未开始", "note": ""},
    {"id": "ATH-05", "phase": "Phase 2", "name": "API/Agent基础验证测试", "team": "ATH", "days": 10, "start": "2024-12-16", "end": "2024-12-27", "status": "未开始", "note": ""},
    {"id": "ATH-06", "phase": "Phase 2", "name": "Cookbook核心用例编写（5个场景）", "team": "ATH", "days": 10, "start": "2025-01-06", "end": "2025-01-17", "status": "未开始", "note": ""},
    {"id": "ATH-07", "phase": "Phase 2", "name": "开发者社区平台搭建", "team": "ATH", "days": 5, "start": "2025-01-06", "end": "2025-01-10", "status": "未开始", "note": ""},
    {"id": "ATH-08", "phase": "Phase 2", "name": "开发者社区内容初始化", "team": "ATH", "days": 5, "start": "2025-01-13", "end": "2025-01-17", "status": "未开始", "note": ""},
    {"id": "ATH-09", "phase": "Phase 2", "name": "Onboarding网站上线 & 测试", "team": "ATH", "days": 5, "start": "2025-01-20", "end": "2025-01-24", "status": "未开始", "note": ""},
    {"id": "ATH-10", "phase": "Phase 2", "name": "开发者社区正式上线", "team": "ATH", "days": 2, "start": "2025-01-27", "end": "2025-01-28", "status": "未开始", "note": ""},

    # Phase 2: KC组任务（复工后）
    {"id": "KC-05", "phase": "Phase 2", "name": "Review ATH组工作成果", "team": "KC", "days": 2, "start": "2025-01-06", "end": "2025-01-07", "status": "未开始", "note": "复工后第一任务"},
    {"id": "KC-06", "phase": "Phase 2", "name": "补充Cookbook高级用例（2个场景）", "team": "KC", "days": 5, "start": "2025-01-08", "end": "2025-01-14", "status": "未开始", "note": ""},
    {"id": "KC-07", "phase": "Phase 2", "name": "编写Agent最佳实践文档", "team": "KC", "days": 5, "start": "2025-01-15", "end": "2025-01-21", "status": "未开始", "note": ""},
    {"id": "KC-08", "phase": "Phase 2", "name": "Phase 2 总结报告", "team": "KC", "days": 2, "start": "2025-01-22", "end": "2025-01-23", "status": "未开始", "note": ""},

    # 过渡期: POC讨论
    {"id": "JOINT-01", "phase": "过渡期", "name": "POC项目主题头脑风暴会议", "team": "KC+ATH", "days": 2, "start": "2025-01-27", "end": "2025-01-28", "status": "未开始", "note": "全员参与"},
    {"id": "JOINT-02", "phase": "过渡期", "name": "收集各方需求和想法", "team": "ATH", "days": 3, "start": "2025-01-27", "end": "2025-01-29", "status": "未开始", "note": ""},
    {"id": "JOINT-03", "phase": "过渡期", "name": "POC项目可行性初评", "team": "KC+ATH", "days": 2, "start": "2025-01-30", "end": "2025-01-31", "status": "未开始", "note": ""},
    {"id": "JOINT-04", "phase": "过渡期", "name": "确定POC项目列表（2-3个）", "team": "KC+ATH", "days": 1, "start": "2025-01-31", "end": "2025-01-31", "status": "未开始", "note": "决策会议"},

    # Phase 3: 海外MFナレッジ探索
    {"id": "ATH-11", "phase": "Phase 3", "name": "收集US/UK/AU等MF的Agent开发案例", "team": "ATH", "days": 10, "start": "2025-02-03", "end": "2025-02-14", "status": "未开始", "note": ""},
    {"id": "ATH-12", "phase": "Phase 3", "name": "整理海外MF技术文档翻译", "team": "ATH", "days": 5, "start": "2025-02-17", "end": "2025-02-21", "status": "未开始", "note": ""},
    {"id": "ATH-13", "phase": "Phase 3", "name": "建立知识复用框架", "team": "ATH", "days": 5, "start": "2025-02-24", "end": "2025-02-28", "status": "未开始", "note": ""},
    {"id": "ATH-14", "phase": "Phase 3", "name": "POC项目A启动", "team": "ATH", "days": 10, "start": "2025-03-03", "end": "2025-03-14", "status": "未开始", "note": ""},
    {"id": "ATH-15", "phase": "Phase 3", "name": "POC项目B启动", "team": "ATH", "days": 10, "start": "2025-03-03", "end": "2025-03-14", "status": "未开始", "note": ""},
    {"id": "KC-09", "phase": "Phase 3", "name": "与海外MF技术交流会议（2次）", "team": "KC", "days": 2, "start": "2025-02-10", "end": "2025-02-11", "status": "未开始", "note": ""},
    {"id": "KC-10", "phase": "Phase 3", "name": "评估海外方案可行性报告", "team": "KC", "days": 5, "start": "2025-02-17", "end": "2025-02-21", "status": "未开始", "note": ""},
    {"id": "KC-11", "phase": "Phase 3", "name": "POC技术支援", "team": "KC", "days": 10, "start": "2025-03-03", "end": "2025-03-14", "status": "未开始", "note": ""},
    {"id": "KC-12", "phase": "Phase 3", "name": "POC中期评审", "team": "KC", "days": 2, "start": "2025-03-17", "end": "2025-03-18", "status": "未开始", "note": ""},

    # Phase 4: スピーディなプロダクトプロセス
    {"id": "ATH-16", "phase": "Phase 4", "name": "POC项目A完成 & 评审", "team": "ATH", "days": 10, "start": "2025-03-17", "end": "2025-03-28", "status": "未开始", "note": ""},
    {"id": "ATH-17", "phase": "Phase 4", "name": "POC项目B完成 & 评审", "team": "ATH", "days": 10, "start": "2025-03-17", "end": "2025-03-28", "status": "未开始", "note": ""},
    {"id": "ATH-18", "phase": "Phase 4", "name": "本番アプリ开发（项目A）", "team": "ATH", "days": 20, "start": "2025-04-01", "end": "2025-04-25", "status": "未开始", "note": ""},
    {"id": "ATH-19", "phase": "Phase 4", "name": "本番アプリ开发（项目B）", "team": "ATH", "days": 20, "start": "2025-04-01", "end": "2025-04-25", "status": "未开始", "note": ""},
    {"id": "ATH-20", "phase": "Phase 4", "name": "UAT测试", "team": "ATH", "days": 10, "start": "2025-04-28", "end": "2025-05-09", "status": "未开始", "note": ""},
    {"id": "ATH-21", "phase": "Phase 4", "name": "本番リリース準備", "team": "ATH", "days": 10, "start": "2025-05-12", "end": "2025-05-23", "status": "未开始", "note": ""},
    {"id": "ATH-22", "phase": "Phase 4", "name": "本番リリース & 监控", "team": "ATH", "days": 5, "start": "2025-05-26", "end": "2025-05-30", "status": "未开始", "note": ""},
    {"id": "ATH-23", "phase": "Phase 4", "name": "新流程文档化", "team": "ATH", "days": 5, "start": "2025-06-02", "end": "2025-06-06", "status": "未开始", "note": ""},
    {"id": "KC-13", "phase": "Phase 4", "name": "本番アプリ架构设计Review", "team": "KC", "days": 5, "start": "2025-04-01", "end": "2025-04-04", "status": "未开始", "note": ""},
    {"id": "KC-14", "phase": "Phase 4", "name": "安全合规审查", "team": "KC", "days": 5, "start": "2025-04-14", "end": "2025-04-18", "status": "未开始", "note": ""},
    {"id": "KC-15", "phase": "Phase 4", "name": "最终发布审批", "team": "KC", "days": 2, "start": "2025-05-26", "end": "2025-05-27", "status": "未开始", "note": ""},
    {"id": "KC-16", "phase": "Phase 4", "name": "项目总结报告", "team": "KC", "days": 5, "start": "2025-06-09", "end": "2025-06-13", "status": "未开始", "note": ""},
]

# Milestones
MILESTONES = [
    {"name": "Phase 2 开始", "target_date": "2024-12-09", "actual_date": "", "deliverable": "团队启动", "status": "未开始"},
    {"name": "KC组休假开始", "target_date": "2024-12-15", "actual_date": "", "deliverable": "交接完成", "status": "未开始"},
    {"name": "ATH组休假开始", "target_date": "2024-12-27", "actual_date": "", "deliverable": "-", "status": "未开始"},
    {"name": "全员复工", "target_date": "2025-01-06", "actual_date": "", "deliverable": "恢复开发", "status": "未开始"},
    {"name": "Onboarding网站上线", "target_date": "2025-01-24", "actual_date": "", "deliverable": "网站可访问", "status": "未开始"},
    {"name": "开发者社区上线", "target_date": "2025-01-28", "actual_date": "", "deliverable": "社区可访问", "status": "未开始"},
    {"name": "POC项目确定", "target_date": "2025-01-31", "actual_date": "", "deliverable": "项目列表", "status": "未开始"},
    {"name": "Phase 3 开始", "target_date": "2025-02-03", "actual_date": "", "deliverable": "知识探索启动", "status": "未开始"},
    {"name": "POC启动", "target_date": "2025-03-03", "actual_date": "", "deliverable": "POC开发开始", "status": "未开始"},
    {"name": "POC完成", "target_date": "2025-03-28", "actual_date": "", "deliverable": "POC评审通过", "status": "未开始"},
    {"name": "Phase 4 开始", "target_date": "2025-04-01", "actual_date": "", "deliverable": "本番开发启动", "status": "未开始"},
    {"name": "UAT完成", "target_date": "2025-05-09", "actual_date": "", "deliverable": "测试通过", "status": "未开始"},
    {"name": "本番リリース", "target_date": "2025-05-30", "actual_date": "", "deliverable": "アプリ上线", "status": "未开始"},
    {"name": "项目完成", "target_date": "2025-06-13", "actual_date": "", "deliverable": "总结报告", "status": "未开始"},
]

# Holiday periods
HOLIDAYS = [
    {"team": "KC", "start": "2024-12-15", "end": "2025-01-05", "note": "年末年始休暇"},
    {"team": "ATH", "start": "2024-12-27", "end": "2025-01-05", "note": "年末年始休暇"},
]


def create_header_style():
    """Create KPMG branded header style"""
    return {
        'font': Font(bold=True, color="FFFFFF", size=11),
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
    """Apply header style to a cell"""
    style = create_header_style()
    cell.font = style['font']
    cell.fill = style['fill']
    cell.alignment = style['alignment']
    cell.border = style['border']


def apply_cell_style(cell, team=None):
    """Apply standard cell style"""
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Team-based coloring for team column
    if team:
        if team == "KC":
            cell.fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
        elif team == "ATH":
            cell.fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
        elif "KC+ATH" in team:
            cell.fill = PatternFill(start_color=JOINT_COLOR, end_color=JOINT_COLOR, fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)


def get_week_start_dates(start_date_str, end_date_str):
    """Generate list of week start dates (Mondays)"""
    start = datetime.strptime(start_date_str, "%Y-%m-%d")
    end = datetime.strptime(end_date_str, "%Y-%m-%d")

    # Find first Monday
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
    """Check if a date is in holiday period for a team"""
    for holiday in HOLIDAYS:
        if holiday['team'] == team or team == "KC+ATH":
            h_start = datetime.strptime(holiday['start'], "%Y-%m-%d")
            h_end = datetime.strptime(holiday['end'], "%Y-%m-%d")
            if h_start <= date <= h_end:
                return True
    return False


def create_gantt_sheet(wb):
    """Create Gantt chart sheet"""
    ws = wb.create_sheet("甘特图总览")

    # Define date range
    weeks = get_week_start_dates("2024-12-02", "2025-06-30")

    # Headers
    headers = ["任务ID", "任务名称", "负责组", "开始", "结束"]

    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # Write week headers
    for col, week in enumerate(weeks, len(headers) + 1):
        cell = ws.cell(row=1, column=col, value=week.strftime("%m/%d"))
        apply_header_style(cell)
        ws.column_dimensions[get_column_letter(col)].width = 6

    # Set column widths
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12

    # Write tasks
    for row, task in enumerate(ALL_TASKS, 2):
        ws.cell(row=row, column=1, value=task['id'])
        ws.cell(row=row, column=2, value=task['name'])

        team_cell = ws.cell(row=row, column=3, value=task['team'])
        apply_cell_style(team_cell, task['team'])

        ws.cell(row=row, column=4, value=task['start'])
        ws.cell(row=row, column=5, value=task['end'])

        # Apply borders to basic cells
        for col in [1, 2, 4, 5]:
            apply_cell_style(ws.cell(row=row, column=col))

        # Draw Gantt bars
        task_start = datetime.strptime(task['start'], "%Y-%m-%d")
        task_end = datetime.strptime(task['end'], "%Y-%m-%d")

        for col, week_start in enumerate(weeks, len(headers) + 1):
            week_end = week_start + timedelta(days=6)
            cell = ws.cell(row=row, column=col)

            # Check if task overlaps with this week
            if task_start <= week_end and task_end >= week_start:
                # Task is active this week
                if task['team'] == "KC":
                    cell.fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
                elif task['team'] == "ATH":
                    cell.fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color=JOINT_COLOR, end_color=JOINT_COLOR, fill_type="solid")
            elif is_in_holiday(week_start, task['team']):
                cell.fill = PatternFill(start_color=HOLIDAY_COLOR, end_color=HOLIDAY_COLOR, fill_type="solid")

            cell.border = Border(
                left=Side(style='thin', color='CCCCCC'),
                right=Side(style='thin', color='CCCCCC'),
                top=Side(style='thin', color='CCCCCC'),
                bottom=Side(style='thin', color='CCCCCC')
            )

    # Add legend
    legend_row = len(ALL_TASKS) + 4
    ws.cell(row=legend_row, column=1, value="凡例:")
    ws.cell(row=legend_row, column=2, value="KC组").fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=2).font = Font(color="FFFFFF")
    ws.cell(row=legend_row, column=3, value="ATH组").fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=3).font = Font(color="FFFFFF")
    ws.cell(row=legend_row, column=4, value="联合").fill = PatternFill(start_color=JOINT_COLOR, end_color=JOINT_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=4).font = Font(color="FFFFFF")
    ws.cell(row=legend_row, column=5, value="休假").fill = PatternFill(start_color=HOLIDAY_COLOR, end_color=HOLIDAY_COLOR, fill_type="solid")

    # Freeze panes
    ws.freeze_panes = 'F2'

    return ws


def create_all_tasks_sheet(wb):
    """Create all tasks detail sheet"""
    ws = wb.create_sheet("全部任务详情")

    headers = ["任务ID", "阶段", "任务名称", "负责组", "时长(天)", "开始日期", "结束日期", "状态", "备注"]

    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # Set column widths
    col_widths = [10, 12, 45, 10, 10, 12, 12, 10, 25]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Write tasks
    for row, task in enumerate(ALL_TASKS, 2):
        ws.cell(row=row, column=1, value=task['id'])
        ws.cell(row=row, column=2, value=task['phase'])
        ws.cell(row=row, column=3, value=task['name'])

        team_cell = ws.cell(row=row, column=4, value=task['team'])
        apply_cell_style(team_cell, task['team'])

        ws.cell(row=row, column=5, value=task['days'])
        ws.cell(row=row, column=6, value=task['start'])
        ws.cell(row=row, column=7, value=task['end'])
        ws.cell(row=row, column=8, value=task['status'])
        ws.cell(row=row, column=9, value=task['note'])

        # Apply borders
        for col in range(1, 10):
            if col != 4:  # Skip team column (already styled)
                apply_cell_style(ws.cell(row=row, column=col))

    # Add data validation for status column
    status_validation = DataValidation(
        type="list",
        formula1='"未开始,进行中,已完成,阻塞"',
        allow_blank=True
    )
    status_validation.add(f'H2:H{len(ALL_TASKS) + 1}')
    ws.add_data_validation(status_validation)

    # Add conditional formatting for status
    for status, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(
            f'H2:H{len(ALL_TASKS) + 1}',
            CellIsRule(
                operator='equal',
                formula=[f'"{status}"'],
                fill=PatternFill(start_color=color, end_color=color, fill_type="solid")
            )
        )

    # Freeze panes
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:I{len(ALL_TASKS) + 1}'

    return ws


def create_team_sheet(wb, team_name, team_filter):
    """Create team-specific task sheet"""
    ws = wb.create_sheet(f"{team_name}组任务")

    headers = ["任务ID", "阶段", "任务名称", "时长(天)", "开始日期", "结束日期", "状态", "备注"]

    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # Set column widths
    col_widths = [10, 12, 45, 10, 12, 12, 10, 25]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Filter tasks
    team_tasks = [t for t in ALL_TASKS if team_filter in t['team']]

    # Write tasks
    for row, task in enumerate(team_tasks, 2):
        ws.cell(row=row, column=1, value=task['id'])
        ws.cell(row=row, column=2, value=task['phase'])
        ws.cell(row=row, column=3, value=task['name'])
        ws.cell(row=row, column=4, value=task['days'])
        ws.cell(row=row, column=5, value=task['start'])
        ws.cell(row=row, column=6, value=task['end'])
        ws.cell(row=row, column=7, value=task['status'])
        ws.cell(row=row, column=8, value=task['note'])

        # Apply borders
        for col in range(1, 9):
            apply_cell_style(ws.cell(row=row, column=col))

    # Add data validation for status
    status_validation = DataValidation(
        type="list",
        formula1='"未开始,进行中,已完成,阻塞"',
        allow_blank=True
    )
    if team_tasks:
        status_validation.add(f'G2:G{len(team_tasks) + 1}')
        ws.add_data_validation(status_validation)

    # Freeze panes
    ws.freeze_panes = 'A2'

    return ws


def create_holiday_sheet(wb):
    """Create holiday calendar sheet"""
    ws = wb.create_sheet("休假日历")

    # Title
    title_cell = ws.cell(row=1, column=1, value="2024-2025 年末年始休暇カレンダー")
    title_cell.font = Font(bold=True, size=14, color=KPMG_BLUE)
    ws.merge_cells('A1:G1')

    # Holiday details
    headers = ["组织", "开始日期", "结束日期", "休假天数", "备注"]
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

    # Visual calendar
    ws.cell(row=7, column=1, value="日历视图:").font = Font(bold=True)

    # Generate December 2024 and January 2025 calendars
    months = [
        ("2024年12月", 2024, 12),
        ("2025年1月", 2025, 1)
    ]

    start_col = 1
    for month_name, year, month in months:
        ws.cell(row=8, column=start_col, value=month_name).font = Font(bold=True)

        # Week headers
        week_headers = ["月", "火", "水", "木", "金", "土", "日"]
        for i, header in enumerate(week_headers):
            ws.cell(row=9, column=start_col + i, value=header)

        # Calculate first day of month
        first_day = datetime(year, month, 1)
        # Get number of days in month
        if month == 12:
            days_in_month = 31
        else:
            days_in_month = 31  # January

        # Fill calendar
        current_row = 10
        current_col = start_col + first_day.weekday()

        for day in range(1, days_in_month + 1):
            date = datetime(year, month, day)
            cell = ws.cell(row=current_row, column=current_col, value=day)

            # Check holidays
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
                cell.font = Font(color="FFFFFF")
            elif is_ath_holiday:
                cell.fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
                cell.font = Font(color="FFFFFF")

            current_col += 1
            if current_col > start_col + 6:
                current_col = start_col
                current_row += 1

        start_col += 9  # Move to next month

    # Set column widths
    for col in range(1, 20):
        ws.column_dimensions[get_column_letter(col)].width = 5

    # Legend
    legend_row = 17
    ws.cell(row=legend_row, column=1, value="凡例:")
    ws.cell(row=legend_row, column=2, value="KC休假").fill = PatternFill(start_color=KC_COLOR, end_color=KC_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=2).font = Font(color="FFFFFF")
    ws.cell(row=legend_row, column=4, value="ATH休假").fill = PatternFill(start_color=ATH_COLOR, end_color=ATH_COLOR, fill_type="solid")
    ws.cell(row=legend_row, column=4).font = Font(color="FFFFFF")
    ws.cell(row=legend_row, column=6, value="両方休假").fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    return ws


def create_milestone_sheet(wb):
    """Create milestone tracking sheet"""
    ws = wb.create_sheet("里程碑追踪")

    headers = ["里程碑", "预期日期", "实际日期", "交付物", "状态"]

    # Write headers
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        apply_header_style(cell)

    # Set column widths
    col_widths = [25, 12, 12, 25, 10]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Write milestones
    for row, milestone in enumerate(MILESTONES, 2):
        ws.cell(row=row, column=1, value=milestone['name'])
        ws.cell(row=row, column=2, value=milestone['target_date'])
        ws.cell(row=row, column=3, value=milestone['actual_date'])
        ws.cell(row=row, column=4, value=milestone['deliverable'])
        ws.cell(row=row, column=5, value=milestone['status'])

        # Apply borders
        for col in range(1, 6):
            apply_cell_style(ws.cell(row=row, column=col))

    # Add data validation for status
    status_validation = DataValidation(
        type="list",
        formula1='"未开始,进行中,已完成,延期"',
        allow_blank=True
    )
    status_validation.add(f'E2:E{len(MILESTONES) + 1}')
    ws.add_data_validation(status_validation)

    # Freeze panes
    ws.freeze_panes = 'A2'

    return ws


def main():
    """Main function to generate the Excel file"""
    print("=" * 60)
    print("KPMG Workbench 开发日程表 Excel 生成器")
    print("=" * 60)

    # Create workbook
    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Create sheets
    print("\n[1/6] 创建甘特图总览...")
    create_gantt_sheet(wb)

    print("[2/6] 创建全部任务详情...")
    create_all_tasks_sheet(wb)

    print("[3/6] 创建KC组任务视图...")
    create_team_sheet(wb, "KC", "KC")

    print("[4/6] 创建ATH组任务视图...")
    create_team_sheet(wb, "ATH", "ATH")

    print("[5/6] 创建休假日历...")
    create_holiday_sheet(wb)

    print("[6/6] 创建里程碑追踪...")
    create_milestone_sheet(wb)

    # Ensure output directory exists
    output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "generated_docs")
    os.makedirs(output_dir, exist_ok=True)

    # Save workbook
    output_path = os.path.join(output_dir, "KPMG_Workbench_Schedule.xlsx")
    wb.save(output_path)

    print("\n" + "=" * 60)
    print(f"Excel文件已生成: {output_path}")
    print("=" * 60)

    # Summary
    kc_tasks = len([t for t in ALL_TASKS if t['team'] == 'KC'])
    ath_tasks = len([t for t in ALL_TASKS if t['team'] == 'ATH'])
    joint_tasks = len([t for t in ALL_TASKS if 'KC+ATH' in t['team']])

    print(f"\n任务统计:")
    print(f"  - KC组任务: {kc_tasks} 个")
    print(f"  - ATH组任务: {ath_tasks} 个")
    print(f"  - 联合任务: {joint_tasks} 个")
    print(f"  - 总计: {len(ALL_TASKS)} 个任务")
    print(f"  - 里程碑: {len(MILESTONES)} 个")

    return output_path


if __name__ == "__main__":
    main()
