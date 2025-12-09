# -*- coding: utf-8 -*-
"""
KPMG Workbench 页面保存脚本
使用 SingleFile CLI 保存所有 Workbench 相关页面为完整 HTML
"""

import subprocess
import os
import sys
import time

# 配置
OUTPUT_DIR = r"C:\Users\junchenma\workbench-research\saved_pages\workbench\html"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA = os.path.expanduser(r"~\AppData\Local\Google\Chrome\User Data")

# 要保存的页面列表
PAGES = [
    ("01_KPMG_Workbench_Home", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB"),
    ("02_KPMG_Workbench_Platform", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-Platform.aspx"),
    ("03_aIQ_Chat_on_KPMG_Workbench", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/aIQ-Chat-on-KPMG-Workbench.aspx"),
    ("04_aIQ_CaseCraft_on_KPMG_Workbench", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/aIQ-CaseCraft-on-KPMG-Workbench.aspx"),
    ("05_KPMG_Workbench_Change_Adoption", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench-Change-&-Adoption.aspx"),
    ("06_KPMG_Workbench_Intro", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench.aspx"),
    ("07_KPMG_Workbench_1", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/KPMG-Workbench(1).aspx"),
    ("08_aIQ_Chat_Technical_Info", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/SitePages/aIQ-Chat-on-KPMG-Workbench(1).aspx"),
]

# PDF/文档资源列表
DOCUMENTS = [
    ("What_AI_tool_to_use_when_guide.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/KPMGWorkbench/DRAFT%20-%20What%20AI%20Chat%20tool%20should%20I%20use_Guidebook%20v3.0.pdf"),
    ("FY25_aiQ-Chat_Getting-started.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/FY25_aiQ-Chat_Getting-started.pdf"),
    ("FY25_aiQ-Chat_Conversation-history.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/FY25_aiQ-Chat_Conversation-history.pdf"),
    ("aIQ_Chat_FAQs.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/aIQ%20Chat%20on%20KPMG%20Workbench_FAQs%20for%20end%20users.pdf"),
    ("FY26_aiQ-Chat_Community.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/FY26_aiQ-Chat_Community.pdf"),
    ("FY25_aIQ-Chat_Persona-creation.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/FY25_aIQ-Chat_Persona-creation.pdf"),
    ("FY25_aiQ-Chat_Personas.pdf", "https://spo-global.kpmg.com/sites/GO-OI-BUS-GTK-WB/aIQChatKPMGWorkbench/FY25_aiQ-Chat_Personas.pdf"),
]


def ensure_dir(path):
    """确保目录存在"""
    if not os.path.exists(path):
        os.makedirs(path)


def save_page_with_singlefile(name, url, output_dir):
    """使用 SingleFile CLI 保存页面"""
    output_file = os.path.join(output_dir, f"{name}.html")

    cmd = [
        "single-file",
        url,
        f"--browser-executable-path={CHROME_PATH}",
        f"--browser-args=[\"--user-data-dir={CHROME_USER_DATA}\", \"--profile-directory=Default\"]",
        "--browser-wait-until=networkidle0",
        "--browser-wait-delay=5000",
        f"--output={output_file}",
    ]

    print(f"\n正在保存: {name}")
    print(f"URL: {url}")
    print(f"输出: {output_file}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode == 0:
            print(f"✓ 成功保存: {name}")
            return True
        else:
            print(f"✗ 保存失败: {name}")
            print(f"错误: {result.stderr}")
            return False
    except subprocess.TimeoutExpired:
        print(f"✗ 超时: {name}")
        return False
    except Exception as e:
        print(f"✗ 异常: {e}")
        return False


def main():
    """主函数"""
    print("=" * 60)
    print("KPMG Workbench 页面保存脚本")
    print("=" * 60)

    # 确保输出目录存在
    ensure_dir(OUTPUT_DIR)
    pdf_dir = os.path.join(os.path.dirname(OUTPUT_DIR), "pdf")
    media_dir = os.path.join(os.path.dirname(OUTPUT_DIR), "media")
    ensure_dir(pdf_dir)
    ensure_dir(media_dir)

    # 保存页面
    success_count = 0
    fail_count = 0

    print(f"\n共有 {len(PAGES)} 个页面需要保存\n")

    for name, url in PAGES:
        if save_page_with_singlefile(name, url, OUTPUT_DIR):
            success_count += 1
        else:
            fail_count += 1
        time.sleep(2)  # 避免请求过快

    # 打印结果
    print("\n" + "=" * 60)
    print(f"保存完成: 成功 {success_count}, 失败 {fail_count}")
    print("=" * 60)

    # 提示文档下载
    print("\n注意: PDF/文档资源需要通过浏览器手动下载:")
    for name, url in DOCUMENTS:
        print(f"  - {name}")
        print(f"    {url}")


if __name__ == "__main__":
    main()
