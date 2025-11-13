# -*- coding: utf-8 -*-
"""
KPMG Workbench 戦略評価フレームワーク - 共通ツール関数
"""

from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def add_hyperlink(paragraph, text, url):
    """段落にハイパーリンクを追加

    Args:
        paragraph: python-docx段落オブジェクト
        text: 表示テキスト
        url: リンクURL

    Returns:
        ハイパーリンク要素
    """
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # ハイパーリンクスタイル（青色下線）を設定
    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0000FF')
    rPr.append(c)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._element.append(hyperlink)

    return hyperlink


def set_heading_style(doc):
    """見出しスタイルを設定

    Args:
        doc: python-docx文書オブジェクト
    """
    styles = doc.styles

    # 見出し1スタイル
    h1 = styles['Heading 1']
    h1.font.name = 'Arial'
    h1.font.size = Pt(18)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 51, 102)
    # 東アジアフォント設定
    rPr = h1._element.rPr
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        h1._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rFonts.set(qn('w:eastAsia'), 'Meiryo UI')

    # 見出し2スタイル
    h2 = styles['Heading 2']
    h2.font.name = 'Arial'
    h2.font.size = Pt(14)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 112, 192)
    # 東アジアフォント設定
    rPr = h2._element.rPr
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        h2._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rFonts.set(qn('w:eastAsia'), 'Meiryo UI')

    # 見出し3スタイル
    h3 = styles['Heading 3']
    h3.font.name = 'Arial'
    h3.font.size = Pt(12)
    h3.font.bold = True
    # 東アジアフォント設定
    rPr = h3._element.rPr
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        h3._element.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rFonts.set(qn('w:eastAsia'), 'Meiryo UI')
