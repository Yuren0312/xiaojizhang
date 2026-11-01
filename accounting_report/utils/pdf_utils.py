# -*- coding: utf-8 -*-
"""PDF 工具：文本/表格提取"""
import re
import fitz
from pathlib import Path


def extract_text_blocks(pdf_path: str) -> list:
    """提取 PDF 文本块，返回 (x0,y0,x1,y1,text,block_no,block_type) 列表"""
    doc = fitz.open(pdf_path)
    blocks = []
    for page in doc:
        for b in page.get_text('blocks'):
            blocks.append(b)
    return blocks


def extract_full_text(pdf_path: str) -> str:
    """提取全部文本，处理编码问题"""
    doc = fitz.open(pdf_path)
    parts = []
    for page in doc:
        text = page.get_text()
        parts.append(text)
    return '\n'.join(parts)


def find_table_words(pdf_path: str, page_no: int = 0) -> list:
    """按坐标位置提取单词（用于银行对账单等）"""
    doc = fitz.open(pdf_path)
    if page_no >= len(doc):
        return []
    page = doc[page_no]
    words = page.get_text('words')  # (x0,y0,x1,y1,word,block_no,line_no,word_no)
    return words


def parse_number(text: str) -> float:
    """把 '49,977.26' 转成浮点数，空字符串返回 0.0"""
    if not text:
        return 0.0
    text = str(text).strip().replace(',', '').replace('，', '')
    if text in ('', '-', '--'):
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def clean_text(text: str) -> str:
    """清理文本中的多余空白和乱码字符"""
    if not text:
        return ''
    text = text.replace('\xa0', ' ').replace('\u3000', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return text
