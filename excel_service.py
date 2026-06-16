# -*- coding: utf-8 -*-
"""엑셀 템플릿 로드 (Streamlit / Flask 공용)."""
import glob
import os

import openpyxl

from excel_builder import prepare_workbook, rename_workbook_sheets

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
SUPPORTED_TERMS = {12, 24, 36, 48, 60}


def find_base_template():
    candidates = [os.path.join(TEMPLATES_DIR, 'base.xlsx')]
    for p in candidates:
        if os.path.exists(p):
            return p
    for p in glob.glob(os.path.join(BASE_DIR, '수익률분석표*.xlsx')):
        return p
    return None


def get_template_path(term: int):
    cached = os.path.join(TEMPLATES_DIR, f'수익률분석표_{term}개월.xlsx')
    if os.path.exists(cached):
        return cached
    return find_base_template()


def load_workbook_for_term(term: int):
    term = int(term)
    path = get_template_path(term)
    if not path:
        raise FileNotFoundError('템플릿 파일을 찾을 수 없습니다. templates/ 폴더를 확인하세요.')

    wb = openpyxl.load_workbook(path)
    cached = path.endswith(f'_{term}개월.xlsx')
    if not cached:
        prepare_workbook(wb, term)
    else:
        rename_workbook_sheets(wb, term)
    return wb
