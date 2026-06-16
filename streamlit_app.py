# -*- coding: utf-8 -*-
"""
SmartQuote Streamlit MVP
실행: streamlit run streamlit_app.py
(기존 HTML/Flask와 별도 — 롤백 시 이 파일만 삭제하면 됨)
"""
from __future__ import annotations

import io
from datetime import datetime

import streamlit as st

from excel_builder import build_bundle_workbook
from excel_service import SUPPORTED_TERMS, load_workbook_for_term
from rental_calc import compute_rental

st.set_page_config(
    page_title='SmartQuote · 수익률 분석',
    page_icon='📊',
    layout='wide',
)

PRODUCT_OPTIONS = [
    '',
    '4족보행로봇',
    '서빙로봇 (기본형)',
    '물류로봇 (AMR)',
    '__custom__',
]
PRODUCT_LABELS = {
    '': '선택안함',
    '__custom__': '직접입력',
}
SUPPORTED_TERMS_LIST = sorted(SUPPORTED_TERMS)


def _init_state():
    if 'scenarios' not in st.session_state:
        st.session_state.scenarios = []


def _resolve_product(product_key: str, custom_name: str) -> str | None:
    if not product_key:
        return None
    if product_key == '__custom__':
        name = (custom_name or '').strip()
        return name or None
    return product_key


def _build_current_inputs(
    product_name: str,
    globals_: dict,
    scenario: dict,
    result: dict,
) -> dict:
    return {
        'productName': product_name,
        'cost': result['cost'],
        'qty': scenario['qty'],
        'term': scenario['term'],
        'totalMonthlyFee': result['total_monthly_fee'],
        'advance': scenario['advance'],
        'sgaRate': globals_['sga_rate_pct'],
        'irr': result['irr'],
    }


def _scenario_to_excel_line(scenario_row: dict) -> dict:
    return {
        'productName': scenario_row['productName'],
        'cost': scenario_row['cost'],
        'qty': scenario_row['qty'],
        'term': scenario_row['term'],
        'totalMonthlyFee': scenario_row['totalMonthlyFee'],
        'advance': scenario_row['advance'],
        'sgaRate': scenario_row['sgaRate'],
    }


def _render_sidebar() -> dict:
    st.sidebar.header('공통 설정')
    irr_type = st.sidebar.radio(
        '수익률 산출',
        options=['unlevered', 'levered'],
        format_func=lambda x: '순수 수익률 (엑셀 동일)' if x == 'unlevered' else '자기자본 수익률',
    )
    timing_mode = st.sidebar.radio(
        '초기대금 시점',
        options=['m1', 'm0'],
        format_func=lambda x: '1회차 청구 시' if x == 'm1' else '계약 시점',
    )
    calc_mode = st.sidebar.radio(
        '목표 계산',
        options=['irr', 'fee', 'cost'],
        format_func=lambda x: {'irr': '수익률(IRR)', 'fee': '월 렌탈료', 'cost': '취득원가'}[x],
    )
    borrow_rate = st.sidebar.number_input('조달 금리 (연 %)', value=6.0, step=0.1)
    sga_rate_pct = st.sidebar.number_input(
        '월별 판관비율 (%)',
        value=0.105555555555,
        step=0.0001,
        format='%.10f',
    )
    residual = st.sidebar.number_input('잔존가치 (원, 공통)', value=0, step=10000)
    manager = st.sidebar.text_input('담당자 (메모)', value='')

    st.sidebar.divider()
    st.sidebar.caption('기존 HTML 버전: `python server.py` → localhost:5000')

    return {
        'irr_type': irr_type,
        'timing_mode': timing_mode,
        'calc_mode': calc_mode,
        'borrow_rate': borrow_rate,
        'sga_rate_pct': sga_rate_pct,
        'residual': residual,
        'manager': manager,
    }


def main():
    _init_state()
    globals_ = _render_sidebar()

    st.title('렌탈 수익률 분석 (Streamlit)')
    st.caption('공급사 협의용 · 시나리오별 수익률분석표 엑셀 묶음 다운로드')

    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.subheader('시나리오 입력')
        product_key = st.selectbox(
            '상품',
            PRODUCT_OPTIONS,
            format_func=lambda x: PRODUCT_LABELS.get(x, x),
        )
        custom_name = ''
        if product_key == '__custom__':
            custom_name = st.text_input('상품명 직접 입력')

        c1, c2 = st.columns(2)
        with c1:
            qty = st.number_input('수량 (대)', min_value=1, value=1, step=1)
        with c2:
            term = st.selectbox('렌탈 기간 (개월)', SUPPORTED_TERMS_LIST, index=SUPPORTED_TERMS_LIST.index(36))

        cost = st.number_input('대당 취득원가 (원)', min_value=0, value=0, step=10000)
        monthly_fee = st.number_input('대당 월 렌탈료 (원)', min_value=0, value=0, step=1000)
        target_irr = st.number_input('목표 수익률 (%)', min_value=0.0, value=0.0, step=0.01)

        st.markdown('**초기/만기 현금흐름**')
        c3, c4 = st.columns(2)
        with c3:
            advance = st.number_input('선수금 (원)', min_value=0, value=0, step=10000)
            deposit = st.number_input('보증금 (원)', min_value=0, value=0, step=10000)
        with c4:
            buyout = st.number_input('인수금 (원)', min_value=0, value=0, step=10000)

    product_name = _resolve_product(product_key, custom_name)
    scenario_inputs = {
        'qty': int(qty),
        'term': int(term),
        'advance': float(advance),
        'deposit': float(deposit),
        'buyout': float(buyout),
    }

    result = compute_rental(
        mode=globals_['calc_mode'],
        irr_type=globals_['irr_type'],
        timing_mode=globals_['timing_mode'],
        cost=cost,
        qty=scenario_inputs['qty'],
        term=scenario_inputs['term'],
        borrow_rate=globals_['borrow_rate'],
        sga_rate_pct=globals_['sga_rate_pct'],
        advance=scenario_inputs['advance'],
        deposit=scenario_inputs['deposit'],
        residual=globals_['residual'],
        buyout=scenario_inputs['buyout'],
        monthly_fee=monthly_fee,
        target_irr=target_irr,
    )

    with col_r:
        st.subheader('계산 결과')
        m1, m2, m3 = st.columns(3)
        irr_display = result['irr']
        irr_text = 'N/A' if irr_display < -100 or irr_display != irr_display else f'{irr_display:.2f}%'
        m1.metric('IRR (연)', irr_text)
        m2.metric('총 월 렌탈료', f"{result['total_monthly_fee']:,}원")
        m3.metric('월 판관비', f"{result['monthly_sga']:,}원")

        if globals_['calc_mode'] == 'fee':
            st.info(f"역산 대당 월렌탈료: **{result['monthly_fee']:,}** 원")
        elif globals_['calc_mode'] == 'cost':
            st.info(f"역산 대당 취득원가: **{result['cost']:,}** 원")
        elif globals_['calc_mode'] == 'irr':
            st.info(f"산출 IRR: **{irr_text}**")

        st.divider()
        add_disabled = product_name is None
        if st.button('＋ 시나리오 목록에 추가', type='primary', disabled=add_disabled, use_container_width=True):
            row = _build_current_inputs(product_name, globals_, scenario_inputs, result)
            st.session_state.scenarios.append(row)
            st.rerun()

        if add_disabled:
            st.caption('상품을 선택하면 목록에 추가할 수 있습니다.')

    st.divider()
    st.subheader(f"시나리오 목록 ({len(st.session_state.scenarios)}건)")

    if st.session_state.scenarios:
        display_rows = []
        for i, s in enumerate(st.session_state.scenarios):
            irr_val = s.get('irr', 0)
            irr_s = 'N/A' if irr_val < -100 else f"{irr_val:.2f}"
            display_rows.append({
                '#': i + 1,
                '상품': s['productName'],
                '기간': f"{s['term']}개월",
                '수량': s['qty'],
                '월렌탈료': s['totalMonthlyFee'],
                'IRR(%)': irr_s,
            })
        st.dataframe(display_rows, use_container_width=True, hide_index=True)

        bc1, bc2, _ = st.columns([1, 1, 2])
        with bc1:
            if st.button('목록 전체 삭제', use_container_width=True):
                st.session_state.scenarios = []
                st.rerun()
        with bc2:
            if st.button('마지막 항목 삭제', use_container_width=True):
                st.session_state.scenarios.pop()
                st.rerun()
    else:
        st.info('목록이 비어 있으면 **현재 조건 1건**만 엑셀에 담깁니다.')

    lines = [_scenario_to_excel_line(s) for s in st.session_state.scenarios]
    if not lines:
        if product_name:
            cur = _build_current_inputs(product_name, globals_, scenario_inputs, result)
            lines = [_scenario_to_excel_line(cur)]
        else:
            lines = []

    st.divider()
    can_download = len(lines) > 0

    if not can_download:
        st.warning('상품을 선택하거나 시나리오를 추가해주세요.')
    else:
        try:
            excel_globals = {
                'borrowRate': globals_['borrow_rate'],
                'sgaRate': globals_['sga_rate_pct'],
                'irrType': globals_['irr_type'],
                'residual': globals_['residual'],
            }
            wb = build_bundle_workbook(lines, excel_globals, load_workbook_for_term)
            buf = io.BytesIO()
            wb.save(buf)
            fname = f"수익률분석_협의_{len(lines)}건_{datetime.now().strftime('%Y%m%d')}.xlsx"
            st.download_button(
                label=f'📊 수익률분석표 다운로드 ({len(lines)}건)',
                data=buf.getvalue(),
                file_name=fname,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
            )
        except Exception as e:
            st.error(f'엑셀 생성 오류: {e}')


if __name__ == '__main__':
    main()
