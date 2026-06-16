# -*- coding: utf-8 -*-
"""
SmartQuote Streamlit MVP
실행: streamlit run streamlit_app.py
"""
from __future__ import annotations

import io
from datetime import datetime

import streamlit as st

from excel_builder import build_bundle_workbook
from excel_service import SUPPORTED_TERMS, load_workbook_for_term
from rental_calc import compute_rental

st.set_page_config(
    page_title='렌탈 수익률 분석',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='collapsed',
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

CALC_LABELS = {'irr': '수익률', 'fee': '월렌탈료', 'cost': '취득원가'}
IRR_LABELS = {'unlevered': '순수수익률', 'levered': '자기자본'}
TIMING_LABELS = {'m1': '1회차 청구', 'm0': '계약 시점'}
MODE_HINTS = {
    'irr': '취득원가 · 월렌탈료 입력 → <strong>수익률</strong> 산출',
    'fee': '취득원가 · 목표 IRR 입력 → <strong>월렌탈료</strong> 산출',
    'cost': '월렌탈료 · 목표 IRR 입력 → <strong>취득원가</strong> 산출',
}
SOLVED_TAGS = {'irr': '수익률 산출', 'fee': '월렌탈료 산출', 'cost': '취득원가 산출'}


def _inject_styles():
    st.markdown(
        """
        <style>
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');

        :root {
            --bg: #F2F4F6;
            --surface: #FFFFFF;
            --text: #191F28;
            --text-sub: #6B7684;
            --border: #E5E8EB;
            --accent: #3182F6;
            --accent-soft: #E8F3FF;
            --radius: 14px;
        }

        .stApp {
            background: var(--bg);
            font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }

        .block-container {
            max-width: 980px;
            padding-top: 1.25rem;
            padding-bottom: 2.5rem;
        }

        header[data-testid="stHeader"] { background: transparent; }

        h1 {
            font-size: 1.55rem !important;
            font-weight: 700 !important;
            letter-spacing: -0.03em;
            color: var(--text) !important;
            margin-bottom: 0.15rem !important;
        }

        h2, h3 {
            font-size: 0.95rem !important;
            font-weight: 600 !important;
            color: var(--text) !important;
            letter-spacing: -0.02em;
        }

        p, label, .stCaption, [data-testid="stMarkdownContainer"] p {
            color: var(--text-sub);
            font-size: 0.82rem;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--surface);
            border: 1px solid var(--border) !important;
            border-radius: var(--radius);
            padding: 1rem 1.1rem 0.6rem;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }

        div[data-testid="stMetric"] {
            background: var(--bg);
            border-radius: 12px;
            padding: 0.65rem 0.75rem;
        }

        div[data-testid="stMetricLabel"] {
            font-size: 0.72rem !important;
            color: var(--text-sub) !important;
        }

        div[data-testid="stMetricValue"] {
            font-size: 1.15rem !important;
            font-weight: 700 !important;
            color: var(--text) !important;
            letter-spacing: -0.02em;
        }

        div[data-testid="stNumberInput"] button,
        div[data-testid="stNumberInput"] [data-testid="stNumberInputStepDown"],
        div[data-testid="stNumberInput"] [data-testid="stNumberInputStepUp"] {
            display: none !important;
        }

        div[data-testid="stNumberInput"] input,
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] > div > div {
            border-radius: 10px !important;
            border-color: var(--border) !important;
            font-size: 0.88rem !important;
            min-height: 2.35rem;
        }

        div[data-testid="stNumberInput"] > div {
            border-radius: 10px !important;
        }

        .stTextInput label, .stNumberInput label, .stSelectbox label {
            font-size: 0.78rem !important;
            font-weight: 500 !important;
            color: var(--text-sub) !important;
        }

        div[data-testid="stSegmentedControl"] {
            background: var(--bg);
            border-radius: 10px;
            padding: 3px;
        }

        div[data-testid="stSegmentedControl"] button {
            font-size: 0.78rem !important;
            font-weight: 500 !important;
            border-radius: 8px !important;
        }

        div[data-testid="stExpander"] {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
        }

        .stButton > button {
            border-radius: 10px !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
            border: none !important;
            padding: 0.55rem 1rem !important;
            transition: opacity 0.15s;
        }

        .stButton > button[kind="primary"] {
            background: var(--accent) !important;
            color: white !important;
        }

        .stButton > button[kind="secondary"] {
            background: var(--surface) !important;
            color: var(--text) !important;
            border: 1px solid var(--border) !important;
        }

        .stDownloadButton > button {
            border-radius: 12px !important;
            background: var(--accent) !important;
            color: white !important;
            font-weight: 600 !important;
            padding: 0.7rem 1rem !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow: hidden;
        }

        .hero-irr {
            background: linear-gradient(135deg, #3182F6 0%, #1B64DA 100%);
            border-radius: 16px;
            padding: 1.25rem 1.4rem;
            color: white;
            margin-bottom: 0.75rem;
        }
        .hero-irr .label { font-size: 0.75rem; opacity: 0.85; margin-bottom: 0.2rem; }
        .hero-irr .value { font-size: 2rem; font-weight: 700; letter-spacing: -0.04em; line-height: 1.1; }
        .hero-irr .sub { font-size: 0.78rem; opacity: 0.8; margin-top: 0.35rem; }

        .mode-hint {
            background: var(--accent-soft);
            color: #1B64DA;
            font-size: 0.8rem;
            font-weight: 500;
            padding: 0.55rem 0.85rem;
            border-radius: 10px;
            margin: 0.25rem 0 0.75rem;
        }

        .solved-tag {
            display: inline-block;
            background: var(--accent-soft);
            color: #1B64DA;
            font-size: 0.72rem;
            font-weight: 600;
            padding: 0.2rem 0.55rem;
            border-radius: 6px;
            margin-bottom: 0.5rem;
        }

        .login-wrap {
            max-width: 360px;
            margin: 4rem auto 0;
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 20px;
            padding: 2rem 1.75rem 1.5rem;
            box-shadow: 0 8px 24px rgba(0,0,0,0.06);
            text-align: center;
        }
        .login-wrap h1 { font-size: 1.35rem !important; margin-bottom: 0.5rem !important; }

        section[data-testid="stSidebar"] {
            background: var(--surface);
            border-right: 1px solid var(--border);
        }

        #MainMenu, footer { visibility: hidden; }
        </style>
        """,
        unsafe_allow_html=True,
    )


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


def _get_app_password() -> str:
    try:
        return st.secrets.get('APP_PASSWORD', '') or ''
    except Exception:
        return ''


def _format_irr(irr_display: float) -> str:
    if irr_display != irr_display or irr_display < -100:
        return 'N/A'
    return f'{irr_display:.2f}%'


def _render_hero(calc_mode: str, result: dict, target_irr: float) -> tuple[str, str, str]:
    if calc_mode == 'fee':
        return (
            '월렌탈료 / 대',
            f'{result["monthly_fee"]:,}원',
            f'목표 IRR {target_irr:.2f}% · 총 월렌탈 {result["total_monthly_fee"]:,}원',
        )
    if calc_mode == 'cost':
        return (
            '취득원가 / 대',
            f'{result["cost"]:,}원',
            f'목표 IRR {target_irr:.2f}% · 월렌탈 {result["monthly_fee"]:,}원/대',
        )
    irr_text = _format_irr(result['irr'])
    return (
        '연 IRR',
        irr_text,
        f'월렌탈 {result["monthly_fee"]:,}원/대 · 총 {result["total_monthly_fee"]:,}원',
    )


def _render_scenario_values(calc_mode: str) -> tuple[float, float, float]:
    """모드별 입력만 노출. 산출 대상 필드는 제외."""
    if calc_mode == 'irr':
        c1, c2 = st.columns(2)
        with c1:
            cost = st.number_input('취득원가/대', min_value=0, value=0, step=10000, format='%d', key='in_cost')
        with c2:
            monthly_fee = st.number_input('월렌탈료/대', min_value=0, value=0, step=1000, format='%d', key='in_fee')
        return cost, monthly_fee, 0.0

    if calc_mode == 'fee':
        c1, c2 = st.columns(2)
        with c1:
            cost = st.number_input('취득원가/대', min_value=0, value=0, step=10000, format='%d', key='in_cost')
        with c2:
            target_irr = st.number_input(
                '목표 IRR (%)', min_value=0.0, value=15.0, step=0.01, format='%.2f', key='in_target_irr',
            )
        return cost, 0.0, target_irr

    c1, c2 = st.columns(2)
    with c1:
        monthly_fee = st.number_input('월렌탈료/대', min_value=0, value=0, step=1000, format='%d', key='in_fee')
    with c2:
        target_irr = st.number_input(
            '목표 IRR (%)', min_value=0.0, value=15.0, step=0.01, format='%.2f', key='in_target_irr',
        )
    return 0.0, monthly_fee, target_irr


def _segmented(label: str, options: list, labels: dict, key: str) -> str:
    try:
        return st.segmented_control(
            label,
            options=options,
            format_func=lambda x: labels[x],
            key=key,
            label_visibility='collapsed',
        )
    except Exception:
        return st.radio(label, options, format_func=lambda x: labels[x], key=key, horizontal=True)


def _render_login() -> bool:
    if st.session_state.get('authenticated'):
        return True

    st.markdown(
        '<div class="login-wrap">'
        '<h1>렌탈 수익률 분석</h1>'
        '<p style="margin:0 0 1.5rem;color:#6B7684;font-size:0.85rem;">담당자 전용</p></div>',
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1, 1.2, 1])
    with center:
        with st.form('login_form'):
            password = st.text_input('비밀번호', type='password', placeholder='비밀번호 입력')
            submitted = st.form_submit_button('로그인', type='primary', use_container_width=True)

    if submitted:
        expected = _get_app_password()
        if not expected:
            st.error('APP_PASSWORD가 설정되지 않았습니다.')
        elif password == expected:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error('비밀번호가 올바르지 않습니다.')

    return False


def _render_settings_bar() -> dict:
    st.markdown('##### 계산 설정')
    c1, c2, c3 = st.columns(3)
    with c1:
        calc_mode = _segmented('목표', ['irr', 'fee', 'cost'], CALC_LABELS, 'calc_mode')
    with c2:
        irr_type = _segmented('수익률', ['unlevered', 'levered'], IRR_LABELS, 'irr_type')
    with c3:
        timing_mode = _segmented('시점', ['m1', 'm0'], TIMING_LABELS, 'timing_mode')

    st.markdown(
        f'<div class="mode-hint">{MODE_HINTS[calc_mode]}</div>',
        unsafe_allow_html=True,
    )

    with st.expander('금리 · 판관비 · 잔존가치', expanded=False):
        g1, g2, g3, g4 = st.columns(4)
        with g1:
            borrow_rate = st.number_input('조달금리 (연%)', value=6.0, step=0.1, format='%.1f')
        with g2:
            sga_rate_pct = st.number_input(
                '판관비율 (월%)',
                value=0.105555555555,
                step=0.0001,
                format='%.6f',
            )
        with g3:
            residual = st.number_input('잔존가치 (원)', value=0, step=10000, format='%d')
        with g4:
            manager = st.text_input('담당자', value='', placeholder='메모')

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
    _inject_styles()

    if not _render_login():
        return

    _init_state()

    top_l, top_r = st.columns([5, 1])
    with top_l:
        st.title('렌탈 수익률 분석')
        st.caption('시나리오별 수익률분석표 엑셀 묶음')
    with top_r:
        if st.button('로그아웃', use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.scenarios = []
            st.rerun()

    globals_ = _render_settings_bar()
    st.markdown('<div style="height:0.5rem"></div>', unsafe_allow_html=True)

    col_l, col_r = st.columns(2, gap='medium')

    with col_l:
        with st.container(border=True):
            st.markdown('##### 시나리오')
            product_key = st.selectbox(
                '상품',
                PRODUCT_OPTIONS,
                format_func=lambda x: PRODUCT_LABELS.get(x, x),
                label_visibility='collapsed',
            )
            custom_name = ''
            if product_key == '__custom__':
                custom_name = st.text_input('상품명', placeholder='직접 입력')

            r1, r2 = st.columns(2)
            with r1:
                qty = st.number_input('수량 (대)', min_value=1, value=1, step=1)
            with r2:
                term = st.selectbox(
                    '기간 (개월)',
                    SUPPORTED_TERMS_LIST,
                    index=SUPPORTED_TERMS_LIST.index(36),
                )

            calc_mode = globals_['calc_mode']
            st.markdown(
                f'<span class="solved-tag">{SOLVED_TAGS[calc_mode]}</span>',
                unsafe_allow_html=True,
            )
            cost, monthly_fee, target_irr = _render_scenario_values(calc_mode)

            with st.expander('선수금 · 보증금 · 인수금', expanded=False):
                f1, f2, f3 = st.columns(3)
                with f1:
                    advance = st.number_input('선수금', min_value=0, value=0, step=10000, format='%d')
                with f2:
                    deposit = st.number_input('보증금', min_value=0, value=0, step=10000, format='%d')
                with f3:
                    buyout = st.number_input('인수금', min_value=0, value=0, step=10000, format='%d')

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

    irr_text = _format_irr(result['irr'])
    calc_mode = globals_['calc_mode']
    hero_label, hero_value, sub_line = _render_hero(calc_mode, result, target_irr)

    with col_r:
        with st.container(border=True):
            st.markdown('##### 결과')

            st.markdown(
                f'<div class="hero-irr">'
                f'<div class="label">{hero_label}</div>'
                f'<div class="value">{hero_value}</div>'
                f'<div class="sub">{sub_line}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

            m1, m2 = st.columns(2)
            if calc_mode == 'irr':
                m1.metric('총 월렌탈료', f'{result["total_monthly_fee"]:,}원')
                m2.metric('월 판관비', f'{result["monthly_sga"]:,}원')
            elif calc_mode == 'fee':
                m1.metric('연 IRR', irr_text)
                m2.metric('월 판관비', f'{result["monthly_sga"]:,}원')
            else:
                m1.metric('연 IRR', irr_text)
                m2.metric('총 월렌탈료', f'{result["total_monthly_fee"]:,}원')

            add_disabled = product_name is None
            if st.button('시나리오 추가', type='primary', disabled=add_disabled, use_container_width=True):
                row = _build_current_inputs(product_name, globals_, scenario_inputs, result)
                st.session_state.scenarios.append(row)
                st.rerun()
            if add_disabled:
                st.caption('상품을 선택하면 추가할 수 있습니다.')

    st.markdown('<div style="height:0.75rem"></div>', unsafe_allow_html=True)

    n = len(st.session_state.scenarios)
    with st.container(border=True):
        st.markdown(f'##### 시나리오 목록 · {n}건')

        if st.session_state.scenarios:
            display_rows = []
            for i, s in enumerate(st.session_state.scenarios):
                irr_val = s.get('irr', 0)
                irr_s = 'N/A' if irr_val < -100 else f'{irr_val:.2f}'
                display_rows.append({
                    '상품': s['productName'],
                    '기간': f"{s['term']}개월",
                    '수량': s['qty'],
                    '월렌탈료': f"{s['totalMonthlyFee']:,}",
                    'IRR': irr_s,
                })
            st.dataframe(display_rows, use_container_width=True, hide_index=True, height=min(44 + n * 35, 220))

            b1, b2, _ = st.columns([1, 1, 2])
            with b1:
                if st.button('전체 삭제', use_container_width=True):
                    st.session_state.scenarios = []
                    st.rerun()
            with b2:
                if st.button('마지막 삭제', use_container_width=True):
                    st.session_state.scenarios.pop()
                    st.rerun()
        else:
            st.caption('비어 있으면 현재 조건 1건이 엑셀에 포함됩니다.')

    lines = [_scenario_to_excel_line(s) for s in st.session_state.scenarios]
    if not lines:
        if product_name:
            cur = _build_current_inputs(product_name, globals_, scenario_inputs, result)
            lines = [_scenario_to_excel_line(cur)]
        else:
            lines = []

    st.markdown('<div style="height:0.5rem"></div>', unsafe_allow_html=True)

    if not lines:
        st.warning('상품을 선택하거나 시나리오를 추가해 주세요.')
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
                label=f'수익률분석표 다운로드 · {len(lines)}건',
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
