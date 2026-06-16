# -*- coding: utf-8 -*-
"""렌탈 IRR·역산 계산 (HTML v2.4 JS 로직 Python 포팅)."""


def compute_rental(
    *,
    mode: str = 'irr',
    irr_type: str = 'unlevered',
    timing_mode: str = 'm1',
    cost: float = 0,
    qty: int = 1,
    term: int = 36,
    borrow_rate: float = 6.0,
    sga_rate_pct: float = 0.105555555555,
    advance: float = 0,
    deposit: float = 0,
    residual: float = 0,
    buyout: float = 0,
    monthly_fee: float = 0,
    target_irr: float = 0,
) -> dict:
    """
    mode: 'irr' | 'fee' | 'cost'
    irr_type: 'unlevered' | 'levered'
    timing_mode: 'm1' | 'm0'
    sga_rate_pct: HTML과 동일, % 단위 (0.1056 = 0.1056%)
    """
    qty = max(int(qty or 1), 1)
    term = max(int(term or 1), 1)
    cost = float(cost or 0)
    monthly_fee = float(monthly_fee or 0)
    target_irr = float(target_irr or 0)
    borrow_rate = float(borrow_rate or 0)
    sga_rate = float(sga_rate_pct or 0) / 100.0
    b_rate_monthly = borrow_rate / 100.0 / 12.0

    advance = float(advance or 0)
    deposit = float(deposit or 0)
    residual = float(residual or 0)
    buyout = float(buyout or 0)

    total_cost = cost * qty
    total_monthly_fee = monthly_fee * qty
    monthly_sga = total_cost * sga_rate

    if mode == 'fee' and target_irr >= 0:
        r = target_irr / 100.0 / 12.0
        pv_interest_and_sga = 0.0
        pv_annuity_factor = 0.0

        for t in range(1, term + 1):
            discount = 1.0 if r == 0 else (1 + r) ** t
            pv_annuity_factor += 1.0 / discount

            principal_remain = total_cost - (total_cost / term) * (t - 1)
            interest_t = (principal_remain * b_rate_monthly) if irr_type == 'levered' else 0.0
            extra_m1 = (advance + deposit) if (t == 1 and timing_mode == 'm1') else 0.0
            pv_interest_and_sga += (interest_t + monthly_sga - extra_m1) / discount

        cf0 = (advance + deposit - total_cost) if timing_mode == 'm0' else (-total_cost)
        if r == 0:
            pv_end = residual + buyout - deposit
        else:
            pv_end = (residual / (1 + r) ** term) + ((buyout - deposit) / (1 + r) ** (term + 1))

        pv_known = cf0 + pv_end - pv_interest_and_sga
        total_monthly_fee = (-pv_known / term) if r == 0 else (-pv_known / pv_annuity_factor)
        monthly_fee = total_monthly_fee / qty

    elif mode == 'cost' and target_irr >= 0:
        r = target_irr / 100.0 / 12.0
        pv_pmt = 0.0
        sum_factor = 0.0

        for t in range(1, term + 1):
            interest_factor = ((1 - (t - 1) / term) * b_rate_monthly) if irr_type == 'levered' else 0.0
            deduction_factor = interest_factor + sga_rate
            discount = 1.0 if r == 0 else (1 + r) ** t
            sum_factor += deduction_factor / discount
            extra_m1 = (advance + deposit) if (t == 1 and timing_mode == 'm1') else 0.0
            pv_pmt += (total_monthly_fee + extra_m1) / discount

        m0_cash = (advance + deposit) if timing_mode == 'm0' else 0.0
        if r == 0:
            pv_end = residual + buyout - deposit
        else:
            pv_end = (residual / (1 + r) ** term) + ((buyout - deposit) / (1 + r) ** (term + 1))

        total_cost = (m0_cash + pv_pmt + pv_end) / (1 + sum_factor)
        cost = total_cost / qty
        monthly_sga = total_cost * sga_rate
        total_monthly_fee = monthly_fee * qty

    calculated_irr = target_irr
    if mode in ('irr', 'fee', 'cost'):
        end_cf = buyout - deposit
        r = 0.1 / 12.0

        for _ in range(100):
            npv = -total_cost
            if timing_mode == 'm0':
                npv += advance + deposit
            dnpv = 0.0

            for t in range(1, term + 1):
                principal_remain = total_cost - (total_cost / term) * (t - 1)
                interest_t = (principal_remain * b_rate_monthly) if irr_type == 'levered' else 0.0
                net_pmt = total_monthly_fee - interest_t - monthly_sga
                if t == 1 and timing_mode == 'm1':
                    net_pmt += advance + deposit

                npv += net_pmt / (1 + r) ** t
                dnpv -= (t * net_pmt) / (1 + r) ** (t + 1)

            npv += residual / (1 + r) ** term
            dnpv -= (term * residual) / (1 + r) ** (term + 1)
            npv += end_cf / (1 + r) ** (term + 1)
            dnpv -= ((term + 1) * end_cf) / (1 + r) ** (term + 2)

            if abs(dnpv) < 1e-10:
                break
            new_r = r - npv / dnpv
            if abs(new_r - r) < 1e-7:
                calculated_irr = new_r * 12.0 * 100.0
                break
            r = new_r

    if mode == 'irr':
        target_irr = calculated_irr

    return {
        'cost': round(cost),
        'monthly_fee': round(monthly_fee),
        'total_monthly_fee': round(total_monthly_fee),
        'monthly_sga': round(monthly_sga),
        'irr': calculated_irr,
        'target_irr': target_irr if mode == 'irr' else target_irr,
    }
