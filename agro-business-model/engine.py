#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Универсальный движок расчёта бизнес-профилей сельхозбизнеса.

Профиль = словарь со статьями CAPEX / выручки / OPEX, параметрами
финансирования и множителями прогнозов. Движок считает EBITDA, денежный
поток по 3 схемам финансирования (свои / грант / кредит) и инвест-метрики
(окупаемость, NPV, IRR) для 3 прогнозов (пессим / база / оптим).

Используется profiles.py (данные) и build_profiles_excel.py (Excel).
Чистый Python, без зависимостей.
"""

FIN_SCENARIOS = {"own": "Свои средства", "grant": "Грант КФХ", "loan": "Льготный кредит"}


def npv(rate, flows):
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(flows))


def irr(flows):
    lo, hi = -0.9, 5.0
    f_lo, f_hi = npv(lo, flows), npv(hi, flows)
    if f_lo * f_hi > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = npv(mid, flows)
        if abs(f_mid) < 1e-2:
            return mid
        if f_lo * f_mid < 0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def payback_period(flows):
    cum = 0.0
    for t, cf in enumerate(flows):
        prev = cum
        cum += cf
        if cum >= 0 and t > 0:
            return (t - 1) + (-prev) / cf if cf else float(t)
    return None


def annuity_payment(principal, rate, years):
    if years <= 0:
        return 0.0
    if rate == 0:
        return principal / years
    return principal * rate / (1 - (1 + rate) ** -years)


def capex_total(profile, mult=1.0):
    sub = sum(v for _, v in profile["capex"])
    return (sub * (1 + profile["capex_reserve_pct"])) * mult


def revenue_total(profile, mult=1.0):
    return sum(v for _, v in profile["revenue"]) * mult


def opex_total(profile, mult=1.0):
    return sum(v for _, v in profile["opex"]) * mult


def build_cashflow(profile, capex_tot, ebitda, scenario, discount_rate):
    n = profile["horizon_years"]
    infl = profile["inflation"]
    tax = profile["tax_rate"]
    fin = profile["financing"]

    own0 = capex_tot
    loan_pmt = 0.0
    if scenario == "grant":
        own0 = max(capex_tot - fin["grant_amount"], capex_tot * fin["grant_cofinance_pct"])
    elif scenario == "loan":
        principal = capex_tot * fin["loan_share"]
        own0 = capex_tot - principal
        loan_pmt = annuity_payment(principal, fin["loan_rate"], fin["loan_term"])

    flows = [-own0]
    for t in range(1, n + 1):
        g = (1 + infl) ** (t - 1)
        e = ebitda * g
        net = e - max(0, e) * tax
        if scenario == "loan" and t <= fin["loan_term"]:
            net -= loan_pmt
        flows.append(net)
    return flows


def compute(profile, forecast="Базовый"):
    f = profile["forecasts"][forecast]
    disc = profile["discount_rate"] + f["discount_add"]
    cap = capex_total(profile, f["capex"])
    rev = revenue_total(profile, f["revenue"])
    opx = opex_total(profile, f["opex"])
    ebitda = rev - opx
    fin = {}
    for k in FIN_SCENARIOS:
        flows = build_cashflow(profile, cap, ebitda, k, disc)
        fin[k] = {
            "flows": flows,
            "npv": npv(disc, flows),
            "irr": irr(flows),
            "payback": payback_period(flows),
            "own_investment": -flows[0] if flows[0] < 0 else 0,
        }
    return {"discount": disc, "capex": cap, "revenue": rev, "opex": opx,
            "ebitda": ebitda, "fin": fin}


def fmt(x):
    return f"{x:,.0f}".replace(",", " ")


if __name__ == "__main__":
    import profiles
    for key, prof in profiles.PROFILES.items():
        print(f"\n=== {prof['name']} ===")
        for fname in prof["forecasts"]:
            r = compute(prof, fname)
            g = r["fin"]["grant"]
            pb = f"{g['payback']:.1f}" if g["payback"] else "—"
            ir = f"{g['irr']*100:.0f}%" if g["irr"] else "—"
            print(f"  {fname:14} CAPEX {fmt(r['capex']):>11}  EBITDA {fmt(r['ebitda']):>10}"
                  f"  | грант: окуп {pb:>4} лет, NPV {fmt(g['npv']):>11}, IRR {ir}")
