#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Финансовая модель сельхозбизнеса (Томская область) — РЕАЛЬНЫЕ ЦЕНЫ 2026.

Считает CAPEX, OPEX, выручку, юнит-экономику, 10-летний денежный поток и
инвест-метрики (окупаемость, NPV, IRR, безубыточность) для ТРЁХ схем
финансирования (свои / грант КФХ / льготный кредит) и ТРЁХ прогнозов
(пессимистичный / базовый / оптимистичный).

Запуск:   python3 model.py          -> печатает отчёт, пишет report.md, cashflow.csv
Зависимости: только стандартная библиотека Python 3.8+.

Источники цен (см. README → «Источники»): рынок участков Томского района,
сельхозземля, строительство, закупочные цены на молоко, цены на КРС,
грант «Агростартап» (Томск 3–6 млн, до 7 млн на КРС), льготный кредит РСХБ
(ставка до 5%, срок до 7 лет). Все суммы в рублях (₽).
"""

import csv
import json
import os
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
#  БАЗОВЫЕ ПАРАМЕТРЫ — реальные ориентиры по Томску (2026). Правьте здесь
#  или в params.json (он переопределяет значения).
# ---------------------------------------------------------------------------
PARAMS = {
    "horizon_years": 10,
    "discount_rate": 0.18,          # ставка дисконтирования (ключевая + премия за риск)
    "inflation": 0.06,              # годовой рост цен и затрат
    "tax_eshn_rate": 0.06,          # ЕСХН: 6% с прибыли

    "land": {
        "izhs_area_sotka": 15,
        "izhs_price_per_sotka": 130_000,   # Томский р-н: 100–300 тыс/сотка, берём средне
        "farm_area_ha": 3.0,
        "farm_price_per_ha": 400_000,      # сельхозземля: ~до 600 тыс/га
    },

    "capex": {
        "house_m2": 80,
        "house_price_per_m2": 50_000,      # Томск: каркас/брус, дешевле Подмосковья
        "utilities": 600_000,              # скважина, э/э, септик, подъезд
        "barn": 1_200_000,
        "hay_storage": 300_000,
        "fencing_infra": 250_000,
        "machinery": 900_000,
        "greenhouses": 200_000,
        "reserve_pct": 0.10,
    },

    "herd": {
        "cows": 5,
        "cow_price": 50_000,               # Томск: дойная 40–53 тыс
        "young_stock": 3,
        "young_price": 45_000,             # нетель ~45 тыс
        "goats": 4, "goat_price": 12_000,
        "pigs": 4, "pig_price": 8_000,
        "hens": 30, "hen_price": 500,
    },

    "production": {
        "milk_l_per_cow_year": 5_500,
        "milk_price": 30,                  # КФХ-реализация (завод даёт меньше, чем крупным)
        "milk_to_cheese_share": 0.30,
        "cheese_yield": 0.10,
        "cheese_price": 900,
        "meat_kg_per_year": 600,
        "meat_price": 380,
        "eggs_per_hen_year": 230,
        "egg_price": 12,
        "veg_revenue_year": 250_000,
    },

    "opex": {
        "feed_purchased_per_cow": 60_000,
        "feed_self_offset_per_ha": 80_000,
        "vet_per_cow": 8_000,
        "other_animals_feed": 120_000,
        "fuel": 120_000,
        "electricity": 90_000,
        "repairs_other": 100_000,
        "hired_labor": 0,
    },

    "financing": {
        "grant_amount": 5_000_000,         # Агростартап Томск: 3–6 млн, до 7 на КРС
        "grant_own_cofinance_pct": 0.10,
        "loan_share_of_capex": 0.60,
        "loan_rate": 0.05,                 # льготная ставка РСХБ до 5%
        "loan_term_years": 7,
    },

    # Множители прогнозов: применяются к ключевым драйверам.
    "forecasts": {
        "Пессимистичный": {
            "milk_price": 0.80, "milk_yield": 0.85, "feed_cost": 1.20,
            "capex": 1.15, "veg_revenue": 0.70, "meat_price": 0.85,
            "discount_add": 0.04,
        },
        "Базовый": {
            "milk_price": 1.00, "milk_yield": 1.00, "feed_cost": 1.00,
            "capex": 1.00, "veg_revenue": 1.00, "meat_price": 1.00,
            "discount_add": 0.00,
        },
        "Оптимистичный": {
            "milk_price": 1.20, "milk_yield": 1.15, "feed_cost": 0.85,
            "capex": 0.90, "veg_revenue": 1.30, "meat_price": 1.15,
            "discount_add": -0.03,
        },
    },
}


def load_params():
    p = json.loads(json.dumps(PARAMS))
    fp = os.path.join(HERE, "params.json")
    if os.path.exists(fp):
        with open(fp, encoding="utf-8") as f:
            _deep_update(p, json.load(f))
    return p


def _deep_update(base, over):
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_update(base[k], v)
        else:
            base[k] = v
    return base


def apply_forecast(p, name):
    """Вернуть копию параметров с применёнными множителями прогноза."""
    pp = json.loads(json.dumps(p))
    f = p["forecasts"][name]
    pp["production"]["milk_price"] *= f["milk_price"]
    pp["production"]["milk_l_per_cow_year"] *= f["milk_yield"]
    pp["production"]["meat_price"] *= f["meat_price"]
    pp["production"]["veg_revenue_year"] *= f["veg_revenue"]
    pp["opex"]["feed_purchased_per_cow"] *= f["feed_cost"]
    pp["_capex_mult"] = f["capex"]
    pp["discount_rate"] = p["discount_rate"] + f["discount_add"]
    return pp


# ---------------------------------------------------------------------------
#  БЛОКИ РАСЧЁТА
# ---------------------------------------------------------------------------
def calc_capex(p):
    land, c, h = p["land"], p["capex"], p["herd"]
    items = {
        "Земля ИЖС (участок под дом)": land["izhs_area_sotka"] * land["izhs_price_per_sotka"],
        "Земля сельхоз (поля/выпас)": land["farm_area_ha"] * land["farm_price_per_ha"],
        "Дом": c["house_m2"] * c["house_price_per_m2"],
        "Коммуникации": c["utilities"],
        "Коровник": c["barn"],
        "Сенник/склад": c["hay_storage"],
        "Ограждение/инфраструктура": c["fencing_infra"],
        "Техника": c["machinery"],
        "Теплицы/огород": c["greenhouses"],
        "Закупка скота (коровы)": h["cows"] * h["cow_price"],
        "Закупка молодняка": h["young_stock"] * h["young_price"],
        "Прочие животные": (h["goats"] * h["goat_price"] + h["pigs"] * h["pig_price"]
                            + h["hens"] * h["hen_price"]),
    }
    subtotal = sum(items.values())
    items["Резерв (непредвиденное)"] = subtotal * c["reserve_pct"]
    total = (subtotal + items["Резерв (непредвиденное)"]) * p.get("_capex_mult", 1.0)
    return items, total


def calc_revenue(p):
    pr, h = p["production"], p["herd"]
    milk_total = h["cows"] * pr["milk_l_per_cow_year"]
    milk_to_cheese = milk_total * pr["milk_to_cheese_share"]
    milk_sold = milk_total - milk_to_cheese
    rev = {
        "Молоко (сырое)": milk_sold * pr["milk_price"],
        "Сыр/творог (переработка)": milk_to_cheese * pr["cheese_yield"] * pr["cheese_price"],
        "Мясо": pr["meat_kg_per_year"] * pr["meat_price"],
        "Яйца": h["hens"] * pr["eggs_per_hen_year"] * pr["egg_price"],
        "Овощи/огород": pr["veg_revenue_year"],
    }
    return rev, sum(rev.values())


def calc_opex(p):
    o, h = p["opex"], p["herd"]
    feed_self_offset = p["land"]["farm_area_ha"] * o["feed_self_offset_per_ha"]
    feed_net = max(0, h["cows"] * o["feed_purchased_per_cow"] - feed_self_offset)
    items = {
        "Корма (за вычетом своих полей)": feed_net,
        "Корм прочим животным": o["other_animals_feed"],
        "Ветеринария/осеменение": h["cows"] * o["vet_per_cow"],
        "ГСМ": o["fuel"],
        "Электричество": o["electricity"],
        "Ремонт/прочее": o["repairs_other"],
        "Наёмный труд": o["hired_labor"],
    }
    return items, sum(items.values())


def unit_economics(p):
    pr, h, o = p["production"], p["herd"], p["opex"]
    cow_rev = pr["milk_l_per_cow_year"] * pr["milk_price"]
    cow_cost = o["feed_purchased_per_cow"] + o["vet_per_cow"]
    return {
        "Выручка с 1 коровы (молоко), ₽/год": cow_rev,
        "Прямые затраты на 1 корову, ₽/год": cow_cost,
        "Маржа на 1 корову, ₽/год": cow_rev - cow_cost,
        "Эффект 1 га кормовых (экономия), ₽/год": o["feed_self_offset_per_ha"],
    }


# ---------------------------------------------------------------------------
#  ИНВЕСТ-МЕТРИКИ
# ---------------------------------------------------------------------------
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


def build_cashflow(p, capex_total, annual_rev, annual_opex, scenario):
    n, infl, tax = p["horizon_years"], p["inflation"], p["tax_eshn_rate"]
    fin = p["financing"]
    # own_capex_y0 = СОБСТВЕННЫЕ деньги, вложенные в год 0 (грант/кредит покрывают остальное).
    own_capex_y0 = capex_total
    loan_pmt = 0.0

    if scenario == "grant":
        # грант покрывает часть CAPEX напрямую; своих — остаток, но не меньше софинансирования
        own_capex_y0 = max(capex_total - fin["grant_amount"],
                           capex_total * fin["grant_own_cofinance_pct"])
    elif scenario == "loan":
        loan_principal = capex_total * fin["loan_share_of_capex"]
        own_capex_y0 = capex_total - loan_principal
        loan_pmt = annuity_payment(loan_principal, fin["loan_rate"], fin["loan_term_years"])

    flows = []
    for t in range(n + 1):
        if t == 0:
            flows.append(-own_capex_y0)
            continue
        g = (1 + infl) ** (t - 1)
        ebitda = annual_rev * g - annual_opex * g
        net = ebitda - max(0, ebitda) * tax
        if scenario == "loan" and t <= fin["loan_term_years"]:
            net -= loan_pmt
        flows.append(net)
    return flows


def scenario_metrics(p, capex_total, annual_rev, annual_opex, scenario):
    flows = build_cashflow(p, capex_total, annual_rev, annual_opex, scenario)
    return {
        "flows": flows,
        "npv": npv(p["discount_rate"], flows),
        "irr": irr(flows),
        "payback": payback_period(flows),
        "own_investment": -flows[0] if flows[0] < 0 else 0,
    }


FIN_SCENARIOS = {"own": "Свои средства (100%)", "grant": "Грант КФХ + свои", "loan": "Льготный кредит"}


def run_all(p):
    """Полный расчёт по всем прогнозам и схемам финансирования."""
    out = {}
    for fname in p["forecasts"]:
        pp = apply_forecast(p, fname)
        capex_items, capex_total = calc_capex(pp)
        rev_items, annual_rev = calc_revenue(pp)
        opex_items, annual_opex = calc_opex(pp)
        out[fname] = {
            "params": pp,
            "capex_items": capex_items, "capex_total": capex_total,
            "rev_items": rev_items, "annual_rev": annual_rev,
            "opex_items": opex_items, "annual_opex": annual_opex,
            "ebitda": annual_rev - annual_opex,
            "unit": unit_economics(pp),
            "fin": {k: scenario_metrics(pp, capex_total, annual_rev, annual_opex, k)
                    for k in FIN_SCENARIOS},
        }
    return out


# ---------------------------------------------------------------------------
#  ТЕКСТОВЫЙ ОТЧЁТ
# ---------------------------------------------------------------------------
def fmt(x):
    return f"{x:,.0f}".replace(",", " ")


def main():
    p = load_params()
    res = run_all(p)
    base = res["Базовый"]

    L = ["# Финмодель сельхозбизнеса (Томск, реальные цены 2026)\n"]
    L.append(f"_Сгенерировано {date.today().isoformat()}. Все суммы в ₽. "
             f"Горизонт {p['horizon_years']} лет._\n")

    L.append("## CAPEX (базовый сценарий)\n| Статья | ₽ |\n|---|---:|")
    for k, v in base["capex_items"].items():
        L.append(f"| {k} | {fmt(v)} |")
    L.append(f"| **ИТОГО CAPEX** | **{fmt(base['capex_total'])}** |\n")

    L.append("## Годовой P&L (базовый)\n| | ₽/год |\n|---|---:|")
    L.append(f"| Выручка | {fmt(base['annual_rev'])} |")
    L.append(f"| OPEX | {fmt(base['annual_opex'])} |")
    L.append(f"| **EBITDA** | **{fmt(base['ebitda'])}** |\n")

    L.append("## Окупаемость по прогнозам × финансированию\n")
    L.append("| Прогноз | Схема | Свои вложения | Окуп., лет | NPV | IRR |")
    L.append("|---|---|---:|---:|---:|---:|")
    for fname in p["forecasts"]:
        for k, name in FIN_SCENARIOS.items():
            m = res[fname]["fin"][k]
            pb = f"{m['payback']:.1f}" if m["payback"] is not None else "—"
            ir = f"{m['irr']*100:.1f}%" if m["irr"] is not None else "—"
            L.append(f"| {fname} | {name} | {fmt(m['own_investment'])} | {pb} | {fmt(m['npv'])} | {ir} |")
    L.append("")

    report = "\n".join(L)
    with open(os.path.join(HERE, "report.md"), "w", encoding="utf-8") as f:
        f.write(report)
    with open(os.path.join(HERE, "cashflow.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["forecast", "financing", "year", "net_flow"])
        for fname in p["forecasts"]:
            for k in FIN_SCENARIOS:
                for t, cf in enumerate(res[fname]["fin"][k]["flows"]):
                    w.writerow([fname, k, t, round(cf)])

    print(report)
    print("\n[OK] report.md и cashflow.csv записаны в", HERE)


if __name__ == "__main__":
    main()
