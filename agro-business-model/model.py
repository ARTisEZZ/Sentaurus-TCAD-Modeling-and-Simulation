#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Финансовая модель сельхозбизнеса (Томская область).

Считает: CAPEX, OPEX, выручку, юнит-экономику (на корову / на га / на сотку),
10-летний денежный поток и инвест-метрики (срок окупаемости, NPV, IRR,
точка безубыточности) для ТРЁХ сценариев финансирования и анализ
чувствительности.

Запуск:   python3 model.py
Зависимости: только стандартная библиотека Python 3.8+.

Все суммы — в рублях (₽), горизонт по умолчанию 10 лет. Цифры ИЛЛЮСТРАТИВНЫЕ
(оценка 2026 г. по Томскому району) — правьте словарь PARAMS под свои данные
или положите рядом params.json (он переопределит значения).
"""

import csv
import json
import os
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))

# ---------------------------------------------------------------------------
#  ПАРАМЕТРЫ (редактируйте здесь или в params.json)
# ---------------------------------------------------------------------------
PARAMS = {
    "horizon_years": 10,
    "discount_rate": 0.18,          # ставка дисконтирования (ключевая + премия за риск)
    "inflation": 0.06,              # ежегодный рост цен и затрат
    "tax_eshn_rate": 0.06,          # ЕСХН: 6% с прибыли (доходы минус расходы)

    # --- Земля и площади ---
    "land": {
        "izhs_area_sotka": 15,      # участок под дом (ИЖС), соток
        "izhs_price_per_sotka": 80_000,
        "farm_area_ha": 3.0,        # сельхозземля под поля/выпас, га
        "farm_price_per_ha": 350_000,
    },

    # --- CAPEX, разовые капитальные вложения (₽) ---
    "capex": {
        "house_m2": 80,
        "house_price_per_m2": 45_000,
        "utilities": 600_000,       # скважина, электричество, септик, дорога
        "barn": 1_200_000,          # коровник на 6-10 голов
        "hay_storage": 300_000,     # сенник/склад
        "fencing_infra": 250_000,   # ограждение, навесы, мелочёвка
        "machinery": 900_000,       # мини-трактор/мотоблок + навесное
        "greenhouses": 200_000,     # теплицы и обустройство огорода
        "reserve_pct": 0.10,        # резерв на непредвиденное от суммы CAPEX
    },

    # --- Поголовье и закупка (₽) ---
    "herd": {
        "cows": 5,
        "cow_price": 120_000,
        "young_stock": 3,           # тёлки/бычки на выращивание
        "young_price": 50_000,
        "goats": 4, "goat_price": 12_000,
        "pigs": 4, "pig_price": 8_000,
        "hens": 30, "hen_price": 500,
    },

    # --- Продуктивность и цены (выручка) ---
    "production": {
        "milk_l_per_cow_year": 5_500,   # надой на корову, л/год
        "milk_price": 38,               # цена реализации молока, ₽/л
        "milk_to_cheese_share": 0.30,   # доля молока в переработку (сыр/творог)
        "cheese_yield": 0.10,           # кг сыра на 1 л молока
        "cheese_price": 900,            # ₽/кг
        "meat_kg_per_year": 600,        # мясо (выбраковка+бычки), кг/год
        "meat_price": 380,              # ₽/кг
        "eggs_per_hen_year": 230,
        "egg_price": 12,                # ₽/шт
        "veg_revenue_year": 250_000,    # овощи/огород, выручка ₽/год
    },

    # --- OPEX, годовые операционные затраты (₽) ---
    "opex": {
        "feed_purchased_per_cow": 60_000,   # покупные корма на корову, ₽/год
        "feed_self_offset_per_ha": 80_000,  # экономия на кормах с 1 га своих полей
        "vet_per_cow": 8_000,               # ветеринария + осеменение на корову
        "other_animals_feed": 120_000,      # корм козам/свиньям/птице
        "fuel": 120_000,                    # ГСМ
        "electricity": 90_000,
        "repairs_other": 100_000,           # ремонт, расходники, прочее
        "hired_labor": 0,                   # наёмный труд (0 = семейный)
    },

    # --- Сценарии финансирования ---
    "financing": {
        # Грант КФХ "Агростартап": разовое поступление, покрывает часть CAPEX
        "grant_amount": 4_000_000,
        "grant_own_cofinance_pct": 0.10,    # обязательное софинансирование грантополучателя
        # Льготный кредит: доля CAPEX в долг, субсидируемая ставка, срок
        "loan_share_of_capex": 0.60,
        "loan_rate": 0.05,                  # льготная ставка, годовых
        "loan_term_years": 7,
    },
}


def load_params():
    p = json.loads(json.dumps(PARAMS))  # глубокая копия
    fp = os.path.join(HERE, "params.json")
    if os.path.exists(fp):
        with open(fp, encoding="utf-8") as f:
            override = json.load(f)
        _deep_update(p, override)
    return p


def _deep_update(base, over):
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_update(base[k], v)
        else:
            base[k] = v
    return base


# ---------------------------------------------------------------------------
#  БЛОКИ РАСЧЁТА
# ---------------------------------------------------------------------------
def calc_capex(p):
    land = p["land"]
    c = p["capex"]
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
    }
    h = p["herd"]
    items["Закупка скота (коровы)"] = h["cows"] * h["cow_price"]
    items["Закупка молодняка"] = h["young_stock"] * h["young_price"]
    items["Прочие животные"] = (h["goats"] * h["goat_price"]
                                + h["pigs"] * h["pig_price"]
                                + h["hens"] * h["hen_price"])
    subtotal = sum(items.values())
    reserve = subtotal * c["reserve_pct"]
    items["Резерв (непредвиденное)"] = reserve
    total = subtotal + reserve
    return items, total


def calc_revenue(p):
    pr = p["production"]
    h = p["herd"]
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
    o = p["opex"]
    h = p["herd"]
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
    pr, h = p["production"], p["herd"]
    o = p["opex"]
    # на 1 корову
    milk = pr["milk_l_per_cow_year"]
    cow_rev = milk * pr["milk_price"]
    cow_cost = o["feed_purchased_per_cow"] + o["vet_per_cow"]
    # на 1 га полей (экономия = выручка-эквивалент)
    ha_value = o["feed_self_offset_per_ha"]
    return {
        "Выручка с 1 коровы (молоко), ₽/год": cow_rev,
        "Прямые затраты на 1 корову, ₽/год": cow_cost,
        "Маржа на 1 корову, ₽/год": cow_rev - cow_cost,
        "Эффект 1 га кормовых (экономия), ₽/год": ha_value,
    }


# ---------------------------------------------------------------------------
#  ИНВЕСТ-МЕТРИКИ
# ---------------------------------------------------------------------------
def npv(rate, flows):
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(flows))


def irr(flows):
    """IRR методом бисекции на [-0.9, 5.0]; None если знак не меняется."""
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
            hi, f_hi = mid, f_mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


def payback_period(flows):
    """Простой срок окупаемости в годах (с дробной частью), None если не окупается."""
    cum = 0.0
    for t, cf in enumerate(flows):
        prev = cum
        cum += cf
        if cum >= 0 and t > 0:
            need = -prev
            frac = need / cf if cf != 0 else 0
            return (t - 1) + frac
    return None


def annuity_payment(principal, rate, years):
    if years <= 0:
        return 0.0
    if rate == 0:
        return principal / years
    return principal * rate / (1 - (1 + rate) ** -years)


def build_cashflow(p, capex_total, annual_rev, annual_opex, scenario):
    """Возвращает список годовых чистых потоков (год 0 = инвестиции)."""
    n = p["horizon_years"]
    infl = p["inflation"]
    tax = p["tax_eshn_rate"]
    fin = p["financing"]

    # Стартовые вложения собственных средств в год 0 и долговое обслуживание
    own_capex_y0 = capex_total
    grant_inflow_y0 = 0.0
    loan_principal = 0.0
    loan_pmt = 0.0

    if scenario == "grant":
        grant_inflow_y0 = fin["grant_amount"]
        # собственные = CAPEX - грант, но не меньше обязательного софинансирования
        own_capex_y0 = max(capex_total - fin["grant_amount"],
                           capex_total * fin["grant_own_cofinance_pct"])
    elif scenario == "loan":
        loan_principal = capex_total * fin["loan_share_of_capex"]
        own_capex_y0 = capex_total - loan_principal
        loan_pmt = annuity_payment(loan_principal, fin["loan_rate"], fin["loan_term_years"])

    flows = []
    for t in range(n + 1):
        if t == 0:
            flows.append(-own_capex_y0 + grant_inflow_y0)
            continue
        g = (1 + infl) ** (t - 1)
        rev = annual_rev * g
        opex = annual_opex * g
        ebitda = rev - opex
        interest = 0.0
        principal_pay = 0.0
        if scenario == "loan" and t <= fin["loan_term_years"]:
            # для налога считаем проценты; тело долга — отток, но не расход
            # упрощённо: проценты = остаток*ставка на начало года
            pass
        taxable = max(0, ebitda)  # амортизацию для ЕСХН-упрощения опускаем
        tax_amt = taxable * tax
        net = ebitda - tax_amt
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


# ---------------------------------------------------------------------------
#  ОТЧЁТ
# ---------------------------------------------------------------------------
def fmt(x):
    return f"{x:,.0f}".replace(",", " ")


def main():
    p = load_params()
    capex_items, capex_total = calc_capex(p)
    rev_items, annual_rev = calc_revenue(p)
    opex_items, annual_opex = calc_opex(p)
    ue = unit_economics(p)
    ebitda = annual_rev - annual_opex

    scenarios = {
        "own": "Свои средства (100%)",
        "grant": "Грант КФХ + свои",
        "loan": "Льготный кредит",
    }
    results = {k: scenario_metrics(p, capex_total, annual_rev, annual_opex, k)
               for k in scenarios}

    # --- Markdown отчёт ---
    L = []
    L.append("# Финансовая модель сельхозбизнеса (Томская область)\n")
    L.append(f"_Сгенерировано {date.today().isoformat()} скриптом `model.py`. "
             f"Все суммы в ₽. Горизонт {p['horizon_years']} лет, "
             f"ставка дисконтирования {p['discount_rate']*100:.0f}%._\n")

    L.append("## 1. Капитальные вложения (CAPEX)\n")
    L.append("| Статья | Сумма, ₽ |\n|---|---:|")
    for k, v in capex_items.items():
        L.append(f"| {k} | {fmt(v)} |")
    L.append(f"| **ИТОГО CAPEX** | **{fmt(capex_total)}** |\n")

    L.append("## 2. Годовая выручка\n")
    L.append("| Источник | ₽/год |\n|---|---:|")
    for k, v in rev_items.items():
        L.append(f"| {k} | {fmt(v)} |")
    L.append(f"| **ИТОГО выручка** | **{fmt(annual_rev)}** |\n")

    L.append("## 3. Годовые операционные затраты (OPEX)\n")
    L.append("| Статья | ₽/год |\n|---|---:|")
    for k, v in opex_items.items():
        L.append(f"| {k} | {fmt(v)} |")
    L.append(f"| **ИТОГО OPEX** | **{fmt(annual_opex)}** |\n")
    L.append(f"**EBITDA (выручка − OPEX): {fmt(ebitda)} ₽/год**\n")

    L.append("## 4. Юнит-экономика\n")
    L.append("| Показатель | Значение |\n|---|---:|")
    for k, v in ue.items():
        L.append(f"| {k} | {fmt(v)} |")
    L.append("")

    L.append("## 5. Сценарии финансирования и окупаемость\n")
    L.append("| Сценарий | Свои вложения, ₽ | Окупаемость, лет | NPV, ₽ | IRR |\n|---|---:|---:|---:|---:|")
    for k, name in scenarios.items():
        r = results[k]
        pb = f"{r['payback']:.1f}" if r["payback"] is not None else "не окупается"
        ir = f"{r['irr']*100:.1f}%" if r["irr"] is not None else "—"
        L.append(f"| {name} | {fmt(r['own_investment'])} | {pb} | {fmt(r['npv'])} | {ir} |")
    L.append("")

    L.append("## 6. Денежный поток по годам (₽)\n")
    header = "| Год | " + " | ".join(scenarios.values()) + " |"
    L.append(header)
    L.append("|---|" + "---:|" * len(scenarios))
    for t in range(p["horizon_years"] + 1):
        row = [str(t)] + [fmt(results[k]["flows"][t]) for k in scenarios]
        L.append("| " + " | ".join(row) + " |")
    L.append("")

    # --- Точка безубыточности по поголовью ---
    pr, o = p["production"], p["opex"]
    margin_per_cow = (pr["milk_l_per_cow_year"] * pr["milk_price"]
                      - o["feed_purchased_per_cow"] - o["vet_per_cow"])
    fixed = (annual_opex - p["herd"]["cows"]
             * (o["feed_purchased_per_cow"] + o["vet_per_cow"]))
    # сколько коров нужно, чтобы покрыть условно-постоянные затраты
    be_cows = fixed / margin_per_cow if margin_per_cow > 0 else None
    L.append("## 7. Точка безубыточности\n")
    if be_cows is not None:
        L.append(f"- Маржа на 1 корову: **{fmt(margin_per_cow)} ₽/год**")
        L.append(f"- Условно-постоянные затраты: **{fmt(fixed)} ₽/год**")
        L.append(f"- Безубыточное поголовье (только дойное стадо vs пост. затраты): "
                 f"**~{be_cows:.1f} коров**\n")

    # --- Чувствительность (NPV, сценарий "свои средства") ---
    L.append("## 8. Анализ чувствительности (NPV, сценарий «свои средства»)\n")
    L.append("Изменяем цену молока и стоимость кормов на ±20%.\n")
    L.append("| | Корма −20% | Корма базовые | Корма +20% |\n|---|---:|---:|---:|")
    base_milk = p["production"]["milk_price"]
    base_feed = p["opex"]["feed_purchased_per_cow"]
    for milk_d, milk_lbl in [(-0.2, "Молоко −20%"), (0.0, "Молоко базовое"), (0.2, "Молоко +20%")]:
        cells = []
        for feed_d in (-0.2, 0.0, 0.2):
            pp = json.loads(json.dumps(p))
            pp["production"]["milk_price"] = base_milk * (1 + milk_d)
            pp["opex"]["feed_purchased_per_cow"] = base_feed * (1 + feed_d)
            _, rv = calc_revenue(pp)
            _, ox = calc_opex(pp)
            m = scenario_metrics(pp, capex_total, rv, ox, "own")
            cells.append(fmt(m["npv"]))
        L.append(f"| **{milk_lbl}** | " + " | ".join(cells) + " |")
    L.append("")

    L.append("---\n_Цифры иллюстративные. Уточняйте через MCP `fetch`/`brave-search` "
             "актуальные цены по Томску и правьте `PARAMS`/`params.json`._")

    report = "\n".join(L)
    with open(os.path.join(HERE, "report.md"), "w", encoding="utf-8") as f:
        f.write(report)

    # --- CSV денежного потока ---
    with open(os.path.join(HERE, "cashflow.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["year"] + list(scenarios.keys()))
        for t in range(p["horizon_years"] + 1):
            w.writerow([t] + [round(results[k]["flows"][t]) for k in scenarios])

    print(report)
    print("\n[OK] Записаны report.md и cashflow.csv в", HERE)


if __name__ == "__main__":
    main()
