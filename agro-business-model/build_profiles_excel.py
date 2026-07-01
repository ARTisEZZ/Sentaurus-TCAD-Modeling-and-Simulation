#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор Excel-книги сравнения бизнес-профилей: agro_business_profiles.xlsx

Листы:
  Сравнение           — сводка по 3 профилям × 3 прогнозам × 3 схемам финанс.
  <профиль>           — CAPEX/Выручка/OPEX на формулах + денежный поток (NPV/IRR)
                        + прогнозы (3 сценария) + заметки
  Земля Томск         — конкретные направления участков (50 км)
  Источники

Запуск: python3 build_profiles_excel.py   (нужен openpyxl)
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import engine as E
import profiles as PR

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "agro_business_profiles.xlsx")

H1 = Font(bold=True, size=14, color="FFFFFF")
H2 = Font(bold=True, size=11, color="FFFFFF")
BOLD = Font(bold=True)
BLUE = Font(color="0000CC")
TITLE_FILL = PatternFill("solid", fgColor="2F5496")
HEAD_FILL = PatternFill("solid", fgColor="4472C4")
GROUP_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
GOOD = PatternFill("solid", fgColor="C6EFCE")
BAD = PatternFill("solid", fgColor="FFC7CE")
RUB = '# ##0 \\₽'
PCT = '0.0%'
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def head(ws, row, ncols, fill=HEAD_FILL, font=H2):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def title(ws, text, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=text)
    c.fill = TITLE_FILL; c.font = H1
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26


def profile_sheet(wb, key, prof):
    short = {"greenhouse": "Теплицы", "jersey": "Джерси+птица", "mixed": "Смеш.ферма"}[key]
    ws = wb.create_sheet(short)
    title(ws, prof["name"], 5)
    rr = 2

    # --- CAPEX ---
    ws.cell(row=rr, column=1, value="CAPEX (разовые вложения)").font = BOLD
    for c in range(1, 3): ws.cell(row=rr, column=c).fill = GROUP_FILL
    rr += 1
    cap_first = rr
    for label, val in prof["capex"]:
        ws.cell(row=rr, column=1, value=label)
        c = ws.cell(row=rr, column=2, value=val); c.number_format = RUB; c.font = BLUE
        rr += 1
    cap_last = rr - 1
    ws.cell(row=rr, column=1, value="Подытог").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"=SUM(B{cap_first}:B{cap_last})"); c.number_format = RUB; c.font = BOLD
    sub = rr; rr += 1
    ws.cell(row=rr, column=1, value=f"Резерв ({prof['capex_reserve_pct']:.0%})").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"=B{sub}*{prof['capex_reserve_pct']}"); c.number_format = RUB
    resrow = rr; rr += 1
    ws.cell(row=rr, column=1, value="ИТОГО CAPEX").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"=B{sub}+B{resrow}"); c.number_format = RUB; c.font = BOLD
    c.fill = TOTAL_FILL; ws.cell(row=rr, column=1).fill = TOTAL_FILL
    CAPEX = f"B{rr}"; rr += 2

    # --- Выручка ---
    ws.cell(row=rr, column=1, value="ВЫРУЧКА (₽/год)").font = BOLD
    for c in range(1, 3): ws.cell(row=rr, column=c).fill = GROUP_FILL
    rr += 1
    rev_first = rr
    for label, val in prof["revenue"]:
        ws.cell(row=rr, column=1, value=label)
        c = ws.cell(row=rr, column=2, value=val); c.number_format = RUB; c.font = BLUE
        rr += 1
    rev_last = rr - 1
    ws.cell(row=rr, column=1, value="ИТОГО выручка").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"=SUM(B{rev_first}:B{rev_last})"); c.number_format = RUB; c.font = BOLD
    REV = f"B{rr}"; rr += 2

    # --- OPEX ---
    ws.cell(row=rr, column=1, value="OPEX (₽/год)").font = BOLD
    for c in range(1, 3): ws.cell(row=rr, column=c).fill = GROUP_FILL
    rr += 1
    op_first = rr
    for label, val in prof["opex"]:
        ws.cell(row=rr, column=1, value=label)
        c = ws.cell(row=rr, column=2, value=val); c.number_format = RUB; c.font = BLUE
        rr += 1
    op_last = rr - 1
    ws.cell(row=rr, column=1, value="ИТОГО OPEX").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"=SUM(B{op_first}:B{op_last})"); c.number_format = RUB; c.font = BOLD
    OPEX = f"B{rr}"; rr += 1
    ws.cell(row=rr, column=1, value="EBITDA (выручка − OPEX)").font = BOLD
    c = ws.cell(row=rr, column=2, value=f"={REV}-{OPEX}"); c.number_format = RUB; c.font = BOLD
    c.fill = TOTAL_FILL; ws.cell(row=rr, column=1).fill = TOTAL_FILL
    EBITDA = f"B{rr}"; rr += 2

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 15

    # --- Денежный поток (база), формулы ---
    ws.cell(row=rr, column=1, value="ДЕНЕЖНЫЙ ПОТОК — базовый прогноз").font = BOLD
    for c in range(1, 6): ws.cell(row=rr, column=c).fill = GROUP_FILL
    rr += 1
    fin = prof["financing"]
    disc = prof["discount_rate"]; infl = prof["inflation"]; tax = prof["tax_rate"]
    n = prof["horizon_years"]
    # параметры справа
    ws.cell(row=rr, column=4, value="Ставка дисконт.").font = BOLD
    ws.cell(row=rr, column=5, value=disc).number_format = PCT
    pmt_r = rr
    ws.cell(row=rr + 1, column=4, value="Платёж по кредиту/год").font = BOLD
    pmt_cell = f"E{rr+1}"
    c = ws.cell(row=rr + 1, column=5,
                value=f"={CAPEX}*{fin['loan_share']}*{fin['loan_rate']}/(1-(1+{fin['loan_rate']})^-{fin['loan_term']})")
    c.number_format = RUB
    disc_cell = f"E{rr}"
    # шапка
    ws.cell(row=rr, column=1, value="Год").font = BOLD
    cols = list(E.FIN_SCENARIOS.values())
    for i, name in enumerate(cols):
        ws.cell(row=rr, column=2 + i, value=name)
    head(ws, rr, 1 + len(cols))
    rr += 1
    keys = list(E.FIN_SCENARIOS.keys())
    own0 = {
        "own": f"=-{CAPEX}",
        "grant": f"=-MAX({CAPEX}-{fin['grant_amount']},{CAPEX}*{fin['grant_cofinance_pct']})",
        "loan": f"=-{CAPEX}*(1-{fin['loan_share']})",
    }
    y0 = rr
    ws.cell(row=rr, column=1, value=0)
    for i, k in enumerate(keys):
        c = ws.cell(row=rr, column=2 + i, value=own0[k]); c.number_format = RUB
    rr += 1
    for t in range(1, n + 1):
        ws.cell(row=rr, column=1, value=t)
        ebt = f"({EBITDA}*(1+{infl})^{t-1})"
        net = f"={ebt}*(1-{tax})"
        ws.cell(row=rr, column=2, value=net).number_format = RUB
        ws.cell(row=rr, column=3, value=net).number_format = RUB
        loan_net = f"={ebt}*(1-{tax})-IF({t}<={fin['loan_term']},${pmt_cell},0)"
        ws.cell(row=rr, column=4, value=loan_net).number_format = RUB
        rr += 1
    ylast = rr - 1
    rr += 1
    ws.cell(row=rr, column=1, value="NPV").font = BOLD
    for i in range(len(keys)):
        col = get_column_letter(2 + i)
        f = f"={col}{y0}+NPV(${disc_cell},{col}{y0+1}:{col}{ylast})"
        c = ws.cell(row=rr, column=2 + i, value=f); c.number_format = RUB; c.font = BOLD
    rr += 1
    ws.cell(row=rr, column=1, value="IRR").font = BOLD
    for i in range(len(keys)):
        col = get_column_letter(2 + i)
        c = ws.cell(row=rr, column=2 + i, value=f"=IFERROR(IRR({col}{y0}:{col}{ylast}),\"—\")")
        c.number_format = PCT; c.font = BOLD
    rr += 2

    # --- Прогнозы (3 сценария), вычисленные значения ---
    ws.cell(row=rr, column=1, value="ПРОГНОЗЫ (3 сценария)").font = BOLD
    for c in range(1, 5): ws.cell(row=rr, column=c).fill = GROUP_FILL
    rr += 1
    fnames = list(prof["forecasts"].keys())
    ws.cell(row=rr, column=1, value="Показатель").font = BOLD
    for i, fn in enumerate(fnames):
        ws.cell(row=rr, column=2 + i, value=fn).font = BOLD
    head(ws, rr, 1 + len(fnames))
    rr += 1
    comp = {fn: E.compute(prof, fn) for fn in fnames}
    def line(label, getter, fmt=RUB, fill_npv=False):
        nonlocal rr
        ws.cell(row=rr, column=1, value=label)
        for i, fn in enumerate(fnames):
            v = getter(comp[fn])
            c = ws.cell(row=rr, column=2 + i, value=v)
            if isinstance(v, (int, float)):
                c.number_format = fmt
                if fill_npv:
                    c.fill = GOOD if v >= 0 else BAD
        rr += 1
    line("CAPEX, ₽", lambda d: round(d["capex"]))
    line("Выручка, ₽/год", lambda d: round(d["revenue"]))
    line("OPEX, ₽/год", lambda d: round(d["opex"]))
    line("EBITDA, ₽/год", lambda d: round(d["ebitda"]))
    for k, kname in E.FIN_SCENARIOS.items():
        ws.cell(row=rr, column=1, value=kname).font = BOLD
        for c in range(1, 2 + len(fnames)): ws.cell(row=rr, column=c).fill = GROUP_FILL
        rr += 1
        line("   Окупаемость, лет",
             lambda d, k=k: d["fin"][k]["payback"] if d["fin"][k]["payback"] else "—", "0.0")
        line("   NPV, ₽", lambda d, k=k: round(d["fin"][k]["npv"]), RUB, fill_npv=True)
        line("   IRR", lambda d, k=k: d["fin"][k]["irr"] if d["fin"][k]["irr"] else "—", PCT)
    rr += 1

    # --- Заметки ---
    ws.cell(row=rr, column=1, value="Заметки").font = BOLD
    rr += 1
    for note in prof.get("notes", []):
        ws.cell(row=rr, column=1, value="• " + note)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
        rr += 1


def comparison_sheet(wb):
    ws = wb.create_sheet("Сравнение", 0)
    title(ws, "СРАВНЕНИЕ ПРОФИЛЕЙ (база; схема «Грант КФХ»)", 7)
    ws.cell(row=2, column=1, value="Профиль").font = BOLD
    headers = ["CAPEX", "Выручка/год", "OPEX/год", "EBITDA/год", "Окуп., лет", "NPV (грант)", "IRR (грант)"]
    for i, h in enumerate(headers):
        ws.cell(row=2, column=2 + i, value=h)
    head(ws, 2, 1 + len(headers))
    rr = 3
    for key, prof in PR.PROFILES.items():
        d = E.compute(prof, "Базовый")
        g = d["fin"]["grant"]
        ws.cell(row=rr, column=1, value=prof["name"])
        vals = [round(d["capex"]), round(d["revenue"]), round(d["opex"]), round(d["ebitda"]),
                g["payback"] if g["payback"] else "—", round(g["npv"]),
                g["irr"] if g["irr"] else "—"]
        for i, v in enumerate(vals):
            c = ws.cell(row=rr, column=2 + i, value=v)
            if i in (0, 1, 2, 3, 5) and isinstance(v, (int, float)):
                c.number_format = RUB
            if i == 4 and isinstance(v, (int, float)):
                c.number_format = "0.0"
            if i == 6 and isinstance(v, (int, float)):
                c.number_format = PCT
            if i == 5 and isinstance(v, (int, float)):
                c.fill = GOOD if v >= 0 else BAD
        rr += 1
    rr += 1
    ws.cell(row=rr, column=1,
            value="Внимание: во все профили включён жилой дом (~1.8–5.5 млн ₽) — это личный актив, "
                  "он раздувает CAPEX и ухудшает NPV «бизнеса». Для чистой бизнес-оценки вынесите дом.")
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=7)
    rr += 2
    ws.cell(row=rr, column=1, value="Полный прогноз (пессим/база/оптим) и денежные потоки — на листах профилей.")
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=7)
    ws.column_dimensions["A"].width = 42
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 15


def land_sheet(wb):
    ws = wb.create_sheet("Земля Томск")
    title(ws, "УЧАСТКИ ВОКРУГ ТОМСКА (≈50 КМ) — НАПРАВЛЕНИЯ", 5)
    cols = ["Локация", "Удалённость", "Ориентир цены", "Чем хорошо", "Подходит для"]
    for i, h in enumerate(cols):
        ws.cell(row=2, column=1 + i, value=h)
    head(ws, 2, len(cols))
    rows = [
        ("Корнилово / Лоскутово", "10–20 км В", "ИЖС 3–3.7 млн/уч.; дорого",
         "близко к городу, газ, инфраструктура, сбыт", "теплицы (сбыт), дом"),
        ("Богашёво / Петухово", "15–20 км Ю-В", "ИЖС 1.5–3 млн",
         "аэропорт, газ, асфальт", "теплицы, ЛПХ"),
        ("Зоркальцево / Кафтанчиково", "12–18 км З", "ИЖС/ЛПХ 1–2.5 млн",
         "трасса, газ, рядом р. Томь", "теплицы, смешанная ферма"),
        ("Моряковский Затон", "35–45 км С-З", "уч. от ~0.55 млн; дёшево",
         "дёшево, заливные луга (корма), р. Обь", "джерси/КРС, поля"),
        ("Турунтаево / Кузовлево", "25–35 км С", "сельхоз/ЛПХ недорого",
         "много земли, тихо", "КРС, кормовые поля"),
        ("Мельниково (Шегарский р-н)", "~50 км З", "сельхоз дёшево, паи",
         "большие массивы, агро-район", "КРС, поля, крупное КФХ"),
        ("Межи / участки КФХ 25+ га", "30–50 км", "сельхоз, оптом за га",
         "масштаб под КФХ/грант", "молочная ферма, кормопроизводство"),
    ]
    rr = 3
    for r in rows:
        for i, v in enumerate(r):
            ws.cell(row=rr, column=1 + i, value=v)
        rr += 1
    rr += 1
    notes = [
        "Под ДОМ нужна категория ИЖС/ЛПХ (в населённом пункте); под ПОЛЯ/выпас — сельхозназначения.",
        "Для ТЕПЛИЦ круглый год критичен ГАЗ на участке (отопление) — выбирайте газифицированные сёла (Корнилово, Богашёво, Зоркальцево).",
        "Для ДЖЕРСИ/КРС важны луга и площадь под корма — дешевле и просторнее севернее/западнее (Моряковка, Турунтаево, Мельниково).",
        "Грант 'Агростартап' требует регистрации КФХ в сельской местности — учитывайте при выборе.",
        "Конкретные лоты смотрите на cian, sibdom, move.ru, neagent, olan, mirkvartir (ссылки на листе «Источники»).",
    ]
    for nt in notes:
        ws.cell(row=rr, column=1, value="• " + nt)
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
        rr += 1
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 30


def sources_sheet(wb):
    ws = wb.create_sheet("Источники")
    title(ws, "ИСТОЧНИКИ ДАННЫХ (Томск, 2026)", 2)
    ws.cell(row=2, column=1, value="Тема").font = BOLD
    ws.cell(row=2, column=2, value="Источник").font = BOLD
    head(ws, 2, 2)
    src = [
        ("Джерси: надой ~4.5–5 тыс л, жирность 5.6–8%", "wikipedia / direct.farm / salkovo.ru"),
        ("Тепличный бизнес: рентаб. 15–25%, урожай огурца 20–30 кг/м²", "polygalvostok.ru / openbusiness.ru / svoefermerstvo.ru"),
        ("Зимняя теплица: цена зимой ×2–3, отопление — главная статья", "technolakpiter.ru / polygalvostok.ru"),
        ("Участки ИЖС/ЛПХ Томский р-н", "cian / sibdom / m2.ru / olan.ru / mirkvartir.ru"),
        ("Сельхозземля, участки КФХ (25+ га)", "move.ru / neagent.info / cian"),
        ("Закупка молока (выс. сорт ~40 ₽/л)", "vtomske.ru / souzmoloko.ru"),
        ("Грант 'Агростартап' Томск (3–6 млн; до 7 на КРС)", "svetich.info / svoefermerstvo.ru"),
        ("Льготный кредит АПК (ставка до 5%, до 7 лет)", "rshb.ru"),
    ]
    rr = 3
    for a, b in src:
        ws.cell(row=rr, column=1, value=a)
        ws.cell(row=rr, column=2, value=b)
        rr += 1
    ws.cell(row=rr + 1, column=1, value="Цены ориентировочные (диапазоны рынка). Проверяйте по конкретному лоту.")
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 46


def build():
    wb = Workbook()
    wb.remove(wb.active)
    for key, prof in PR.PROFILES.items():
        profile_sheet(wb, key, prof)
    comparison_sheet(wb)
    land_sheet(wb)
    sources_sheet(wb)
    wb.save(OUT)
    print("[OK] Excel сохранён:", OUT)


if __name__ == "__main__":
    build()
