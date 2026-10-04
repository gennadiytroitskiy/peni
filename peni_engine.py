#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Движок расчёта пеней по ЖКУ.
- FIFO-гашение платежей (платёж гасит самый старый непогашенный период)
- Два режима ставки: 'fixed95' (всегда 9.5%) и 'historical' (ставка ЦБ по датам + cap с 27.02.2022)
- Автоматический срок оплаты: до 01.02.2026 — 10-е число; с 01.02.2026 — 15-е число следующего месяца
"""

from datetime import date, timedelta
from bisect import bisect_right

KEY_RATE_HISTORY = [
    (date(2019, 6, 17), 7.50), (date(2019, 7, 29), 7.25), (date(2019, 9, 9), 7.00),
    (date(2019, 10, 28), 6.50), (date(2019, 12, 16), 6.25), (date(2020, 2, 10), 6.00),
    (date(2020, 4, 27), 5.50), (date(2020, 6, 22), 4.50), (date(2020, 7, 27), 4.25),
    (date(2021, 3, 22), 4.50), (date(2021, 4, 26), 5.00), (date(2021, 6, 14), 5.50),
    (date(2021, 7, 26), 6.50), (date(2021, 9, 13), 6.75), (date(2021, 10, 25), 7.50),
    (date(2021, 12, 20), 8.50), (date(2022, 2, 14), 9.50), (date(2022, 2, 28), 20.00),
    (date(2022, 4, 11), 17.00), (date(2022, 5, 16), 14.00), (date(2022, 6, 14), 11.00),
    (date(2022, 7, 25), 9.50), (date(2022, 8, 1), 8.00), (date(2022, 9, 19), 7.50),
    (date(2023, 2, 13), 7.50), (date(2023, 7, 24), 8.50), (date(2023, 8, 15), 12.00),
    (date(2023, 9, 18), 13.00), (date(2023, 10, 30), 15.00), (date(2023, 12, 18), 16.00),
    (date(2024, 7, 29), 18.00), (date(2024, 9, 16), 19.00), (date(2024, 10, 28), 21.00),
    (date(2025, 2, 17), 20.00), (date(2025, 3, 24), 19.00), (date(2025, 4, 28), 18.00),
    (date(2025, 6, 23), 17.00), (date(2025, 7, 28), 16.00), (date(2025, 9, 15), 15.50),
    (date(2025, 10, 27), 15.00), (date(2025, 12, 22), 16.00), (date(2026, 2, 16), 15.50),
    (date(2026, 3, 23), 15.00), (date(2026, 4, 27), 14.50), (date(2026, 6, 22), 14.25),
    (date(2026, 7, 27), 14.00),
]
RATE_DATES = [r[0] for r in KEY_RATE_HISTORY]
RATE_VALUES = [r[1] for r in KEY_RATE_HISTORY]

MORATORIUM_START = date(2022, 2, 27)
MORATORIUM_CAP = 9.5

# ФЗ от 24.06.2025 № 177-ФЗ вступил в силу 01.03.2026.
# Поэтому для срока оплаты, приходящегося на 01.03.2026 или позднее,
# устанавливаем 15-е число; до этого сохраняется 10-е число.
DEADLINE_CHANGE_EFFECTIVE_DATE = date(2026, 3, 1)


def get_rate_on_date(d, mode="historical"):
    if mode == "fixed95":
        return MORATORIUM_CAP
    idx = bisect_right(RATE_DATES, d) - 1
    rate = RATE_VALUES[idx] if idx >= 0 else 7.5
    if d >= MORATORIUM_START:
        return min(MORATORIUM_CAP, rate)
    return rate


def due_date_for_period(period_start):
    """Срок оплаты по ч. 1 ст. 155 ЖК РФ.

    ФЗ от 24.06.2025 № 177-ФЗ вступил в силу 01.03.2026 и изменил срок
    оплаты на 15-е число месяца, следующего за истекшим. Изменение относится
    к платежам, срок которых наступает с 01.03.2026; для более ранних сроков
    применяется прежнее правило (до 10-го числа).
    """
    y, m = period_start.year, period_start.month
    ny, nm = (y, m + 1) if m < 12 else (y + 1, 1)
    due_month_start = date(ny, nm, 1)
    day = 15 if due_month_start >= DEADLINE_CHANGE_EFFECTIVE_DATE else 10
    return date(ny, nm, day)


def peni_for_segment(amount, due_date, seg_start, seg_end, kapremont=False, rate_mode="historical"):
    if seg_end <= seg_start or amount <= 0:
        return 0.0
    peni_start = due_date + timedelta(days=31)
    start = max(seg_start, peni_start)
    if start >= seg_end:
        return 0.0
    breakpoints = [start, seg_end]
    breakpoints.extend(d for d in RATE_DATES if start < d < seg_end)
    # День 91 просрочки переключает 1/300 на 1/130; делим интервал и здесь.
    day_91 = due_date + timedelta(days=91)
    if start < day_91 < seg_end and not kapremont:
        breakpoints.append(day_91)
    breakpoints = sorted(set(breakpoints))
    total = 0.0
    for i in range(len(breakpoints) - 1):
        p_start, p_end = breakpoints[i], breakpoints[i + 1]
        days = (p_end - p_start).days
        if days <= 0:
            continue
        rate = get_rate_on_date(p_start, rate_mode)
        days_from_due = (p_start - due_date).days
        mult = 300 if (kapremont or days_from_due <= 90) else 130
        total += amount * rate / 100 / mult * days
    return total


def calculate_fifo(charges, payments, calc_date, kapremont=False, rate_mode="historical"):
    """
    charges: [{period: date, amount: float, deadline: date}]
    payments: [{date: date, amount: float}]
    """
    debts = [
        {"period": c["period"], "deadline": c["deadline"], "original": c["amount"],
         "remaining": c["amount"], "history": []}
        for c in sorted(charges, key=lambda x: x["deadline"])
    ]
    # Платежи, совершённые после даты расчёта, не влияют на расчёт на эту дату.
    pays = sorted((p for p in payments if p["date"] <= calc_date), key=lambda x: x["date"])

    for pay in pays:
        amt = pay["amount"]
        for d in debts:
            if amt <= 0:
                break
            if d["remaining"] <= 0:
                continue
            pay_amt = min(amt, d["remaining"])
            d["history"].append({"date": pay["date"], "paid": pay_amt,
                                  "remaining_after": d["remaining"] - pay_amt})
            d["remaining"] -= pay_amt
            amt -= pay_amt

    results = []
    for d in debts:
        total_peni = 0.0
        cur_amount = d["original"]
        cur_start = d["deadline"]
        for h in d["history"]:
            # День оплаты включается в период начисления на сумму до платежа.
            payment_day_after = h["date"] + timedelta(days=1)
            total_peni += peni_for_segment(cur_amount, d["deadline"], cur_start, payment_day_after,
                                            kapremont, rate_mode)
            cur_amount = h["remaining_after"]
            cur_start = payment_day_after
        # Дата расчёта также включается, если долг на эту дату не погашен.
        total_peni += peni_for_segment(cur_amount, d["deadline"], cur_start, calc_date + timedelta(days=1),
                                        kapremont, rate_mode)
        results.append({
            "period": d["period"], "deadline": d["deadline"],
            "original": d["original"], "remaining": round(d["remaining"], 2),
            "peni": round(total_peni, 2),
        })

    total_debt = sum(r["remaining"] for r in results)
    total_peni = sum(r["peni"] for r in results)
    return {
        "rows": results,
        "total_debt": round(total_debt, 2),
        "total_peni": round(total_peni, 2),
        "total_all": round(total_debt + total_peni, 2),
        "calc_date": calc_date,
        "rate_mode": rate_mode,
    }
