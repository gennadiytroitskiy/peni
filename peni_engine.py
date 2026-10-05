#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Движок расчёта пеней по ЖКУ — v3 (проверяемый).

Возможности:
- FIFO-гашение платежей: платёж закрывает самый ранний непогашенный период.
- Полная разбивка по отрезкам: база, доля ставки (1/300 или 1/130), ставка, дни.
- Три метода ставки (параметр rate_mode):
    'historical'       — ставка Банка России по каждому дню; с 27.02.2022 применяется
                         минимум из неё и 9,5 % (постановления Правительства № 474 и № 329).
    'fixed95'          — всегда 9,5 % на всю глубину долга.
    'on_payment_date'  — одна ставка на весь долг: ставка, действующая на дату оплаты
                         (для непогашенного долга — на дату расчёта), с ограничением 9,5 %
                         для дат с 27.02.2022.
- Срок оплаты: до 28.02.2026 — 10-е число месяца, следующего за истекшим;
  с 01.03.2026 — 15-е число (ФЗ от 24.06.2025 № 177-ФЗ).

ОГОВОРКА. Расчёт носит справочный характер. Применение ставки к задолженности,
возникшей до 27.02.2022, спорно: часть практики применяет ставку на день фактической
оплаты, часть — ставку, действовавшую в соответствующий период. Выбор метода фиксируйте
внутренним регламентом организации и согласуйте с юристом до использования в суде.
"""

from datetime import date, timedelta
from bisect import bisect_right

# ==================== КЛЮЧЕВАЯ СТАВКА БАНКА РОССИИ ====================
# Проверяйте по официальному архиву Банка России перед юридическим использованием.

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
DEADLINE_CHANGE_EFFECTIVE_DATE = date(2026, 3, 1)

RATE_MODES = ("historical", "fixed95", "on_payment_date")
RATE_MODE_LABELS = {
    "historical": "Историческая ставка Банка России по дням (ограничение 9,5 % с 27.02.2022)",
    "fixed95": "Фиксированная ставка 9,5 % на весь период",
    "on_payment_date": "Ставка на день фактической оплаты (для непогашенного долга — на дату расчёта)",
}


def rate_on_date(d):
    """Ставка Банка России на дату с учётом ограничения 9,5 % с 27.02.2022."""
    idx = bisect_right(RATE_DATES, d) - 1
    rate = RATE_VALUES[idx] if idx >= 0 else 7.5
    if d >= MORATORIUM_START:
        return min(MORATORIUM_CAP, rate)
    return rate


def due_date_for_period(period_start):
    """Срок оплаты по ч. 1 ст. 155 ЖК РФ с учётом ФЗ № 177-ФЗ."""
    y, m = period_start.year, period_start.month
    ny, nm = (y, m + 1) if m < 12 else (y + 1, 1)
    due_month_start = date(ny, nm, 1)
    day = 15 if due_month_start >= DEADLINE_CHANGE_EFFECTIVE_DATE else 10
    return date(ny, nm, day)


def _make_rate_fn(rate_mode, single_rate=None):
    if rate_mode == "fixed95":
        return lambda d: MORATORIUM_CAP
    if rate_mode == "on_payment_date":
        r = MORATORIUM_CAP if single_rate is None else single_rate
        return lambda d: r
    return rate_on_date


def _split_segment(base, due, seg_start, seg_end, kapremont, rate_fn):
    """Делит отрезок на подинтервалы по смене ставки и по границе 1/300 → 1/130."""
    if seg_end <= seg_start or base <= 0:
        return []
    peni_start = due + timedelta(days=31)          # первые 30 дней без пеней
    start = max(seg_start, peni_start)
    if start >= seg_end:
        return []
    points = [start, seg_end]
    points.extend(d for d in RATE_DATES if start < d < seg_end)
    day_91 = due + timedelta(days=91)
    if not kapremont and start < day_91 < seg_end:
        points.append(day_91)
    points = sorted(set(points))

    out = []
    for i in range(len(points) - 1):
        a, b = points[i], points[i + 1]
        days = (b - a).days
        if days <= 0:
            continue
        days_from_due = (a - due).days
        divid = 300 if (kapremont or days_from_due <= 90) else 130
        rate = rate_fn(a)
        value = base * rate / 100 / divid * days
        out.append({
            "start": a, "end": b, "days": days, "base": round(base, 2),
            "dividend": divid, "rate": rate, "peni": value,
        })
    return out


def calculate_fifo(charges, payments, calc_date, kapremont=False, rate_mode="historical"):
    """charges: [{period, amount, deadline}]  payments: [{date, amount}]"""
    if rate_mode not in RATE_MODES:
        rate_mode = "historical"

    debts = [
        {"period": c["period"], "deadline": c["deadline"], "original": float(c["amount"]),
         "remaining": float(c["amount"]), "history": []}
        for c in sorted(charges, key=lambda x: x["deadline"])
    ]
    # Платежи после даты расчёта в расчёт не попадают.
    pays = sorted((p for p in payments if p["date"] <= calc_date), key=lambda x: x["date"])

    # Распределение FIFO. Аванс — платёж, поступивший раньше срока оплаты периода,
    # который он закрывает: тогда пеня на этот период вообще не должна возникнуть.
    movements = []
    for pay in pays:
        amount = float(pay["amount"])
        left = amount
        for d in debts:
            if left <= 0:
                break
            if d["remaining"] <= 0:
                continue
            paid = min(left, d["remaining"])
            is_advance = pay["date"] < d["deadline"]
            d["history"].append({
                "date": pay["date"], "paid": round(paid, 2),
                "remaining_after": round(d["remaining"] - paid, 2),
                "is_advance": is_advance,
            })
            d["remaining"] = round(d["remaining"] - paid, 2)
            left = round(left - paid, 2)
        movements.append({"date": pay["date"], "amount": amount,
                          "unallocated": round(left, 2) if left > 0.005 else 0.0})

    # Неразнесённый остаток = переплата: заплатили больше, чем начислено.
    unallocated = round(sum(m["unallocated"] for m in movements), 2)
    overpayment = round(unallocated, 2) if unallocated > 0.005 else 0.0
    advance_total = round(sum(h["paid"] for d in debts for h in d["history"] if h["is_advance"]), 2)

    rows = []
    for d in debts:
        single_rate = None
        if rate_mode == "on_payment_date":
            pay_date = d["history"][-1]["date"] if d["history"] else calc_date
            single_rate = rate_on_date(pay_date)
        rate_fn = _make_rate_fn(rate_mode, single_rate)

        segments = []
        cur_base = d["original"]
        cur_start = d["deadline"]
        for h in d["history"]:
            seg_end = h["date"] + timedelta(days=1)      # день оплаты относится к периоду до платежа
            segments += _split_segment(cur_base, d["deadline"], cur_start, seg_end, kapremont, rate_fn)
            cur_base = h["remaining_after"]
            cur_start = seg_end
        segments += _split_segment(cur_base, d["deadline"], cur_start,
                                   calc_date + timedelta(days=1), kapremont, rate_fn)

        peni = round(sum(s["peni"] for s in segments), 2)
        days_total = (calc_date - d["deadline"]).days
        days_peni = sum(s["days"] for s in segments)
        paid_total = round(sum(h["paid"] for h in d["history"]), 2)
        advance_paid = round(sum(h["paid"] for h in d["history"] if h["is_advance"]), 2)

        rates_used = sorted({s["rate"] for s in segments})
        notes = []
        if peni > 0 and d["remaining"] > 0 and peni > d["remaining"]:
            notes.append("пени превышают остаток долга")
        if not d["history"] and d["remaining"] > 0:
            notes.append("платежей не поступало")
        if kapremont:
            notes.append("капремонт: применена только 1/300")
        if advance_paid > 0:
            notes.append(f"аванс: {advance_paid:,.2f} ₽ внесено до срока оплаты — "
                         f"пеня начислена только на неоплаченную часть")
        if d["deadline"] > calc_date:
            notes.append("срок оплаты ещё не наступил")
        if d["remaining"] <= 0 and peni > 0:
            notes.append("долг погашен, пени начислены за дни до оплаты")
        if d["remaining"] <= 0 and not d["history"]:
            notes.append("остаток равен нулю без платежей — проверьте данные")

        rows.append({
            "period": d["period"], "deadline": d["deadline"],
            "original": round(d["original"], 2), "remaining": round(d["remaining"], 2),
            "paid_total": paid_total, "advance_paid": advance_paid,
            "days_total": days_total, "days_peni": days_peni,
            "peni": peni, "k_pay": round(d["remaining"] + peni, 2),
            "segments": segments,
            "payments": [{"date": h["date"], "paid": round(h["paid"], 2),
                          "remaining_after": h["remaining_after"],
                          "is_advance": h["is_advance"],
                          "kind": ("аванс" if h["is_advance"] else
                                   ("в срок оплаты" if h["date"] <= d["deadline"] else "с просрочкой"))
                          } for h in d["history"]],
            "rates_used": rates_used,
            "rate_mode": rate_mode, "kapremont": kapremont,
            "note": "; ".join(notes),
        })

    total_debt = round(sum(r["remaining"] for r in rows), 2)
    total_peni = round(sum(r["peni"] for r in rows), 2)
    check_peni = round(sum(s["peni"] for r in rows for s in r["segments"]), 2)
    return {
        "rows": rows,
        "total_debt": total_debt,
        "total_peni": total_peni,
        "total_all": round(total_debt + total_peni, 2),
        "calc_date": calc_date, "rate_mode": rate_mode, "kapremont": kapremont,
        "overpayment": overpayment, "advance_total": advance_total,
        "movements": movements,
        "control": {
            "sum_of_segments": check_peni,
            "matches": abs(check_peni - total_peni) < 0.01,
        },
    }


def check_reconciliation(charges, payments, strict=False):
    """Сверяет входные данные до расчёта.

    Различает два случая превышения платежей над начислениями:
    - авансовые платежи и переплата — нормальная ситуация, только предупреждение;
    - остаток долга в листе «Начисления» вместе с платежами — ошибка выгрузки,
      суммы вычитаются дважды.

    Отличить их по одним суммам нельзя, поэтому по умолчанию превышение даёт
    предупреждение («warning») и не останавливает расчёт. При strict=True
    проверка становится блокирующей.
    """
    billed = round(sum(float(c["amount"]) for c in charges), 2)
    paid = round(sum(float(p["amount"]) for p in payments), 2)
    over = round(paid - billed, 2)

    if over > 0.005:
        return {
            "ok": not strict,
            "severity": "error" if strict else "warning",
            "code": "overpayment", "billed": billed, "paid": paid,
            "diff": round(billed - paid, 2),
            "overpayment": over,
            "message": (
                (f"платежей больше, чем начислено, на {over:,.2f} ₽ "
                 f"(начислено {billed:,.2f}, оплачено {paid:,.2f}). "
                 f"Разница учтена как аванс/переплата и по периодам не распределена.")
                if not strict else
                (f"платежей больше, чем начислено, на {over:,.2f} ₽ "
                 f"(начислено {billed:,.2f}, оплачено {paid:,.2f}). "
                 f"Проверьте лист «Начисления»: если там указан остаток долга, "
                 f"а не полное начисление, суммы вычтутся дважды. "
                 f"Если это действительно аванс или переплата — выберите режим "
                 f"«Разрешить авансы».")),
        }
    if billed > 0 and paid == 0:
        return {"ok": True, "severity": "ok", "code": "no_payments",
                "billed": billed, "paid": paid, "diff": round(billed - paid, 2),
                "overpayment": 0.0,
                "message": "платежи не указаны — пени считаются на всю сумму начислений."}
    return {"ok": True, "severity": "ok", "code": "ok", "billed": billed, "paid": paid,
            "diff": round(billed - paid, 2), "overpayment": 0.0,
            "message": "данные согласованы."}


def batch_calculate(groups, calc_date, rate_mode="historical"):
    """groups: {account: {"charges": [...], "payments": [...], "kapremont": bool}}"""
    accounts = []
    for acc in sorted(groups):
        g = groups[acc]
        res = calculate_fifo(g["charges"], g["payments"], calc_date,
                             kapremont=g.get("kapremont", False), rate_mode=rate_mode)
        res["account"] = acc
        res["reconciliation"] = check_reconciliation(g["charges"], g["payments"],
                                                    strict=g.get("strict", False))
        accounts.append(res)
    return {
        "accounts": accounts,
        "total_debt": round(sum(a["total_debt"] for a in accounts), 2),
        "total_peni": round(sum(a["total_peni"] for a in accounts), 2),
        "total_all": round(sum(a["total_all"] for a in accounts), 2),
        "rate_mode": rate_mode, "calc_date": calc_date,
        "control_ok": all(a["control"]["matches"] for a in accounts),
        "reconciliation_ok": all(a["reconciliation"]["ok"] for a in accounts),
        "total_overpayment": round(sum(a["overpayment"] for a in accounts), 2),
        "total_advance": round(sum(a["advance_total"] for a in accounts), 2),
    }
