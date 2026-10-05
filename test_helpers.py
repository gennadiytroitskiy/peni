#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Общие помощники для тестов: чтение отчёта по названиям колонок."""


def rows_of(wb, sheet, key_fragment="Лицевой счёт"):
    """Находит строку заголовка по ключевой колонке и возвращает список словарей
    «заголовок -> значение». Устойчиво к сдвигу шапки и порядку столбцов."""
    ws = wb[sheet]
    data = [r for r in ws.iter_rows(values_only=True)]
    header_idx = None
    for i, row in enumerate(data):
        if row and any(v and key_fragment.lower() in str(v).lower() for v in row):
            header_idx = i
            break
    if header_idx is None:
        return []
    header = data[header_idx]
    out = []
    for r in data[header_idx + 1:]:
        if not r or all(v is None for v in r):
            continue
        out.append({h: v for h, v in zip(header, r) if h})
    return out


def find(row, *fragments):
    """Значение по фрагментам названия колонки (все фрагменты должны встречаться)."""
    for key, value in row.items():
        low = str(key).lower()
        if all(f.lower() in low for f in fragments):
            return value
    return None


def total_row(wb, sheet, prefix="ИТОГО"):
    rows = rows_of(wb, sheet)
    for row in rows:
        first = row.get("Лицевой счёт")
        if first and str(first).upper().startswith(prefix.upper()):
            return row
    return None
