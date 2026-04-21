# -*- coding: utf-8 -*-
"""
Сравнение выгрузок ЖурналДокументов_БП.xlsx и ЖурналДокументов_УТ.xlsx.

Исходный лист (первая страница) не меняется.

На второй странице «Результат» выводится «чистая» таблица: заголовок с первой
строки, без пустых строк сверху, плюс колонка статуса; оформлена как умная
таблица Excel: tbl_BP (файл БП) и tbl_UT (файл УТ).

«Полное сходство» — в другой таблице есть строка с тем же ИНН, той же календарной
датой, той же суммой документа (±0.01) и тем же направлением («поступление» /
«списание»), см. classify_direction() по колонке вида документа.

«Есть расхождения» — полного ключа нет, но есть строка с тем же ИНН и той же
датой ИЛИ с тем же ИНН и той же суммой (±0.01).

«Строки нет» — нет ни полного совпадения, ни такого частичного совпадения.

Запуск из папки со скриптом:
  python compare_journals_bp_ut.py
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

BP_NAME = "ЖурналДокументов_БП.xlsx"
UT_NAME = "ЖурналДокументов_УТ.xlsx"
# Выгрузка БП после запроса с полем ВидОперации: 9 колонок до статуса включительно.
BP_SOURCE_DATA_COLS = 9
UT_SOURCE_DATA_COLS = 8
AMOUNT_TOL = 0.01
HEADER_ROW = 2
DATA_START_ROW = 3
RESULT_SHEET_NAME = "Результат"


def norm_inn(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\D", "", str(value).strip())


def norm_date(value):
    dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.normalize()


def norm_amount(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def classify_direction(kind: str) -> str:
    """«поступление» | «списание» | «» по тексту вида документа."""
    if kind is None or (isinstance(kind, float) and pd.isna(kind)):
        return ""
    s = str(kind).strip()
    if not s:
        return ""
    low = s.lower()
    compact = "".join(low.split())

    def has_sub(*parts: str) -> bool:
        return any(p in low or p in compact for p in parts)

    # Сначала узкие исключения (корректировки)
    if has_sub("корректировкаприобретения", "корректировка приобретения"):
        return "поступление"
    if has_sub("корректировкареализации", "корректировка реализации"):
        return "списание"

    if has_sub("реализация", "отгрузк", "списание", "выдача"):
        return "списание"
    if has_sub("возврат", "поставщик"):
        return "списание"
    if has_sub("отчеткомитенту", "отчёткомитенту", "комитенту"):
        return "списание"

    if has_sub("приобретение", "поступление"):
        return "поступление"
    if has_sub("отклиента", "от клиента", "покупател", "возвраттоваровот"):
        return "поступление"
    if has_sub("отчеткомиссионера", "отчёткомиссионера", "комиссионера"):
        return "поступление"

    return ""


def load_table(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, header=None)


def extract_dataframe(raw: pd.DataFrame, n_source_cols: int) -> pd.DataFrame:
    """
    n_source_cols=8 (УТ): дата, контрагент, ИНН, договор, вид документа, номер, сумма, статус.
    n_source_cols=9 (БП): то же + колонка «Вид операции» (ВидОперации) между видом документа и номером.
    """
    body = raw.iloc[DATA_START_ROW:, :n_source_cols].copy()
    body.columns = list(range(n_source_cols))
    if n_source_cols == 8:
        rename = {
            0: "date",
            1: "client",
            2: "inn",
            3: "contract",
            4: "kind",
            5: "num",
            6: "amount",
            7: "status",
        }
    elif n_source_cols == 9:
        rename = {
            0: "date",
            1: "client",
            2: "inn",
            3: "contract",
            4: "kind",
            5: "op_kind",
            6: "num",
            7: "amount",
            8: "status",
        }
    else:
        raise ValueError("n_source_cols must be 8 or 9")
    return body.rename(columns=rename).reset_index(drop=True)


def build_indexes(body: pd.DataFrame):
    full_keys: set[tuple] = set()
    by_inn_date: set[tuple] = set()
    by_inn_amt: set[tuple] = set()

    for _, r in body.iterrows():
        inn = norm_inn(r.get("inn"))
        d = norm_date(r.get("date"))
        a = norm_amount(r.get("amount"))
        dr = classify_direction(r.get("kind"))
        if inn == "" or d is None or a is None or not dr:
            continue
        full_keys.add((inn, d, a, dr))
        by_inn_date.add((inn, d))
        by_inn_amt.add((inn, a))

    return full_keys, by_inn_date, by_inn_amt


def status_bp_vs_ut(r, ut_full, ut_id, ut_ia) -> str:
    inn = norm_inn(r.get("inn"))
    d = norm_date(r.get("date"))
    a = norm_amount(r.get("amount"))
    dr = classify_direction(r.get("kind"))

    if inn == "" or d is None or a is None or not dr:
        return "нет данных для сравнения (ИНН/дата/сумма/направление)"

    if (inn, d, a, dr) in ut_full:
        return "полное сходство"

    if (inn, d) in ut_id or (inn, a) in ut_ia:
        return "есть расхождения"
    return "строки нет"


def status_ut_vs_bp(r, bp_full, bp_id, bp_ia) -> str:
    inn = norm_inn(r.get("inn"))
    d = norm_date(r.get("date"))
    a = norm_amount(r.get("amount"))
    dr = classify_direction(r.get("kind"))

    if inn == "" or d is None or a is None or not dr:
        return "нет данных для сравнения (ИНН/дата/сумма/направление)"

    if (inn, d, a, dr) in bp_full:
        return "полное сходство"

    if (inn, d) in bp_id or (inn, a) in bp_ia:
        return "есть расхождения"
    return "строки нет"


def _excel_cell(value):
    """Значение для ячейки openpyxl (без NaN/NaT в «сыром» виде)."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def write_result_smart_table(
    workbook_path: Path,
    raw: pd.DataFrame,
    statuses: list[str],
    status_header: str,
    table_display_name: str,
    n_source_cols: int,
) -> None:
    """
    Добавляет/обновляет лист RESULT_SHEET_NAME (вторая позиция в книге):
    строка 1 — заголовки из исходной строки HEADER_ROW + status_header;
    далее данные из колонок 0..n_source_cols-1 исходного листа + статус;
    оформление — таблица Excel с именем table_display_name (tbl_BP / tbl_UT).
    """
    wb = load_workbook(workbook_path, read_only=False, data_only=False)
    if RESULT_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[RESULT_SHEET_NAME])
    ws = wb.create_sheet(RESULT_SHEET_NAME, 1)

    n = len(statuses)
    header: list = []
    for j in range(n_source_cols):
        h = raw.iat[HEADER_ROW, j]
        if isinstance(h, float) and pd.isna(h):
            h = ""
        header.append(h)
    header.append(status_header)

    ws.append([_excel_cell(x) for x in header])
    for i in range(n):
        row = [_excel_cell(raw.iat[DATA_START_ROW + i, j]) for j in range(n_source_cols)]
        row.append(statuses[i])
        ws.append(row)

    ncols = len(header)
    last_row = 1 + n
    last_col = get_column_letter(ncols)
    ref = f"A1:{last_col}{last_row}"
    tab = Table(displayName=table_display_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)
    wb.save(workbook_path)


def main() -> int:
    base = Path(__file__).resolve().parent
    bp_path = base / BP_NAME
    ut_path = base / UT_NAME
    if not bp_path.is_file() or not ut_path.is_file():
        print("Нужны файлы:", BP_NAME, "и", UT_NAME, "в", base, file=sys.stderr)
        return 1

    raw_bp = load_table(bp_path)
    raw_ut = load_table(ut_path)
    if raw_bp.shape[1] < 8 or raw_ut.shape[1] < 8:
        print("Error: each xlsx needs at least 8 data columns.", file=sys.stderr)
        return 1
    bp_cols = min(BP_SOURCE_DATA_COLS, raw_bp.shape[1])
    ut_cols = min(UT_SOURCE_DATA_COLS, raw_ut.shape[1])
    if bp_cols < BP_SOURCE_DATA_COLS:
        print(
            "Warning: BP xlsx has",
            bp_cols,
            "columns (expected",
            str(BP_SOURCE_DATA_COLS) + " after query update); using",
            bp_cols,
            "for compare.",
            file=sys.stderr,
        )
    bp_body = extract_dataframe(raw_bp, bp_cols)
    ut_body = extract_dataframe(raw_ut, ut_cols)

    ut_full, ut_id, ut_ia = build_indexes(ut_body)
    bp_full, bp_id, bp_ia = build_indexes(bp_body)

    bp_statuses = [status_bp_vs_ut(bp_body.iloc[i], ut_full, ut_id, ut_ia) for i in range(len(bp_body))]
    ut_statuses = [status_ut_vs_bp(ut_body.iloc[i], bp_full, bp_id, bp_ia) for i in range(len(ut_body))]

    write_result_smart_table(
        bp_path, raw_bp, bp_statuses, "Статус в УТ", "tbl_BP", bp_cols
    )
    write_result_smart_table(
        ut_path, raw_ut, ut_statuses, "Статус в БП", "tbl_UT", ut_cols
    )

    print("OK.")
    print(" ", bp_path.name, "- sheet", RESULT_SHEET_NAME, "table tbl_BP")
    print(" ", ut_path.name, "- sheet", RESULT_SHEET_NAME, "table tbl_UT")
    print("\nSummary BP:")
    for k, v in Counter(bp_statuses).most_common():
        print(" ", v, repr(k))
    print("Summary UT:")
    for k, v in Counter(ut_statuses).most_common():
        print(" ", v, repr(k))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
