from __future__ import annotations

import argparse
import math
from pathlib import Path
from tkinter import Tk, filedialog

import pandas as pd


CANONICAL_COLUMNS = [
    "Период",
    "Склад",
    "КодНоменклатуры",
    "Номенклатура",
    "Регистратор",
    "Операция",
    "Количество",
]

EPS = 1e-9
ROUND_STEP_TON_10KG = 0.01


def _ceil_abs_to_10kg_tons(value: float) -> float:
    if pd.isna(value):
        return value
    sign = -1.0 if float(value) < 0 else 1.0
    rounded_abs = math.ceil(abs(float(value)) / ROUND_STEP_TON_10KG) * ROUND_STEP_TON_10KG
    return sign * rounded_abs


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    if len(df.columns) < 7:
        raise ValueError(f"Ожидалось минимум 7 колонок, найдено: {len(df.columns)}")
    df = df.iloc[:, :7].copy()
    df.columns = CANONICAL_COLUMNS
    return df


def _signed_quantity(df: pd.DataFrame) -> pd.Series:
    qty = pd.to_numeric(df["Количество"], errors="coerce").fillna(0.0).round(3)
    op = df["Операция"].astype(str).str.lower().str.strip()

    signed = qty.copy()
    is_expense = op.str.contains("расход", na=False)
    is_income = op.str.contains("приход", na=False)
    signed.loc[is_expense] = -qty.loc[is_expense].abs()
    signed.loc[is_income] = qty.loc[is_income].abs()
    return signed


def _prepare_daily(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = _normalize_columns(df_raw)
    df["Период"] = pd.to_datetime(df["Период"], errors="coerce", dayfirst=True)
    df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce").fillna(0.0).round(3)
    df["Изменение"] = _signed_quantity(df)
    df["Дата"] = df["Период"].dt.normalize()

    group_keys = ["Дата", "Склад", "КодНоменклатуры", "Номенклатура"]
    daily = (
        df.groupby(group_keys, dropna=False, as_index=False)
        .agg(ИзменениеДня=("Изменение", "sum"), КоличествоДвижений=("Изменение", "size"))
    )
    daily["ИзменениеДня"] = daily["ИзменениеДня"].round(3)
    return daily


def find_negative_moments_daily(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    daily = _prepare_daily(df_raw)
    daily = daily.sort_values(["Склад", "КодНоменклатуры", "Дата"], kind="mergesort").reset_index(drop=True)

    keys = ["Склад", "КодНоменклатуры"]
    daily["ОстатокПослеДня"] = daily.groupby(keys, dropna=False)["ИзменениеДня"].cumsum().round(3)
    daily["ОстатокПослеДняОкругл10кгВверх"] = daily["ОстатокПослеДня"].apply(_ceil_abs_to_10kg_tons)

    negatives = daily[daily["ОстатокПослеДня"] < 0].copy()
    first_negative = negatives.groupby(keys, dropna=False, as_index=False).first()
    return negatives, first_negative


def build_yearly_adjustment_plan_daily(
    df_raw: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[int, pd.DataFrame]]:
    daily = _prepare_daily(df_raw)
    daily = daily.sort_values(["Склад", "КодНоменклатуры", "Дата"], kind="mergesort").reset_index(drop=True)
    daily["Год"] = daily["Дата"].dt.year

    rows: list[dict] = []
    years_set: set[int] = set()

    keys = ["Склад", "КодНоменклатуры"]
    for (warehouse, item_code), grp in daily.groupby(keys, dropna=False, sort=False):
        grp = grp.copy()
        grp["Год"] = grp["Год"].astype("Int64")
        grp = grp[grp["Год"].notna()]
        if grp.empty:
            continue

        item_name = grp["Номенклатура"].dropna().iloc[0] if grp["Номенклатура"].notna().any() else ""
        reserve_start = 0.0
        actual_cum_start = 0.0

        year_min = int(grp["Год"].min())
        year_max = int(grp["Год"].max())

        for year in range(year_min, year_max + 1):
            year_rows = grp[grp["Год"] == year]
            if year_rows.empty:
                year_total = 0.0
                min_prefix = 0.0
            else:
                changes = year_rows["ИзменениеДня"].astype(float)
                year_total = float(changes.sum())
                min_prefix = float(min(0.0, changes.cumsum().min()))

            needed_receipt = max(0.0, -(actual_cum_start + reserve_start + min_prefix))
            reserve_after_receipt = reserve_start + needed_receipt

            actual_cum_end = actual_cum_start + year_total
            max_safe_writeoff = actual_cum_end + reserve_after_receipt
            year_writeoff = max(0.0, min(reserve_after_receipt, max_safe_writeoff))
            reserve_end = reserve_after_receipt - year_writeoff

            needed_receipt = 0.0 if abs(needed_receipt) < EPS else round(needed_receipt, 3)
            year_writeoff = 0.0 if abs(year_writeoff) < EPS else round(year_writeoff, 3)
            reserve_start = 0.0 if abs(reserve_start) < EPS else round(reserve_start, 3)
            reserve_end = 0.0 if abs(reserve_end) < EPS else round(reserve_end, 3)
            actual_cum_start = 0.0 if abs(actual_cum_start) < EPS else round(actual_cum_start, 3)
            actual_cum_end = 0.0 if abs(actual_cum_end) < EPS else round(actual_cum_end, 3)

            years_set.add(year)
            rows.append(
                {
                    "Год": year,
                    "Склад": warehouse,
                    "КодНоменклатуры": item_code,
                    "Номенклатура": item_name,
                    "РезервНаНачалоГода": reserve_start,
                    "ОприходоватьНа01_01": needed_receipt,
                    "СписатьНа31_12": year_writeoff,
                    "РезервНаКонецГода": reserve_end,
                    "ФактБалансНаНачалоГода": actual_cum_start,
                    "ФактБалансНаКонецГода": actual_cum_end,
                    "ТипРасчета": "ГруппировкаПоДням",
                }
            )

            reserve_start = reserve_end
            actual_cum_start = actual_cum_end

    plan_df = pd.DataFrame(rows)
    if plan_df.empty:
        return plan_df, {}

    plan_df = plan_df.sort_values(["Год", "Склад", "КодНоменклатуры"], kind="mergesort").reset_index(drop=True)
    for col in ["РезервНаНачалоГода", "ОприходоватьНа01_01", "СписатьНа31_12", "РезервНаКонецГода"]:
        plan_df[f"{col}_Округл10кгВверх"] = plan_df[col].apply(_ceil_abs_to_10kg_tons)

    by_year: dict[int, pd.DataFrame] = {}
    for year in sorted(years_set):
        year_df = plan_df[plan_df["Год"] == year].copy()
        year_df = year_df[
            (year_df["ОприходоватьНа01_01"] > 0)
            | (year_df["СписатьНа31_12"] > 0)
            | (year_df["РезервНаКонецГода"] > 0)
        ]
        by_year[year] = year_df.reset_index(drop=True)

    return plan_df, by_year


def choose_input_file() -> Path:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title="Выберите файл выгрузки движений (вариант группировка по дням)",
        initialdir=str(Path.cwd()),
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()

    if not selected:
        raise SystemExit("Файл не выбран. Обработка остановлена.")
    return Path(selected)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Вариант 'группировка по дням': контроль минусов на конец дня "
            "и построение годового плана оприходования/списания."
        )
    )
    parser.add_argument(
        "--input",
        default=None,
        help="Путь к входному xlsx/csv. Если не указан, откроется выбор файла.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Путь к итоговому xlsx. По умолчанию создается рядом с входным файлом.",
    )
    args = parser.parse_args()

    input_path = Path(args.input) if args.input else choose_input_file()
    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(f"{input_path.stem}_минуса_группировка_по_дням.xlsx")
    )

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if input_path.suffix.lower() == ".csv":
        raw = pd.read_csv(input_path)
    else:
        raw = pd.read_excel(input_path, sheet_name=0)

    negatives, first_negative = find_negative_moments_daily(raw)
    yearly_plan, plan_by_year = build_yearly_adjustment_plan_daily(raw)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        negatives.to_excel(writer, sheet_name="ВсеМинусы_КонецДня", index=False)
        first_negative.to_excel(writer, sheet_name="ПервыйМинус_КонецДня", index=False)
        yearly_plan.to_excel(writer, sheet_name="ПланПоГодам_Свод_Дни", index=False)
        for year, year_df in plan_by_year.items():
            year_df.to_excel(writer, sheet_name=f"ПланДни_{year}", index=False)

    print("Тип расчета: группировка по дням")
    print(f"Входной файл: {input_path}")
    print(f"Всего строк (сырой файл): {len(raw):,}".replace(",", " "))
    print(f"Минусовых дней: {len(negatives):,}".replace(",", " "))
    print(f"Позиции с первым минусом по дням: {len(first_negative):,}".replace(",", " "))
    print(f"Строк в сводном плане по годам: {len(yearly_plan):,}".replace(",", " "))
    print(f"Листов по годам: {len(plan_by_year):,}".replace(",", " "))
    print(f"Результат записан: {output_path}")


if __name__ == "__main__":
    main()
