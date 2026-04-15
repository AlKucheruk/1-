import re
import sys
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import difflib

SCRIPT_DIR = Path(__file__).resolve().parent
MAPPING_FILE = SCRIPT_DIR / "mapping_operations_ut.xlsx"
DDS_MAPPING_FILE = SCRIPT_DIR / "dds_mapping_ut_bp.xlsx"

EXPECTED_COLUMNS = [
    "УИД",
    "ТипДокумента",
    "Номер1С",
    "Дата1С",
    "НомерБанк",
    "ДатаБанк",
    "Организация",
    "ИНН_Орг",
    "КПП_Орг",
    "НашСчет",
    "НашБанк",
    "НашБИК",
    "ВидОперации",
    "Контрагент",
    "ИНН_Контр",
    "КПП_Контр",
    "КодКонтр",
    "Договор",
    "ДокОснование",
    "Сумма",
    "Валюта",
    "Назначение",
    "СчетРасчетов",
    "СчетКонтр",
    "БанкКонтр",
    "БИК_Контр",
    "СтатьяДДС",
    "Комментарий",
    "Ответственный",
]

MAPPING_REQUIRED_COLUMNS = [
    "ТипДокументаБП",
    "ВидОперацииБП",
    "ДокументУТ",
    "ОперацияУТ",
    "ГруппаУчета",
    "НаправлениеДенег",
    "ТребуетПроверки",
    "КомментарийМаппинга",
]


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_key(value) -> str:
    return normalize_text(value).lower()


def ask_file_path() -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title="Выберите Excel-файл банковской выгрузки из БП",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
    )

    root.destroy()

    if not file_path:
        raise ValueError("Файл не выбран.")

    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    return path


def load_excel_first_sheet(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if not xls.sheet_names:
        raise ValueError("В Excel-файле не найдено ни одного листа.")
    sheet_name = xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet_name)


def validate_bank_file(df: pd.DataFrame):
    missing = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            "Файл не соответствует ожидаемому формату.\n"
            f"Не хватает колонок: {missing}"
        )


def validate_mapping_file(df_mapping: pd.DataFrame):
    missing = [col for col in MAPPING_REQUIRED_COLUMNS if col not in df_mapping.columns]
    if missing:
        raise ValueError(
            "Файл маппинга не соответствует ожидаемому формату.\n"
            f"Не хватает колонок: {missing}"
        )


def load_mapping(mapping_path: Path) -> pd.DataFrame:
    if not mapping_path.exists():
        raise FileNotFoundError(
            f"Не найден файл маппинга: {mapping_path}\n"
            f"Сначала запустите 01_create_mapping_table.py"
        )

    xls = pd.ExcelFile(mapping_path)
    sheet_name = "mapping" if "mapping" in xls.sheet_names else xls.sheet_names[0]
    df_mapping = pd.read_excel(mapping_path, sheet_name=sheet_name)

    validate_mapping_file(df_mapping)

    df_mapping = df_mapping.copy()
    df_mapping["ТипДокументаБП_norm"] = df_mapping["ТипДокументаБП"].apply(normalize_key)
    df_mapping["ВидОперацииБП_norm"] = df_mapping["ВидОперацииБП"].apply(normalize_key)

    return df_mapping


def load_dds_mappings(dds_path: Path) -> dict:
    """
    Загружает mаппинг статей ДДС из файла dds_mapping_ut_bp.xlsx и возвращает словарь ключ->список записей.
    Попытается найти колонки с названиями 'Код' и 'Наименование' (и вариациями).
    """
    dds_map = {}
    if not dds_path.exists():
        # не фатально — будем работать без маппинга по статьям
        return dds_map

    try:
        xls = pd.ExcelFile(dds_path)
        sheet = "mapping" if "mapping" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(dds_path, sheet_name=sheet)
    except Exception:
        return dds_map

    # Попытка определить колонки для кода/наименования
    cols = [c.lower() for c in df.columns]
    code_col = None
    name_col = None
    for c_orig, c_low in zip(df.columns, cols):
        if "код" in c_low or "code" in c_low:
            code_col = c_orig
        if "наимен" in c_low or "name" in c_low or "статья" in c_low:
            # prefer "наименование" like columns
            name_col = c_orig

    # если не нашли, ставим первые две колонки
    if not name_col and len(df.columns) >= 1:
        name_col = df.columns[0]
    if not code_col and len(df.columns) >= 2:
        code_col = df.columns[1]

    # Составляем map: нормализованный ключ -> список записей
    def add_entry(code, name, source="unknown"):
        if pd.isna(name) and pd.isna(code):
            return
        # primary key by normalized name and by code
        if not pd.isna(name):
            nk = normalize_key(name)
            dds_map.setdefault(nk, []).append({'source': source, 'code': str(code) if not pd.isna(code) else '', 'name': name})
        if not pd.isna(code):
            kc = normalize_key(str(code))
            dds_map.setdefault(kc, []).append({'source': source, 'code': str(code), 'name': name})

    for _, r in df.iterrows():
        code = r.get(code_col) if code_col in df.columns else None
        name = r.get(name_col) if name_col in df.columns else None
        # Попытка понять тип (income/expense) по наличию слова "доход" или "расход" в наименовании листа или в строке
        src = "unknown"
        if isinstance(name, str):
            nl = name.lower()
            if "доход" in nl:
                src = "income"
            elif "расход" in nl:
                src = "expense"
        add_entry(code, name, source=src)

    return dds_map


def find_best_dds_match(statia_text: str, dds_map: dict):
    """
    Попытка найти подходящую статью по тексту назначения (СтатьяДДС).
    - Сначала точное совпадение по нормализованной строке.
    - Потом поиск вхождения ключей в тексте.
    - Потом fuzzy match (difflib).
    """
    if not statia_text or not dds_map:
        return None
    s = normalize_key(statia_text)

    # точное совпадение
    if s in dds_map:
        return dds_map[s][0]

    # поиск ключа как подстроки
    best = None
    best_len = 0
    for key, cand_list in dds_map.items():
        if key and key in s:
            if len(key) > best_len:
                best = cand_list[0]
                best_len = len(key)
    if best:
        return best

    # fuzzy
    keys = list(dds_map.keys())
    if not keys:
        return None
    matches = difflib.get_close_matches(s, keys, n=1, cutoff=0.75)
    if matches:
        return dds_map[matches[0]][0]

    return None


def clean_source_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all").reset_index(drop=True)

    for col in ["ТипДокумента", "ВидОперации", "Назначение", "ДокОснование"]:
        if col in df.columns:
            df[col] = df[col].apply(normalize_text)

    def keep_row(row):
        return any([
            normalize_text(row.get("ТипДокумента", "")),
            normalize_text(row.get("ВидОперации", "")),
            not pd.isna(row.get("Сумма", None)),
        ])

    df = df[df.apply(keep_row, axis=1)].reset_index(drop=True)
    return df


def build_mapping_dict(df_mapping: pd.DataFrame):
    mapping_dict = {}
    duplicates = []

    for _, row in df_mapping.iterrows():
        key = (row["ТипДокументаБП_norm"], row["ВидОперацииБП_norm"])
        if key in mapping_dict:
            duplicates.append(key)
        mapping_dict[key] = {
            "ДокументУТ": row["ДокументУТ"],
            "ОперацияУТ": row["ОперацияУТ"],
            "ГруппаУчета": row["ГруппаУчета"],
            "НаправлениеДенег": row["НаправлениеДенег"],
            "ТребуетПроверки": row["ТребуетПроверки"],
            "КомментарийМаппинга": row["КомментарийМаппинга"],
        }

    if duplicates:
        raise ValueError(
            "В файле маппинга найдены дубли по ключу "
            "(ТипДокументаБП + ВидОперацииБП): "
            f"{duplicates}"
        )

    return mapping_dict


def parse_amount_str_to_float(text: str):
    if not text:
        return None

    cleaned = text.strip()
    cleaned = cleaned.replace(" ", "")
    cleaned = cleaned.replace("-", ".")
    cleaned = cleaned.replace(",", ".")

    match = re.search(r"(\d+(?:\.\d{1,2})?)", cleaned)
    if not match:
        return None

    try:
        return float(match.group(1))
    except Exception:
        return None


def parse_document_date(value):
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().date()

    if isinstance(value, datetime):
        return value.date()

    text = normalize_text(value)
    if not text:
        return None

    formats = [
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    return None


def get_default_vat_by_date(doc_date):
    border_date = datetime(2019, 1, 1).date()
    if doc_date and doc_date < border_date:
        return "18%"
    return "20%"


def detect_vat(text: str, doc_date=None):
    source = normalize_text(text)
    source_lower = source.lower()

    if not source:
        return {
            "НДС_ИзТекста": "Без НДС",
            "СуммаНДС_ИзТекста": None,
            "ИсточникНДС": "Нет упоминания НДС в тексте",
        }

    no_vat_patterns = [
        r"\bбез\s+ндс\b",
        r"ндс\s+не\s+облагается",
        r"без\s+налога\s*\(?\s*ндс\s*\)?",
        r"без\s+налога",
        r"не\s+облагается",
    ]

    for pattern in no_vat_patterns:
        if re.search(pattern, source_lower, flags=re.IGNORECASE):
            return {
                "НДС_ИзТекста": "Без НДС",
                "СуммаНДС_ИзТекста": None,
                "ИсточникНДС": "Распознано по тексту",
            }

    explicit_patterns = [
        r"ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
        r"в\s*т\.?\s*ч\.?\s*ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
        r"в\s*том\s*числе\s*ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
    ]

    for pattern in explicit_patterns:
        m = re.search(pattern, source_lower, flags=re.IGNORECASE)
        if m:
            rate = f"{m.group(1)}%"
            amount = parse_amount_str_to_float(m.group(2))
            return {
                "НДС_ИзТекста": rate,
                "СуммаНДС_ИзТекста": amount,
                "ИсточникНДС": "Распознано по ставке и сумме",
            }

    rate_patterns = [
        r"ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?",
        r"в\s*т\.?\s*ч\.?\s*ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?",
        r"в\s*том\s*числе\s*ндс[^\d]{0,15}\(?\s*(20|18|10|7|5|0)\s*%\s*\)?",
    ]

    for pattern in rate_patterns:
        m = re.search(pattern, source_lower, flags=re.IGNORECASE)
        if m:
            rate = f"{m.group(1)}%"
            return {
                "НДС_ИзТекста": rate,
                "СуммаНДС_ИзТекста": None,
                "ИсточникНДС": "Распознано по ставке",
            }

    amount_only_patterns = [
        r"ндс\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
        r"в\s*т\.?\s*ч\.?\s*ндс\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
        r"в\s*том\s*числе\s*ндс\s*[-:–—]?\s*([\d\s]+(?:[.,-]\d{2})?)",
    ]

    for pattern in amount_only_patterns:
        m = re.search(pattern, source_lower, flags=re.IGNORECASE)
        if m:
            amount = parse_amount_str_to_float(m.group(1))
            default_rate = get_default_vat_by_date(doc_date)
            return {
                "НДС_ИзТекста": default_rate,
                "СуммаНДС_ИзТекста": amount,
                "ИсточникНДС": "Есть сумма НДС, ставка определена по дате документа",
            }

    if "ндс" in source_lower and "взимается дополнительно" in source_lower:
        default_rate = get_default_vat_by_date(doc_date)
        return {
            "НДС_ИзТекста": default_rate,
            "СуммаНДС_ИзТекста": None,
            "ИсточникНДС": "НДС указан без ставки, ставка определена по дате документа",
        }

    if "ндс" in source_lower:
        default_rate = get_default_vat_by_date(doc_date)
        return {
            "НДС_ИзТекста": default_rate,
            "СуммаНДС_ИзТекста": None,
            "ИсточникНДС": "Есть упоминание НДС, ставка определена по дате документа",
        }

    return {
        "НДС_ИзТекста": "Без НДС",
        "СуммаНДС_ИзТекста": None,
        "ИсточникНДС": "Нет упоминания НДС в тексте",
    }


def apply_mapping(df: pd.DataFrame, mapping_dict: dict, dds_map: dict) -> pd.DataFrame:
    df = df.copy()

    new_cols = [
        "ДокументУТ",
        "ОперацияУТ",
        "ГруппаУчета",
        "НаправлениеДенег",
        "ТребуетПроверки",
        "КомментарийМаппинга",
        "НДС_ИзТекста",
        "СуммаНДС_ИзТекста",
        "ИсточникНДС",
        "МаппингНайден",
        "СтатьяДДС_УТ",
    ]

    for col in new_cols:
        if col not in df.columns:
            df[col] = None

    for idx, row in df.iterrows():
        doc_type = normalize_key(row.get("ТипДокумента", ""))
        operation = normalize_key(row.get("ВидОперации", ""))
        key = (doc_type, operation)

        if key in mapping_dict:
            mapped = mapping_dict[key]
            for col, value in mapped.items():
                df.at[idx, col] = value
            df.at[idx, "МаппингНайден"] = "Да"
        else:
            # Попытка fallback по СтатьяДДС (из файла dds_mapping)
            statia = normalize_text(row.get("СтатьяДДС", "")) or normalize_text(row.get("Назначение", ""))
            best = find_best_dds_match(statia, dds_map) if dds_map else None
            if best:
                df.at[idx, "СтатьяДДС_УТ"] = best.get('name')
                # Эвристика по типу статьи (income/expense/unknown)
                src = best.get('source', 'unknown')
                if src == 'income':
                    df.at[idx, "ДокументУТ"] = "Поступление безналичных ДС"
                    df.at[idx, "ОперацияУТ"] = "Прочее поступление"
                    df.at[idx, "ГруппаУчета"] = "Покупатели"
                    df.at[idx, "НаправлениеДенег"] = "Приход"
                elif src == 'expense':
                    df.at[idx, "ДокументУТ"] = "Списание безналичных ДС"
                    df.at[idx, "ОперацияУТ"] = "Прочее списание"
                    df.at[idx, "ГруппаУчета"] = "Прочее"
                    df.at[idx, "НаправлениеДенег"] = "Расход"
                else:
                    # если неизвестно — пробуем догадаться по сумме (плюс/минус)
                    try:
                        s = float(row.get("Сумма") or 0)
                        if s >= 0:
                            df.at[idx, "ДокументУТ"] = "Поступление безналичных ДС"
                            df.at[idx, "ОперацияУТ"] = "Прочее поступление"
                            df.at[idx, "НаправлениеДенег"] = "Приход"
                        else:
                            df.at[idx, "ДокументУТ"] = "Списание безналичных ДС"
                            df.at[idx, "ОперацияУТ"] = "Прочее списание"
                            df.at[idx, "НаправлениеДенег"] = "Расход"
                    except Exception:
                        df.at[idx, "ДокументУТ"] = None
                        df.at[idx, "ОперацияУТ"] = None
                        df.at[idx, "НаправлениеДенег"] = None

                df.at[idx, "МаппингНайден"] = "Да(поСтатьеДДС)"
                df.at[idx, "КомментарийМаппинга"] = f"Найдено по СтатьеДДС: {best.get('name')} (код {best.get('code')})"
                # Оставляем ТребуетПроверки пустым (оператор может проверить при необходимости)
            else:
                df.at[idx, "ДокументУТ"] = None
                df.at[idx, "ОперацияУТ"] = None
                df.at[idx, "ГруппаУчета"] = None
                df.at[idx, "НаправлениеДенег"] = None
                df.at[idx, "ТребуетПроверки"] = "Да"
                df.at[idx, "КомментарийМаппинга"] = "Не найдено соответствие в mapping_operations_ut.xlsx и по СтатьеДДС"
                df.at[idx, "МаппингНайден"] = "Нет"

        vat_source_text = " ".join([
            normalize_text(row.get("Назначение", "")),
            normalize_text(row.get("ДокОснование", "")),
        ]).strip()

        doc_date = parse_document_date(row.get("ДатаБанк")) or parse_document_date(row.get("Дата1С"))
        vat_result = detect_vat(vat_source_text, doc_date=doc_date)

        df.at[idx, "НДС_ИзТекста"] = vat_result["НДС_ИзТекста"]
        df.at[idx, "СуммаНДС_ИзТекста"] = vat_result["СуммаНДС_ИзТекста"]
        df.at[idx, "ИсточникНДС"] = vat_result["ИсточникНДС"]

    return df


def build_unmapped_report(df: pd.DataFrame) -> pd.DataFrame:
    report = df[df["МаппингНайден"] == "Нет"][["ТипДокумента", "ВидОперации"]].copy()
    if report.empty:
        return pd.DataFrame(columns=["ТипДокумента", "ВидОперации", "Количество"])
    report["Количество"] = 1
    report = (
        report.groupby(["ТипДокумента", "ВидОперации"], dropna=False)["Количество"]
        .sum()
        .reset_index()
        .sort_values(["ТипДокумента", "ВидОперации"])
    )
    return report


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = [
        {"Показатель": "Всего строк", "Значение": len(df)},
        {"Показатель": "Маппинг найден", "Значение": int((df["МаппингНайден"] == "Да").sum())},
        {"Показатель": "Маппинг не найден", "Значение": int((df["МаппингНайден"] == "Нет").sum())},
        {"Показатель": "Требует проверки = Да", "Значение": int((df["ТребуетПроверки"] == "Да").sum())},
        {"Показатель": "НДС = Без НДС", "Значение": int((df["НДС_ИзТекста"] == "Без НДС").sum())},
        {"Показатель": "НДС = 20%", "Значение": int((df["НДС_ИзТекста"] == "20%").sum())},
        {"Показатель": "НДС = 18%", "Значение": int((df["НДС_ИзТекста"] == "18%").sum())},
        {"Показатель": "НДС = 10%", "Значение": int((df["НДС_ИзТекста"] == "10%").sum())},
        {"Показатель": "НДС = 7%", "Значение": int((df["НДС_ИзТекста"] == "7%").sum())},
        {"Показатель": "НДС = 5%", "Значение": int((df["НДС_ИзТекста"] == "5%").sum())},
    ]
    return pd.DataFrame(summary)


def save_result(input_path: Path, df_result: pd.DataFrame):
    output_path = input_path.with_name(f"{input_path.stem}_prepared.xlsx")
    unmapped_report = build_unmapped_report(df_result)
    summary = build_summary(df_result)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_result.to_excel(writer, sheet_name="Данные", index=False)
        unmapped_report.to_excel(writer, sheet_name="НераспознанныеОперации", index=False)
        summary.to_excel(writer, sheet_name="Сводка", index=False)

    return output_path


def main():
    try:
        input_path = ask_file_path()
        df_source = load_excel_first_sheet(input_path)
        validate_bank_file(df_source)

        df_source = clean_source_dataframe(df_source)
        df_mapping = load_mapping(MAPPING_FILE)
        mapping_dict = build_mapping_dict(df_mapping)

        # загрузка маппинга статей ДДС (если есть)
        dds_map = load_dds_mappings(DDS_MAPPING_FILE)
        if not dds_map:
            # не фатально, просто логируем в комментариях (можно показать пользователю)
            print("Предупреждение: файл маппинга статей ДДС не найден или пуст. Будет использоваться только mapping_operations_ut.xlsx")

        df_result = apply_mapping(df_source, mapping_dict, dds_map)
        output_path = save_result(input_path, df_result)

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        messagebox.showinfo(
            "Обработка завершена",
            f"Файл успешно обработан.\n\nРезультат сохранён:\n{output_path}"
        )

        root.destroy()

        print(f"Результат сохранен: {output_path}")

    except Exception as e:
        # При исключении показываем окно и выходим
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showerror("Ошибка", str(e))
            root.destroy()
        except Exception:
            pass

        print("\nОШИБКА:")
        print(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()