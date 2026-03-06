import io
import os
import re
import sys
import tempfile
import unicodedata
import logging
import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st
from xlsxwriter.utility import xl_rowcol_to_cell
logger = logging.getLogger(__name__)

def norm(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).replace("\u00A0", " ")
    s = unicodedata.normalize("NFKC", s)
    return s.strip().strip(":").lower()

# ЛОГИЧЕСКИЕ КОЛОНКИ ДЛЯ БД
COL_CONTRACT = norm("№ Ген. договора")
COL_OIV = norm("ОИВ")
COL_SAPR_DT = norm("Дата загрузки в САПР МГГТ")
COL_AGR_DT = norm("Дата согласования границ")
COL_ASU_DT = norm("Дата загрузки в АСУ ОДС (АСД)")
COL_ASU_APPROVE_DT = norm("Дата утверждения в АСУ ОДС (МГГТ)")
COL_HA = norm("Сумма Объем заказа, га")
COL_ORDER = norm("№ Заказа МГГТ")
COL_STATE = norm("Состояние (действующий / приостановлен / аннулирован)")
COL_STATUS = norm("Статус загрузки")
COL_FIELD_DT = norm("Дата изготовления полевых работ по факту")

PREV_REPORTS_FILE = "prev_reports.pkl"


#  ХРАНЕНИЕ ПРОШЛЫХ ОТЧЁТОВ С ДОПОЛНИТЕЛЬНЫМИ ДАННЫМИ
def load_prev_reports():
    """Загружает все предыдущие отчёты с их дифференциальными данными"""
    if os.path.isfile(PREV_REPORTS_FILE):
        try:
            data = pd.read_pickle(PREV_REPORTS_FILE)
            if isinstance(data, dict):
                # Проверяем структуру данных и дополняем при необходимости
                for year in data:
                    if isinstance(data[year], dict):
                        # Гарантируем наличие всех необходимых полей
                        if "comparison_values_pieces" not in data[year]:
                            data[year]["comparison_values_pieces"] = {}
                        if "comparison_values_hect" not in data[year]:
                            data[year]["comparison_values_hect"] = {}
                        if "delta_values_pieces" not in data[year]:
                            data[year]["delta_values_pieces"] = {}
                        if "delta_values_hect" not in data[year]:
                            data[year]["delta_values_hect"] = {}
                return data
        except Exception:
            logger.exception("Failed to load previous reports")
    return {}


def save_prev_reports(data: dict):
    """Сохраняет предыдущие отчёты со всей структурой данных"""
    try:
        pd.to_pickle(data, PREV_REPORTS_FILE)
    except Exception:
        logger.exception("Failed to save previous reports")


def calculate_comparison_values(current_totals_df, prev_totals_df, plan_col1, plan_col3):
    """
    Рассчитывает значения для строки 'изменение с отчетом неделей ранее'

    Args:
        current_totals_df: DataFrame текущих итогов (последняя строка с 'Итого:')
        prev_totals_df: DataFrame прошлых итогов или None
        plan_col1: Название колонки 'Утвержденный график {year} года'
        plan_col3: Название колонки 'Всего план (скорректированный график {year} года)'

    Returns:
        dict: Словарь с рассчитанными значениями изменений
    """
    comparison_dict = {}

    columns_to_compare = [
        plan_col1,
        plan_col3,
        "Выполнено полевое обследование",
        "Загружено в САПР",
        "Согласовано в САПР",
        "Загружено в АСУ ОДС",
        "Не Загружено в АСУ ОДС",
        "Отклонено",
        "Не утверждено БД",
        "Утверждено",
    ]

    if prev_totals_df is not None and not prev_totals_df.empty:
        # Есть прошлые данные = рассчитываем разницу
        for col in columns_to_compare:
            if col in current_totals_df.columns and col in prev_totals_df.columns:
                current_val = current_totals_df[col].iloc[0] if col in current_totals_df else 0
                prev_val = prev_totals_df[col].iloc[0] if col in prev_totals_df else 0

                # Для числовых значений считаем разницу
                try:
                    comparison_dict[col] = float(current_val) - float(prev_val)
                except (ValueError, TypeError):
                    comparison_dict[col] = 0
            else:
                comparison_dict[col] = 0
    else:
        # Нет прошлых данных = все изменения равны 0
        for col in columns_to_compare:
            comparison_dict[col] = 0

    return comparison_dict


def calculate_delta_values(df_totals):
    """
    Рассчитывает значения стрелочек ▲ (разницы) для блока сравнения

    Args:
        df_totals: DataFrame с итоговой строкой ('Итого:')

    Returns:
        dict: Словарь с значениями ▲ (разниц)
    """
    delta_dict = {}

    if df_totals is not None and not df_totals.empty and "Итого:" in df_totals["ОИВ"].values:
        totals_row = df_totals[df_totals["ОИВ"] == "Итого:"].iloc[0]

        # Рассчитываем разницы
        delta_dict["delta_sapr"] = (
                totals_row.get("Согласовано в САПР", 0) -
                totals_row.get("Загружено в САПР", 0)
        )

        delta_dict["delta_agr"] = (
                totals_row.get("Загружено в АСУ ОДС", 0) -
                totals_row.get("Согласовано в САПР", 0)
        )

        delta_dict["delta_asu"] = (
                totals_row.get("Утверждено", 0) -
                totals_row.get("Загружено в АСУ ОДС", 0)
        )
    else:
        delta_dict["delta_sapr"] = 0
        delta_dict["delta_agr"] = 0
        delta_dict["delta_asu"] = 0

    return delta_dict


# УТИЛИТЫ


SYNONYMS = {
    "оив": {"оив"},
    "№ ген. договора": {
        "№ ген. договора", "номер ген. договора", "номер договора",
        "ген. договор", "ген договор", "ген.договора", "ген.договор"
    },
    "дата загрузки в сапр мггт": {
        "дата загрузки в сапр мггт", "дата загрузки в сапр", "дата_загрузки_в_сапр"
    },
    "согласование границ в сапр": {
        "согласование границ в сапр", "согласовано в сапр", "согласование в сапр",
        "дата согласования в сапр", "дата согласования границ"
    },
    "дата загрузки в асу одс (асд)": {
        "дата загрузки в асу одс (асд)", "дата загрузки в асу одс",
        "дата_загрузки_в_асу_одс", "загружено в асу одс"
    },
    "дата утверждения в асу одс (мггт)": {
        "дата утверждения в асу одс (мггт)", "дата утверждения в асу одс",
        "утверждено в асу одс", "дата утверждения асу одс"
    },
    "сумма объем заказа, га": {
        "сумма объем заказа, га","сумма объём заказа, га","объем, га","объём, га","объем (га)",
        "объём (га)","sum of объем заказа, га","sum of объём заказа, га"
    },

    "№ заказа мггт": {
        "№ заказа мггт", "номер заказа мггт", "заказ мггт", "№ заказа", "номер заказа"
    },
    "состояние (действующий / приостановлен / аннулирован)": {
        "состояние (действующий / приостановлен / аннулирован)", "состояние", "статус"
    },
    "статус загрузки": {"статус загрузки", "статус"},
    "дата изготовления полевых работ по факту": {
        "дата изготовления полевых работ по факту",
        "дата изготовления полевых работ",
        "дата полевых работ",
        "дата фактических полевых работ",
    },
}


def find_similar(logical_name: str, columns) -> str | None:
    key = norm(logical_name)
    candidates = SYNONYMS.get(key, {key})
    for real in columns:
        if norm(real) in candidates:
            return real
    return None

def add_year_to_corr_plan_column(df: pd.DataFrame, selected_year: int) -> pd.DataFrame:
    """
    В прошлом отчёте колонка может называться
    'Всего план (скорректированный график)'
    → приводим к
    'Всего план (скорректированный график {year} года)'
    """
    if df is None or df.empty:
        return df

    target_col = f"Всего план (скорректированный график {selected_year} года)"

    for c in df.columns:
        nc = norm(c)
        if "всего" in nc and "план" in nc and "скоррект" in nc:
            if c != target_col:
                df = df.rename(columns={c: target_col})
            break

    return df



def find_plan_col_fuzzy_cols(cols, year: int, must_words: list[str]):
    """
    Ищем колонку плана по смысловым словам, год в заголовке игнорируем.
    year оставлен в сигнатуре для совместимости с текущими вызовами.
    """
    candidates = []
    for idx, c in enumerate(cols):
        nc = norm(c)
        if all(w in nc for w in must_words):
            candidates.append((idx, nc))

    if not candidates:
        return None

    preferred_tokens = ["скоррект", "всего", "план", "график", "утвержден"]
    def score(nc: str) -> int:
        return sum(1 for t in preferred_tokens if t in nc)

    candidates.sort(key=lambda x: score(x[1]), reverse=True)
    return candidates[0][0]



OIV_ALIASES = {
    norm("Комитет ветеринарии города Москвы (Москомвет)"): norm("москомвет"),
    norm("Департамент строительства города Москвы"): norm("департамент стоительства города москвы"),
}


def load_single_plan_sheet(plan_path: str, sheet_index: int, year: int, label: str):
    xls = pd.ExcelFile(plan_path)
    if sheet_index >= len(xls.sheet_names):
        return {}
    df = pd.read_excel(plan_path, sheet_name=xls.sheet_names[sheet_index])
    if df.empty:
        return {}

    norm_cols = [norm(c) for c in df.columns]

    # колонка ОИВ
    oiv_idx = None
    for i, nc in enumerate(norm_cols):
        if ("оив" in nc or "департамент" in nc or "префектура" in nc
                or "комитет" in nc or "москомвет" in nc):
            oiv_idx = i
            break
    if oiv_idx is None:
        oiv_idx = 0

    # утвержденный график
    c1_idx = find_plan_col_fuzzy_cols(df.columns, year, ["утвержден", "график"])
    if c1_idx is None:
        start = oiv_idx + 1
        c1_idx = start if start < len(df.columns) else None

    if c1_idx is None:
        logger.warning(
            "Could not detect approved plan column for sheet '%s'",
            label,
        )
        return {}

    plan_dict = {}
    for _, row in df.iterrows():
        oiv_val = row.iloc[oiv_idx]
        if pd.isna(oiv_val):
            continue
        key = norm(str(oiv_val))
        if not key:
            continue
        v1 = float(row.iloc[c1_idx]) if pd.notna(row.iloc[c1_idx]) else 0.0
        plan_dict[key] = v1
    logger.info("Loaded %s plan rows for sheet '%s'", len(plan_dict), label)
    return plan_dict


def load_plan_dicts(plan_path: str | None, year: int):
    if not plan_path or not os.path.isfile(plan_path):
        return {}, {}
    try:
        plan_pieces = load_single_plan_sheet(plan_path, 0, year, "ШТ")
        plan_hect = load_single_plan_sheet(plan_path, 1, year, "ГА")
        return plan_pieces, plan_hect
    except Exception:
        logger.exception("Failed to load previous reports")
        return {}, {}

def make_output_path(filename: str) -> str:
    """
    Пытаемся сохранять в папку 'готовые отчеты'.
    """
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    preferred_dir = os.path.join(base_dir, "готовые отчеты")

    try:
        os.makedirs(preferred_dir, exist_ok=True)
        return os.path.join(preferred_dir, filename)
    except Exception:
        return os.path.join(os.getcwd(), filename)


def build_header_display_map(selected_year: int):
    return {
        "ОИВ": "ОИВ",
        f"Утвержденный график {selected_year} года": f"Утвержденный\nграфик {selected_year} года",
        f"Всего план (скорректированный график {selected_year} года)": "Всего план\n(скорректированный график)",
        "Выполнено полевое обследование": "Выполнено\nполевое обследование",
        "Выполнено полевое обследование (от плана) %": "Выполнено полевое\n(от плана) %",
        "Загружено в САПР": "Загружено\nв САПР",
        "Загружено в САПР %": "Загружено в САПР,\n%",
        "Согласовано в САПР": "Согласовано\nв САПР",
        "Согласовано в САПР (от плана) %": "Согласовано в САПР\n(от плана) %",
        "Загружено в АСУ ОДС": "Загружено\nв АСУ ОДС",
        "Загружено в АСУ ОДС (от плана) %": "Загружено в АСУ ОДС\n(от плана) %",
        "Не Загружено в АСУ ОДС": "Не загружено\nв АСУ ОДС",
        "Отклонено": "Отклонено",
        "Отклонено (от плана) %": "Отклонено\n(от плана) %",
        "Не утверждено БД": "Не утверждено\nБД",
        "Не утверждено БД (от плана) %": "Не утверждено БД\n(от плана) %",
        "Утверждено": "Утверждено",
        "Утверждено (от плана ) %": "Утверждено\n(от плана ) %",
    }


def wrap_long_header(text: str, max_len: int = 25) -> str:
    if "\n" in text or len(text) <= max_len:
        return text
    mid = len(text) // 2
    left_space = text.rfind(" ", 0, mid)
    right_space = text.find(" ", mid)
    if left_space == -1 and right_space == -1:
        return text[:mid] + "\n" + text[mid:]
    if left_space == -1:
        split = right_space
    elif right_space == -1:
        split = left_space
    else:
        split = left_space if (mid - left_space) <= (right_space - mid) else right_space
    return text[:split] + "\n" + text[split + 1:]

def ensure_percent_columns(df: pd.DataFrame, plan_col3: str) -> pd.DataFrame:
    """
    Гарантирует наличие % колонок в df.
    Если колонок нет/они кривые — пересчитывает.
    """
    if df is None or df.empty:
        return df

    df = df.copy()

    # подстраховка: нужные числовые колонки должны существовать
    needed = [
        "Выполнено полевое обследование",
        "Загружено в САПР",
        "Согласовано в САПР",
        "Загружено в АСУ ОДС",
        "Отклонено",
        "Не утверждено БД",
        "Утверждено",
        plan_col3,
    ]
    for c in needed:
        if c not in df.columns:
            df[c] = 0

    plan = pd.to_numeric(df[plan_col3], errors="coerce").replace(0, np.nan)

    df["Выполнено полевое обследование (от плана) %"] = (
        (pd.to_numeric(df["Выполнено полевое обследование"], errors="coerce") / plan) * 100
    ).round().fillna(0).astype(int)

    df["Загружено в САПР %"] = (
        (pd.to_numeric(df["Загружено в САПР"], errors="coerce") / plan) * 100
    ).round().fillna(0).astype(int)

    df["Согласовано в САПР (от плана) %"] = (
        (pd.to_numeric(df["Согласовано в САПР"], errors="coerce") / plan) * 100
    ).round().fillna(0).astype(int)

    df["Загружено в АСУ ОДС (от плана) %"] = (
        (pd.to_numeric(df["Загружено в АСУ ОДС"], errors="coerce") / plan) * 100
    ).round().fillna(0).astype(int)

    loaded = pd.to_numeric(df["Загружено в АСУ ОДС"], errors="coerce").replace(0, np.nan)
    df["Отклонено (от плана) %"] = (
        (pd.to_numeric(df["Отклонено"], errors="coerce") / loaded) * 100
    ).round().fillna(0).astype(int)

    df["Не утверждено БД (от плана) %"] = (
        (pd.to_numeric(df["Не утверждено БД"], errors="coerce") / loaded) * 100
    ).round().fillna(0).astype(int)

    df["Утверждено (от плана ) %"] = (
        (pd.to_numeric(df["Утверждено"], errors="coerce") / loaded) * 100
    ).round().fillna(0).astype(int)

    return df


def n(s: str) -> str:
    return norm(s)


OTKLONENO_STATUSES = {
    n("Запрос обрабатывается"),
    n("Получен ответ об ошибке"),
    n("Проект был отклонен"),
    n("Проект утвержден"),
    n(""),
    n("Задача отправлена в АСУ ОДС"),
    n("Ошибка обработки в АСУ ОДС"),
}
NE_UTV_BD_STATUSES = {
    n("Акт подписан"),
    n("Отправлен на согласование"),
    n("Получен ответ"),
    n("Согласован с внешней системой"),
    n("Объект создан в АСУ ОДС"),
}


def get_status_groups(series: pd.Series):
    s = series.fillna("").astype(str).map(norm)
    return s.isin(OTKLONENO_STATUSES), s.isin(NE_UTV_BD_STATUSES)

def normalize_excel_headers(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    def clean(c):
        s = "" if c is None else str(c)
        s = s.replace("\n", " ").replace("\r", " ")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    df = df.copy()
    df.columns = [clean(c) for c in df.columns]
    return df

def extract_date_from_filename(path: str) -> str:
    """
    Ищет в имени файла дату формата 11.12 или 11-12 или 11_12.
    Возвращает строку вида 11.12.
    Если не найдено — возвращает 'baseline'.
    """
    filename = os.path.basename(path)

    match = re.search(r'(\d{1,2})[.\-_](\d{1,2})', filename)

    if match:
        day = match.group(1).zfill(2)
        month = match.group(2).zfill(2)
        return f"{day}.{month}"

    return "baseline"


def load_prev_report_from_xlsx(report_xlsx_path: str, selected_year: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(report_xlsx_path, sheet_name="Отчёт", header=None)

    def norm_cell(x):
        if pd.isna(x):
            return ""
        return norm(str(x))

    year = selected_year

    # Ищем заголовок первой таблицы (шт): строка, где первая ячейка = "ОИВ"
    header_row_1 = None
    for r in range(min(2000, len(raw))):
        if norm_cell(raw.iat[r, 0]) == norm("ОИВ"):
            header_row_1 = r
            break
    if header_row_1 is None:
        raise ValueError("Не удалось найти заголовок таблицы ШТ (строка с 'ОИВ') на листе 'Отчёт'.")

    cols_1 = [str(c).strip() if not pd.isna(c) else "" for c in raw.iloc[header_row_1].tolist()]
    df1 = raw.iloc[header_row_1 + 1:].copy()
    df1.columns = cols_1

    stop_idx = None
    for i in range(len(df1)):
        first = norm_cell(df1.iloc[i, 0])
        if first == "":
            stop_idx = i
            break
        if first.startswith(norm(f"ГЗ {year}, га")):
            stop_idx = i
            break
    if stop_idx is not None:
        df1 = df1.iloc[:stop_idx]

    df1 = df1.dropna(axis=1, how="all").dropna(axis=0, how="all")

    # ищем строку заголовка блока "га"
    title_ga_row = None
    for r in range(header_row_1 + 1, min(header_row_1 + 2000, len(raw))):
        s = norm_cell(raw.iat[r, 0])
        if s.startswith("гз") and "га" in s:
            title_ga_row = r
            break

    if title_ga_row is None:
        raise ValueError(
            "Не удалось найти заголовок блока 'ГЗ ..., га' на листе 'Отчёт' (ищу строку начинающуюся с 'ГЗ' и содержащую 'га').")

    header_row_2 = None
    for r in range(title_ga_row + 1, min(title_ga_row + 50, len(raw))):
        if norm_cell(raw.iat[r, 0]) == norm("ОИВ"):
            header_row_2 = r
            break
    if header_row_2 is None:
        raise ValueError("Не удалось найти заголовок таблицы ГА (строка с 'ОИВ') после блока 'ГЗ ..., га'.")

    cols_2 = [str(c).strip() if not pd.isna(c) else "" for c in raw.iloc[header_row_2].tolist()]
    df2 = raw.iloc[header_row_2 + 1:].copy()
    df2.columns = cols_2

    stop_idx2 = None
    for i in range(len(df2)):
        if norm_cell(df2.iloc[i, 0]) == "":
            stop_idx2 = i
            break
    if stop_idx2 is not None:
        df2 = df2.iloc[:stop_idx2]

    df2 = df2.dropna(axis=1, how="all").dropna(axis=0, how="all")

    for df in (df1, df2):
        for c in df.columns:
            if norm(c) == norm("ОИВ"):
                continue
            df[c] = pd.to_numeric(df[c], errors="ignore")

    return df1, df2



# ПОСТРОЕНИЕ БЛОКА ШТ
def build_pieces_block(
        dfp: pd.DataFrame,
        oiv_order,
        plan_pieces,
        corr_plan_pieces: pd.Series,
        selected_year: int,
        col_oiv,
        col_sapr,
        col_agr,
        col_asu,
        col_asu_appr,
        col_field,
        col_ha,
        col_status,
):
    col1 = f"Утвержденный график {selected_year} года"
    col3 = f"Всего план (скорректированный график {selected_year} года)"

    base = (
        pd.DataFrame({col_oiv: oiv_order})
        .dropna()
        .drop_duplicates()
        .set_index(col_oiv)
    )
    df_out = base.copy()
    df_out[col1] = 0
    df_out[col3] = 0
    df_out["Выполнено полевое обследование"] = 0

    for oiv_val in df_out.index:
        key = norm(str(oiv_val))
        plan_key = OIV_ALIASES.get(key, key)
        if plan_key in plan_pieces:
            p1 = plan_pieces[plan_key]
            df_out.at[oiv_val, col1] = int(p1)

    if corr_plan_pieces is not None and len(corr_plan_pieces) > 0:
        df_out[col3] = (
            corr_plan_pieces.reindex(df_out.index)
            .fillna(0)
            .astype(int)
        )

    def count_non_empty(col):
        if col not in dfp.columns:
            return pd.Series(0, index=base.index)
        return (
            dfp[col].notna()
            .groupby(dfp[col_oiv])
            .sum()
            .reindex(base.index)
            .fillna(0)
            .astype(int)
        )

    def count_empty(col):
        if col not in dfp.columns:
            return pd.Series(0, index=base.index)
        return (
            dfp[col].isna()
            .groupby(dfp[col_oiv])
            .sum()
            .reindex(base.index)
            .fillna(0)
            .astype(int)
        )

    df_out["Выполнено полевое обследование"] = count_non_empty(col_field)
    df_out["Загружено в САПР"] = count_non_empty(col_sapr)
    df_out["Согласовано в САПР"] = count_non_empty(col_agr)
    df_out["Загружено в АСУ ОДС"] = count_non_empty(col_asu)
    df_out["Не Загружено в АСУ ОДС"] = count_empty(col_asu)

    status_otk, status_ne = get_status_groups(dfp[col_status])


    gate_mask = dfp[col_asu].notna() & dfp[col_asu_appr].isna()
    approved_mask = dfp[col_asu].notna() & dfp[col_asu_appr].notna()

    otk_counts = (
        (gate_mask & status_otk)
        .groupby(dfp[col_oiv])
        .sum()
        .reindex(df_out.index)
        .fillna(0)
        .astype(int)
    )
    ne_counts = (
        (gate_mask & status_ne)
        .groupby(dfp[col_oiv])
        .sum()
        .reindex(df_out.index)
        .fillna(0)
        .astype(int)
    )
    appr_counts = (
        approved_mask
        .groupby(dfp[col_oiv])
        .sum()
        .reindex(df_out.index)
        .fillna(0)
        .astype(int)
    )

    df_out["Отклонено"] = otk_counts
    df_out["Не утверждено БД"] = ne_counts
    df_out["Утверждено"] = appr_counts

    total = pd.DataFrame(df_out.sum(numeric_only=True)).T
    total.index = ["Итого:"]
    df_final = pd.concat([df_out, total])

    plan_series = df_final[col3].replace(0, np.nan)

    df_final["Выполнено полевое обследование (от плана) %"] = (
            (df_final["Выполнено полевое обследование"] / plan_series) * 100
    ).round().fillna(0).astype(int)

    df_final["Загружено в САПР %"] = (
            (df_final["Загружено в САПР"] / plan_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Согласовано в САПР (от плана) %"] = (
            (df_final["Согласовано в САПР"] / plan_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Загружено в АСУ ОДС (от плана) %"] = (
            (df_final["Загружено в АСУ ОДС"] / plan_series) * 100
    ).round().fillna(0).astype(int)

    loaded_series = df_final["Загружено в АСУ ОДС"].replace(0, np.nan)
    df_final["Отклонено (от плана) %"] = (
            (df_final["Отклонено"] / loaded_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Не утверждено БД (от плана) %"] = (
            (df_final["Не утверждено БД"] / loaded_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Утверждено (от плана ) %"] = (
            (df_final["Утверждено"] / loaded_series) * 100
    ).round().fillna(0).astype(int)

    cols_out = [
        col1, col3,
        "Выполнено полевое обследование",
        "Выполнено полевое обследование (от плана) %",
        "Загружено в САПР", "Загружено в САПР %",
        "Согласовано в САПР", "Согласовано в САПР (от плана) %",
        "Загружено в АСУ ОДС", "Загружено в АСУ ОДС (от плана) %",
        "Не Загружено в АСУ ОДС",
        "Отклонено", "Отклонено (от плана) %",
        "Не утверждено БД", "Не утверждено БД (от плана) %",
        "Утверждено", "Утверждено (от плана ) %",
    ]
    df_final = df_final[cols_out]
    df_final = df_final.reset_index().rename(columns={"index": "ОИВ"})
    return df_final


# ПОСТРОЕНИЕ БЛОКА ГА
def build_hectares_block(
        dfp: pd.DataFrame,
        oiv_order,
        plan_hect,
        corr_plan_hect: pd.Series,
        selected_year: int,
        col_oiv,
        col_sapr,
        col_agr,
        col_asu,
        col_asu_appr,
        col_field,
        col_ha,
        col_status,
):
    col1 = f"Утвержденный график {selected_year} года"
    col3 = f"Всего план (скорректированный график {selected_year} года)"

    base = (
        pd.DataFrame({col_oiv: oiv_order})
        .dropna()
        .drop_duplicates()
        .set_index(col_oiv)
    )
    df_out = base.copy()
    df_out[col1] = 0.0
    df_out[col3] = 0.0
    df_out["Выполнено полевое обследование"] = 0.0

    for oiv_val in df_out.index:
        key = norm(str(oiv_val))
        plan_key = OIV_ALIASES.get(key, key)
        if plan_key in plan_hect:
            p1 = plan_hect[plan_key]
            df_out.at[oiv_val, col1] = float(p1)

    if corr_plan_hect is not None and len(corr_plan_hect) > 0:
        df_out[col3] = (
            corr_plan_hect.reindex(df_out.index)
            .fillna(0.0)
            .astype(float)
        )

    def sum_ha_for(col):
        if col not in dfp.columns:
            return pd.Series(0.0, index=base.index)
        mask = dfp[col].notna()
        s = (
            dfp.loc[mask]
            .groupby(dfp[col_oiv])[col_ha]
            .sum(min_count=1)
        )
        return s.reindex(base.index).fillna(0.0)

    def sum_ha_missing(col):
        if col not in dfp.columns:
            return pd.Series(0.0, index=base.index)
        mask = dfp[col].isna() & dfp[col_ha].notna()
        s = (
            dfp.loc[mask]
            .groupby(dfp[col_oiv])[col_ha]
            .sum(min_count=1)
        )
        return s.reindex(base.index).fillna(0.0)

    df_out["Выполнено полевое обследование"] = sum_ha_for(col_field)
    df_out["Загружено в САПР"] = sum_ha_for(col_sapr)
    df_out["Согласовано в САПР"] = sum_ha_for(col_agr)
    df_out["Загружено в АСУ ОДС"] = sum_ha_for(col_asu)
    df_out["Не Загружено в АСУ ОДС"] = sum_ha_missing(col_asu)

    status_otk, status_ne = get_status_groups(dfp[col_status])

    gate_mask = dfp[col_asu].notna() & dfp[col_asu_appr].isna()
    approved_mask = dfp[col_asu].notna() & dfp[col_asu_appr].notna()

    if len(dfp) > 0:
        otk_ha = (
            dfp.loc[gate_mask & status_otk & dfp[col_ha].notna()]
            .groupby(dfp[col_oiv])[col_ha]
            .sum(min_count=1)
            .reindex(df_out.index)
            .fillna(0.0)
        )
        ne_ha = (
            dfp.loc[gate_mask & status_ne & dfp[col_ha].notna()]
            .groupby(dfp[col_oiv])[col_ha]
            .sum(min_count=1)
            .reindex(df_out.index)
            .fillna(0.0)
        )
        appr_ha = (
            dfp.loc[approved_mask & dfp[col_ha].notna()]
            .groupby(dfp[col_oiv])[col_ha]
            .sum(min_count=1)
            .reindex(df_out.index)
            .fillna(0.0)
        )
    else:
        otk_ha = pd.Series(0.0, index=df_out.index)
        ne_ha = pd.Series(0.0, index=df_out.index)
        appr_ha = pd.Series(0.0, index=df_out.index)

    df_out["Отклонено"] = otk_ha
    df_out["Не утверждено БД"] = ne_ha
    df_out["Утверждено"] = appr_ha

    total = pd.DataFrame(df_out.sum(numeric_only=True)).T
    total.index = ["Итого:"]
    df_final = pd.concat([df_out, total])

    plan_series = df_final[col3].replace(0, np.nan)

    df_final["Выполнено полевое обследование (от плана) %"] = (
            (df_final["Выполнено полевое обследование"] / plan_series) * 100
    ).round().fillna(0).astype(int)

    df_final["Загружено в САПР %"] = (
            (df_final["Загружено в САПР"] / plan_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Согласовано в САПР (от плана) %"] = (
            (df_final["Согласовано в САПР"] / plan_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Загружено в АСУ ОДС (от плана) %"] = (
            (df_final["Загружено в АСУ ОДС"] / plan_series) * 100
    ).round().fillna(0).astype(int)

    loaded_series = df_final["Загружено в АСУ ОДС"].replace(0, np.nan)
    df_final["Отклонено (от плана) %"] = (
            (df_final["Отклонено"] / loaded_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Не утверждено БД (от плана) %"] = (
            (df_final["Не утверждено БД"] / loaded_series) * 100
    ).round().fillna(0).astype(int)
    df_final["Утверждено (от плана ) %"] = (
            (df_final["Утверждено"] / loaded_series) * 100
    ).round().fillna(0).astype(int)

    for c in [
        col1,
        col3,
        "Выполнено полевое обследование",
        "Загружено в САПР",
        "Согласовано в САПР",
        "Загружено в АСУ ОДС",
        "Не Загружено в АСУ ОДС",
        "Отклонено",
        "Не утверждено БД",
        "Утверждено",
    ]:
        df_final[c] = df_final[c].round(2)

    cols_out = [
        col1, col3,
        "Выполнено полевое обследование",
        "Выполнено полевое обследование (от плана) %",
        "Загружено в САПР", "Загружено в САПР %",
        "Согласовано в САПР", "Согласовано в САПР (от плана) %",
        "Загружено в АСУ ОДС", "Загружено в АСУ ОДС (от плана) %",
        "Не Загружено в АСУ ОДС",
        "Отклонено", "Отклонено (от плана) %",
        "Не утверждено БД", "Не утверждено БД (от плана) %",
        "Утверждено", "Утверждено (от плана ) %",
    ]
    df_final = df_final[cols_out]
    df_final = df_final.reset_index().rename(columns={"index": "ОИВ"})
    return df_final


def make_hectares_display_df(df: pd.DataFrame, selected_year: int) -> pd.DataFrame:
    return df.copy()


#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ EXCEL
def write_comparison_values_row(worksheet_obj, df_local, last_data_row,
                                start_row_compare, comparison_values, plan_col1, plan_col3, dyn_cell_fmt):
    """
    Записывает сохранённые значения для строки 'изменение с отчетом неделей ранее'
    в прошлый отчёт.

    Args:
        worksheet_obj: Объект листа Excel
        df_local: DataFrame с данными текущего листа
        last_data_row: Строка с итогами
        start_row_compare: Начальная строка блока сравнения
        comparison_values: dict с рассчитанными значениями изменений
        plan_col1: Название колонки плана 1
        plan_col3: Название колонки плана 3
        dyn_cell_fmt: Формат для ячеек
    """
    if start_row_compare is None:
        return

    target_row = start_row_compare + 7

    cols_to_use = [
        plan_col1,
        plan_col3,
        "Выполнено полевое обследование",
        "Загружено в САПР",
        "Согласовано в САПР",
        "Загружено в АСУ ОДС",
        "Не Загружено в АСУ ОДС",
        "Отклонено",
        "Не утверждено БД",
        "Утверждено",
    ]

    # Записываем сохранённые значения
    for col_name in cols_to_use:
        if col_name not in df_local.columns:
            continue

        idx = df_local.columns.get_loc(col_name)

        if col_name in comparison_values:
            # Записываем сохранённое значение
            value = comparison_values[col_name]

            # Определяем формат в зависимости от типа данных
            if col_name == plan_col1 or col_name == plan_col3:
                # Для плановых значений используем обычный формат
                worksheet_obj.write_number(target_row, idx, float(value), dyn_cell_fmt)
            elif "га" in col_name.lower():
                # Для гектаров округляем до 2 знаков
                worksheet_obj.write_number(target_row, idx, round(float(value), 2), dyn_cell_fmt)
            else:
                # Для штук - целые числа
                try:
                    worksheet_obj.write_number(target_row, idx, int(value), dyn_cell_fmt)
                except (ValueError, TypeError):
                    worksheet_obj.write(target_row, idx, value, dyn_cell_fmt)
        else:
            # Если значения нет в словаре, пишем 0
            worksheet_obj.write(target_row, idx, 0, dyn_cell_fmt)

def write_delta_comparison_cells(
    worksheet,
    df_main,
    last_data_row,
    start_row_compare,
    delta_values_curr,
    delta_values_prev,
    dyn_cols,
    start_col_dyn,
    cell_fmt,
    is_prev_report=False,
):
    """
    Записывает три стрелочные строки (▲) под персиковой таблицей.

    Формулы:
        = (текущее значение строки под основной таблицей)
        - (значение этой же строки из прошлого отчёта)

    Работает и для текущего отчёта, и для прошлого (через флаг is_prev_report).
    """

    if start_row_compare is None:
        return

    row_sapr = start_row_compare + 1   # ▲ от загруженного в САПР
    row_agr  = start_row_compare + 3   # ▲ от согласованного в САПР
    row_asu  = start_row_compare + 5   # ▲ от загруженного в АСУ ОДС

    try:
        col_sogl = start_col_dyn + dyn_cols.index("согласовано в САПР")
        col_otkl = start_col_dyn + dyn_cols.index("отклонено")
        col_utv  = start_col_dyn + dyn_cols.index("утверждено")
    except ValueError:
        return

    # РЕЖИМ ПРОШЛОГО ОТЧЁТА
    if is_prev_report:
        worksheet.write_number(row_sapr, col_sogl, delta_values_curr.get("delta_sapr", 0), cell_fmt)
        worksheet.write_number(row_agr,  col_otkl, delta_values_curr.get("delta_agr", 0),  cell_fmt)
        worksheet.write_number(row_asu,  col_utv,  delta_values_curr.get("delta_asu", 0),  cell_fmt)
        return\

    if not delta_values_prev:
        worksheet.write_number(row_sapr, col_sogl, 0, cell_fmt)
        worksheet.write_number(row_agr, col_otkl, 0, cell_fmt)
        worksheet.write_number(row_asu, col_utv, 0, cell_fmt)
        return

    # РЕЖИМ ТЕКУЩЕГО ОТЧЁТА

    # Текущие значения
    sapr_curr = delta_values_curr.get("delta_sapr", 0)
    agr_curr  = delta_values_curr.get("delta_agr", 0)
    asu_curr  = delta_values_curr.get("delta_asu", 0)

    # Прошлые значения (если есть)
    sapr_prev = delta_values_prev.get("delta_sapr", 0) if delta_values_prev else 0
    agr_prev  = delta_values_prev.get("delta_agr", 0)  if delta_values_prev else 0
    asu_prev  = delta_values_prev.get("delta_asu", 0)  if delta_values_prev else 0

    # Формулы = текущее – прошлое
    formula_sapr = f"={sapr_curr}-{sapr_prev}"
    formula_agr  = f"={agr_curr}-{agr_prev}"
    formula_asu  = f"={asu_curr}-{asu_prev}"

    worksheet.write_formula(row_sapr, col_sogl, formula_sapr, cell_fmt)
    worksheet.write_formula(row_agr,  col_otkl, formula_agr,  cell_fmt)
    worksheet.write_formula(row_asu,  col_utv,  formula_asu,  cell_fmt)


def add_green_numbers_under_persic(
    worksheet,
    workbook,
    df_main,
    last_data_row,
    start_row_compare,
    start_col_dyn,
    dyn_cols,
    selected_year,
    is_hectares: bool = False,
):
    """
    Зелёная строка под персиковой таблицей (динамика недели).

    Для каждой колонки персиковой таблицы (например, "согласовано в САПР") считаем:

        зелёная_ячейка =
            = (ячейка "Итого" в ПЕРСИКОВОЙ таблице по этой колонке)
            - (ячейка из строки "изменение с отчётом за предыдущую неделю"
               в этой же колонке ОСНОВНОЙ таблицы)

    Для гектаров дополнительно округляем результат через ROUND(...;2),
    чтобы не было хвостов типа -3,64E-13.
    Всё делается формулами Excel.
    """

    if start_row_compare is None:
        return None

    # Строка "Итого:" в персиковой таблице (она же строка "Итого" основной таблицы)
    persic_total_row = last_data_row

    # Строка "изменение с отчётом за предыдущую неделю" в блоке сравнения
    compare_row_8 = start_row_compare + 7

    # Зелёная строка — сразу под "Итого:" в персиковой таблице
    green_row = last_data_row + 1

    green_number_fmt = workbook.add_format({
        "font_color": "#00B050",
        "align": "center",
        "valign": "vcenter",
        "bold": True,
    })

    # Соответствие колонок ПЕРСИКОВОЙ таблицы ↔ колонкам ОСНОВНОЙ таблицы
    persic_to_main_map = {
        "Изменение плана скорректированного": (
            f"Всего план (скорректированный график {selected_year} года)"
        ),
        "Выполнено полевое обследование": "Выполнено полевое обследование",
        "загружено в САПР": "Загружено в САПР",
        "согласовано в САПР": "Согласовано в САПР",
        "загружено в АСУ ОДС": "Загружено в АСУ ОДС",
        "не утверждено в АСУ ОДС": "Не утверждено БД",
        "отклонено": "Отклонено",
        "утверждено": "Утверждено",
    }

    # Для каждой колонки персиковой таблицы:
    # зелёная_ячейка = Итого_персиковой - строка_«изменение…» в ОСНОВНОЙ таблице
    for j, persic_col in enumerate(dyn_cols):
        if persic_col not in persic_to_main_map:
            continue

        main_col_name = persic_to_main_map[persic_col]
        if main_col_name not in df_main.columns:
            continue

        # Индекс колонки в основной таблице
        main_col_idx = df_main.columns.get_loc(main_col_name)

        # Ячейка "Итого" в ПЕРСИКОВОЙ таблице под этой колонкой
        persic_cell = xl_rowcol_to_cell(persic_total_row, start_col_dyn + j)

        # Ячейка строки "изменение с отчётом за предыдущую неделю"
        # в соответствующей колонке ОСНОВНОЙ таблицы
        compare_cell = xl_rowcol_to_cell(compare_row_8, main_col_idx)

        if is_hectares:
            formula = f"=ROUND({persic_cell}-{compare_cell},2)"
        else:

            formula = f"={persic_cell}-{compare_cell}"

        worksheet.write_formula(
            green_row,
            start_col_dyn + j,
            formula,
            green_number_fmt,
        )

    return green_row


def set_columns_by_header(ws, df, min_width=14, max_width=50, padding=2):
    """
    Делает ширину колонок по максимуму из:
    - длины заголовка
    - длины данных
    """
    for i, col in enumerate(df.columns):
        header_len = len(str(col))
        data_len = df[col].astype(str).map(len).max()
        width = max(header_len, data_len) + padding
        width = max(min_width, min(width, max_width))
        ws.set_column(i, i, width)


# ГЕНЕРАЦИЯ ОТЧЁТА
def generate_report(
        df_input: pd.DataFrame,
        plan_path: str | None,
        report_type: str,
        selected_year: int,
        baseline_prev_path: str | None = None,
        return_excel_bytes: bool = True,
        preview: bool = False,
) -> tuple[bytes, str] | str | dict:


    #  планы
    plan_pieces, plan_hect = load_plan_dicts(plan_path, selected_year)

    #  прошлые отчёты
    prev_reports = load_prev_reports()
    prev_for_year = prev_reports.get(selected_year) if isinstance(prev_reports, dict) else None

    use_baseline_prev = bool(baseline_prev_path)

    baseline_prev_pieces = pd.DataFrame()
    baseline_prev_hectares = pd.DataFrame()

    if use_baseline_prev:
        res = load_prev_report_from_xlsx(baseline_prev_path, selected_year)
        if res is None:
            raise ValueError("load_prev_report_from_xlsx вернула None (ожидалось (df_pieces, df_hectares))")
        baseline_prev_pieces, baseline_prev_hectares = res


        baseline_prev_pieces = normalize_excel_headers(baseline_prev_pieces)
        baseline_prev_hectares = normalize_excel_headers(baseline_prev_hectares)

        # добавляем год к колонке "Всего план (скорректированный график)"
        baseline_prev_pieces = add_year_to_corr_plan_column(baseline_prev_pieces, selected_year)
        baseline_prev_hectares = add_year_to_corr_plan_column(baseline_prev_hectares, selected_year)

        baseline_prev_pieces = baseline_prev_pieces.loc[:, ~baseline_prev_pieces.columns.duplicated()]
        baseline_prev_hectares = baseline_prev_hectares.loc[:, ~baseline_prev_hectares.columns.duplicated()]

        plan_col3 = f"Всего план (скорректированный график {selected_year} года)"

        baseline_prev_pieces = ensure_percent_columns(baseline_prev_pieces, plan_col3)
        baseline_prev_hectares = ensure_percent_columns(baseline_prev_hectares, plan_col3)

        logger.info(
            "Loaded baseline report: pieces=%s, hectares=%s",
            len(baseline_prev_pieces),
            len(baseline_prev_hectares),
        )


    # Извлекаем все данные из прошлого отчёта
    auto_prev_pieces_df = None
    auto_prev_hectares_df = None
    auto_prev_dyn_pieces_df = None
    auto_prev_dyn_hect_df = None
    auto_prev_comparison_pieces = {}
    auto_prev_comparison_hect = {}
    auto_prev_delta_pieces = {}
    auto_prev_delta_hect = {}

    if isinstance(prev_for_year, dict):
        auto_prev_pieces_df = prev_for_year.get("pieces")
        auto_prev_hectares_df = prev_for_year.get("hectares")
        auto_prev_dyn_pieces_df = prev_for_year.get("dyn_pieces")
        auto_prev_dyn_hect_df = prev_for_year.get("dyn_hect")
        auto_prev_comparison_pieces = prev_for_year.get("comparison_values_pieces", {})
        auto_prev_comparison_hect = prev_for_year.get("comparison_values_hect", {})
        auto_prev_delta_pieces = prev_for_year.get("delta_values_pieces", {})
        auto_prev_delta_hect = prev_for_year.get("delta_values_hect", {})

    auto_prev_date_short = None
    if isinstance(prev_for_year, dict):
        auto_prev_date_short = prev_for_year.get("report_date_short")

    if not auto_prev_date_short:
        auto_prev_date_short = "00.00"

    # что показываем на листе "прошлый отчёт" (витрина)
    if use_baseline_prev:
        display_prev_pieces_df = baseline_prev_pieces
        display_prev_hectares_df = baseline_prev_hectares

        #  дата берётся из имени прикреплённого файла
        display_prev_date_short = extract_date_from_filename(baseline_prev_path)
    else:
        display_prev_pieces_df = auto_prev_pieces_df
        display_prev_hectares_df = auto_prev_hectares_df
        display_prev_date_short = auto_prev_date_short

    if use_baseline_prev:
        prev_sheet_name = f"отчет {display_prev_date_short} ({selected_year})"
    else:
        prev_sheet_name = "прошлый отчет"

    prev_sheet_name = re.sub(r'[:\\/*?\[\]]', "_", prev_sheet_name)[:31]

    # заранее длины прошлых таблиц (витрина)
    n1_prev = len(display_prev_pieces_df) if isinstance(display_prev_pieces_df, pd.DataFrame) else 0
    n2_prev = len(display_prev_hectares_df) if isinstance(display_prev_hectares_df, pd.DataFrame) else 0

    COMPARE_BLOCK_ROWS = 8
    GAP_AFTER_COMPARE = 2

    header_prev_row_1 = 1
    last_prev_data_row_1 = header_prev_row_1 + n1_prev

    title_prev_row_2 = last_prev_data_row_1 + 2 + COMPARE_BLOCK_ROWS + GAP_AFTER_COMPARE
    header_prev_row_2 = title_prev_row_2 + 1
    last_prev_data_row_2 = header_prev_row_2 + n2_prev

    # prev, который участвует в расчётах
    if use_baseline_prev:
        calc_prev_pieces_df = baseline_prev_pieces
        calc_prev_hectares_df = baseline_prev_hectares

        # динамика прошлого для расчётов не нужна
        calc_prev_dyn_pieces_df = pd.DataFrame()
        calc_prev_dyn_hect_df = pd.DataFrame()

        # ВАЖНО:  прошлого отчёта берём из baseline
        calc_prev_delta_pieces = calculate_delta_values(calc_prev_pieces_df)
        calc_prev_delta_hect = calculate_delta_values(calc_prev_hectares_df)

    else:
        # baseline НЕ выбран → сравнение выключено
        calc_prev_pieces_df = pd.DataFrame()
        calc_prev_hectares_df = pd.DataFrame()
        calc_prev_dyn_pieces_df = pd.DataFrame()
        calc_prev_dyn_hect_df = pd.DataFrame()
        calc_prev_delta_pieces = {}
        calc_prev_delta_hect = {}

    df_raw = df_input.copy()
    df_raw = normalize_excel_headers(df_raw)
    df_raw.columns = [norm(c) for c in df_raw.columns]
    cols = list(df_raw.columns)

    col_oiv = find_similar(COL_OIV, cols)
    col_contract = find_similar(COL_CONTRACT, cols)
    col_sapr = find_similar(COL_SAPR_DT, cols)
    col_agr = find_similar(COL_AGR_DT, cols)
    col_asu = find_similar(COL_ASU_DT, cols)
    col_asu_appr = find_similar(COL_ASU_APPROVE_DT, cols)
    col_ha = find_similar(COL_HA, cols)
    col_order = find_similar(COL_ORDER, cols)
    col_state = find_similar(COL_STATE, cols)
    col_status = find_similar(COL_STATUS, cols)
    col_field = find_similar(COL_FIELD_DT, cols)

    # подстраховка для № ген. договора
    if col_contract is None:
        for c in cols:
            nc = str(c)
            if ("догов" in nc and "ген" in nc) or ("номер" in nc and "догов" in nc):
                col_contract = c
                break
    if col_contract is None:
        for c in cols:
            nc = str(c)
            if "догов" in nc:
                col_contract = c
                break

    if col_oiv is None:
        raise ValueError("Не найдена колонка с ОИВ.")

    # убираем заказы с Р/К/В
    if col_order is not None and col_order in df_raw.columns:
        bad = df_raw[col_order].astype(str).str.contains(r"[РКВЮ]", case=False, na=False)
        df_raw = df_raw[~bad].copy()

    if col_state is not None and col_state in df_raw.columns:
        state_series = df_raw[col_state].astype(str).str.lower()

        allowed_mask = (
                state_series.str.contains("действ", na=False) |
                state_series.str.contains("приостанов", na=False)
        )

        df_raw = df_raw[allowed_mask].copy()

    # ОБЯЗАТЕЛЬНЫЕ ПОЛЯ - СОХРАНЯЕМ ИСХОДНУЮ ЛОГИКУ
    if col_sapr is None or col_sapr not in df_raw.columns:
        col_sapr = "missing_sapr";
        df_raw[col_sapr] = pd.NaT
    if col_agr is None or col_agr not in df_raw.columns:
        col_agr = "missing_agr";
        df_raw[col_agr] = pd.NaT
    if col_asu is None or col_asu not in df_raw.columns:
        col_asu = "missing_asu";
        df_raw[col_asu] = pd.NaT
    if col_asu_appr is None or col_asu_appr not in df_raw.columns:
        col_asu_appr = "missing_asu_appr";
        df_raw[col_asu_appr] = pd.NaT
    if col_ha is None or col_ha not in df_raw.columns:
        col_ha = "сумма объем заказа, га (auto)";
        df_raw[col_ha] = np.nan
    if col_status is None or col_status not in df_raw.columns:
        col_status = "missing_status";
        df_raw[col_status] = ""
    if col_field is None or col_field not in df_raw.columns:
        col_field = "missing_field_date";
        df_raw[col_field] = pd.NaT

    df = df_raw.copy()
    for c in [col_sapr, col_agr, col_asu, col_asu_appr, col_field]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df[col_ha] = pd.to_numeric(df[col_ha], errors="coerce")

    # нормализация ОИВ (объединение ДОНМ)
    def normalize_oiv_name(v):
        if not isinstance(v, str):
            return v
        v_norm = v.strip()
        if v_norm == "ДОНМ":
            return "Департамент образования и науки города Москвы (ДОНМ)"
        return v_norm

    df[col_oiv] = df[col_oiv].apply(normalize_oiv_name)

    def year_from_contract(s):
        if not isinstance(s, str):
            return None
        m = re.findall(r"(\d{2})(?!.*\d)", s)
        return 2000 + int(m[-1]) if m else None

    # фильтр по году
    if col_contract is not None and col_contract in df.columns:
        df["__year"] = df[col_contract].apply(year_from_contract)
        dfy = df[df["__year"] == selected_year].copy()
        if dfy.empty:
            dfy = df.iloc[0:0].copy()
    else:
        df["__year"] = selected_year
        dfy = df.copy()

    oiv_order = pd.Index(dfy[col_oiv].dropna().unique())

    # скорректированный план по БД
    corr_plan_pieces = dfy.groupby(col_oiv).size() if not dfy.empty else pd.Series(dtype=int)
    if not dfy.empty:
        corr_plan_hect = dfy.groupby(col_oiv)[col_ha].sum(min_count=1)
    else:
        corr_plan_hect = pd.Series(dtype=float)

    # период дат
    date_candidates = pd.concat(
        [dfy[col_sapr], dfy[col_agr], dfy[col_asu], dfy[col_asu_appr], dfy[col_field]],
        axis=0, ignore_index=True
    ).dropna()
    if not date_candidates.empty:
        date_from = date_candidates.min().normalize()
    else:
        date_from = pd.to_datetime(f"{selected_year}-01-01")
    today = pd.to_datetime(dt.date.today()).normalize()
    date_to = min(today, pd.to_datetime(f"{selected_year}-12-31"))

    def filter_period(dfx):
        dtmp = dfx.copy()
        for c in [col_sapr, col_agr, col_asu, col_asu_appr, col_field]:
            mask = (dtmp[c] >= date_from) & (dtmp[c] <= date_to)
            dtmp.loc[~mask, c] = pd.NaT
        return dtmp

    df_today = filter_period(dfy)

    # расчёт таблиц
    pieces_today = build_pieces_block(
        df_today, oiv_order, plan_pieces, corr_plan_pieces, selected_year,
        col_oiv, col_sapr, col_agr, col_asu, col_asu_appr,
        col_field, col_ha, col_status
    )
    hectares_today = build_hectares_block(
        df_today, oiv_order, plan_hect, corr_plan_hect, selected_year,
        col_oiv, col_sapr, col_agr, col_asu, col_asu_appr,
        col_field, col_ha, col_status
    )
    hectares_display = make_hectares_display_df(hectares_today, selected_year)

    # превью для Streamlit
    preview_pieces = pieces_today.copy()
    preview_hectares = hectares_display.copy()

    # РАСЧЁТ ЗНАЧЕНИЙ ДЛЯ СРАВНЕНИЙ
    plan_col1 = f"Утвержденный график {selected_year} года"
    plan_col3 = f"Всего план (скорректированный график {selected_year} года)"

    # Получаем итоговые строки для текущего отчёта
    current_totals_pieces = pieces_today[pieces_today["ОИВ"] == "Итого:"].copy()
    current_totals_hect = hectares_today[hectares_today["ОИВ"] == "Итого:"].copy()

    # Получаем итоговые строки из прошлого отчёта
    prev_totals_pieces = pd.DataFrame()
    prev_totals_hect = pd.DataFrame()

    if isinstance(calc_prev_pieces_df, pd.DataFrame) and not calc_prev_pieces_df.empty:
        prev_totals_pieces = calc_prev_pieces_df[
            calc_prev_pieces_df["ОИВ"] == "Итого:"
            ].copy()

    if isinstance(calc_prev_hectares_df, pd.DataFrame) and not calc_prev_hectares_df.empty:
        prev_totals_hect = calc_prev_hectares_df[
            calc_prev_hectares_df["ОИВ"] == "Итого:"
            ].copy()

    # Рассчитываем изменения для текущего отчёта
    comparison_values_pieces = calculate_comparison_values(
        current_totals_pieces,
        prev_totals_pieces,
        plan_col1,
        plan_col3
    )

    comparison_values_hect = calculate_comparison_values(
        current_totals_hect,
        prev_totals_hect,
        plan_col1,
        plan_col3
    )

    if not use_baseline_prev:
        comparison_values_pieces = {k: 0 for k in comparison_values_pieces}
        comparison_values_hect = {k: 0 for k in comparison_values_hect}

    # Рассчитываем значения ▲ (стрелочек) для текущего отчёт
    delta_values_pieces_curr = calculate_delta_values(pieces_today)
    delta_values_hect_curr = calculate_delta_values(hectares_today)

    # динамика по отчётам
    def build_dynamics_by_reports(today_df, prev_df, is_hect=False):
        t = today_df.set_index("ОИВ")
        if prev_df is None or prev_df.empty:
            p = pd.DataFrame(0, index=t.index, columns=t.columns)
        else:
            p = prev_df.set_index("ОИВ").reindex(t.index).fillna(0)

        col_plan = f"Всего план (скорректированный график {selected_year} года)"

        # список метрик
        metrics = [
            ("Изменение плана скорректированного", col_plan),
            ("Выполнено полевое обследование", "Выполнено полевое обследование"),
            ("загружено в САПР", "Загружено в САПР"),
            ("согласовано в САПР", "Согласовано в САПР"),
            ("загружено в АСУ ОДС", "Загружено в АСУ ОДС"),
            ("не утверждено в АСУ ОДС", "Не утверждено БД"),
            ("отклонено", "Отклонено"),
            ("утверждено", "Утверждено"),
        ]

        dyn = pd.DataFrame(index=t.index)
        for out_name, src_col in metrics:
            dyn[out_name] = t[src_col] - p[src_col]

        if is_hect:
            dyn = dyn.round(2)
        else:
            dyn = dyn.astype(int)

        return dyn.reset_index()

    dyn_pieces = build_dynamics_by_reports(pieces_today, calc_prev_pieces_df, is_hect=False)
    dyn_hect = build_dynamics_by_reports(hectares_today, calc_prev_hectares_df, is_hect=True)

    if not use_baseline_prev:
        for df_dyn in (dyn_pieces, dyn_hect):
            if isinstance(df_dyn, pd.DataFrame) and not df_dyn.empty:
                for c in df_dyn.columns:
                    if norm(str(c)) != norm("ОИВ"):
                        df_dyn[c] = 0

    if preview:
        return {
            "pieces": pieces_today,
            "hectares": hectares_display,
            "dyn_pieces": dyn_pieces,
            "dyn_hect": dyn_hect,
        }

    # подготовка Excel
    now = dt.datetime.now()
    timestamp = now.strftime("%d.%m.%Y_%H-%M")
    filename = f"отчет по ОИВ {timestamp}.xlsx"
    header_display = build_header_display_map(selected_year)

    gray_headers = {
        "Выполнено полевое обследование",
        "Загружено в САПР",
        "Согласовано в САПР",
        "Загружено в АСУ ОДС",
        "Отклонено",
    }
    lightgreen_headers = {
        "Выполнено полевое обследование (от плана) %",
        "Загружено в САПР %",
        "Согласовано в САПР (от плана) %",
        "Загружено в АСУ ОДС (от плана) %",
        "Не Загружено в АСУ ОДС",
        "Не утверждено БД",
        "Не утверждено БД (от плана) %",
        "Отклонено (от плана) %",
        "Утверждено (от плана ) %",
    }

    # Streamlit: пишем в память, локально: пишем на диск
    if return_excel_bytes:
        output = io.BytesIO()
        writer_target = output
    else:
        out_name = make_output_path(filename)
        writer_target = out_name

    with pd.ExcelWriter(writer_target, engine="xlsxwriter") as writer:
        sheet = "Отчёт"

        n1 = len(pieces_today)
        n2 = len(hectares_today)

        title_row_1 = 0
        header_row_1 = 1
        first_data_row_1 = header_row_1 + 1
        last_data_row_1 = header_row_1 + n1

        COMPARE_BLOCK_ROWS = 8
        GAP_AFTER_COMPARE = 2

        title_row_2 = last_data_row_1 + 2 + COMPARE_BLOCK_ROWS + GAP_AFTER_COMPARE
        header_row_2 = title_row_2 + 1
        first_data_row_2 = header_row_2 + 1
        last_data_row_2 = header_row_2 + n2

        pieces_today.to_excel(writer, sheet_name=sheet, index=False, startrow=header_row_1)
        hectares_display.to_excel(writer, sheet_name=sheet, index=False, startrow=header_row_2)

        workbook = writer.book
        worksheet = writer.sheets[sheet]
        worksheet.activate()

        header_default_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center", "bg_color": "#C6E0B4",
            "font_name": "Times New Roman",
        })
        header_oiv_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center",
        })
        header_plan_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center", "bg_color": "#BDD7EE", "font_color": "#FF0000",
        })
        header_approved_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center", "bg_color": "#BDD7EE",
        })
        header_gray_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center", "bg_color": "#D9D9D9",
        })
        header_lightgreen_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "vcenter",
            "align": "center", "bg_color": "#E2EFDA",
        })

        def get_header_fmt(col_name: str):
            if col_name in {plan_col1, plan_col3}:
                return header_plan_fmt
            if col_name == "Утверждено":
                return header_approved_fmt
            if col_name in gray_headers:
                return header_gray_fmt
            if col_name in lightgreen_headers:
                return header_lightgreen_fmt
            return header_default_fmt

        # Форматы строки ИТОГО для основных таблиц
        ito_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#FFFFFF",
            "font_color": "#FF0000",
            "align": "center",
            "valign": "vcenter",
            "font_name": "Times New Roman",
            "font_size": 12,
        })

        total_blue_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#DDEBF7",
            "align": "center",
            "valign": "vcenter",
            "font_name": "Times New Roman",
            "font_size": 12,
        })

        total_blue_percent_fmt: object = workbook.add_format({
            "bold": True,
            "bg_color": "#DDEBF7",
            "align": "center",
            "valign": "vcenter",
            "font_name": "Times New Roman",
            "font_size": 12,
            'num_format': '0.0"%"',
        })


        percent_fmt = workbook.add_format({
            "left": 1,
            "top": 1,
            "bottom": 1,
            "align": "center",
            "valign": "vcenter",
            'num_format': '0.0"%"',
        })

        percent_lastcol_fmt = workbook.add_format({
            "left": 1,
            "top": 1,
            "bottom": 1,
            "right": 2,
            "align": "center",
            "valign": "vcenter",
            'num_format': '0.0"%"',
        })
        title_fmt = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#FFFF00", "border": 1,
        })
        data_center_fmt = workbook.add_format({
            "border": 1,
            "align": "center", "valign": "vcenter",
        })
        format_hectares = workbook.add_format({
            "border": 1,
            "align": "center", "valign": "vcenter",
            "num_format": "0.00",
        })
        dyn_total_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F8CBAD",
            "border": 1,
            "bold": True,
        })
        dyn_header_fmt = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "bg_color": "#F8CBAD",
            "border": 1,
        })
        dyn_cell_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F8CBAD",
            "border": 1,
        })

        dyn_total_cell_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F8CBAD",
            "border": 1,
            "bold": True,
        })
        dyn_blank_cell_fmt = workbook.add_format({
            "align": "center", "valign": "vcenter",
            "bg_color": "#FCE4D6",
            "border": 1,
        })
        dyn_title_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#F8CBAD",
            "border": 1,
            "top": 2,
            "left": 2,
            "right": 2,
            "bold": True,
        })
        #   форматы для рамки и ИТОГО
        dyn_hdr_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "top": 2,
            "bottom": 2,
            "left": 2,
            "right": 2,
            "align": "center",
            "valign": "vcenter",
            "bold": True,
            "text_wrap": True
        })

        dyn_cell_mid_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        dyn_cell_left_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "left": 2,
            "align": "center",
            "valign": "vcenter"
        })

        dyn_cell_right_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "right": 2,
            "align": "center",
            "valign": "vcenter"
        })

        dyn_total_mid_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "top": 2,
            "bottom": 2,
            "align": "center",
            "valign": "vcenter",
            "bold": True
        })

        dyn_total_left_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "top": 2,
            "bottom": 2,
            "left": 2,
            "align": "center",
            "valign": "vcenter",
            "bold": True
        })

        dyn_total_right_fmt = workbook.add_format({
            "bg_color": "#F8CBAD",
            "border": 1,
            "top": 2,
            "bottom": 2,
            "right": 2,
            "align": "center",
            "valign": "vcenter",
            "bold": True
        })

        compare_label_fmt = workbook.add_format({
            "align": "left",
            "valign": "vcenter",
        })
        compare_delta_label_fmt = workbook.add_format({
            "align": "left",
            "valign": "vcenter",
            "font_color": "#0070C0",
        })
        compare_delta_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "font_color": "#0070C0",
        })
        center_only_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
        })
        #  жирная рамка только для “числовых” ячеек под таблицей
        percent_thick_fmt = workbook.add_format({
            "border": 2,
            "align": "center",
            "valign": "vcenter",
            'num_format': '0.0"%"',
        })

        compare_delta_thick_fmt = workbook.add_format({
            "border": 2,
            "align": "center",
            "valign": "vcenter",
            "font_color": "#0070C0",
        })
        header_edges_fmt = workbook.add_format({
            "border": 1,
            "bottom": 2,

        })
        oiv_prev_col_fmt = workbook.add_format({
            "border": 1,
            "left": 2,
        })
        header_bottom_prev_fmt = workbook.add_format({
            "bottom": 2,
        })
        header_under_prev_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
        })
        thick_top_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
        })
        thick_bottom_fmt = workbook.add_format({
            "border": 1,
            "bottom": 2,
        })
        thick_left_fmt = workbook.add_format({
            "border": 1,
            "left": 2,
        })
        thick_right_fmt = workbook.add_format({
            "border": 1,
            "right": 2,
        })
        hdr_thick_vlines_fmt = workbook.add_format({"right": 2})
        outer_thick_right_fmt = workbook.add_format({"right": 2})
        corner_tl_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "left": 2,
        })
        corner_tr_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "right": 2,
        })
        corner_bl_fmt = workbook.add_format({
            "border": 1,
            "bottom": 2,
            "left": 2,
        })
        corner_br_fmt = workbook.add_format({
            "border": 1,
            "bottom": 2,
            "right": 2,
        })
        header_thick_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "bottom": 2,
            "left": 2,
            "right": 2
        })

        percent_display_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            'num_format': '0.0"%"'
        })
        worksheet.write(
            title_row_1, 0,
            f"ГЗ {selected_year}, шт", title_fmt
        )
        worksheet.write(
            title_row_2, 0,
            f"ГЗ {selected_year}, га", title_fmt
        )

        # шапки
        for col_num, value in enumerate(pieces_today.columns.values):
            disp = header_display.get(value, value)
            disp = wrap_long_header(disp)
            fmt = get_header_fmt(value)
            worksheet.write(header_row_1, col_num, disp, fmt)

        for col_num, value in enumerate(hectares_today.columns.values):
            disp = header_display.get(value, value)
            disp = wrap_long_header(disp)
            fmt = get_header_fmt(value)
            worksheet.write(header_row_2, col_num, disp, fmt)

        worksheet.set_row(header_row_1, 100)
        worksheet.set_row(header_row_2, 100)

        # ширины по заголовкам/данным (ШТ и ГА)
        set_columns_by_header(worksheet, pieces_today, min_width=25, max_width=60, padding=4)
        set_columns_by_header(worksheet, hectares_display, min_width=16, max_width=55, padding=3)

        for row_idx in range(first_data_row_1, last_data_row_1 + 1):
            worksheet.set_row(row_idx, 18)
        for row_idx in range(first_data_row_2, last_data_row_2 + 1):
            worksheet.set_row(row_idx, 18)

        worksheet.set_row(last_data_row_1, 20)
        worksheet.set_row(last_data_row_2, 20)

        percent_cols = [
            "Выполнено полевое обследование (от плана) %",
            "Загружено в САПР %",
            "Согласовано в САПР (от плана) %",
            "Загружено в АСУ ОДС (от плана) %",
            "Отклонено (от плана) %",
            "Не утверждено БД (от плана) %",
            "Утверждено (от плана ) %",
        ]
        percent_col_width = 14

        # Только ширина, без формата (рамок)
        for name in percent_cols:
            if name in pieces_today.columns:
                ci = pieces_today.columns.get_loc(name)
                worksheet.set_column(ci, ci, percent_col_width)
            if name in hectares_today.columns:
                ci = hectares_today.columns.get_loc(name)
                worksheet.set_column(ci, ci, percent_col_width)

        NUM_W = 16  # ширина для количественных колонок

        for i, col_name in enumerate(pieces_today.columns):
            if col_name == "ОИВ":
                continue
            if col_name in percent_cols:
                continue
            worksheet.set_column(i, i, NUM_W,center_only_fmt)

        for i, col_name in enumerate(hectares_display.columns):
            if col_name == "ОИВ":
                continue
            if col_name in percent_cols:
                continue
            worksheet.set_column(i, i, NUM_W,center_only_fmt)

        def write_percent_formulas(worksheet_obj, df_local, first_data_row, last_data_row_exclusive):
            if df_local.empty:
                return

            plan_col = f"Всего план (скорректированный график {selected_year} года)"

            specs = [
                ("Выполнено полевое обследование (от плана) %",
                 "Выполнено полевое обследование", plan_col),

                ("Загружено в САПР %",
                 "Загружено в САПР", plan_col),

                ("Согласовано в САПР (от плана) %",
                 "Согласовано в САПР", plan_col),

                ("Загружено в АСУ ОДС (от плана) %",
                 "Загружено в АСУ ОДС", plan_col),

                ("Отклонено (от плана) %",
                 "Отклонено", "Загружено в АСУ ОДС"),

                ("Не утверждено БД (от плана) %",
                 "Не утверждено БД", "Загружено в АСУ ОДС"),

                ("Утверждено (от плана ) %",
                 "Утверждено", "Загружено в АСУ ОДС"),
            ]

            col_idx_map = {name: df_local.columns.get_loc(name) for name in df_local.columns}

            for pct_col, num_col, den_col in specs:
                if pct_col not in df_local.columns or num_col not in df_local.columns or den_col not in df_local.columns:
                    continue

                pct_c = col_idx_map[pct_col]
                num_c = col_idx_map[num_col]
                den_c = col_idx_map[den_col]

                for r in range(first_data_row, last_data_row_exclusive):
                    num_cell = xl_rowcol_to_cell(r, num_c)
                    den_cell = xl_rowcol_to_cell(r, den_c)

                    formula = f"=IF({den_cell}=0,0,ROUND(100*{num_cell}/{den_cell},1))"
                    worksheet_obj.write_formula(r, pct_c, formula, percent_display_fmt)

        write_percent_formulas(
            worksheet,
            pieces_today,
            first_data_row_1,
            last_data_row_1
        )
        write_percent_formulas(
            worksheet,
            hectares_today,
            first_data_row_2,
            last_data_row_2
        )

        # Итого
        for col_idx, col_name in enumerate(pieces_today.columns):
            value = pieces_today.iloc[-1, col_idx]
            if col_idx == 0:
                fmt = ito_fmt
            else:
                fmt = total_blue_percent_fmt if col_name in percent_cols else total_blue_fmt
            worksheet.write(last_data_row_1, col_idx, value, fmt)

        for col_idx, col_name in enumerate(hectares_display.columns):
            value = hectares_display.iloc[-1, col_idx]
            if col_idx == 0:
                fmt = ito_fmt
            else:
                fmt = total_blue_percent_fmt if col_name in percent_cols else total_blue_fmt
            worksheet.write(last_data_row_2, col_idx, value, fmt)

        #  тонкая сетка + ОИВ + ИТОГО + ВНЕШНЯЯ РАМКА

        thin_grid_mid_fmt = workbook.add_format({
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })

        thin_grid_left_oiv_sep_fmt = workbook.add_format({
            "border": 1,
            "left": 2,
            "right": 2,
            "align": "left",
            "valign": "vcenter",
        })

        # последняя колонка: жирная правая рамка + центрирование
        thin_grid_right_outline_fmt = workbook.add_format({
            "border": 1,
            "right": 2,
            "align": "center",
            "valign": "vcenter",
        })

        # строка ИТОГО
        thin_grid_total_mid_bottom_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "bottom": 2,
            "align": "center",
            "valign": "vcenter",
        })

        # строка ИТОГО - первый столбец (ОИВ)
        thin_grid_total_left_oiv_sep_bottom_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "bottom": 2,
            "left": 2,
            "right": 2,
            "align": "left",
            "valign": "vcenter",
        })

        # строка ИТОГО - последняя колонка + центрирование
        thin_grid_total_right_outline_bottom_fmt = workbook.add_format({
            "border": 1,
            "top": 2,
            "bottom": 2,
            "right": 2,
            "align": "center",
            "valign": "vcenter",
        })
        center_only_fmt = workbook.add_format({
            "align": "center",
            "valign": "vcenter",
        })
        thick_cell_fmt = workbook.add_format({"border": 2})

        # ШАПКА (ТЕКУЩИЙ ОТЧЁТ)

        if n1 > 0:
            # жирная линия под шапкой
            worksheet.conditional_format(
                header_row_1, 0,
                header_row_1, pieces_today.shape[1] - 1,
                {"type": "formula", "criteria": "=TRUE", "format": header_thick_fmt}
            )
            # жирная вертикаль после ОИВ
            worksheet.conditional_format(
                header_row_1, 0, header_row_1, 0,
                {"type": "formula", "criteria": "=TRUE", "format": thin_grid_left_oiv_sep_fmt}
            )
            # правая жирная рамка на последней колонке
            last_col_1 = pieces_today.shape[1] - 1
            worksheet.conditional_format(
                header_row_1, last_col_1, header_row_1, last_col_1,
                {"type": "formula", "criteria": "=TRUE", "format": thin_grid_right_outline_fmt}
            )

        if n2 > 0:
            worksheet.conditional_format(
                header_row_2, 0,
                header_row_2, hectares_display.shape[1] - 1,
                {"type": "formula", "criteria": "=TRUE", "format": header_thick_fmt}
            )
            worksheet.conditional_format(
                header_row_2, 0, header_row_2, 0,
                {"type": "formula", "criteria": "=TRUE", "format": thin_grid_left_oiv_sep_fmt}
            )
            last_col_2 = hectares_display.shape[1] - 1
            worksheet.conditional_format(
                header_row_2, last_col_2, header_row_2, last_col_2,
                {"type": "formula", "criteria": "=TRUE", "format": thin_grid_right_outline_fmt}
            )

        def apply_main_grid_with_outline(ws, first_data_row, df_rows, ncols):
            """
            Основная таблица:
              - внутри: тонкая сетка
              - col=0 (ОИВ): жирная рамка слева + жирный разделитель справа
              - последняя колонка: жирная правая рамка
              - строка ИТОГО: жирная линия сверху + жирная нижняя рамка
            """
            if df_rows <= 0 or ncols <= 0:
                return

            total_row = first_data_row + df_rows - 1
            last_col = ncols - 1

            # тело таблицы (ВСЁ КРОМЕ ИТОГО)
            if total_row - 1 >= first_data_row:

                # ОИВ
                ws.conditional_format(
                    first_data_row, 0,
                    total_row - 1, 0,
                    {"type": "no_errors", "format": thin_grid_left_oiv_sep_fmt}
                )

                # середина
                if last_col >= 2:
                    ws.conditional_format(
                        first_data_row, 1,
                        total_row - 1, last_col - 1,
                        {"type": "no_errors", "format": thin_grid_mid_fmt}
                    )

                # последняя колонка
                if last_col >= 1:
                    ws.conditional_format(
                        first_data_row, last_col,
                        total_row - 1, last_col,
                        {"type": "no_errors", "format": thin_grid_right_outline_fmt}
                    )

            #  строка ИТОГО
            # ОИВ
            ws.conditional_format(
                total_row, 0,
                total_row, 0,
                {"type": "no_errors", "format": thin_grid_total_left_oiv_sep_bottom_fmt}
            )

            # середина
            if last_col >= 2:
                ws.conditional_format(
                    total_row, 1,
                    total_row, last_col - 1,
                    {"type": "no_errors", "format": thin_grid_total_mid_bottom_fmt}
                )

            # последняя колонка
            if last_col >= 1:
                ws.conditional_format(
                    total_row, last_col,
                    total_row, last_col,
                    {"type": "no_errors", "format": thin_grid_total_right_outline_bottom_fmt}
                )

        #ЛИНИИ ОСНОВНОЙ ТАБЛИЦЫ (ТЕКУЩИЙ ОТЧЁТ)
        if n1 > 0:
            apply_main_grid_with_outline(worksheet, first_data_row_1, pieces_today.shape[0], pieces_today.shape[1])

        if n2 > 0:
            apply_main_grid_with_outline(worksheet, first_data_row_2, hectares_display.shape[0],
                                         hectares_display.shape[1])

        #  БЛОК СРАВНЕНИЯ
        def write_compare_block(worksheet_obj, df_local, last_data_row):
            cols_local = list(df_local.columns)
            needed = [
                "Загружено в САПР",
                "Согласовано в САПР",
                "Загружено в АСУ ОДС",
                "Утверждено",
                "Отклонено",
                "Не утверждено БД",
            ]
            if any(name not in cols_local for name in needed):
                return None

            idx_zagr_sapr = cols_local.index("Загружено в САПР")
            idx_sogl_sapr = cols_local.index("Согласовано в САПР")
            idx_zagr_asu = cols_local.index("Загружено в АСУ ОДС")
            idx_utv = cols_local.index("Утверждено")
            idx_otk = cols_local.index("Отклонено")
            idx_neutv = cols_local.index("Не утверждено БД")

            total_row = last_data_row

            cell_zagr_sapr = xl_rowcol_to_cell(total_row, idx_zagr_sapr)
            cell_sogl_sapr = xl_rowcol_to_cell(total_row, idx_sogl_sapr)
            cell_zagr_asu = xl_rowcol_to_cell(total_row, idx_zagr_asu)
            cell_utv = xl_rowcol_to_cell(total_row, idx_utv)
            cell_otk = xl_rowcol_to_cell(total_row, idx_otk)
            cell_neutv = xl_rowcol_to_cell(total_row, idx_neutv)

            start_row = last_data_row + 2
            label_col = 0

            # % от загруженного в САПР
            row1 = start_row
            worksheet_obj.write(row1, label_col, "% от загруженного в САПР", compare_label_fmt)
            formula_1 = f"=IF({cell_zagr_sapr}=0,0,100*{cell_sogl_sapr}/{cell_zagr_sapr})"
            worksheet_obj.write_formula(row1, idx_sogl_sapr, formula_1,  percent_thick_fmt)

            # ▲ от загруженного в САПР
            row2 = start_row + 1
            worksheet_obj.write(row2, label_col, "▲ от загруженного в САПР", compare_delta_label_fmt)
            formula_2 = f"={cell_sogl_sapr}-{cell_zagr_sapr}"
            worksheet_obj.write_formula(row2, idx_sogl_sapr, formula_2, compare_delta_thick_fmt)

            # % от согласованного в САПР
            row3 = start_row + 2
            worksheet_obj.write(row3, label_col, "% от согласованного в САПР", compare_label_fmt)
            formula_3 = f"=IF({cell_sogl_sapr}=0,0,100*{cell_zagr_asu}/{cell_sogl_sapr})"
            worksheet_obj.write_formula(row3, idx_zagr_asu, formula_3, percent_thick_fmt)

            # ▲ от согласованного в САПР
            row4 = start_row + 3
            worksheet_obj.write(row4, label_col, "▲ от согласованного в САПР", compare_delta_label_fmt)
            formula_4 = f"={cell_zagr_asu}-{cell_sogl_sapr}"
            worksheet_obj.write_formula(row4, idx_zagr_asu, formula_4, compare_delta_thick_fmt)

            # % от загруженного в АСУ ОДС
            row5 = start_row + 4
            worksheet_obj.write(row5, label_col, "% от загруженного в АСУ ОДС", compare_label_fmt)
            formula_5 = f"=IF({cell_zagr_asu}=0,0,100*{cell_utv}/{cell_zagr_asu})"
            worksheet_obj.write_formula(row5, idx_utv, formula_5, percent_thick_fmt)

            # ▲ от загруженного в АСУ ОДС
            row6 = start_row + 5
            worksheet_obj.write(row6, label_col, "▲ от загруженного в АСУ ОДС", compare_delta_label_fmt)
            formula_6 = f"={cell_utv}-{cell_zagr_asu}"
            worksheet_obj.write_formula(row6, idx_utv, formula_6, compare_delta_thick_fmt)

            # (Отклонено + Не утв. БД + Утв) – Загружено в АСУ ОДС
            row7 = start_row + 6
            formula_7 = f"={cell_otk}+{cell_neutv}+{cell_utv}-{cell_zagr_asu}"
            worksheet_obj.write_formula(row7, idx_utv, formula_7, compare_delta_fmt)

            row8 = start_row + 7
            worksheet_obj.write(row8, label_col, "изменение с отчетом за предыдущую неделю", dyn_cell_fmt)

            return start_row

        start_row_compare_1 = write_compare_block(worksheet, pieces_today, last_data_row_1)
        start_row_compare_2 = write_compare_block(worksheet, hectares_today, last_data_row_2)

        #  ТЕКУЩИЙ ОТЧЁТ:  "изменение с отчетом за предыдущую неделю"
        write_comparison_values_row(
            worksheet,
            pieces_today,
            last_data_row_1,
            start_row_compare_1,
            comparison_values_pieces,
            plan_col1,
            plan_col3,
            dyn_cell_fmt
        )

        write_comparison_values_row(
            worksheet,
            hectares_today,
            last_data_row_2,
            start_row_compare_2,
            comparison_values_hect,
            plan_col1,
            plan_col3,
            dyn_cell_fmt
        )

        #  Персиковая таблица (динамика) справа от блоков
        dyn_cols = [
            "Изменение плана скорректированного",
            "Выполнено полевое обследование",
            "загружено в САПР",
            "согласовано в САПР",
            "загружено в АСУ ОДС",
            "отклонено",
            "не утверждено в АСУ ОДС",
            "утверждено",
        ]

        #  Ш Т У К И (ШТ)

        start_col_dyn1 = pieces_today.shape[1] + 2

        # Заголовок над персиковой таблицей (ШТ)
        worksheet.merge_range(
            title_row_1,
            start_col_dyn1,
            title_row_1,
            start_col_dyn1 + len(dyn_cols) - 1,
            "Динамика за неделю",
            dyn_title_fmt,
        )

        # Шапка персиковой таблицы (ШТ)
        for j, col_name in enumerate(dyn_cols):
            worksheet.write(
                header_row_1,
                start_col_dyn1 + j,
                col_name,
                dyn_hdr_fmt,
            )
            worksheet.set_column(start_col_dyn1 + j, start_col_dyn1 + j, 14)

        # Данные персиковой таблицы (ШТ) — с толстой рамкой по краям и жирным ИТОГО
        for i in range(n1):
            r = first_data_row_1 + i
            is_total = (i == n1 - 1)
            last_j = len(dyn_cols) - 1

            for j, col_name in enumerate(dyn_cols):
                c = start_col_dyn1 + j
                val = dyn_pieces.iloc[i][col_name]

                if is_total:
                    if j == 0:
                        fmt = dyn_total_left_fmt
                    elif j == last_j:
                        fmt = dyn_total_right_fmt
                    else:
                        fmt = dyn_total_mid_fmt
                else:
                    if j == 0:
                        fmt = dyn_cell_left_fmt
                    elif j == last_j:
                        fmt = dyn_cell_right_fmt
                    else:
                        fmt = dyn_cell_mid_fmt

                worksheet.write(r, c, val, fmt)

        #  Г Е К Т А Р Ы (ГА)

        start_col_dyn2 = hectares_today.shape[1] + 2

        # Заголовок над персиковой таблицей (ГА)
        worksheet.merge_range(
            title_row_2,
            start_col_dyn2,
            title_row_2,
            start_col_dyn2 + len(dyn_cols) - 1,
            "Динамика за неделю",
            dyn_title_fmt,
        )

        # Шапка персиковой таблицы (ГА)
        for j, col_name in enumerate(dyn_cols):
            worksheet.write(
                header_row_2,
                start_col_dyn2 + j,
                col_name,
                dyn_hdr_fmt,
            )
            worksheet.set_column(start_col_dyn2 + j, start_col_dyn2 + j, 14)

        # Данные персиковой таблицы (ГА) - С ТОЛСТОЙ РАМКОЙ ПО КРАЯМ И ЖИРНЫМ ИТОГО (как в ШТ)
        for i in range(n2):
            r = first_data_row_2 + i
            is_total = (i == n2 - 1)
            last_j = len(dyn_cols) - 1

            for j, col_name in enumerate(dyn_cols):
                c = start_col_dyn2 + j
                val = dyn_hect.iloc[i][col_name]

                if is_total:
                    if j == 0:
                        fmt = dyn_total_left_fmt
                    elif j == last_j:
                        fmt = dyn_total_right_fmt
                    else:
                        fmt = dyn_total_mid_fmt
                else:
                    if j == 0:
                        fmt = dyn_cell_left_fmt
                    elif j == last_j:
                        fmt = dyn_cell_right_fmt
                    else:
                        fmt = dyn_cell_mid_fmt

                worksheet.write(r, c, val, fmt)

        #  НИЖНЯЯ ПЕРСИКОВАЯ СТРОКА: ИТОГО_новый - ИТОГО_старый
        def write_bottom_diff_row(df_local, last_data_row, prev_total_row, start_row_compare):
            if prev_total_row <= 1:
                return
            if start_row_compare is None:
                return

            target_row = start_row_compare + 7  # это row8 в блоке сравнения

            cols_to_use = [
                plan_col1,
                plan_col3,
                "Выполнено полевое обследование",
                "Загружено в САПР",
                "Согласовано в САПР",
                "Загружено в АСУ ОДС",
                "Не Загружено в АСУ ОДС",
                "Отклонено",
                "Не утверждено БД",
                "Утверждено",
            ]

            # подпись слева уже есть, но на всякий случай перезапишем
            worksheet.write(target_row, 0, "изменение с отчетом за предыдущую неделю", dyn_cell_fmt)

            for col_name in cols_to_use:
                if col_name not in df_local.columns:
                    continue
                idx = df_local.columns.get_loc(col_name)
                cell_new = xl_rowcol_to_cell(last_data_row, idx)
                cell_prev = f"'{prev_sheet_name}'!{xl_rowcol_to_cell(prev_total_row, idx)}"
                formula = f"={cell_new}-{cell_prev}"
                worksheet.write_formula(target_row, idx, formula, dyn_cell_fmt)

        # ДЛЯ ТАБЛИЦЫ ШТ
        if use_baseline_prev:
            write_bottom_diff_row(
                pieces_today,
                last_data_row_1,
                last_prev_data_row_1,
                start_row_compare_1,
            )

        # ДЛЯ ТАБЛИЦЫ ГА
        if use_baseline_prev:
            write_bottom_diff_row(
                hectares_today,
                last_data_row_2,
                last_prev_data_row_2,
                start_row_compare_2,
            )

        #  ЯЧЕЙКИ ▲ ПОД ПЕРСИКОВОЙ ТАБЛИЦЕЙ
        if use_baseline_prev:
            prev_delta_pieces_for_calc = calc_prev_delta_pieces
            prev_delta_hect_for_calc = calc_prev_delta_hect
        else:
            prev_delta_pieces_for_calc = {}
            prev_delta_hect_for_calc = {}

        # Для штук
        write_delta_comparison_cells(
            worksheet, pieces_today, last_data_row_1, start_row_compare_1,
            delta_values_pieces_curr, prev_delta_pieces_for_calc,
            dyn_cols, start_col_dyn1, dyn_cell_fmt, is_prev_report=False
        )

        # Для гектаров
        write_delta_comparison_cells(
            worksheet, hectares_today, last_data_row_2, start_row_compare_2,
            delta_values_hect_curr, prev_delta_hect_for_calc,
            dyn_cols, start_col_dyn2, dyn_cell_fmt, is_prev_report=False
        )

        #  ДОБАВЛЯЕМ ЗЕЛЁНЫЕ ЧИСЛА ПОД ПЕРСИКОВОЙ ТАБЛИЦЕЙ
        # Для ШТ
        green_row_pieces = add_green_numbers_under_persic(
            worksheet=worksheet,
            workbook=workbook,
            df_main=pieces_today,
            last_data_row=last_data_row_1,
            start_row_compare=start_row_compare_1,
            start_col_dyn=start_col_dyn1,
            dyn_cols=dyn_cols,
            selected_year=selected_year,
            is_hectares=False
        )

        # Для ГА
        green_row_hect = add_green_numbers_under_persic(
            worksheet=worksheet,
            workbook=workbook,
            df_main=hectares_today,
            last_data_row=last_data_row_2,
            start_row_compare=start_row_compare_2,
            start_col_dyn=start_col_dyn2,
            dyn_cols=dyn_cols,
            selected_year=selected_year,
            is_hectares=True
        )

        #  Заполнить пустые ячейки персиковой сеткой (текущий отчёт)
        if n1 > 0:
            worksheet.conditional_format(
                first_data_row_1, start_col_dyn1,
                last_data_row_1, start_col_dyn1 + len(dyn_cols) - 1,
                {"type": "blanks", "format": dyn_blank_cell_fmt}
            )

        if n2 > 0:
            worksheet.conditional_format(
                first_data_row_2, start_col_dyn2,
                last_data_row_2, start_col_dyn2 + len(dyn_cols) - 1,
                {"type": "blanks", "format": dyn_blank_cell_fmt}
            )

        #  ЗЕЛЁНАЯ ПРОКРАСКА ПРИ 100%
        green_fill_fmt = workbook.add_format({
            "bg_color": "#00B050",
        })

        #  зелёная заливка по строкам (ШТ)
        if not pieces_today.empty:
            col_oiv_idx = pieces_today.columns.get_loc("ОИВ")
            col_plan1_idx = pieces_today.columns.get_loc(plan_col1)
            col_plan3_idx = pieces_today.columns.get_loc(plan_col3)
            col_pole_pct_idx = pieces_today.columns.get_loc("Выполнено полевое обследование (от плана) %")
            col_sapr_pct_idx = pieces_today.columns.get_loc("Загружено в САПР %")
            col_agr_pct_idx = pieces_today.columns.get_loc("Согласовано в САПР (от плана) %")
            col_asu_pct_idx = pieces_today.columns.get_loc("Загружено в АСУ ОДС (от плана) %")

            # адреса процентных ячеек (фиксируем только колонку, строка будет меняться)
            cell_pole_pct = xl_rowcol_to_cell(first_data_row_1, col_pole_pct_idx, row_abs=False, col_abs=True)
            cell_sapr_pct = xl_rowcol_to_cell(first_data_row_1, col_sapr_pct_idx, row_abs=False, col_abs=True)
            cell_agr_pct = xl_rowcol_to_cell(first_data_row_1, col_agr_pct_idx, row_abs=False, col_abs=True)
            cell_asu_pct = xl_rowcol_to_cell(first_data_row_1, col_asu_pct_idx, row_abs=False, col_abs=True)

            # диапазон строк для ОИВ (без строки "Итого")
            row1_data = first_data_row_1
            row1_last = last_data_row_1 - 1

            #  ОИВ зелёный только если все 4 процента = 100
            formula_oiv = f"=AND({cell_pole_pct}=100,{cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
            worksheet.conditional_format(
                row1_data, col_oiv_idx, row1_last, col_oiv_idx,
                {"type": "formula", "criteria": formula_oiv, "format": green_fill_fmt}
            )

            #  от Утв. графика до Выполнено % - если хотя бы один из 4 процентов =100
            formula_block1 = (
                f"=OR({cell_pole_pct}=100,{cell_sapr_pct}=100,"
                f"{cell_agr_pct}=100,{cell_asu_pct}=100)"
            )
            worksheet.conditional_format(
                row1_data, col_plan1_idx, row1_last, col_pole_pct_idx,
                {"type": "formula", "criteria": formula_block1, "format": green_fill_fmt}
            )

            #  от столбца после Выполнено % до Загружено в САПР % - если хоть один из 3 правых =100
            if col_pole_pct_idx + 1 <= col_sapr_pct_idx:
                formula_block2 = (
                    f"=OR({cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                )
                worksheet.conditional_format(
                    row1_data, col_pole_pct_idx + 1, row1_last, col_sapr_pct_idx,
                    {"type": "formula", "criteria": formula_block2, "format": green_fill_fmt}
                )

            #  от столбца после Загружено в САПР % до Согласовано % -  если один из 2 правых =100
            if col_sapr_pct_idx + 1 <= col_agr_pct_idx:
                formula_block3 = f"=OR({cell_agr_pct}=100,{cell_asu_pct}=100)"
                worksheet.conditional_format(
                    row1_data, col_sapr_pct_idx + 1, row1_last, col_agr_pct_idx,
                    {"type": "formula", "criteria": formula_block3, "format": green_fill_fmt}
                )

            #  от столбца после Согласовано % до Загружено в АСУ ОДС (от плана) % - если именно этот последний =100
            if col_agr_pct_idx + 1 <= col_asu_pct_idx:
                formula_block4 = f"={cell_asu_pct}=100"
                worksheet.conditional_format(
                    row1_data, col_agr_pct_idx + 1, row1_last, col_asu_pct_idx,
                    {"type": "formula", "criteria": formula_block4, "format": green_fill_fmt}
                )

        # зелёная заливка по строкам (ГА)
        if not hectares_today.empty:
            col_oiv_idx_ha = hectares_today.columns.get_loc("ОИВ")
            col_plan1_idx_ha = hectares_today.columns.get_loc(plan_col1)
            col_plan3_idx_ha = hectares_today.columns.get_loc(plan_col3)
            col_pole_pct_idx_ha = hectares_today.columns.get_loc("Выполнено полевое обследование (от плана) %")
            col_sapr_pct_idx_ha = hectares_today.columns.get_loc("Загружено в САПР %")
            col_agr_pct_idx_ha = hectares_today.columns.get_loc("Согласовано в САПР (от плана) %")
            col_asu_pct_idx_ha = hectares_today.columns.get_loc("Загружено в АСУ ОДС (от плана) %")

            cell_pole_pct_ha = xl_rowcol_to_cell(first_data_row_2, col_pole_pct_idx_ha, row_abs=False, col_abs=True)
            cell_sapr_pct_ha = xl_rowcol_to_cell(first_data_row_2, col_sapr_pct_idx_ha, row_abs=False, col_abs=True)
            cell_agr_pct_ha = xl_rowcol_to_cell(first_data_row_2, col_agr_pct_idx_ha, row_abs=False, col_abs=True)
            cell_asu_pct_ha = xl_rowcol_to_cell(first_data_row_2, col_asu_pct_idx_ha, row_abs=False, col_abs=True)

            row2_data = first_data_row_2
            row2_last = last_data_row_2 - 1

            # ОИВ
            formula_oiv_ha = (
                f"=AND({cell_pole_pct_ha}=100,{cell_sapr_pct_ha}=100,"
                f"{cell_agr_pct_ha}=100,{cell_asu_pct_ha}=100)"
            )
            worksheet.conditional_format(
                row2_data, col_oiv_idx_ha, row2_last, col_oiv_idx_ha,
                {"type": "formula", "criteria": formula_oiv_ha, "format": green_fill_fmt}
            )

            # блок 1
            formula_block1_ha = (
                f"=OR({cell_pole_pct_ha}=100,{cell_sapr_pct_ha}=100,"
                f"{cell_agr_pct_ha}=100,{cell_asu_pct_ha}=100)"
            )
            worksheet.conditional_format(
                row2_data, col_plan1_idx_ha, row2_last, col_pole_pct_idx_ha,
                {"type": "formula", "criteria": formula_block1_ha, "format": green_fill_fmt}
            )

            # блок 2
            if col_pole_pct_idx_ha + 1 <= col_sapr_pct_idx_ha:
                formula_block2_ha = (
                    f"=OR({cell_sapr_pct_ha}=100,{cell_agr_pct_ha}=100,{cell_asu_pct_ha}=100)"
                )
                worksheet.conditional_format(
                    row2_data, col_pole_pct_idx_ha + 1, row2_last, col_sapr_pct_idx_ha,
                    {"type": "formula", "criteria": formula_block2_ha, "format": green_fill_fmt}
                )

            # блок 3
            if col_sapr_pct_idx_ha + 1 <= col_agr_pct_idx_ha:
                formula_block3_ha = f"=OR({cell_agr_pct_ha}=100,{cell_asu_pct_ha}=100)"
                worksheet.conditional_format(
                    row2_data, col_sapr_pct_idx_ha + 1, row2_last, col_agr_pct_idx_ha,
                    {"type": "formula", "criteria": formula_block3_ha, "format": green_fill_fmt}
                )

            # блок 4
            if col_agr_pct_idx_ha + 1 <= col_asu_pct_idx_ha:
                formula_block4_ha = f"={cell_asu_pct_ha}=100"
                worksheet.conditional_format(
                    row2_data, col_agr_pct_idx_ha + 1, row2_last, col_asu_pct_idx_ha,
                    {"type": "formula", "criteria": formula_block4_ha, "format": green_fill_fmt}
                )

        #  ЛИСТ "ПРОШЛЫЙ ОТЧЁТ {ГОД}"
        if auto_prev_pieces_df is not None or auto_prev_hectares_df is not None:
            prev_pieces = display_prev_pieces_df if isinstance(display_prev_pieces_df,
                                                               pd.DataFrame) else pieces_today.iloc[0:0].copy()
            prev_hect = display_prev_hectares_df if isinstance(display_prev_hectares_df,
                                                               pd.DataFrame) else hectares_today.iloc[0:0].copy()
            prev_hect_display = make_hectares_display_df(prev_hect, selected_year)

            prev_dyn_pieces = auto_prev_dyn_pieces_df.copy() if isinstance(auto_prev_dyn_pieces_df,pd.DataFrame) else pieces_today.iloc[0:0].copy()
            prev_dyn_hect = auto_prev_dyn_hect_df.copy() if isinstance(auto_prev_dyn_hect_df,pd.DataFrame) else hectares_today.iloc[0:0].copy()

            n1_prev = len(prev_pieces)
            n2_prev = len(prev_hect)

            green_prev_row_pieces = None
            green_prev_row_hect = None

            title_prev_row_1 = 0
            header_prev_row_1 = 1
            first_prev_data_row_1 = header_prev_row_1 + 1
            last_prev_data_row_1 = header_prev_row_1 + n1_prev

            COMPARE_BLOCK_ROWS = 8
            GAP_AFTER_COMPARE = 2

            title_prev_row_2 = last_prev_data_row_1 + 2 + COMPARE_BLOCK_ROWS + GAP_AFTER_COMPARE
            header_prev_row_2 = title_prev_row_2 + 1
            first_prev_data_row_2 = header_prev_row_2 + 1
            last_prev_data_row_2 = header_prev_row_2 + n2_prev

            prev_pieces.to_excel(
                writer,
                sheet_name=prev_sheet_name,
                index=False,
                startrow=header_prev_row_1
            )
            prev_hect_display.to_excel(
                writer,
                sheet_name=prev_sheet_name,
                index=False,
                startrow=header_prev_row_2
            )


            ws_prev = writer.sheets[prev_sheet_name]

            #  ПРОШЛЫЙ ОТЧЁТ: проставить проценты со знаком %
            write_percent_formulas(
                ws_prev,
                prev_pieces,
                first_prev_data_row_1,
                last_prev_data_row_1
            )

            write_percent_formulas(
                ws_prev,
                prev_hect,
                first_prev_data_row_2,
                last_prev_data_row_2
            )

            #  ПРОШЛЫЙ ОТЧЁТ: блоки под таблицами (как в новом)
            start_row_compare_1 = write_compare_block(ws_prev, prev_pieces, last_prev_data_row_1)
            start_row_compare_2 = write_compare_block(ws_prev, prev_hect, last_prev_data_row_2)

            # ПРОШЛЫЙ ОТЧЁТ: зелёная прокраска по строкам при 100% (как в новом)

            # зелёный только фон (как ты сейчас сделал в новом отчёте)
            green_fill_fmt_prev = workbook.add_format({"bg_color": "#00B050"})

            # зелёная заливка по строкам (ШТ)
            if not prev_pieces.empty:
                col_oiv_idx = prev_pieces.columns.get_loc("ОИВ")
                col_plan1_idx = prev_pieces.columns.get_loc(plan_col1)
                col_plan3_idx = prev_pieces.columns.get_loc(plan_col3)
                col_pole_pct_idx = prev_pieces.columns.get_loc("Выполнено полевое обследование (от плана) %")
                col_sapr_pct_idx = prev_pieces.columns.get_loc("Загружено в САПР %")
                col_agr_pct_idx = prev_pieces.columns.get_loc("Согласовано в САПР (от плана) %")
                col_asu_pct_idx = prev_pieces.columns.get_loc("Загружено в АСУ ОДС (от плана) %")

                cell_pole_pct = xl_rowcol_to_cell(first_prev_data_row_1, col_pole_pct_idx, row_abs=False, col_abs=True)
                cell_sapr_pct = xl_rowcol_to_cell(first_prev_data_row_1, col_sapr_pct_idx, row_abs=False, col_abs=True)
                cell_agr_pct = xl_rowcol_to_cell(first_prev_data_row_1, col_agr_pct_idx, row_abs=False, col_abs=True)
                cell_asu_pct = xl_rowcol_to_cell(first_prev_data_row_1, col_asu_pct_idx, row_abs=False, col_abs=True)

                row_data = first_prev_data_row_1
                row_last = last_prev_data_row_1 - 1

                formula_oiv = f"=AND({cell_pole_pct}=100,{cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                ws_prev.conditional_format(row_data, col_oiv_idx, row_last, col_oiv_idx,
                                           {"type": "formula", "criteria": formula_oiv, "format": green_fill_fmt_prev})

                formula_block1 = f"=OR({cell_pole_pct}=100,{cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                ws_prev.conditional_format(row_data, col_plan1_idx, row_last, col_pole_pct_idx,
                                           {"type": "formula", "criteria": formula_block1,
                                            "format": green_fill_fmt_prev})

                if col_pole_pct_idx + 1 <= col_sapr_pct_idx:
                    formula_block2 = f"=OR({cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                    ws_prev.conditional_format(row_data, col_pole_pct_idx + 1, row_last, col_sapr_pct_idx,
                                               {"type": "formula", "criteria": formula_block2,
                                                "format": green_fill_fmt_prev})

                if col_sapr_pct_idx + 1 <= col_agr_pct_idx:
                    formula_block3 = f"=OR({cell_agr_pct}=100,{cell_asu_pct}=100)"
                    ws_prev.conditional_format(row_data, col_sapr_pct_idx + 1, row_last, col_agr_pct_idx,
                                               {"type": "formula", "criteria": formula_block3,
                                                "format": green_fill_fmt_prev})

                if col_agr_pct_idx + 1 <= col_asu_pct_idx:
                    formula_block4 = f"={cell_asu_pct}=100"
                    ws_prev.conditional_format(row_data, col_agr_pct_idx + 1, row_last, col_asu_pct_idx,
                                               {"type": "formula", "criteria": formula_block4,
                                                "format": green_fill_fmt_prev})

            #  зелёная заливка по строкам (ГА)
            if not prev_hect_display.empty:
                col_oiv_idx = prev_hect_display.columns.get_loc("ОИВ")
                col_plan1_idx = prev_hect_display.columns.get_loc(plan_col1)
                col_plan3_idx = prev_hect_display.columns.get_loc(plan_col3)
                col_pole_pct_idx = prev_hect_display.columns.get_loc("Выполнено полевое обследование (от плана) %")
                col_sapr_pct_idx = prev_hect_display.columns.get_loc("Загружено в САПР %")
                col_agr_pct_idx = prev_hect_display.columns.get_loc("Согласовано в САПР (от плана) %")
                col_asu_pct_idx = prev_hect_display.columns.get_loc("Загружено в АСУ ОДС (от плана) %")

                cell_pole_pct = xl_rowcol_to_cell(first_prev_data_row_2, col_pole_pct_idx, row_abs=False, col_abs=True)
                cell_sapr_pct = xl_rowcol_to_cell(first_prev_data_row_2, col_sapr_pct_idx, row_abs=False, col_abs=True)
                cell_agr_pct = xl_rowcol_to_cell(first_prev_data_row_2, col_agr_pct_idx, row_abs=False, col_abs=True)
                cell_asu_pct = xl_rowcol_to_cell(first_prev_data_row_2, col_asu_pct_idx, row_abs=False, col_abs=True)

                row_data = first_prev_data_row_2
                row_last = last_prev_data_row_2 - 1

                formula_oiv = f"=AND({cell_pole_pct}=100,{cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                ws_prev.conditional_format(row_data, col_oiv_idx, row_last, col_oiv_idx,
                                           {"type": "formula", "criteria": formula_oiv, "format": green_fill_fmt_prev})

                formula_block1 = f"=OR({cell_pole_pct}=100,{cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                ws_prev.conditional_format(row_data, col_plan1_idx, row_last, col_pole_pct_idx,
                                           {"type": "formula", "criteria": formula_block1,
                                            "format": green_fill_fmt_prev})

                if col_pole_pct_idx + 1 <= col_sapr_pct_idx:
                    formula_block2 = f"=OR({cell_sapr_pct}=100,{cell_agr_pct}=100,{cell_asu_pct}=100)"
                    ws_prev.conditional_format(row_data, col_pole_pct_idx + 1, row_last, col_sapr_pct_idx,
                                               {"type": "formula", "criteria": formula_block2,
                                                "format": green_fill_fmt_prev})

                if col_sapr_pct_idx + 1 <= col_agr_pct_idx:
                    formula_block3 = f"=OR({cell_agr_pct}=100,{cell_asu_pct}=100)"
                    ws_prev.conditional_format(row_data, col_sapr_pct_idx + 1, row_last, col_agr_pct_idx,
                                               {"type": "formula", "criteria": formula_block3,
                                                "format": green_fill_fmt_prev})

                if col_agr_pct_idx + 1 <= col_asu_pct_idx:
                    formula_block4 = f"={cell_asu_pct}=100"
                    ws_prev.conditional_format(row_data, col_agr_pct_idx + 1, row_last, col_asu_pct_idx,
                                               {"type": "formula", "criteria": formula_block4,
                                                "format": green_fill_fmt_prev})

            header_display_prev = build_header_display_map(selected_year)

            ws_prev.write(
                title_prev_row_1, 0,
                f"ГЗ {selected_year}, шт (прошлый отчёт)", title_fmt
            )
            ws_prev.write(
                title_prev_row_2, 0,
                f"ГЗ {selected_year}, га (прошлый отчёт)", title_fmt
            )

            for col_num, value in enumerate(prev_pieces.columns.values):
                disp = header_display_prev.get(value, value)
                disp = wrap_long_header(disp)
                fmt = get_header_fmt(value)
                ws_prev.write(header_prev_row_1, col_num, disp, fmt)

            for col_num, value in enumerate(prev_hect.columns.values):
                disp = header_display_prev.get(value, value)
                disp = wrap_long_header(disp)
                fmt = get_header_fmt(value)
                ws_prev.write(header_prev_row_2, col_num, disp, fmt)

            ws_prev.set_row(header_prev_row_1, 100)
            ws_prev.set_row(header_prev_row_2, 100)

            set_columns_by_header(ws_prev, prev_pieces, min_width=25, max_width=60, padding=4)
            set_columns_by_header(ws_prev, prev_hect_display, min_width=16, max_width=55, padding=4)

            NUM_W = 16
            PCT_W = percent_col_width
            OIV_W = 50

            for i, col_name in enumerate(prev_pieces.columns):
                if col_name == "ОИВ":
                    ws_prev.set_column(i, i, OIV_W)
                elif col_name in percent_cols:
                    ws_prev.set_column(i, i, PCT_W)
                else:
                    ws_prev.set_column(i, i, NUM_W)

            for i, col_name in enumerate(prev_hect_display.columns):
                if col_name == "ОИВ":
                    ws_prev.set_column(i, i, OIV_W)
                elif col_name in percent_cols:
                    ws_prev.set_column(i, i, PCT_W)
                else:
                    ws_prev.set_column(i, i, NUM_W)

            def _apply_main_like_widths(ws, df, oiv_w=50):
                for i, col_name in enumerate(df.columns):
                    if col_name == "ОИВ":
                        ws.set_column(i, i, oiv_w)
                    elif col_name in percent_cols:
                        ws.set_column(i, i, PCT_W)
                    else:
                        ws.set_column(i, i, NUM_W)  # только ширина (форматы ты ниже уже проставляешь)

            _apply_main_like_widths(ws_prev, prev_pieces, oiv_w=OIV_W)
            _apply_main_like_widths(ws_prev, prev_hect_display, oiv_w=OIV_W)

            #  ПРОШЛЫЙ ОТЧЁТ: центрирование чисел в НЕ%-колонках (как в новом)

            # ШТ
            for i, col_name in enumerate(prev_pieces.columns):
                if col_name == "ОИВ":
                    continue
                if col_name in percent_cols:
                    continue
                ws_prev.set_column(i, i, NUM_W, center_only_fmt)

            # ГА
            for i, col_name in enumerate(prev_hect_display.columns):
                if col_name == "ОИВ":
                    continue
                if col_name in percent_cols:
                    continue
                ws_prev.set_column(i, i, NUM_W, center_only_fmt)

            for row_idx in range(first_prev_data_row_1, last_prev_data_row_1 + 1):
                ws_prev.set_row(row_idx, 18)
            for row_idx in range(first_prev_data_row_2, last_prev_data_row_2 + 1):
                ws_prev.set_row(row_idx, 18)
            if n1_prev > 0:
                ws_prev.set_row(last_prev_data_row_1, 20)
            if n2_prev > 0:
                ws_prev.set_row(last_prev_data_row_2, 20)

            # ПРОШЛЫЙ ОТЧЁТ: оформить строку ИТОГО как в новом

            if n1_prev > 0:
                for col_idx, col_name in enumerate(prev_pieces.columns):
                    value = prev_pieces.iloc[-1, col_idx]
                    if col_idx == 0:
                        fmt = ito_fmt
                    else:
                        fmt = total_blue_percent_fmt if col_name in percent_cols else total_blue_fmt
                    ws_prev.write(last_prev_data_row_1, col_idx, value, fmt)

            if n2_prev > 0:
                for col_idx, col_name in enumerate(prev_hect_display.columns):
                    value = prev_hect_display.iloc[-1, col_idx]
                    if col_idx == 0:
                        fmt = ito_fmt
                    else:
                        fmt = total_blue_percent_fmt if col_name in percent_cols else total_blue_fmt
                    ws_prev.write(last_prev_data_row_2, col_idx, value, fmt)

            #  ПРОШЛЫЙ ОТЧЁТ: ширины + центрирование

            # ШТ
            for i, col_name in enumerate(prev_pieces.columns):
                if col_name == "ОИВ":
                    continue
                if col_name in percent_cols:
                    continue
                ws_prev.set_column(i, i, NUM_W, center_only_fmt)

            # ГА (важно: prev_hect_display)
            for i, col_name in enumerate(prev_hect_display.columns):
                if col_name == "ОИВ":
                    continue
                if col_name in percent_cols:
                    continue
                ws_prev.set_column(i, i, NUM_W, center_only_fmt)

            # MAIN: тонкая сетка + ОИВ + ИТОГО + ВНЕШНЯЯ РАМКА (прошлый отчёт)

            thin_grid_mid_fmt = workbook.add_format({
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            })

            thin_grid_left_oiv_sep_fmt = workbook.add_format({
                "border": 1,
                "left": 2,
                "right": 2,
                "align": "center",
                "valign": "vcenter",
            })

            thin_grid_right_outline_fmt = workbook.add_format({
                "border": 1,
                "right": 2,
                "align": "center",
                "valign": "vcenter",
            })

            # строка ИТОГО (верх+низ жирные) — тоже центрируем
            thin_grid_total_mid_bottom_fmt = workbook.add_format({
                "border": 1,
                "top": 2,
                "bottom": 2,
                "align": "center",
                "valign": "vcenter",
            })

            thin_grid_total_left_oiv_sep_bottom_fmt = workbook.add_format({
                "border": 1,
                "top": 2,
                "bottom": 2,
                "left": 2,
                "right": 2,
                "align": "center",
                "valign": "vcenter",
            })

            thin_grid_total_right_outline_bottom_fmt = workbook.add_format({
                "border": 1,
                "top": 2,
                "bottom": 2,
                "right": 2,
                "align": "center",
                "valign": "vcenter",
            })

            # нижняя рамка (на строке ИТОГО)
            thin_grid_total_mid_bottom_fmt = workbook.add_format({"border": 1, "top": 2, "bottom": 2})
            thin_grid_total_left_oiv_sep_bottom_fmt = workbook.add_format(
                {"border": 1, "top": 2, "bottom": 2, "left": 2, "right": 2})
            thin_grid_total_right_outline_bottom_fmt = workbook.add_format(
                {"border": 1, "top": 2, "bottom": 2, "right": 2})

            def apply_main_grid_with_outline(ws, first_data_row, df_rows, ncols):
                """
                Рисуем ОДИН слой границ для данных:
                  - внутри: тонкая сетка
                  - col=0: слева жирная рамка + справа жирный разделитель ОИВ
                  - последняя колонка: правая жирная рамка
                  - строка ИТОГО: жирная линия сверху + нижняя жирная рамка
                """
                if df_rows <= 0 or ncols <= 0:
                    return

                total_row = first_data_row + (df_rows - 1)
                last_col = ncols - 1

                # тело таблицы (все строки ДО ИТОГО)
                if total_row - 1 >= first_data_row:
                    ws.conditional_format(
                        first_data_row, 0,
                        total_row - 1, 0,
                        {"type": "no_errors", "format": thin_grid_left_oiv_sep_fmt}
                    )

                    # середина (если есть)
                    if last_col >= 2:
                        ws.conditional_format(
                            first_data_row, 1,
                            total_row - 1, last_col - 1,
                            {"type": "no_errors", "format": thin_grid_mid_fmt}
                        )

                    # последняя колонка: right=2
                    if last_col >= 1:
                        ws.conditional_format(
                            first_data_row, last_col,
                            total_row - 1, last_col,
                            {"type": "no_errors", "format": thin_grid_right_outline_fmt}
                        )

                # строка ИТОГО: top=2 и bottom=2 (низ рамки)
                ws.conditional_format(
                    total_row, 0,
                    total_row, 0,
                    {"type": "no_errors", "format": thin_grid_total_left_oiv_sep_bottom_fmt}
                )

                # середина
                if last_col >= 2:
                    ws.conditional_format(
                        total_row, 1,
                        total_row, last_col - 1,
                        {"type": "no_errors", "format": thin_grid_total_mid_bottom_fmt}
                    )

                # последняя колонка
                if last_col >= 1:
                    ws.conditional_format(
                        total_row, last_col,
                        total_row, last_col,
                        {"type": "no_errors", "format": thin_grid_total_right_outline_bottom_fmt}
                    )

            #  ШТ
            if n1_prev > 0:
                apply_main_grid_with_outline(
                    ws_prev,
                    first_prev_data_row_1,
                    prev_pieces.shape[0],
                    prev_pieces.shape[1]
                )

            #  ГА
            if n2_prev > 0:
                apply_main_grid_with_outline(
                    ws_prev,
                    first_prev_data_row_2,
                    prev_hect_display.shape[0],
                    prev_hect_display.shape[1]
                )

            # ШАПКА ШТ: жирные вертикали + линия под шапкой
            if n1_prev > 0:
                ws_prev.conditional_format(
                    header_prev_row_1, 0,
                    header_prev_row_1, prev_pieces.shape[1] - 1,
                    {
                        "type": "no_blanks",
                        "format": header_thick_fmt
                    }
                )
            # ШАПКА ГА: жирные вертикали + линия под шапкой
            if n2_prev > 0:
                ws_prev.conditional_format(
                    header_prev_row_2, 0,
                    header_prev_row_2, prev_hect_display.shape[1] - 1,
                    {
                        "type": "no_blanks",
                        "format": header_thick_fmt
                    }
                )

                # Персиковая таблица (динамика) справа от блоков в прошлом отчёте

                # Персиковая: рамка (толстая), жирная линия под шапкой, жирная линия над ИТОГО и низ рамки на ИТОГО.
                PEACH_BG = "#F8CBAD"

                # "Динамика за неделю"
                dyn_title_cell_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                    "top": 2,
                })
                dyn_title_left_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                    "top": 2,
                    "left": 2,
                })
                dyn_title_right_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                    "top": 2,
                    "right": 2,
                })

                #  жирная линия снизу + верх рамки
                dyn_hdr_mid_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "text_wrap": True,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                })
                dyn_hdr_left_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "text_wrap": True,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                    "left": 2,
                })
                dyn_hdr_right_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "text_wrap": True,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                    "right": 2,
                })

                #  жирная линия сверху + низ рамки
                dyn_total_mid_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                })
                dyn_total_left_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                    "left": 2,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                })
                dyn_total_right_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "top": 2,
                    "bottom": 2,
                    "right": 2,
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                })
                dyn_mid_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter"
                })

                dyn_left_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "left": 2,
                    "align": "center",
                    "valign": "vcenter"
                })

                dyn_right_fmt = workbook.add_format({
                    "bg_color": PEACH_BG,
                    "border": 1,
                    "right": 2,
                    "align": "center",
                    "valign": "vcenter"
                })

                def _write_prev_peach_block(ws, title_row, header_row, first_data_row, start_col, dyn_cols, df_dyn,
                                            n_rows):
                    """Рисует персиковую динамику в прошлом отчёте без CF-рамок (чтобы не конфликтовало)."""
                    if n_rows <= 0:
                        return

                    last_j = len(dyn_cols) - 1
                    last_col = start_col + last_j

                    for j, col_name in enumerate(dyn_cols):
                        fmt = dyn_title_cell_fmt
                        if j == 0:
                            fmt = dyn_title_left_fmt
                        elif j == last_j:
                            fmt = dyn_title_right_fmt
                        ws.write(title_row, start_col + j, "" if j != 0 else "Динамика за неделю", fmt)
                    ws.merge_range(title_row, start_col, title_row, last_col, "Динамика за неделю", dyn_title_cell_fmt)

                    for j, col_name in enumerate(dyn_cols):
                        if j == 0:
                            fmt = dyn_hdr_left_fmt
                        elif j == last_j:
                            fmt = dyn_hdr_right_fmt
                        else:
                            fmt = dyn_hdr_mid_fmt
                        ws.write(header_row, start_col + j, col_name, fmt)
                        ws.set_column(start_col + j, start_col + j, 14)

                    for i in range(n_rows):
                        r = first_data_row + i
                        is_total = (i == n_rows - 1)

                        for j, col_name in enumerate(dyn_cols):
                            c = start_col + j

                            # значение
                            val = 0
                            if df_dyn is not None and col_name in df_dyn.columns and i < len(df_dyn):
                                val = df_dyn.iloc[i][col_name]

                            # формат
                            if is_total:
                                if j == 0:
                                    fmt = dyn_total_left_fmt
                                elif j == last_j:
                                    fmt = dyn_total_right_fmt
                                else:
                                    fmt = dyn_total_mid_fmt
                            else:
                                if j == 0:
                                    fmt = dyn_left_fmt
                                elif j == last_j:
                                    fmt = dyn_right_fmt
                                else:
                                    fmt = dyn_mid_fmt

                            ws.write(r, c, val, fmt)

                #ШТ (прошлый отчёт)
                if n1_prev > 0:
                    start_col_prev_dyn1 = prev_pieces.shape[1] + 2

                    _write_prev_peach_block(
                        ws=ws_prev,
                        title_row=title_prev_row_1,
                        header_row=header_prev_row_1,
                        first_data_row=first_prev_data_row_1,
                        start_col=start_col_prev_dyn1,
                        dyn_cols=dyn_cols,
                        df_dyn=prev_dyn_pieces,
                        n_rows=n1_prev
                    )

                # ГА (прошлый отчёт)
                if n2_prev > 0:
                    start_col_prev_dyn2 = prev_hect.shape[1] + 2

                    _write_prev_peach_block(
                        ws=ws_prev,
                        title_row=title_prev_row_2,
                        header_row=header_prev_row_2,
                        first_data_row=first_prev_data_row_2,
                        start_col=start_col_prev_dyn2,
                        dyn_cols=dyn_cols,
                        df_dyn=prev_dyn_hect,
                        n_rows=n2_prev
                    )

                # ВИТРИНА ПРОШЛОГО ОТЧЁТА
                if use_baseline_prev:
                    display_prev_pieces_df = baseline_prev_pieces
                    display_prev_hectares_df = baseline_prev_hectares

                    display_prev_comparison_pieces = {}
                    display_prev_comparison_hect = {}
                    display_prev_delta_pieces = {}
                    display_prev_delta_hect = {}
                else:
                    display_prev_pieces_df = auto_prev_pieces_df
                    display_prev_hectares_df = auto_prev_hectares_df

                    display_prev_comparison_pieces = auto_prev_comparison_pieces
                    display_prev_comparison_hect = auto_prev_comparison_hect
                    display_prev_delta_pieces = auto_prev_delta_pieces
                    display_prev_delta_hect = auto_prev_delta_hect

                # БЛОК СРАВНЕНИЯ В ПРОШЛОМ ОТЧЕТЕ

            # Заполнить пустые ячейки персиковой сеткой (прошлый отчёт)
            if n1_prev > 0:
                ws_prev.conditional_format(
                    first_prev_data_row_1, start_col_prev_dyn1,
                    last_prev_data_row_1, start_col_prev_dyn1 + len(dyn_cols) - 1,
                    {"type": "blanks", "format": dyn_blank_cell_fmt}
                )

            if n2_prev > 0:
                ws_prev.conditional_format(
                    first_prev_data_row_2, start_col_prev_dyn2,
                    last_prev_data_row_2, start_col_prev_dyn2 + len(dyn_cols) - 1,
                    {"type": "blanks", "format": dyn_blank_cell_fmt}
                )

            # БЛОК СРАВНЕНИЯ В ПРОШЛОМ ОТЧЕТЕ
            if n1_prev > 0:
                start_prev_compare_1 = write_compare_block(
                    ws_prev,
                    prev_pieces,
                    last_prev_data_row_1,
                )

                if start_prev_compare_1:
                    write_comparison_values_row(
                        ws_prev,
                        prev_pieces,
                        last_prev_data_row_1,
                        start_prev_compare_1,
                        display_prev_comparison_pieces,
                        plan_col1,
                        plan_col3,
                        dyn_cell_fmt
                    )

                    # Ячейки для сравнения ▲ под персиковой таблицей в прошлом отчёте
                    write_delta_comparison_cells(
                        ws_prev, prev_pieces, last_prev_data_row_1, start_prev_compare_1,
                        display_prev_delta_pieces, {}, dyn_cols, start_col_prev_dyn1,
                        dyn_cell_fmt, is_prev_report=True
                    )
                    # Зелёная строка под персиковой таблицей (прошлый отчёт, ШТ)
                    green_prev_row_pieces = add_green_numbers_under_persic(
                        worksheet=ws_prev,
                        workbook=workbook,
                        df_main=prev_pieces,
                        last_data_row=last_prev_data_row_1,
                        start_row_compare=start_prev_compare_1,
                        start_col_dyn=start_col_prev_dyn1,
                        dyn_cols=dyn_cols,
                        selected_year=selected_year,
                        is_hectares=False,
                    )


            if n2_prev > 0:
                start_prev_compare_2 = write_compare_block(
                    ws_prev,
                    prev_hect,
                    last_prev_data_row_2,
                )

                if start_prev_compare_2:
                    write_comparison_values_row(
                        ws_prev,
                        prev_hect,
                        last_prev_data_row_2,
                        start_prev_compare_2,
                        display_prev_comparison_hect,
                        plan_col1,
                        plan_col3,
                        dyn_cell_fmt
                    )

                    write_delta_comparison_cells(
                        ws_prev, prev_hect, last_prev_data_row_2, start_prev_compare_2,
                        display_prev_delta_hect, {}, dyn_cols, start_col_prev_dyn2,
                        dyn_cell_fmt, is_prev_report=True
                    )
                    green_prev_row_hect = add_green_numbers_under_persic(
                        worksheet=ws_prev,
                        workbook=workbook,
                        df_main=prev_hect,
                        last_data_row=last_prev_data_row_2,
                        start_row_compare=start_prev_compare_2,
                        start_col_dyn=start_col_prev_dyn2,
                        dyn_cols=dyn_cols,
                        selected_year=selected_year,
                        is_hectares=True,
                    )

    # сохраняем текущий отчёт как "прошлый"
    if not isinstance(prev_reports, dict):
        prev_reports = {}

    prev_reports[selected_year] = {
        "pieces": pieces_today,
        "hectares": hectares_today,
        "dyn_pieces": dyn_pieces,
        "dyn_hect": dyn_hect,
        "comparison_values_pieces": comparison_values_pieces,
        "comparison_values_hect": comparison_values_hect,
        "delta_values_pieces": delta_values_pieces_curr,
        "delta_values_hect": delta_values_hect_curr,

        "report_date_short": dt.date.today().strftime("%d.%m"),
        "report_date_full": dt.date.today().strftime("%d.%m.%Y"),
    }
    save_prev_reports(prev_reports)
    # возврат результата для Streamlit
    if return_excel_bytes:
        output.seek(0)
        return output.getvalue(), filename, preview_pieces, preview_hectares

        # локальный режим
    abs_path = os.path.abspath(out_name)
    return abs_path

def generate_report_from_df(
        df_input: pd.DataFrame,
        plan_path: str | None,
        report_type: str,
        selected_year: int,
        baseline_prev_path: str | None = None,
        return_excel_bytes: bool = True,
):
    return generate_report(
        df_input=df_input,
        plan_path=plan_path,
        report_type=report_type,
        selected_year=selected_year,
        baseline_prev_path=baseline_prev_path,
        return_excel_bytes=return_excel_bytes,
    )

@st.cache_data(show_spinner=False)
def _cached_generate_oiv_report(
    df_hash: str,
    df_input: pd.DataFrame,
    plan_bytes: bytes | None,
    prev_bytes: bytes | None,
    selected_year: int,
):

    plan_path = None
    prev_path = None

    if plan_bytes:
        tmp_plan = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_plan.write(plan_bytes)
        tmp_plan.close()
        plan_path = tmp_plan.name

    if prev_bytes:
        tmp_prev = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_prev.write(prev_bytes)
        tmp_prev.close()
        prev_path = tmp_prev.name

    excel_bytes, filename, preview_pieces, preview_hectares = generate_report(
        df_input=df_input,
        plan_path=plan_path,
        report_type="Ежедневный",
        selected_year=selected_year,
        baseline_prev_path=prev_path,
        return_excel_bytes=True,
    )
    return excel_bytes, filename, preview_pieces, preview_hectares


# STREAMLIT PAGE
def OIV_otchet(df: pd.DataFrame):
    st.header("📊 Отчёт по ОИВ")

    # ИНИЦИАЛИЗАЦИЯ SESSION STATE
    if "oiv_preview_pieces" not in st.session_state:
        st.session_state["oiv_preview_pieces"] = None
    if "oiv_preview_hectares" not in st.session_state:
        st.session_state["oiv_preview_hectares"] = None
    if "oiv_excel_bytes" not in st.session_state:
        st.session_state["oiv_excel_bytes"] = None
    if "oiv_excel_name" not in st.session_state:
        st.session_state["oiv_excel_name"] = None
    if "oiv_last_run" not in st.session_state:
        st.session_state["oiv_last_run"] = None

    # Показываем дату последней генерации
    if st.session_state["oiv_last_run"]:
        st.caption(f"Последняя генерация: {st.session_state['oiv_last_run']}")

    # SIDEBAR — ПАРАМЕТРЫ

    st.sidebar.subheader("Параметры отчёта")

    selected_year = st.sidebar.selectbox(
        "Год",
        [2024, 2025, 2026, 2027, 2028, 2029, 2030],
        index=1
    )

    plan_file = st.sidebar.file_uploader(
        "Файл графика (план) — необязательно",
        type=["xlsx", "xls"],
        key="oiv_plan"
    )

    baseline_prev_file = st.sidebar.file_uploader(
        "Прошлый отчёт для сравнения (baseline) — необязательно",
        type=["xlsx", "xls"],
        key="oiv_prev"
    )


    # КНОПКА ГЕНЕРАЦИИ
    if st.button("Сформировать отчёт"):
        try:
            plan_bytes = plan_file.getvalue() if plan_file else None
            prev_bytes = baseline_prev_file.getvalue() if baseline_prev_file else None

            # Небольшой "хэш" чтобы кэш обновлялся при изменении структуры df
            df_hash = f"{df.shape}-{hash(tuple(map(str, df.columns)))}"

            excel_bytes, filename, preview_pieces, preview_hectares = _cached_generate_oiv_report(
                df_hash=df_hash,
                df_input=df,
                plan_bytes=plan_bytes,
                prev_bytes=prev_bytes,
                selected_year=selected_year,
            )

            # Сохраняем результат в session_state
            st.session_state["oiv_excel_bytes"] = excel_bytes
            st.session_state["oiv_excel_name"] = filename
            st.session_state["oiv_preview_pieces"] = preview_pieces
            st.session_state["oiv_preview_hectares"] = preview_hectares
            st.session_state["oiv_last_run"] = dt.datetime.now().strftime("%d.%m.%Y %H:%M:%S")

            st.success("Отчёт сформирован")

        except Exception as e:
            st.error(f"Ошибка: {e}")
            logger.exception("Failed to save previous reports")

    # КНОПКА СБРОСА
    if st.button("🧹 Сбросить превью"):
        st.session_state["oiv_preview_pieces"] = None
        st.session_state["oiv_preview_hectares"] = None
        st.session_state["oiv_excel_bytes"] = None
        st.session_state["oiv_excel_name"] = None
        st.session_state["oiv_last_run"] = None
        st.rerun()

    # ПРЕВЬЮ
    view_mode = st.radio("Превью", ["ШТ", "ГА"], horizontal=True)

    show_totals = st.checkbox("Показывать строку Итого", value=True)
    search = st.text_input("Поиск по ОИВ", value="").strip().lower()

    preview = (
        st.session_state["oiv_preview_pieces"]
        if view_mode == "ШТ"
        else st.session_state["oiv_preview_hectares"]
    )

    if preview is not None:
        df_view = preview.copy()

        # Убираем строку Итого при необходимости
        if not show_totals and "ОИВ" in df_view.columns:
            df_view = df_view[df_view["ОИВ"] != "Итого:"]

        # Поиск по названию ОИВ
        if search and "ОИВ" in df_view.columns:
            df_view = df_view[
                df_view["ОИВ"].astype(str).str.lower().str.contains(search, na=False)
            ]

        # Мультивыбор ОИВ
        if "ОИВ" in df_view.columns:
            oiv_list = [
                x for x in df_view["ОИВ"].dropna().unique().tolist()
                if x != "Итого:"
            ]
            selected_oiv = st.multiselect("ОИВ", oiv_list, default=oiv_list)
            if selected_oiv:
                df_view = df_view[
                    df_view["ОИВ"].isin(selected_oiv)
                    | (df_view["ОИВ"] == "Итого:")
                ]

        # KPI (берём значения из строки "Итого:")
        if "ОИВ" in df_view.columns:
            totals = df_view[df_view["ОИВ"] == "Итого:"]
            if not totals.empty:
                t = totals.iloc[0]
                cols = st.columns(4)

                def get_val(name):
                    return t[name] if name in df_view.columns else None

                cols[0].metric("Утверждено", get_val("Утверждено"))
                cols[1].metric("Отклонено", get_val("Отклонено"))
                cols[2].metric("Не утверждено БД", get_val("Не утверждено БД"))

                pct_name = "Утверждено (от плана ) %"
                cols[3].metric(
                    "% Утверждено",
                    get_val(pct_name) if get_val(pct_name) is not None else "—"
                )

        # Мини-график
        pct_candidates = [
            "Утверждено (от плана ) %",
            "Загружено в АСУ ОДС (от плана) %",
            "Согласовано в САПР (от плана) %",
            "Выполнено полевое обследование (от плана) %",
        ]

        pct_col = next((c for c in pct_candidates if c in df_view.columns), None)

        if pct_col and "ОИВ" in df_view.columns:
            chart_df = df_view[df_view["ОИВ"] != "Итого:"].copy()
            chart_df[pct_col] = pd.to_numeric(chart_df[pct_col], errors="coerce")
            chart_df = chart_df.dropna(subset=[pct_col])

            if not chart_df.empty:
                st.caption(f"Топ-10 по показателю: {pct_col}")
                st.bar_chart(
                    chart_df.sort_values(pct_col, ascending=False)
                            .head(10)
                            .set_index("ОИВ")[pct_col]
                )

        # Таблица
        st.dataframe(df_view, use_container_width=True)

    # КНОПКА СКАЧИВАНИЯ
    if st.session_state["oiv_excel_bytes"] is not None:
        st.download_button(
            label="📥 Скачать отчёт",
            data=st.session_state["oiv_excel_bytes"],
            file_name=st.session_state["oiv_excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )