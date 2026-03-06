import pandas as pd
import streamlit as st
import psycopg2

# util.py
COLUMN_MAP = {
    "ORD_NO": "№ Заказа МГГТ",
    "STG_NO": "Номер этапа МГГТ",
    "ORD_GEN_NO": "№ Ген. договора",
    "ORD_OIV_NAME": "ОИВ",
    "ORD_CLN_NAME": "Балансодержатель",
    "ORD_CLN_INN": "ИНН Балансодержателя",
    "OBJECT_NAME": "Наименование объекта",
    "ITP_CR": "Отдел исполнитель создания ИТП",

    "DT_FACT_MADE_FLD_WRK": "Дата изготовления полевых работ по факту",
    "SHORTNAME_FIELD_WORK": "Исполнитель полевых работ (АСД сводки)",
    "DT_AFTER_KORR": "Дата операции Исправление после корректуры",
    "SHORTNAME_OCIFR": "Исполнитель оцифровки ИТП (АСД сводки)",

    "REK_STATUS": "Статус оценки качества",
    "REK_SEND_DATE": "Дата отправки на исправление",
    "REK_RETURN_DATE": "Дата возврата исправленного материала",
    "REK_ACC_DATE": "Дата приёма без замечаний",

    "DT_GEO_FACT": "Дата изготовления геоподосновы по факту",
    "SHORTNAME_AREA_QUANT_CHAR": "Исполнитель определения площадных и количественных характеристик (АСД сводки)",

    "DT_LOAD_SAPR": "Дата загрузки в САПР МГГТ",
    "DT_AGREE_BORDER": "Дата согласования границ",
    "SHORTNAME_AGREE_BORDER_FACT": "Исполнитель согласования границ факт (АСД сводки)",

    "DT_LOAD_ODS_NO_CONFIRM": "Дата загрузки в АСУ ОДС (АСД)",
    "EXEC_LOAD_EMP": "Исполнитель загрузки - из сводок АСД",
    "DT_LOAD_ODS": "Дата утверждения в АСУ ОДС (МГГТ)",
    "EXEC_AGREE_EMP": "Исполнитель утверждения (АСД)",

    "VAL_PASSP_PLAN_SCHED_VOL_S": "Сумма Объем заказа, га",
}

DATE_COLS_RU = [
    "Дата изготовления полевых работ по факту",
    "Дата операции Исправление после корректуры",
    "Дата отправки на исправление",
    "Дата возврата исправленного материала",
    "Дата приёма без замечаний",
    "Дата изготовления геоподосновы по факту",
    "Дата загрузки в САПР МГГТ",
    "Дата согласования границ",
    "Дата загрузки в АСУ ОДС (АСД)",
    "Дата утверждения в АСУ ОДС (МГГТ)",
]


@st.cache_data(ttl=3600, show_spinner="Загружаем данные из БД")
def load_data():

    conn = psycopg2.connect(
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT"),
    )

    query = """
    SELECT *
    FROM public.v_mr_passport_volume_now;
    """
    df = pd.read_sql(query, conn)
    conn.close()

    # ✅ тут переименовали
    df = df.rename(columns=COLUMN_MAP)

    # ✅ тут привели даты
    for c in DATE_COLS_RU:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    return df


def clear_cache():
    """Сброс кеша данных."""
    load_data.clear()

# COLUMN_MAP и DATE_COLS_RU оставляем как есть

@st.cache_data(show_spinner="Загружаем данные из Excel")
def load_data_from_excel(xlsx_path: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path)

    # приводим колонки к тем, что ждут страницы
    df = df.rename(columns=COLUMN_MAP)

    for c in DATE_COLS_RU:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    return df
