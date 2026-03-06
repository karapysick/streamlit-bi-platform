import os
import sys
from typing import Callable, Optional

import pandas as pd
import streamlit as st

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from util import load_data
from util import load_data_from_excel
from pages_.page_1_quality_assessment import show_quality_assessment
from pages_.page_2_area_characteristics import show_area_characteristics
from pages_.page_3_border_coordination import show_border_coordination
from pages_.page_4_OIV_otchet import show_OIV_otchet
from pages_.page_5_approval_status import show_approval_status
from pages_.page_6_dynamics import show_dynamics


st.set_page_config(
    page_title="Интеграция данных",
    layout="wide",
)


def load_source_data() -> tuple[Optional[pd.DataFrame], Optional[str]]:
    """Load data either from a local Excel file or from the database."""
    st.sidebar.header("Источник данных")

    mode = st.sidebar.radio(
        "Выберите источник",
        ["Excel (локально)", "БД"],
        index=0,
    )

    if mode == "Excel (локально)":
        uploaded_file = st.sidebar.file_uploader(
            "Загрузите Excel-файл",
            type=["xlsx"],
        )

        if uploaded_file is None:
            st.info("Загрузите Excel-файл для работы приложения.")
            st.stop()

        try:
            st.session_state["db_excel_bytes"] = uploaded_file.getvalue()
            df = load_data_from_excel(uploaded_file)
            return df, None
        except Exception as exc:
            return None, str(exc)

    try:
        df = load_data()
        st.session_state["db_excel_bytes"] = None
        return df, None
    except Exception as exc:
        return None, str(exc)


def with_data(
    page_fn: Callable[[pd.DataFrame], None],
    df: Optional[pd.DataFrame],
    error_message: Optional[str],
) -> Callable[[], None]:
    """Wrap page rendering with source data validation."""

    def render() -> None:
        if error_message:
            st.error(f"Ошибка загрузки данных: {error_message}")
            return

        if df is None:
            st.info("Данные не загружены.")
            return

        page_fn(df)

    return render


def build_navigation(
    df: Optional[pd.DataFrame],
    error_message: Optional[str],
) -> st.navigation:
    """Create Streamlit navigation with grouped pages."""
    operations_pages = [
        st.Page(
            with_data(show_quality_assessment, df, error_message),
            title="🔍 Оценка качества УГП",
            url_path="quality_assessment",
            default=True,
        ),
        st.Page(
            with_data(show_area_characteristics, df, error_message),
            title="📐 Площадные характеристики",
            url_path="area_characteristics",
        ),
        st.Page(
            with_data(show_border_coordination, df, error_message),
            title="🧭 Согласование границ ОГХ",
            url_path="border_coordination",
        ),
    ]

    report_pages = [
        st.Page(
            with_data(show_OIV_otchet, df, error_message),
            title="📊 Отчёт по ОИВ",
            url_path="oiv_report",
        ),
        st.Page(
            with_data(show_approval_status, df, error_message),
            title="✅ Статус утверждения",
            url_path="approval_status",
        ),
        st.Page(
            with_data(show_dynamics, df, error_message),
            title="📈 Динамика",
            url_path="dynamics",
        ),
    ]

    return st.navigation(
        {
            "Операции": operations_pages,
            "Отчёты": report_pages,
        }
    )


def main() -> None:
    df, error_message = load_source_data()
    navigation = build_navigation(df, error_message)
    navigation.run()


if __name__ == "__main__":
    main()