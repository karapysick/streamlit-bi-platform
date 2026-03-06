import pandas as pd
import plotly.express as px
import streamlit as st


DEPT_COL = "Отдел исполнитель создания ИТП"
DEPT_DATE_COL = "Дата отправки на исправление"
EXEC_COL = "Исполнитель оцифровки ИТП (АСД сводки)"


def _prepare_return_stats(df: pd.DataFrame, group_col: str, date_col: str) -> pd.DataFrame:
    """Расчет процента возвратов по подразделениям/исполнителям.

    Для упрощения считаем, что любая строка с непустой датой отправки на исправление — это возврат.
    """
    data = df.copy()

    # Приводим названия колонок к ожидаемым, если они есть в данных
    missing = [c for c in [group_col, date_col] if c not in data.columns]
    if missing:
        raise ValueError(f"В данных отсутствуют необходимые столбцы: {', '.join(missing)}")

    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    data["is_return"] = data[date_col].notna()

    grouped = (
        data.groupby(group_col)
        .agg(
            total_orders=("is_return", "count"),
            returns=("is_return", "sum"),
        )
        .reset_index()
    )
    grouped["return_pct"] = grouped["returns"] / grouped["total_orders"] * 100
    return grouped


def _show_table_and_chart(stats: pd.DataFrame, group_col: str, title_prefix: str):
    st.subheader(f"{title_prefix} — сводная таблица")
    st.dataframe(
        stats.rename(
            columns={
                group_col: "Группа",
                "total_orders": "Всего заказов",
                "returns": "Возвраты",
                "return_pct": "Процент возврата (%)",
            }
        ),
        use_container_width=True,
    )

    st.subheader(f"{title_prefix} — график процента возврата")
    fig = px.bar(
        stats,
        x=group_col,
        y="return_pct",
        labels={group_col: "Группа", "return_pct": "Процент возврата, %"},
        title=title_prefix,
    )
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)


def show_quality_assessment(df: pd.DataFrame):
    """Страница «Оценка качества УГП»."""
    st.header("Оценка качества УГП")

    tab_dept, tab_exec = st.tabs(["По подразделениям", "По исполнителям"])

    with tab_dept:
        try:
            stats_dept = _prepare_return_stats(df, DEPT_COL, DEPT_DATE_COL)
            _show_table_and_chart(stats_dept, DEPT_COL, "Возвраты по подразделениям")
        except Exception as e:
            st.error(f"Невозможно рассчитать возвраты по подразделениям: {e}")

    with tab_exec:
        try:
            stats_exec = _prepare_return_stats(df, EXEC_COL, DEPT_DATE_COL)
            _show_table_and_chart(stats_exec, EXEC_COL, "Возвраты по исполнителям")
        except Exception as e:
            st.error(f"Невозможно рассчитать возвраты по исполнителям: {e}")


