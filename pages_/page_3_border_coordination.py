import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

FACT_DATE_COL = "Дата изготовления геоподосновы по факту"
SAPR_LOAD_COL = "Дата загрузки в САПР МГГТ"
APPROVAL_DATE_COL = "Дата согласования границ"
AREA_COL = "Сумма Объем заказа, га"
EXEC_BORDER_COL = "Исполнитель согласования границ факт (АСД сводки)"


def _filter_by_period(df: pd.DataFrame, date_col: str, period_label: str) -> pd.DataFrame:
    if date_col not in df.columns:
        return df

    data = df.copy()
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    data = data.dropna(subset=[date_col])

    today = pd.Timestamp.today().normalize()

    if period_label == "За сегодня":
        mask = data[date_col].dt.date == today.date()
    elif period_label == "Последние 7 дней":
        start = today - pd.Timedelta(days=6)
        mask = (data[date_col] >= start) & (data[date_col] <= today)
    elif period_label == "Месяц":
        start = today.replace(day=1)
        mask = (data[date_col] >= start) & (data[date_col] <= today)
    elif period_label == "Год":
        start = today.replace(month=1, day=1)
        mask = (data[date_col] >= start) & (data[date_col] <= today)
    else:
        return data

    return data.loc[mask]


def _prepare_time_series(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    data = df.copy()
    data[APPROVAL_DATE_COL] = pd.to_datetime(data[APPROVAL_DATE_COL], errors="coerce")
    data = data.dropna(subset=[APPROVAL_DATE_COL])

    if mode == "шт":
        grouped = (
            data.groupby(data[APPROVAL_DATE_COL].dt.date)
            .size()
            .reset_index(name="value")
        )
        value_label = "Количество заказов"
    else:
        if AREA_COL not in data.columns:
            raise ValueError(f"В данных отсутствует столбец '{AREA_COL}'")
        data[AREA_COL] = pd.to_numeric(data[AREA_COL], errors="coerce").fillna(0)
        grouped = (
            data.groupby(data[APPROVAL_DATE_COL].dt.date)[AREA_COL]
            .sum()
            .reset_index(name="value")
        )
        value_label = "Суммарная площадь, га"

    grouped.rename(columns={APPROVAL_DATE_COL: "date"}, inplace=True)
    return grouped, value_label


def _prepare_monthly_summary(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    data = df.copy()
    data[APPROVAL_DATE_COL] = pd.to_datetime(data[APPROVAL_DATE_COL], errors="coerce")
    data = data.dropna(subset=[APPROVAL_DATE_COL])

    data["month"] = data[APPROVAL_DATE_COL].dt.month
    data["year"] = data[APPROVAL_DATE_COL].dt.year
    month_names_ru = {
        1: "Январь",
        2: "Февраль",
        3: "Март",
        4: "Апрель",
        5: "Май",
        6: "Июнь",
        7: "Июль",
        8: "Август",
        9: "Сентябрь",
        10: "Октябрь",
        11: "Ноябрь",
        12: "Декабрь",
    }
    data["month_name"] = data["month"].map(month_names_ru) + " " + data["year"].astype(str)

    if mode == "шт":
        monthly = (
            data.groupby(["year", "month", "month_name"])
            .size()
            .reset_index(name="value")
        )
    else:
        if AREA_COL not in data.columns:
            raise ValueError(f"В данных отсутствует столбец '{AREA_COL}'")
        data[AREA_COL] = pd.to_numeric(data[AREA_COL], errors="coerce").fillna(0)
        monthly = (
            data.groupby(["year", "month", "month_name"])[AREA_COL]
            .sum()
            .reset_index(name="value")
        )

    monthly = monthly.sort_values(["year", "month"])
    return monthly[["month_name", "value"]]


def _prepare_executor_stats(df: pd.DataFrame, mode: str) -> pd.DataFrame:
    if EXEC_BORDER_COL not in df.columns:
        raise ValueError(f"В данных отсутствует столбец '{EXEC_BORDER_COL}'")

    data = df.copy()
    data = data.dropna(subset=[EXEC_BORDER_COL])

    if mode == "шт":
        grouped = (
            data.groupby(EXEC_BORDER_COL)
            .size()
            .reset_index(name="value")
            .sort_values("value", ascending=True)
        )
    else:
        if AREA_COL not in data.columns:
            raise ValueError(f"В данных отсутствует столбец '{AREA_COL}'")
        data[AREA_COL] = pd.to_numeric(data[AREA_COL], errors="coerce").fillna(0)
        grouped = (
            data.groupby(EXEC_BORDER_COL)[AREA_COL]
            .sum()
            .reset_index(name="value")
            .sort_values("value", ascending=True)
        )

    return grouped


def _get_ready_for_upload(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    if FACT_DATE_COL not in data.columns or SAPR_LOAD_COL not in data.columns:
        return pd.DataFrame()

    data[FACT_DATE_COL] = pd.to_datetime(data[FACT_DATE_COL], errors="coerce")
    mask = data[FACT_DATE_COL].notna() & data[SAPR_LOAD_COL].isna()
    return data.loc[mask]


def show_border_coordination(df: pd.DataFrame):
    """Страница «Согласование границ ОГХ»."""
    st.header("Согласование границ ОГХ")

    ready_df = _get_ready_for_upload(df)
    ready_count = len(ready_df)

    st.subheader("Заказы, готовые к загрузке")
    st.metric("Количество заказов готовых к загрузке", ready_count)

    if ready_count > 0:
        st.dataframe(ready_df, use_container_width=True, hide_index=True)
    else:
        st.info("Нет заказов, готовых к загрузке по заданным условиям.")

    st.markdown("---")

    # Фильтры
    st.sidebar.subheader("Фильтры")
    period = st.sidebar.selectbox(
        "Период:",
        ["Все", "За сегодня", "Последние 7 дней", "Месяц", "Год"],
        index=0,
    )

    st.sidebar.subheader("Период (ручной выбор)")
    use_custom_period = st.sidebar.checkbox("Использовать ручной выбор периода")
    date_from = None
    date_to = None
    if use_custom_period:
        default_from = pd.Timestamp.today().replace(day=1).date()
        default_to = pd.Timestamp.today().date()
        date_from = st.sidebar.date_input("Дата от:", value=default_from)
        date_to = st.sidebar.date_input("Дата до:", value=default_to)

    executor_filter = None
    if EXEC_BORDER_COL in df.columns:
        executors = sorted(df[EXEC_BORDER_COL].dropna().unique().tolist())
        if executors:
            executor_filter = st.sidebar.multiselect(
                "Исполнитель:",
                options=executors,
                default=[],
            )

    filtered_df = df.copy()
    if period != "Все" and not use_custom_period:
        filtered_df = _filter_by_period(filtered_df, APPROVAL_DATE_COL, period)

    if use_custom_period and APPROVAL_DATE_COL in filtered_df.columns:
        filtered_df[APPROVAL_DATE_COL] = pd.to_datetime(
            filtered_df[APPROVAL_DATE_COL], errors="coerce"
        )
        if date_from:
            filtered_df = filtered_df[
                filtered_df[APPROVAL_DATE_COL].dt.date >= date_from
            ]
        if date_to:
            filtered_df = filtered_df[filtered_df[APPROVAL_DATE_COL].dt.date <= date_to]

    if executor_filter:
        filtered_df = filtered_df[filtered_df[EXEC_BORDER_COL].isin(executor_filter)]

    mode = st.segmented_control(
        "Режим отображения:",
        options=["шт", "площадь"],
        default="шт",
    )

    if filtered_df.empty:
        st.info("Нет данных для выбранных фильтров.")
        return

    # Временной ряд согласований
    try:
        time_series, value_label = _prepare_time_series(filtered_df, mode)
        fig_ts = px.line(
            time_series,
            x="date",
            y="value",
            markers=True,
            labels={"date": "Дата согласования", "value": value_label},
            title=f"Динамика согласования границ ({mode})",
        )
        st.plotly_chart(fig_ts, use_container_width=True)
    except Exception as exc:
        st.error(f"Не удалось построить график динамики: {exc}")

    # Сводная таблица по месяцам
    try:
        monthly_summary = _prepare_monthly_summary(filtered_df, mode)
        st.subheader(f"Сводная таблица по месяцам ({mode})")
        if mode == "шт":
            monthly_summary.columns = ["Месяц", "Количество, шт"]
        else:
            monthly_summary.columns = ["Месяц", "Площадь, га"]
        st.dataframe(monthly_summary, use_container_width=True, hide_index=True)
    except Exception as exc:
        st.error(f"Не удалось подготовить сводную таблицу по месяцам: {exc}")

    # График по исполнителям
    st.subheader("График согласований по исполнителям")
    if EXEC_BORDER_COL not in filtered_df.columns:
        st.info("В данных отсутствует столбец с исполнителями.")
        return

    exec_data = filtered_df.dropna(subset=[EXEC_BORDER_COL])
    if exec_data.empty:
        st.info("Нет данных по исполнителям для выбранных фильтров.")
        return

    try:
        exec_stats = _prepare_executor_stats(exec_data, mode)
        text_format = (lambda x: f"{x:.0f}") if mode == "шт" else (lambda x: f"{x:.2f}")
        fig_exec = go.Figure(
            data=[
                go.Bar(
                    y=exec_stats[EXEC_BORDER_COL],
                    x=exec_stats["value"],
                    orientation="h",
                    text=[text_format(v) for v in exec_stats["value"]],
                    textposition="auto",
                )
            ]
        )
        xaxis_title = "Количество заказов, шт" if mode == "шт" else "Площадь, га"
        fig_exec.update_layout(
            title=f"График согласований по исполнителям ({mode})",
            xaxis_title=xaxis_title,
            yaxis_title="Исполнитель",
            height=max(400, len(exec_stats) * 40),
        )
        st.plotly_chart(fig_exec, use_container_width=True)
    except Exception as exc:
        st.error(f"Не удалось подготовить график по исполнителям: {exc}")

