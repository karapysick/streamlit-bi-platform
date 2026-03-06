import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


FACT_DATE_COL = "Дата изготовления геоподосновы по факту"
AREA_COL = "Сумма Объем заказа, га"
EXEC_AREA_COL = "Исполнитель определения площадных и количественных характеристик (АСД сводки)"


def _prepare_daily_counts(df: pd.DataFrame, date_col: str) -> pd.DataFrame:
    if date_col not in df.columns:
        raise ValueError(f"В данных отсутствует столбец '{date_col}'")

    data = df.copy()
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    data = data.dropna(subset=[date_col])

    daily = data.groupby(data[date_col].dt.date).size().reset_index(name="orders_count")
    daily.rename(columns={date_col: "date"}, inplace=True)
    return daily


def _prepare_daily_area(df: pd.DataFrame, date_col: str, area_col: str) -> pd.DataFrame:
    missing = [c for c in [date_col, area_col] if c not in df.columns]
    if missing:
        raise ValueError(f"В данных отсутствуют столбцы: {', '.join(missing)}")

    data = df.copy()
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    data = data.dropna(subset=[date_col])

    data[area_col] = pd.to_numeric(data[area_col], errors="coerce").fillna(0)
    daily = data.groupby(data[date_col].dt.date)[area_col].sum().reset_index()
    daily.rename(columns={date_col: "date", area_col: "area_sum"}, inplace=True)
    return daily


def _prepare_monthly_summary(df: pd.DataFrame, date_col: str, area_col: str, mode: str) -> pd.DataFrame:
    """Подготовка сводной таблицы по месяцам."""
    # Словарь для русских названий месяцев
    month_names_ru = {
        1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
        5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
        9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
    }
    
    data = df.copy()
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    data = data.dropna(subset=[date_col])
    
    data['month'] = data[date_col].dt.month
    data['year'] = data[date_col].dt.year
    data['month_name_ru'] = data['month'].map(month_names_ru)
    data['month_name'] = data['month_name_ru'] + ' ' + data['year'].astype(str)
    
    if mode == "шт":
        monthly = data.groupby(['year', 'month', 'month_name']).size().reset_index(name='value')
    else:  # площадь
        data[area_col] = pd.to_numeric(data[area_col], errors="coerce").fillna(0)
        monthly = data.groupby(['year', 'month', 'month_name'])[area_col].sum().reset_index()
        monthly.rename(columns={area_col: 'value'}, inplace=True)
    
    monthly = monthly.sort_values(['year', 'month'])
    return monthly[['month_name', 'value']]


def _prepare_date_status_counts(df: pd.DataFrame, date_col: str, area_col: str, mode: str) -> pd.DataFrame:
    """Подготовка данных для диаграммы: количество или площадь с датой и без даты."""
    data = df.copy()
    
    # Проверяем наличие даты
    data[date_col] = pd.to_datetime(data[date_col], errors="coerce")
    
    if mode == "шт":
        # Считаем количество записей
        with_date = data[date_col].notna().sum()
        without_date = data[date_col].isna().sum()
        value_col = "Количество"
        unit = "шт"
    else:  # площадь
        # Считаем сумму площадей
        if area_col not in data.columns:
            raise ValueError(f"В данных отсутствует столбец '{area_col}'")
        data[area_col] = pd.to_numeric(data[area_col], errors="coerce").fillna(0)
        with_date = data[data[date_col].notna()][area_col].sum()
        without_date = data[data[date_col].isna()][area_col].sum()
        value_col = "Площадь"
        unit = "га"
    
    result = pd.DataFrame({
        'Статус': ['Залито', 'Осталось заполнить'],
        value_col: [with_date, without_date]
    })
    
    return result, unit


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


def show_area_characteristics(df: pd.DataFrame):
    """Страница «Определение площадных характеристик»."""
    st.header("Определение площадных характеристик")

    # Фильтры в sidebar
    st.sidebar.subheader("Фильтры")
    
    # Фильтр по периоду (предустановленные периоды)
    period = st.sidebar.selectbox(
        "Период:",
        ["Все", "За сегодня", "Последние 7 дней", "Месяц", "Год"],
        index=0,
    )

    # Фильтр по датам (ручной выбор)
    st.sidebar.subheader("Период (ручной выбор)")
    use_custom_period = st.sidebar.checkbox("Использовать ручной выбор периода")
    
    date_from = None
    date_to = None
    if use_custom_period:
        date_from = st.sidebar.date_input("Дата от:", value=None)
        date_to = st.sidebar.date_input("Дата до:", value=None)

    # Фильтр по исполнителю
    executor_filter = None
    if EXEC_AREA_COL in df.columns:
        executors = sorted(df[EXEC_AREA_COL].dropna().unique().tolist())
        if executors:
            executor_filter = st.sidebar.multiselect(
                "Исполнитель:",
                options=executors,
                default=[]
            )

    # Применяем фильтры
    filtered_df = df.copy()
    
    # Фильтр по периоду (предустановленные)
    if period != "Все" and not use_custom_period:
        filtered_df = _filter_by_period(filtered_df, FACT_DATE_COL, period)
    
    # Фильтр по датам (ручной выбор)
    if use_custom_period:
        if FACT_DATE_COL in filtered_df.columns:
            filtered_df[FACT_DATE_COL] = pd.to_datetime(filtered_df[FACT_DATE_COL], errors="coerce")
            if date_from:
                filtered_df = filtered_df[filtered_df[FACT_DATE_COL].dt.date >= date_from]
            if date_to:
                filtered_df = filtered_df[filtered_df[FACT_DATE_COL].dt.date <= date_to]
    
    # Фильтр по исполнителю
    if executor_filter and len(executor_filter) > 0:
        filtered_df = filtered_df[filtered_df[EXEC_AREA_COL].isin(executor_filter)]

    # Переключение между шт и площадями через segmented_control
    mode = st.segmented_control(
        "Режим отображения:",
        options=["шт", "площадь"],
        default="шт"
    )

    # Диаграмма с двумя показателями: количество или площадь с датой и без даты
    try:
        date_status_data, unit = _prepare_date_status_counts(filtered_df, FACT_DATE_COL, AREA_COL, mode)
        
        # Определяем название столбца со значениями
        value_col = "Количество" if mode == "шт" else "Площадь"
        
        # Форматирование значений в зависимости от режима
        if mode == "шт":
            text_template = f'%{{label}}<br>%{{value:.0f}} {unit}<br>(%{{percent}})'
        else:
            text_template = f'%{{label}}<br>%{{value:.2f}} {unit}<br>(%{{percent}})'
        
        fig_status = go.Figure(data=[go.Pie(
            labels=date_status_data['Статус'],
            values=date_status_data[value_col],
            hole=0.3,
            marker=dict(colors=['#1f77b4', '#ff7f0e']),
            textinfo='label+percent+value',
            texttemplate=text_template
        )])
        
        title_suffix = f" ({mode})" if mode == "площадь" else " (шт)"
        fig_status.update_layout(
            title=f"Количество {mode} с датой и без даты по столбцу 'Дата изготовления геоподосновы по факту'{title_suffix}",
            showlegend=True
        )
        
        st.plotly_chart(fig_status, use_container_width=True)
    except Exception as e:
        st.error(f"Невозможно построить диаграмму статуса дат: {e}")

    # Сводная таблица по месяцам
    try:
        monthly_summary = _prepare_monthly_summary(filtered_df, FACT_DATE_COL, AREA_COL, mode)
        st.subheader(f"Сводная таблица по месяцам ({mode})")
        
        if mode == "шт":
            monthly_summary.columns = ["Месяц", "Количество, шт"]
        else:
            monthly_summary.columns = ["Месяц", "Площадь, га"]
        
        st.dataframe(monthly_summary, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Невозможно подготовить сводную таблицу по месяцам: {e}")

    # График по исполнителям (перенесен под диаграмму)
    st.subheader("График закрытия операций по исполнителю")
    
    if EXEC_AREA_COL not in df.columns:
        st.error(f"В данных отсутствует столбец '{EXEC_AREA_COL}'")
    else:
        if filtered_df.empty:
            st.info("Нет данных для выбранного периода.")
        else:
            data = filtered_df.copy()
            data[FACT_DATE_COL] = pd.to_datetime(data[FACT_DATE_COL], errors="coerce")
            data = data.dropna(subset=[FACT_DATE_COL])

            # Группируем по исполнителям в зависимости от режима
            if mode == "шт":
                # Считаем количество заказов
                exec_grouped = (
                    data.groupby(EXEC_AREA_COL)
                    .size()
                    .reset_index(name="value")
                    .sort_values("value", ascending=True)
                )
                xaxis_title = "Количество заказов, шт"
                text_format = lambda x: f"{x:.0f}"
            else:  # площадь
                # Считаем сумму площадей
                if AREA_COL not in data.columns:
                    st.error(f"В данных отсутствует столбец '{AREA_COL}'")
                    return
                data[AREA_COL] = pd.to_numeric(data[AREA_COL], errors="coerce").fillna(0)
                exec_grouped = (
                    data.groupby(EXEC_AREA_COL)[AREA_COL]
                    .sum()
                    .reset_index(name="value")
                    .sort_values("value", ascending=True)
                )
                xaxis_title = "Площадь, га"
                text_format = lambda x: f"{x:.2f}"
            
            # Используем plotly для горизонтального bar chart
            fig_exec = go.Figure(data=[go.Bar(
                y=exec_grouped[EXEC_AREA_COL],
                x=exec_grouped["value"],
                orientation='h',
                text=[text_format(v) for v in exec_grouped["value"]],
                textposition='auto',
            )])
            
            fig_exec.update_layout(
                title=f"График закрытия операций по исполнителю ({mode})",
                xaxis_title=xaxis_title,
                yaxis_title="Исполнитель",
                height=max(400, len(exec_grouped) * 40)
            )
            
            st.plotly_chart(fig_exec, use_container_width=True)


