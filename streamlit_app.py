import os
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime
from dateutil.relativedelta import relativedelta


# Set the title and favicon that appear in the Browser's tab bar.
st.set_page_config(
   page_title = "spd_statistic",
   layout="centered"
)

st.sidebar.image("galery/logo.jpg")
st.title(":blue[SPD-EX statistics]")


def select_files(repo_path='.'):
    files = []
    for root, _, filenames in os.walk(repo_path):  # Проходим по всем папкам репо
        for file in filenames:
            if file.endswith('.xlsx'):
                files.append(os.path.join(root, file))
    return files

list_data = select_files() #список файлов к выгрузке


def data_load(files):
    """Загружает и объединяет данные из списка Excel-файлов."""
    data_frames = []
    for file in files:
        try:
            df = pd.read_excel(file, engine="openpyxl")
            data_frames.append(df)
        except Exception as e:
            print(f"Ошибка при загрузке {file}: {e}")
    return pd.concat(data_frames, ignore_index=True) if data_frames else pd.DataFrame()

sub_data = data_load(list_data)  #Исходные данные

@st.cache_data
def processed_data():
    df_clean = sub_data[sub_data['Штрих-код клиента'].notna()]
    df_clean = df_clean.copy()
    df_clean['Type'] = np.where(
        (df_clean['Режим'].isin(['М BOX', 'М DOC'])) | (df_clean['Заказчик'] == 'MAJOR'), 'РФ/Major', 'Local')
    df_clean['Месяц'] = df_clean['Дата заказа'].dt.strftime('%B')
    df_clean['Статус трекинга'] = df_clean['Статус трекинга'].str.replace('(Возврат/Отмена)', 'Возврат/Отмена',
                                                                          regex=False)
    df_clean['City_track_status'] = df_clean['Статус трекинга'].str.extract(r'\((.*?)\)')
    df_clean['Status'] = df_clean['Статус трекинга'].str.split('(').str[0].str.rstrip()

    df_clean_2 = sub_data.copy()
    df_clean_2 = df_clean_2[df_clean_2['Штрих-код клиента'].isna() & (df_clean_2['Заказчик'] == 'HALYK FINSERVICE ТОО')][
        ['Шифр', 'Дата доставки']]
    df_clean_2['Штрих-код клиента'] = df_clean_2['Шифр'].str.replace('pickup_', '', regex=False)
    df_clean_2 = df_clean_2.rename(columns={'Дата доставки': 'Дата доставки HFS'}).drop(columns=['Шифр'])

    final_clean_data = df_clean.merge(df_clean_2, on='Штрих-код клиента', how='left')
    return final_clean_data

sub_data_new = processed_data() #общая таблица без фильтров


Month = np.unique(sub_data_new['Месяц'])
Type = np.unique(sub_data_new['Type'])
Branch = np.unique(sub_data_new['Ответственный филиал'])


with st.sidebar:
    chose_month = st.selectbox('Выберите месяц', Month, index=None)
    chose_type = st.selectbox('Выберите тип доставки', Type, index=None)
    st.caption('Local - Доставка по РК')
    st.caption('РФ/Major - Доставка по РФ и Major')
    chose_branch = st.selectbox('Выберите Филиал', Branch, index=None)



def group_first_table():
    if chose_branch is None:
        month_order = sorted(sub_data_new['Месяц'].unique(),
                             key=lambda x: ['January', 'February', 'March', 'April', 'May', 'June',
                                            'July', 'August', 'September', 'October', 'November',
                                            'December'].index(x))
        group_table = sub_data_new.pivot_table(index=['Type', 'Status'], columns='Месяц', values='Количество',
                                  aggfunc='count', fill_value=0).reset_index()
        return group_table,month_order
    else:
        group_table = sub_data_new[sub_data_new['Ответственный филиал']==chose_branch]
        month_order = sorted(group_table['Месяц'].unique(),
                             key=lambda x: ['January', 'February', 'March', 'April', 'May', 'June',
                                            'July', 'August', 'September', 'October', 'November',
                                            'December'].index(x))
        group_table = group_table.pivot_table(index=['Type', 'Status'], columns='Месяц', values='Количество',
                                aggfunc='count', fill_value=0).reset_index()
        return group_table,month_order


group_data, month_order = group_first_table() #Группированная таблица по филиалу

def group_second_table():
    second_group_table = group_data.copy()
    second_group_table['Total'] = second_group_table.iloc[:, 2:].sum(axis=1)
    second_group_table = second_group_table.sort_values(by=['Type', 'Total'], ascending=[True, False])
    second_group_table = second_group_table.drop(columns=['Total']).reset_index(drop=True)
    order = ['Type', 'Status'] + month_order
    second_group_table = second_group_table[order].reset_index(drop=True)
    if chose_type is None:
        return second_group_table
    else:
        second_group_table = second_group_table[second_group_table['Type'] == chose_type]
        return second_group_table

final_group_data = group_second_table() #Группирвоаная итогоая таблица с учетом филиала и типа доставки


def sla_calculation():
    sla = sub_data_new.copy()
    if chose_type is None:
        sla['Metric'] = np.where(sla['Дата доставки'].isna(), 'Нет даты доставки',
                        np.where(sla['Дата дост. план'].isna(), 'Нет даты плана',
                        np.where(sla['Дата дост. план'] >= sla['Дата доставки'], 'Достигнут SLA',
                        'Превышен SLA')))
        sla_agg = sla.groupby(['Месяц', 'Ответственный филиал', 'Metric']) \
            .agg({'Заказ': 'count'}) \
            .rename(columns={'Заказ': 'Количество'}).reset_index()
        sla_agg_total = sla.groupby(['Месяц', 'Ответственный филиал']) \
            .agg({'Заказ': 'count'}) \
            .rename(columns={'Заказ': 'total_count'}).reset_index()
        sla_agg = sla_agg.merge(sla_agg_total, how='left', on=['Месяц', 'Ответственный филиал'])
        sla_agg['Доля'] = (sla_agg['Количество'] / sla_agg['total_count']).round(2)
        sla_agg = sla_agg.drop(columns=['total_count'])
        month_order = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ]
        sla_agg['Месяц'] = pd.Categorical(sla_agg['Месяц'], categories=month_order, ordered=True)
        sla_agg = sla_agg.sort_values(by=['Ответственный филиал', 'Месяц']).reset_index(drop=True)
        if chose_branch is None and chose_month is None:
            return sla_agg
        elif chose_branch is None and chose_month is not None:
            sla_agg = sla_agg[sla_agg['Месяц'] == chose_month]
            return sla_agg
        elif chose_branch is not None and chose_month is None:
            sla_agg = sla_agg[sla_agg['Ответственный филиал'] == chose_branch]
            return sla_agg
        else:
            sla_agg = sla_agg[sla_agg['Месяц'] == chose_month]
            sla_agg = sla_agg[sla_agg['Ответственный филиал'] == chose_branch]
            return sla_agg
    else:
        sla = sla[sla['Type'] == chose_type]
        sla['Metric'] = np.where(sla['Дата доставки'].isna(), 'Нет даты доставки',
                        np.where(sla['Дата дост. план'].isna(), 'Нет даты плана',
                        np.where(sla['Дата дост. план'] >= sla['Дата доставки'], 'Достигнут SLA',
                        'Превышен SLA')))
        sla_agg = sla.groupby(['Месяц', 'Ответственный филиал', 'Metric']) \
            .agg({'Заказ': 'count'}) \
            .rename(columns={'Заказ': 'Количество'}).reset_index()
        sla_agg_total = sla.groupby(['Месяц', 'Ответственный филиал']) \
            .agg({'Заказ': 'count'}) \
            .rename(columns={'Заказ': 'total_count'}).reset_index()
        sla_agg = sla_agg.merge(sla_agg_total, how='left', on=['Месяц', 'Ответственный филиал'])
        sla_agg['Доля'] = (sla_agg['Количество'] / sla_agg['total_count']).round(2)
        sla_agg = sla_agg.drop(columns=['total_count'])
        month_order = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ]
        sla_agg['Месяц'] = pd.Categorical(sla_agg['Месяц'], categories=month_order, ordered=True)
        sla_agg = sla_agg.sort_values(by=['Ответственный филиал', 'Месяц']).reset_index(drop=True)
        if chose_branch is None and chose_month is None:
            return sla_agg
        elif chose_branch is None and chose_month is not None:
            sla_agg = sla_agg[sla_agg['Месяц'] == chose_month]
            return sla_agg
        elif chose_branch is not None and chose_month is None:
            sla_agg = sla_agg[sla_agg['Ответственный филиал'] == chose_branch]
            return sla_agg
        else:
            sla_agg = sla_agg[sla_agg['Месяц'] == chose_month]
            sla_agg = sla_agg[sla_agg['Ответственный филиал'] == chose_branch]
            return sla_agg

data_sla = sla_calculation() #Таблица SLA с учетом всех фильтров

def grapfic_sla():
    show_sla = data_sla.groupby(['Ответственный филиал', 'Metric']) \
        .agg({'Количество': 'sum'}) \
        .rename(columns={'Количество': 'total_sum'}).reset_index()
    show_sla_sec = data_sla.groupby(['Ответственный филиал']) \
        .agg({'Количество': 'sum'}) \
        .rename(columns={'Количество': 'total_sum_total'}).reset_index()
    show_sla = show_sla.merge(show_sla_sec, how='left', on='Ответственный филиал')
    show_sla['Доля'] = (show_sla['total_sum'] / show_sla['total_sum_total']).round(2)
    show_sla = show_sla.sort_values(by=['total_sum_total', 'Ответственный филиал'], ascending=False).drop(
        columns=['total_sum_total'])
    return show_sla

def kpi_metric():
    current_month = datetime.now().strftime('%B')
    last_month = (datetime.now() - relativedelta(months=1)).strftime('%B')
    kpi_data = data_sla[(data_sla['Месяц']==current_month) & (data_sla['Metric'].isin(['Достигнут SLA', 'Превышен SLA']))] \
        .drop(columns=['Месяц', 'Доля'])
    kpi_data_2 = data_sla[(data_sla['Месяц']==last_month) & (data_sla['Metric'].isin(['Достигнут SLA', 'Превышен SLA']))] \
        .drop(columns=['Месяц', 'Доля']) \
        .rename(columns={'Количество': 'last_Количество'})
    kpi_data = kpi_data.merge(kpi_data_2, how='left', on = ['Ответственный филиал', 'Metric'])
    kpi_data_group = kpi_data.groupby('Ответственный филиал') \
        .agg({'Количество': 'sum','last_Количество':'sum'}) \
        .rename(columns={'Количество': 'Total_sum', 'last_Количество':'Total_sum_last'}).reset_index()
    kpi_data = kpi_data.merge(kpi_data_group, how='left', on='Ответственный филиал')
    kpi_data['SLA'] = ((kpi_data['Количество']/kpi_data['Total_sum']).round(2)*100).round(2)
    kpi_data['SLA_last'] = ((kpi_data['last_Количество'] / kpi_data['Total_sum_last']).round(2)*100).round(2)
    kpi_data['Change'] = ((kpi_data['SLA'] / kpi_data['SLA_last'] - 1).round(2)*100).round(2)
    kpi_data = kpi_data[kpi_data['Metric']=='Достигнут SLA'][['Ответственный филиал','SLA','Change']] \
        .sort_values(by='SLA', ascending=False).reset_index(drop=True)
    return kpi_data

data_KPI = kpi_metric()

st.caption('Ниже приведены показатели SLA за текущий месяц в сравнений с прошлым месяцем. В расчете учитываются только досталенные заказы по которым указана инфомрация:')
st.caption(':orange[Дата Доставки и Дата Плана доставки] Для сокращения выборки выберите определенный филиал')
num_cols = 5
rows = [data_KPI.iloc[i:i+num_cols] for i in range(0, len(data_KPI), num_cols)]

for row in rows:
    cols = st.columns(len(row))
    for i, (_, metric) in enumerate(row.iterrows()):
        with cols[i]:
            st.metric(label=metric["Ответственный филиал"], value=f"{metric['SLA']}%", delta=f"{metric['Change']}%")

if chose_branch is None:
    st.markdown(':blue[Статусы по всем доставкам]')
else:
    st.markdown(f':blue[Статусы {chose_branch}]')
st.caption('Ниже приведены статусы по доставкам. Для сокращения выборки выберите определенный филиал')
final_group_data

st.header('Все метрики по филиалам')
st.caption('Ниже приведены метрики по филиалам. Указаны средние значения, для выбора точного значения выберите расчетный месяц. :orange[Информация представлена в %]')
grafic_data = grapfic_sla()
grafic_source=grafic_data[['Ответственный филиал', 'Metric','Доля']]
chart = alt.Chart(grafic_source).mark_bar(size=15).encode(
    x='Metric:N',
    y='Доля:Q',
    color='Ответственный филиал:N',
    column=alt.Column('Ответственный филиал:N', title=None)
).properties(width=100)
st.altair_chart(chart)

st.header(':blue[Таблица SLA]')
st.caption('Ниже представлена агрегированная информция по метрикам и филиалам')
data_sla


def show_orders():
    if chose_branch is not None and chose_month is not None:
        df = sub_data_new[sub_data_new['Ответственный филиал']==chose_branch]
        df = sub_data_new[sub_data_new['Месяц'] == chose_month]
        df['Start'] = df[['Дата доставки', 'Дата доставки HFS']].bfill(axis=1).iloc[:, 0]
        df = df[
            ['Заказ', 'Дата заказа', 'Штрих-код клиента', 'Заказчик', 'Город-отправитель', 'Город-получатель', 'Start',
             'Дата доставки', 'Дата/время изменения', 'Status', 'Type']]
        df['Дата заказа'] = pd.to_datetime(df['Дата заказа'], errors='coerce').dt.date
        df['Дата доставки'] = pd.to_datetime(df['Дата доставки'], errors='coerce').dt.date
        df['Start'] = pd.to_datetime(df['Start'], errors='coerce').dt.date
        df['Статус не менялся'] = np.where(
            df['Дата доставки'].notna(),
            np.nan,
            (pd.Timestamp.now().normalize() - df['Дата/время изменения']).dt.days)
        return df
    elif chose_branch is not None and chose_month is None:
        df = sub_data_new[sub_data_new['Ответственный филиал'] == chose_branch]
        df['Start'] = df[['Дата доставки', 'Дата доставки HFS']].bfill(axis=1).iloc[:, 0]
        df = df[
            ['Заказ', 'Дата заказа', 'Штрих-код клиента', 'Заказчик', 'Город-отправитель', 'Город-получатель', 'Start',
             'Дата доставки', 'Дата/время изменения', 'Status', 'Type']]
        df['Дата заказа'] = pd.to_datetime(df['Дата заказа'], errors='coerce').dt.date
        df['Дата доставки'] = pd.to_datetime(df['Дата доставки'], errors='coerce').dt.date
        df['Start'] = pd.to_datetime(df['Start'], errors='coerce').dt.date
        df['Статус не менялся'] = np.where(
            df['Дата доставки'].notna(),
            np.nan,
            (pd.Timestamp.now().normalize() - df['Дата/время изменения']).dt.days)
        return df
    else:
        pass

data_orders = show_orders()
st.header('Таблица для выгрузки заказов')
if chose_branch is None:
    st.markdown(f'Выберите филиал для отображения заказов')
    st.table(data=None)
else:
    st.markdown(f'Доставки филиала: {chose_branch}')
    data_orders
