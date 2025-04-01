import os
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from datetime import datetime
from dateutil.relativedelta import relativedelta

#new
# Set the title and favicon that appear in the Browser's tab bar.
st.set_page_config(
   page_title = "spd_statistic",
   layout="centered"
)

st.sidebar.image("galery/logo.jpg")
st.title(":blue[SPD-EX statistics]")


def select_files(repo_path='.'):
    files = []
    for root, _, filenames in os.walk(repo_path):
        for file in filenames:
            if file.endswith('.xlsx'):
                files.append(os.path.join(root, file))
    return files

list_data = select_files() #список файлов к выгрузке


def data_load(files):
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
    df_clean['Неделя'] = df_clean['Дата заказа'].dt.strftime('%Y-%W')
    df_clean['Статус трекинга'] = df_clean['Статус трекинга'].str.replace('(Возврат/Отмена)', 'Возврат/Отмена',
                                                                          regex=False)
    df_clean['City_track_status'] = df_clean['Статус трекинга'].str.extract(r'\((.*?)\)')
    df_clean['Status'] = df_clean['Статус трекинга'].str.split('(').str[0].str.rstrip()

    # df_clean_2 = sub_data.copy()
    # df_clean_2 = df_clean_2[df_clean_2['Штрих-код клиента'].isna() & (df_clean_2['Заказчик'] == 'HALYK FINSERVICE ТОО')][
    #     ['Шифр', 'Дата доставки']]
    # df_clean_2['Штрих-код клиента'] = df_clean_2['Шифр'].str.replace('pickup_', '', regex=False)
    # df_clean_2 = df_clean_2.rename(columns={'Дата доставки': 'Дата доставки HFS'}).drop(columns=['Шифр'])

    #final_clean_data = df_clean.merge(df_clean_2, on='Штрих-код клиента', how='left')
    df_clean['Metric'] = np.where(df_clean['Дата дост. план'].isna(), 'Нет даты плана',
                    np.where(df_clean['Дата доставки'].isna(), 'Нет даты доставки',
                    np.where(df_clean['Дата дост. план'] >= df_clean['Дата доставки'], 'Достигнут SLA',
                    'Превышен SLA')))
    df_clean['Ответственный филиал'] = np.where(df_clean['Ответственный филиал'] == 'Web-службы', 'Авто',
                                                np.where(df_clean['Ответственный филиал'] == 'Передано по районам РК',
                                                         'Аутсорс',
                                                         df_clean['Ответственный филиал']))
    df_delivery = df_clean[~df_clean['Штрих-код клиента'].astype(str).str.endswith('/55')]
    df_return = df_clean[df_clean['Штрих-код клиента'].astype(str).str.endswith('/55')]

    return df_delivery, df_return

sub_data_new, return_data = processed_data() #общая таблица без фильтров


Month = np.unique(sub_data_new['Месяц'])
Type = np.unique(sub_data_new['Type'])
Branch = np.unique(sub_data_new['Ответственный филиал'])
Status = np.unique(sub_data_new['Status'])
Metric = np.unique(sub_data_new['Metric'])


with st.sidebar:
    chose_month = st.selectbox('Выберите месяц', Month, index=None)
    chose_type = st.selectbox('Выберите тип доставки', Type, index=None)
    st.caption('Local - Доставка по РК')
    st.caption('РФ/Major - Доставка по РФ и Major')
    chose_branch = st.selectbox('Выберите Филиал', Branch, index=None)
    chose_status = st.selectbox('Выберите Статус', Status, index=None)
    chose_metric = st.selectbox('Выберите метрику', Metric, index=None)

def group_table(table, group_column:str):
    group_table = table.copy()
    filters = {
        'Metric': chose_metric,
        'Type': chose_type,
        'Ответственный филиал': chose_branch,
        'Status': chose_status
    }
    for column, value in filters.items():
        if value is not None:
            group_table = group_table[group_table[column] == value]
    group_table = group_table.pivot_table(index=['Type', 'Status'], columns=group_column, values='Количество',
                                                  aggfunc='count', fill_value=0).reset_index()
    month_order = [m for m in ['January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December']
                   if m in group_table.columns]
    group_table['Total'] = group_table.iloc[:, 2:].sum(axis=1)
    group_table = group_table.sort_values(by=['Type', 'Total'], ascending=[True, False])
    group_table = group_table.drop(columns=['Total']).reset_index(drop=True)
    if group_column == 'Месяц':
        order = ['Type', 'Status'] + month_order
        group_table = group_table[order].reset_index(drop=True)

    return group_table

group_data = group_table(sub_data_new,'Месяц') #Группированная таблица по филиалу
group_data_week = group_table(sub_data_new,'Неделя')

def sla_calculation():
    sla = sub_data_new.copy()
    filters = {
        'Metric': chose_metric,
        'Месяц':chose_month,
        'Type': chose_type,
        'Ответственный филиал': chose_branch,
        'Status': chose_status
    }
    for column, value in filters.items():
        if value is not None:
            sla = sla[sla[column] == value]
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
    #current_month = datetime.now().strftime('%B') #поставить верный фильтр
    current_month = (datetime.now() - relativedelta(months=1)).strftime('%B')
    last_month = (datetime.now() - relativedelta(months=2)).strftime('%B')
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

st.caption('Ниже приведены показатели SLA за текущий месяц в сравнений с прошлым месяцем. В расчете учитываются только досталенные заказы по которым указана информация: :orange[Дата Доставки и Дата Плана доставки] ')
st.caption('Для сокращения выборки выберите определенный филиал и уберите фильтр месяца для сравнения с прошлым месяцем')
num_cols = 5
rows = [data_KPI.iloc[i:i+num_cols] for i in range(0, len(data_KPI), num_cols)]

for row in rows:
    cols = st.columns(len(row))
    for i, (_, metric) in enumerate(row.iterrows()):
        with cols[i]:
            st.metric(label=metric["Ответственный филиал"], value=f"{metric['SLA']}%", delta=f"{metric['Change']}%")

if chose_branch is None:
    st.header(':blue[Статусы по всем доставкам]')
else:
    st.header(f':blue[Статусы {chose_branch}]')

if "show_week_info" not in st.session_state:
    st.session_state.show_week_info = False
if st.button('Переключение Неделя/Месяц'):
    st.session_state.show_week_info = not st.session_state.show_week_info

if st.session_state.show_week_info:
    st.write("Показаны данные по неделям")
    st.write(group_data_week)  # Отображение данных по неделям
else:
    st.write("Показаны общие данные")
    st.write(group_data)  # Отображение общих данных

st.header('Все метрики по филиалам')
st.caption('Ниже приведены метрики по филиалам. Указаны средние значения, для выбора точного значения выберите расчетный месяц. :orange[Информация представлена в %]')
grafic_data = grapfic_sla()
if "show_count" not in st.session_state:
    st.session_state.show_count = False
if st.button('Переключение Доля/Количество'):
    st.session_state.show_count = not st.session_state.show_count
if st.session_state.show_count:
    st.caption("Приведена доля заказов")
    grafic_source=grafic_data[['Ответственный филиал', 'Metric','Доля']]
    chart = alt.Chart(grafic_source).mark_bar(size=15).encode(
        x='Metric:N',
        y='Доля:Q',
        color='Ответственный филиал:N',
        column=alt.Column('Ответственный филиал:N', title=None)
    ).properties(width=100)
    st.altair_chart(chart)
else:
    st.caption("Приведено количество заказов")
    grafic_source=grafic_data[['Ответственный филиал', 'Metric','total_sum']]
    chart = alt.Chart(grafic_source).mark_bar(size=15).encode(
        x='Metric:N',
        y='total_sum:Q',
        color='Ответственный филиал:N',
        column=alt.Column('Ответственный филиал:N', title=None)
    ).properties(width=100)
    st.altair_chart(chart)


st.header(':blue[Таблица SLA]')
st.caption('Ниже представлена агрегированная информция по метрикам и филиалам')
st.write(data_sla)


def show_orders():
    df = sub_data_new.copy()
    filters = {
        'Metric': chose_metric,
        'Месяц': chose_month,
        'Type': chose_type,
        'Ответственный филиал': chose_branch,
        'Status': chose_status
    }
    for column, value in filters.items():
        if value is not None:
            df = df[df[column] == value]
    df = df[
        ['Заказ', 'Дата заказа', 'Штрих-код клиента', 'Заказчик', 'Город-отправитель', 'Город-получатель',
         'Дата доставки', 'Дата/время изменения', 'Status', 'Type', 'Metric', 'Ответственный филиал']]
    df['Дата заказа'] = pd.to_datetime(df['Дата заказа'], errors='coerce').dt.date
    df['Дата доставки'] = pd.to_datetime(df['Дата доставки'], errors='coerce').dt.date
    df['Статус не менялся'] = np.where(
        df['Дата доставки'].notna(),
        np.nan,
        (pd.Timestamp.now().normalize() - df['Дата/время изменения']).dt.days)
    return df


data_orders = show_orders()
st.header('Таблица для выгрузки заказов')
st.write(data_orders)



def return_table():
    df_return = return_data.copy()
    filters = {
        'Metric': chose_metric,
        'Месяц': chose_month,
        'Type': chose_type,
        'Ответственный филиал': chose_branch,
        'Status': chose_status
    }
    for column, value in filters.items():
        if value is not None:
            df_return = df_return[df_return[column] == value]
    df_return = df_return[
        ['Заказ', 'Дата заказа', 'Штрих-код клиента', 'Заказчик', 'Город-отправитель', 'Город-получатель',
         'Дата доставки', 'Дата/время изменения', 'Status', 'Type', 'Metric', 'Ответственный филиал']]
    df_return['Дата заказа'] = pd.to_datetime(df_return['Дата заказа'], errors='coerce').dt.date
    df_return['Дата доставки'] = pd.to_datetime(df_return['Дата доставки'], errors='coerce').dt.date
    df_return['Статус не менялся'] = np.where(
        df_return['Дата доставки'].notna(),
        np.nan,
        (pd.Timestamp.now().normalize() - df_return['Дата/время изменения']).dt.days)
    return df_return

data_return_proccesed = return_table()

def returns():
    df_1 = data_orders.groupby('Ответственный филиал') \
        .agg({'Заказ': 'count'}) \
        .rename(columns={'Заказ': 'D_count'}).reset_index()
    df_2 = data_return_proccesed.groupby('Ответственный филиал') \
        .agg({'Заказ': 'count'}) \
        .rename(columns={'Заказ': 'R_count'}).reset_index()
    df_reverse = df_1.merge(df_2, how='left', on='Ответственный филиал')
    df_reverse['Доля возвратов'] = (df_reverse['R_count'] / df_reverse['D_count'] * 100).round(2)
    df_reverse = df_reverse.sort_values('Доля возвратов', ascending=False)
    df_reverse = df_reverse.rename(columns={'D_count': 'Количество заказов',
                                   'R_count': 'Количество возвратов'})
    return df_reverse

data_reverse = returns()
st.header('Доля возвратов')
st.write(data_reverse)

return_df_all = group_table(return_data,'Месяц') #Группированная таблица по филиалу
return_df_week = group_table(return_data,'Неделя')
st.header('Статусы по возвратам')
if st.session_state.show_week_info:
    st.write("Показаны данные по неделям")
    st.write(return_df_week)  # Отображение данных по неделям
else:
    st.write("Показаны общие данные")
    st.write(return_df_all)  # Отображение общих данных

st.header('Таблица возвратов')
st.write(data_return_proccesed)
