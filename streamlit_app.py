import os
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time


st.set_page_config(
   page_title = "spd_statistic",
   layout="centered"
)

st.sidebar.image("galery/logo.jpg")
st.title(":blue[Полнота выдачи Алматы]")


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

today = datetime.today().date()
ten_am = time(10, 0, 0)
time_result = datetime.combine(today, ten_am)
sub_data


@st.cache_data
def processed_data():
    df_pre_clean = sub_data[sub_data['Штрих-код клиента'].notna()]
    df_clean = df_pre_clean.copy()
    df_clean = df_clean[df_clean['Дата доставки'].isna()]
    df_clean = df_clean[~df_clean['Режим'].isin(['М BOX', 'М DOC'])]
    df_clean = df_clean[df_clean['Заказчик'] != 'MAJOR']
    df_clean = df_clean[(df_clean['Регион-отправитель'].isin(['Алматы город','Алматинская область'])) |
                        (df_clean['Регион-получатель'].isin(['Алматы город', 'Алматинская область'])) ]
    df_clean = df_clean[~df_clean['Регион-получатель'].str.contains(r'[A-Za-z]', na=False)]
    df_clean['Статус трекинга'] = df_clean['Статус трекинга'].str.replace('(Возврат/Отмена)', 'Возврат/Отмена',
                                                                          regex=False)
    df_clean['Status'] = df_clean['Статус трекинга'].str.split('(').str[0].str.rstrip()
    df_clean = df_clean[df_clean['Status'].isin(['Забран у отправителя', 'Забран у перевозчика',
              'Курьер вернул на склад','Не доставлен','Неудачная попытка доставки со слов курьера',
              'Перенос','Планируется отправка','Получен складом'])]
    df_clean['City_track_status'] = df_clean['Статус трекинга'].str.extract(r'\((.*?)\)')
    df_clean = df_clean[df_clean['City_track_status'].isin(['Web-службы','Офис в Алматы',
               'Склад в Алматы','Передано по районам РК'])]
    df_clean['Вид отправки'] = np.where(
        df_clean['Регион-получатель'].isin(['Алматы город', 'Алматинская область']), 'Город', 'Межгород')
    df_clean = df_clean[df_clean['Дата/время изменения'] <= time_result]
    df_clean['Месяц'] = df_clean['Дата заказа'].dt.strftime('%B')
    df_clean['Месяц_изменения'] = df_clean['Дата/время изменения'].dt.strftime('%B')
    df_clean['Неделя'] = df_clean['Дата заказа'].dt.strftime('%Y-%W')

    return df_clean

sub_data_new = processed_data() #общая таблица без фильтров

Month = np.unique(sub_data_new['Месяц'])
Status = np.unique(sub_data_new['Status'])
Delivery_type = np.unique(sub_data_new['Вид отправки'])
Type = np.unique(sub_data_new['Режим'])
Region_pickup = np.unique(sub_data_new['Регион-отправитель'])
Region_delivery = np.unique(sub_data_new['Регион-получатель'])
Month_change = np.unique(sub_data_new['Месяц_изменения'])

with st.sidebar:
    chose_month = st.selectbox('Выберите месяц', Month, index=None)
    chose_deliviry_type = st.selectbox('Выберите тип доставки', Delivery_type, index=None)
    chose_status = st.selectbox('Выберите Статус', Status, index=None)
    chose_type = st.selectbox('Выберите вид доставки', Type, index=None)
    chose_pick_up = st.selectbox('Выберите Регион-отправитель', Region_pickup, index=None)
    chose_delivery = st.selectbox('Выберите Регион-получатель', Region_delivery, index=None)
    chose_month_change = st.selectbox('Месяц_изменения', Month_change, index=None)

#Группировка заказов по статусам
inside_city = sub_data_new[(sub_data_new['Вид отправки'] == 'Город') &
                                (sub_data_new['Регион-получатель'] == 'Алматы город')]
if chose_month_change is not None:
    inside_city = inside_city[inside_city['Месяц_изменения']==chose_month_change]
count_inside_city = inside_city.shape[0]
inside_city = pd.pivot_table(inside_city, values='Вид отправки', index='Status', columns='Месяц_изменения',
                          aggfunc='count').reset_index()
month_order = [m for m in ['January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December']
                   if m in inside_city.columns]
order = ['Status']+month_order
inside_city = inside_city[order].reset_index(drop=True)
##########
out_of_city = sub_data_new[(sub_data_new['Вид отправки'] == 'Город') &
                                (sub_data_new['Регион-получатель'] == 'Алматинская область')]
if chose_month_change is not None:
    out_of_city = out_of_city[out_of_city['Месяц_изменения']==chose_month_change]
count_out_of_city = out_of_city.shape[0]
out_of_city = pd.pivot_table(out_of_city, values='Вид отправки', index='Город-получатель', columns='Месяц_изменения',
                          aggfunc='count').reset_index()
month_order = [m for m in ['January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December']
                   if m in out_of_city.columns]
order = ['Город-получатель']+month_order
out_of_city = out_of_city[order].reset_index(drop=True)
############
to_regoins = sub_data_new[sub_data_new['Вид отправки'] == 'Межгород']
if chose_month_change is not None:
    to_regoins = to_regoins[to_regoins['Месяц_изменения']==chose_month_change]
count_to_regoins = to_regoins.shape[0]
to_regoins = pd.pivot_table(to_regoins, values='Вид отправки', index='Регион-получатель',
                        columns='Месяц_изменения', aggfunc='count').reset_index()
month_order = [m for m in ['January', 'February', 'March', 'April', 'May', 'June',
                                'July', 'August', 'September', 'October', 'November', 'December']
                   if m in to_regoins.columns]
order = ['Регион-получатель']+month_order
to_regoins = to_regoins[order].reset_index(drop=True)


col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="Доставки по городу", value=count_inside_city)
    with st.expander("Посмотреть таблицу"):
        st.dataframe(inside_city)
with col2:
    st.metric(label="Доставки по области", value=count_out_of_city)
    with st.expander("Посмотреть таблицу"):
        st.dataframe(out_of_city)
with col3:
    st.metric(label="Доставки со склада", value=count_to_regoins)
    with st.expander("Посмотреть таблицу"):
        st.dataframe(to_regoins)


filters = {
    'Месяц': chose_month,
    'Режим': chose_type,
    'Вид отправки':chose_deliviry_type,
    'Status': chose_status,
    'Регион-отправитель': chose_pick_up,
    'Регион-получатель': chose_delivery,
    'Месяц_изменения': chose_month_change
}

filtered_table = sub_data_new.copy()
for column, value in filters.items():
    if value is not None:
       filtered_table = filtered_table[filtered_table[column] == value]


st.header('Таблица для выгрузки заказов')
st.caption(f"Отфильтровано заказов: {len(filtered_table):,}".replace(",", " "))
filtered_table




