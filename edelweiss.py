import streamlit as st
import pandas as pd
import datetime
from datetime import date, datetime
import re
import os
import io

st.set_page_config(
    page_title="Движение товара",
    layout="wide")
def save_goods(new_rows, old_df):
    # если файл уже существует — читаем его
    if old_df is not None:
        final_df = pd.concat([old_df, new_rows], ignore_index=True)
    else:
        final_df = new_rows

    return final_df

def normalize_number(x):
    if pd.isna(x):
        return None

    x = str(x).strip()
    if x == "":
        return None

    # убираем пробелы тысяч
    x = re.sub(r"\s+", "", x)

    # меняем запятую на точку
    x = x.replace(",", ".")

    try:
        return float(x)
    except ValueError:
        return None

    
st.header('Движение товара')
regime = st.segmented_control('', options=['Добавить данные', 'Посмотреть отчёт'], selection_mode='single')




if regime == 'Добавить данные':
    goods = st.file_uploader ('Загрузите goods_data')
    workfile = st.file_uploader ('Загрузите файл')
    if goods: goods_data = pd.read_excel(goods)
    if workfile is not None:
        try:
            table = pd.read_excel (workfile)
            colnames = list(table.columns)
            table.drop(colnames[2], axis =1, inplace=True)
            list_to_find = ['ООО "ТКМ"', 'ООО Статус Юг','накладная', 'поручение', 'Расшифровка', 'Задолженность', 'Итог', 'Документ', 'Номенклатура']
            table = table[~table[colnames[8]].isin(list_to_find)]
            pattern = '|'.join(map(str, list_to_find))
            table = table[~(table[colnames[1]].str.contains(pattern) == True)]
            table = table[~(table[colnames[13]].str.contains('Кол-во') == True)]
            table.dropna(axis=0, inplace=True, how = 'all')
            table.dropna(axis=1, inplace=True, how = 'all')

            colnames = (['good', 'shop','balance_num', 'balance_sum', 'supply_num', 'supply_sum', 'sold_num', 'sold_sum', 'sold_advance_num', 'sold_advance_sum', 'write_off_num', 'write_off_sum', 'end_balance_num', 'end_balance_sum'])
            table.columns = (colnames)
            table['shop'] = table['shop'].ffill()
            table = table[table.notna().sum(axis=1) > 1]
            table.reset_index(drop=True, inplace=True)

            date_start = table['balance_num'][0]
            date_end = table ['end_balance_num'][0]
            table = table[~(table['supply_num'].str.contains('Кол-во') == True)]
            
            st.divider()
            st.subheader('Предпросмотр')

            colnames = ['good', 'shop', 'date_begin', 'date_end', 'balance_num', 'balance_sum', 'sold_num', 'sold_sum', 'supply_num', 'supply_sum', 'end_balance_num', 'end_balance_sum']
            numeric_cols = ['balance_num', 'balance_sum', 'sold_num', 'sold_sum', 'supply_num', 'supply_sum', 'end_balance_num', 'end_balance_sum']
            st.dataframe(table)
            if goods_data is not None:
                df = goods_data
                if date_start in df['date_begin'].values:
                    st.error ('Внимание! Кажется, отчёты за эти даты уже добавлены. Пожалуйста, проверьте файл')


            result = table.copy()
                # нормализуем числа
            for col in numeric_cols:
                    result[col] = result[col].apply(normalize_number)

            result['date_begin'] = date_start
            result['date_end']   = date_end

                # приводим к правильному порядку колонок
            final_rows = result[colnames]
            
            filename = "goods_data.xlsx"
            table = save_goods(final_rows, goods_data)
            buffer = io.BytesIO()
            table.to_excel(buffer, index=False)
            buffer.seek(0)
            if st.download_button(
                label="Скачать",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            ):
                st.success('✅ Данные успешно добавлены в таблицу')

        except Exception as e: 
            st.error (f'❌ Ошибка в файле или данных. Попробуйте загрузить другую таблицу.')
            st.error (f'{e}')





if regime == 'Посмотреть отчёт':
    goods_data = st.file_uploader ('Загрузите общую таблицу')
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    df = pd.read_excel("goods_data.xlsx")
    shop_names = list(df['shop'].unique())
    with col1:
        selected_shop = st.multiselect('Выберите магазин', shop_names)

    col1, col2, col3, col4, col5, col6, col7, col8, col9, col10 = st.columns(10)
    with col1:
        min_date = datetime.strptime("01.01.2000", "%d.%m.%Y").date()
        max_date = date.today()
    
        date_start = st.date_input ('С:', format = 'DD/MM/YYYY',min_value=min_date)
        date_end = st.date_input ('По:', format = 'DD/MM/YYYY', max_value=max_date)
    shop_name = selected_shop
    date_start_str = date_start.strftime("%d.%m.%Y")
    date_end_str   = date_end.strftime("%d.%m.%Y")
    if st.button('Просмотр'):
        st.divider()

        

        # приводим даты
        df['date_begin'] = pd.to_datetime(df['date_begin'], format="%d.%m.%Y")
        df['date_end']   = pd.to_datetime(df['date_end'], format="%d.%m.%Y")

        # фильтрация по магазину и пересечению периодов
        filtered = df[
            (df['shop'].isin(selected_shop)) &
            (df['date_end'] >= pd.to_datetime(date_start)) &
            (df['date_begin'] <= pd.to_datetime(date_end))
        ]

        if filtered.empty:
            st.warning("Нет данных за выбранный период")
        else:
            # 🔥 фактический период
            real_start = filtered['date_begin'].min().strftime("%d.%m.%Y")
            real_end   = filtered['date_end'].max().strftime("%d.%m.%Y")

            st.subheader(
                f'Общий отчёт о движении товаров в {", ".join(map(str, shop_name))} '
                f'за {real_start} – {real_end}'
            )

            # сортируем, чтобы first / last работали корректно
            filtered = filtered.sort_values(['good', 'date_begin'])

            # группируем сначала по good и shop
            tmp = (
                filtered
                .groupby(['good','shop'], as_index=False)
                .agg(
                    balance_num=('balance_num', 'first'),
                    balance_sum=('balance_sum', 'first'),

                    sold_num=('sold_num', 'sum'),
                    sold_sum=('sold_sum', 'sum'),

                    supply_num=('supply_num', 'sum'),
                    supply_sum=('supply_sum', 'sum'),

                    end_balance_num=('end_balance_num', 'last'),
                    end_balance_sum=('end_balance_sum', 'last'),
                )
            )

            # затем объединяем по good
            report = (
                tmp
                .groupby('good', as_index=False)
                .agg(
                    balance_num=('balance_num', 'sum'),
                    balance_sum=('balance_sum', 'sum'),

                    sold_num=('sold_num', 'sum'),
                    sold_sum=('sold_sum', 'sum'),

                    supply_num=('supply_num', 'sum'),
                    supply_sum=('supply_sum', 'sum'),

                    end_balance_num=('end_balance_num', 'sum'),
                    end_balance_sum=('end_balance_sum', 'sum'),
                )
            )

            report = report.rename(columns={
                'good': 'Товар',
                'balance_num': 'Начальный остаток (кол-во)',
                'balance_sum': 'Начальный остаток (сумма)',
                'sold_num': 'Реализация (кол-во)',
                'sold_sum': 'Реализация (сумма)',
                'supply_num': 'Поступление (кол-во)',
                'supply_sum': 'Поступление (сумма)',
                'end_balance_num': 'Конечный остаток (кол-во)',
                'end_balance_sum': 'Конечный остаток (сумма)',
            })
            filename = f'Общий отчёт {real_start}-{real_end} {selected_shop}.xlsx'
            st.dataframe(report, use_container_width=True, height=600)

            buffer = io.BytesIO()
            report.to_excel(buffer, index=False)
            buffer.seek(0)
            if st.download_button(
                label="Скачать",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            ):
                st.success('✅ Отчет сохранен')

            st.subheader(
                f'Отчёт о движении товаров за {real_start} – {real_end}')
            
            tmp = tmp.sort_values('shop')
            filename = f'Отчёт {real_start}-{real_end} {selected_shop}.xlsx'
            tmp = tmp.rename(columns={
                'good': 'Товар',
                'shop': 'Магазин',
                'balance_num': 'Начальный остаток (кол-во)',
                'balance_sum': 'Начальный остаток (сумма)',
                'sold_num': 'Реализация (кол-во)',
                'sold_sum': 'Реализация (сумма)',
                'supply_num': 'Поступление (кол-во)',
                'supply_sum': 'Поступление (сумма)',
                'end_balance_num': 'Конечный остаток (кол-во)',
                'end_balance_sum': 'Конечный остаток (сумма)',
            })

            st.dataframe (tmp.reset_index(drop=True, inplace=True))

            buffer = io.BytesIO()
            tmp.to_excel(buffer, index=False)
            buffer.seek(0)
            if st.download_button(
                label="Скачать",
                data=buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            ):
                st.success('✅ Отчет сохранен')
