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





        if st.button('Просмотр'):
            st.divider()

            # Приводим даты к datetime (Excel может хранить их в разных форматах)
            df['date_begin'] = pd.to_datetime(df['date_begin'], errors='coerce', dayfirst=True)
            df['date_end']   = pd.to_datetime(df['date_end'], errors='coerce', dayfirst=True)

            # Парсим введённые пользователем даты
            date_start_dt = pd.to_datetime(date_start, errors='coerce', dayfirst=True)
            date_end_dt   = pd.to_datetime(date_end, errors='coerce', dayfirst=True)

            # Фильтрация по выбранным магазинам и пересечению периодов
            filtered = df[
                (df['shop'].isin(selected_shop)) &
                (df['date_end'] >= date_start_dt) &
                (df['date_begin'] <= date_end_dt)
            ]

            if filtered.empty:
                st.warning("Нет данных за выбранный период")
            else:
                # Фактический период в фильтрованных данных
                real_start = filtered['date_begin'].min().strftime("%d.%m.%Y")
                real_end   = filtered['date_end'].max().strftime("%d.%m.%Y")

                st.subheader(
                    f'Общий отчёт о движении товаров в {", ".join(map(str, selected_shop))} '
                    f'за {real_start} – {real_end}'
                )

                # Сортируем для корректной агрегации
                filtered = filtered.sort_values(['good', 'date_begin'])

                # Сначала группируем по товару и магазину
                tmp = (
                    filtered
                    .groupby(['good', 'shop'], as_index=False)
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

                # Затем агрегируем по товару
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

                # Переименовываем колонки для финального отчёта
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

                # Формируем имя файла
                filename_total = f'Общий отчёт {real_start}-{real_end} {"_".join(selected_shop)}.xlsx'
                if len(filename_total) > 100:
                    filename_total = f'Общий отчёт {real_start}-{real_end} {selected_shop[0][:3]}_итд.xlsx'

                # Отображаем итоговый отчёт
                st.dataframe(report, use_container_width=True, height=600)

                # Создаём Excel в памяти
                buffer = io.BytesIO()
                report.to_excel(buffer, index=False)
                buffer.seek(0)

                # Кнопка скачивания
                if st.download_button(
                    label="Скачать",
                    data=buffer,
                    file_name=filename_total,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                ):
                    st.success('✅ Отчет сохранен')

                # Отдельный отчёт по магазинам
                tmp = tmp.sort_values('shop')
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
                tmp.reset_index(drop=True, inplace=True)

                st.subheader(f'Отчёт о движении товаров за {real_start} – {real_end}')
                st.dataframe(tmp)

                # Кнопка скачивания с возможностью задать название
                if "filename_tmp" not in st.session_state:
                    st.session_state.filename_tmp = f'Отчёт {real_start}-{real_end}'

                st.session_state.filename_tmp = st.text_input(
                    "Введите название для сохранения",
                    value=st.session_state.filename_tmp
                )

                buffer_tmp = io.BytesIO()
                tmp.to_excel(buffer_tmp, index=False)
                buffer_tmp.seek(0)

                if st.download_button(
                    label="Скачать",
                    data=buffer_tmp,
                    file_name=st.session_state.filename_tmp + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                ):
                    st.success('✅ Отчет сохранен')

