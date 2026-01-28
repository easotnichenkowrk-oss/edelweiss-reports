    df = pd.read_excel(goods_data)
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
                filename_total = f'Общий отчёт {real_start}-{real_end} {selected_shop}.xlsx'
                st.dataframe(report, use_container_width=True, height=600)

                if len(filename_total) > 100:
                    filename_total = f'Общий отчёт {real_start}-{real_end} {selected_shop[:3]} итд.xlsx'

                buffer = io.BytesIO()
                report.to_excel(buffer, index=False)
                buffer.seek(0)
                if st.download_button(
                    label="Скачать",
                    data=buffer,
                    file_name=filename_total,
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
                
                tmp.reset_index(drop=True, inplace=True)
                st.dataframe (tmp)

                if len(filename) > 100:
                    filename = f'Отчёт {real_start}-{real_end} {selected_shop[:3]} итд.xlsx'
                buffer = io.BytesIO()
                tmp.to_excel(buffer, index=False)
                buffer.seek(0)
                if len (selected_shop) > 3:
                    filename = st.textinput ('Ввведите название для сохранения')
                if st.download_button(
                    label="Скачать",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                ):
                    st.success('✅ Отчет сохранен')
