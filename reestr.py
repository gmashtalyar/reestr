import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


def format_zakupki():
    segodna = datetime.today().strftime('%Y-%m-%d')
    segodna_2 = datetime.today().strftime('%Y-%m-%d')
    segodna_3 = datetime.today().strftime('%Y.%m.%d')

    raw_data = read_excel(f'input/pre-filled.xlsx')  # читаем данные
    raw_data2 = raw_data.melt(id_vars=['Поставщик', 'Договор'])
    df = raw_data2.dropna()
    df['variable'] = pd.to_datetime(df['variable']).dt.date
    df.rename(columns={
        'variable': 'Срок оплаты',
        'value': 'Сумма оплаты по сроку'}, inplace=True)
    df['Назначение платежа'] = 'Оплата за металлопрокат'
    df['Лицо, ответственное за договор (счет)'] = 'Чернышова Светлана Эдуардовна'
    df['Оплачено всего на нач. опер дня'] = ''
    df['Текущая сумма задолженности'] = df['Сумма оплаты по сроку']
    df['Остаток задолженности после оплат'] = df['Сумма оплаты по сроку']
    df['Сумма к оплате на сегодня'] = ''
    df = df[['Поставщик',
             'Назначение платежа',
             'Лицо, ответственное за договор (счет)',
             'Договор',
             'Текущая сумма задолженности',
             'Оплачено всего на нач. опер дня',
             'Срок оплаты',
             'Сумма оплаты по сроку',
             'Сумма к оплате на сегодня',
             'Остаток задолженности после оплат']]
    df.to_excel(f'рабочий файл {segodna}.xlsx')


def format_zakupki_2():
    segodna = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    segodna_2 = datetime.today().strftime('%Y-%m-%d')

    datelist = []
    datelist.append('Поставщик')
    datelist.append('Договор')

    df_column = pd.read_excel(f'input/Закупки_{segodna_2}.xlsx', sheet_name='график', skiprows=7, nrows=57).columns
    for i in df_column:
        if isinstance(i, datetime) and i >= segodna:
            datelist.append(i)

    raw_data = read_excel(f'input/Закупки_{segodna_2}.xlsx', sheet_name='график', skiprows=7, nrows=57, dtype=str)  # читаем данные
    dates = raw_data[datelist]
    raw_data2 = dates.melt(id_vars=['Поставщик', 'Договор'])
    df = raw_data2.dropna()
    df['variable'] = pd.to_datetime(df['variable']).dt.date
    df.rename(columns={
        'variable': 'Срок оплаты',
        'value': 'Сумма оплаты по сроку'}, inplace=True)
    df['Назначение платежа'] = 'Оплата за металлопрокат'
    df['Лицо, ответственное за договор (счет)'] = 'Чернышова Светлана Эдуардовна'
    df['Оплачено всего на нач. опер дня'] = ''
    df['Текущая сумма задолженности'] = df['Сумма оплаты по сроку']
    df['Остаток задолженности после оплат'] = df['Сумма оплаты по сроку']
    df['Сумма к оплате на сегодня'] = ''
    df['Наличие подтверждения от поставщика'] = ''
    df['Договора дата, номер'] = ''
    df = df[['Поставщик',
             'Назначение платежа',
             'Лицо, ответственное за договор (счет)',
             'Наличие подтверждения от поставщика',
             'Договор',
             'Договора дата, номер',
             'Текущая сумма задолженности',
             'Оплачено всего на нач. опер дня',
             'Срок оплаты',
             'Сумма оплаты по сроку',
             'Сумма к оплате на сегодня',
             'Остаток задолженности после оплат']]
    df["Сумма оплаты по сроку"] = pd.to_numeric(df["Сумма оплаты по сроку"], downcast="float")
    df["Остаток задолженности после оплат"] = pd.to_numeric(df["Остаток задолженности после оплат"], downcast="float")

    df['Срок оплаты'] = pd.to_datetime(df['Срок оплаты'], format='%Y-%m-%d')
    df.loc[(df['Срок оплаты'] == segodna_2), 'Сумма к оплате на сегодня'] = df['Остаток задолженности после оплат']
    df.loc[(df['Срок оплаты'] == segodna_2), 'Остаток задолженности после оплат'] = 0
    df['Срок оплаты'] = pd.to_datetime(df['Срок оплаты'], format='%d-%m-%Y')
    df['Срок оплаты'] = df['Срок оплаты'].astype(str)
    df['Срок оплаты'] = df['Срок оплаты'].str.replace('-', '.')

    df.to_excel(f'Бюджет закупок {segodna_2}.xlsx')


# format_zakupki_2()
