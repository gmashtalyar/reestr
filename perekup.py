import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


def perekup():
    segodna = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    segodna_2 = datetime.today().strftime('%Y-%m-%d')
    # df = pd.read_excel('input/05.04.2023 предварительно.xlsx')
    df = pd.read_excel(f'input/Закупки_{segodna_2}.xlsx')

    if len(df.columns) > 5:
        raise ValueError('not perekup')

    new_df = df[df['Контрагент'].where(df['Контрагент'] == 'Итого без СБФ', np.nan).ffill() != 'Итого без СБФ']
    new_df_2 = new_df[new_df['Контрагент'].where(new_df['Контрагент'] == 'Отдел складских закупок', np.nan).bfill() != 'Отдел складских закупок']

    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент Тгн')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент ТГН')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент СМР')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент СМР')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Отдел складских закупок СПБ')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Отдел складских закупок МСК')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Отдел складских закупок ТГН')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Отдел складских закупок СМР')]

    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != '')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Факторинг')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент МСК')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент СПБ')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент Таганрог')]
    new_df_2 = new_df_2.loc[(new_df_2['Контрагент'] != 'Коммерческий департамент Самара')]

    new_df_2['Срок оплаты'] = segodna_2
    new_df_2.rename(columns={
        'Сумма': 'Сумма к оплате на сегодня',
        'Контрагент': 'Поставщик'}, inplace=True)
    new_df_2['Срок оплаты'] = new_df_2['Срок оплаты'].astype(str)

    """
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА В СОРТИРОВКЕ (STR)
    """
    new_df_2['Срок оплаты'] = new_df_2['Срок оплаты'].str.replace('-', '.')
    new_df_2['Назначение платежа'] = 'Оплата за металлопрокат'
    new_df_2['Лицо, ответственное за договор (счет)'] = 'Чернышова Светлана Эдуардовна'
    new_df_2['Наличие подтверждения от поставщика'] = ''
    new_df_2['Договора дата, номер'] = ''
    new_df_2['Оплачено всего на нач. опер дня'] = ''
    new_df_2['Остаток задолженности после оплат'] = 0
    new_df_2['Текущая сумма задолженности'] = new_df_2['Сумма к оплате на сегодня']
    new_df_2['Сумма оплаты по сроку'] = new_df_2['Сумма к оплате на сегодня']

    reestr = pd.read_excel(f'Бюджет закупок {segodna_2}.xlsx')
    # reestr = reestr.append(new_df_2, ignore_index=True)
    frames = [reestr, new_df_2]
    reestr_2 = pd.concat(frames)

    reestr_2 = reestr_2.drop(columns=['Unnamed: 0'])


    reestr_2['Срок оплаты'] = pd.to_datetime(reestr_2['Срок оплаты']).dt.strftime('%d-%m-%Y')
    reestr_2 = reestr_2.sort_values(by='Срок оплаты')
    reestr_2['Срок оплаты'] = reestr_2['Срок оплаты'].astype(str)
    """
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    ЗДЕСЬ ОШИБКА #2 В СОРТИРОВКЕ (STR)
    """
    reestr_2['Срок оплаты'] = reestr_2['Срок оплаты'].str.replace('-', '.')

    reestr_2.to_excel(f'input2/Бюджет закупок {segodna_2}.xlsx')

