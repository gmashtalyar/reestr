import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


def add_factoring():
    df = pd.read_excel(f'input/Факторинг.xlsx')
    segodna_2 = datetime.today().strftime('%Y-%m-%d')
    segodna = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    df = df[(df['Срок оплаты'] >= segodna)]
    df.loc[(df['Срок оплаты'] == segodna), 'Сумма к оплате на сегодня'] = df['Остаток задолженности после оплат']
    df.loc[(df['Срок оплаты'] == segodna), 'Остаток задолженности после оплат'] = 0

    reestr = pd.read_excel(f'input2/Бюджет закупок {segodna_2}.xlsx')
    # reestr = reestr.append(df, ignore_index=True)

    frames = [reestr, df]
    reestr_2 = pd.concat(frames)

    reestr_2 = reestr_2.drop(columns=['Unnamed: 0'])
    reestr_2['Срок оплаты'] = pd.to_datetime(reestr_2['Срок оплаты'], dayfirst=True).dt.strftime('%d-%m-%Y')
    reestr_2 = reestr_2.sort_values(by='Срок оплаты')
    reestr_2['Срок оплаты'] = reestr_2['Срок оплаты'].astype(str)
    reestr_2['Срок оплаты'] = reestr_2['Срок оплаты'].str.replace('-', '.')
    reestr_2.to_excel(f'input3/Бюджет закупок {segodna_2}.xlsx')


    # reestr = reestr.drop(columns=['Unnamed: 0'])
    # reestr['Срок оплаты'] = pd.to_datetime(reestr['Срок оплаты'], dayfirst=True).dt.strftime('%d-%m-%Y')
    # reestr = reestr.sort_values(by='Срок оплаты')
    # reestr['Срок оплаты'] = reestr['Срок оплаты'].astype(str)
    # reestr['Срок оплаты'] = reestr['Срок оплаты'].str.replace('-', '.')
    # reestr.to_excel(f'input3/Бюджет закупок {segodna_2}.xlsx')


