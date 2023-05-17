import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


def get_reestr():
    segodna_2 = datetime.today().strftime('%Y-%m-%d')

    model_E = load_workbook('data/Реестр Ариэль Металл_formatted.xlsx')
    sheet = model_E['реестр платежей']
    start_cell = 262
    ending_cell = 803

    df = pd.read_excel(f'input3/Бюджет закупок {segodna_2}.xlsx')
    for y in range(len(df)):
        sheet[f'A{start_cell + y}'] = df.iloc[y, 1]
        sheet[f'B{start_cell + y}'] = df.iloc[y, 2]
        sheet[f'C{start_cell + y}'] = df.iloc[y, 3]
        sheet[f'D{start_cell + y}'] = df.iloc[y, 4]
        sheet[f'E{start_cell + y}'] = df.iloc[y, 5]
        sheet[f'F{start_cell + y}'] = df.iloc[y, 6]
        sheet[f'G{start_cell + y}'] = df.iloc[y, 7]
        sheet[f'H{start_cell + y}'] = df.iloc[y, 8]
        sheet[f'I{start_cell + y}'] = df.iloc[y, 9]
        sheet[f'J{start_cell + y}'] = df.iloc[y, 10]
        sheet[f'K{start_cell + y}'] = df.iloc[y, 11]
        sheet[f'L{start_cell + y}'] = df.iloc[y, 12]

    for i in range(start_cell+len(df), ending_cell):
        sheet.row_dimensions[i].hidden = True


    model_E.save(f'input3/Реестр Ариэль Металл {segodna_2}.xlsx')
