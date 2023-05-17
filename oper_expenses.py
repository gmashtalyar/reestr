import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


def paste_expenses(start_cell, exp_dataframe, ending_cell):
    segodna_2 = datetime.today().strftime('%Y-%m-%d')
    model_E = load_workbook(f'input3/Реестр Ариэль Металл {segodna_2}.xlsx')
    sheet = model_E['реестр платежей']

    for y in range(len(exp_dataframe)):
        sheet[f'A{start_cell + y}'] = exp_dataframe.iloc[y, 2]
        sheet[f'B{start_cell + y}'] = exp_dataframe.iloc[y, 3]
        sheet[f'C{start_cell + y}'] = exp_dataframe.iloc[y, 4]
        sheet[f'D{start_cell + y}'] = exp_dataframe.iloc[y, 5]
        sheet[f'E{start_cell + y}'] = exp_dataframe.iloc[y, 6]
        sheet[f'F{start_cell + y}'] = exp_dataframe.iloc[y, 7]
        sheet[f'G{start_cell + y}'] = exp_dataframe.iloc[y, 8]
        sheet[f'H{start_cell + y}'] = exp_dataframe.iloc[y, 9]
        sheet[f'I{start_cell + y}'] = exp_dataframe.iloc[y, 10]
        sheet[f'J{start_cell + y}'] = exp_dataframe.iloc[y, 11]
        sheet[f'K{start_cell + y}'] = exp_dataframe.iloc[y, 12]
        sheet[f'L{start_cell + y}'] = exp_dataframe.iloc[y, 13]

    for i in range(start_cell+len(exp_dataframe), ending_cell):
        sheet.row_dimensions[i].hidden = True

    model_E.save(f'input3/Реестр Ариэль Металл {segodna_2}.xlsx')


def oper_exp():
    segodna_1 = datetime.today().strftime('%Y-%m-%d')

    raw_data = pd.read_excel(f'input/Закупки_{segodna_1}.xlsx')
    raw_data['Дата'] = pd.to_datetime(raw_data['Дата'], dayfirst=True)
    segodna = datetime.today().strftime('%d.%m.%Y')
    segodna_3 = datetime.today().strftime('%Y.%m.%d')
    segodna_2 = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    raw_data = raw_data[(raw_data['Дата'] >= segodna_2)]

    raw_data['Дата'] = raw_data['Дата'].astype(str)
    raw_data['Дата'] = raw_data['Дата'].str.replace('-', '.')


    raw_data = raw_data.loc[raw_data['Организация'] == 'Ариэль Металл']
    raw_data['Текущая сумма задолженности'] = raw_data['Сумма']
    raw_data['Остаток задолженности после оплат'] = raw_data['Сумма']
    raw_data['Сумма оплаты по сроку'] = raw_data['Сумма']
    raw_data['Наличие подтверждения от поставщика'] = ''
    raw_data['Договора дата, номер'] = ''
    raw_data['Оплачено всего на нач. опер дня'] = ''
    raw_data['Сумма к оплате на сегодня'] = ''
    raw_data = raw_data[['ЦФО', 'Статья оборотов', 'Сокращенное юр. наименование',
                         'Назначение платежа',
                         'Ответственный',
                         'Наличие подтверждения от поставщика',
                         'Договор',
                         'Договора дата, номер',
                         'Текущая сумма задолженности',
                         'Оплачено всего на нач. опер дня',
                         'Дата',
                         'Сумма оплаты по сроку',
                         'Сумма к оплате на сегодня',
                         'Остаток задолженности после оплат']]

    raw_data.loc[(raw_data['Дата'] == segodna_3), 'Сумма к оплате на сегодня'] = raw_data['Остаток задолженности после оплат']
    raw_data.loc[(raw_data['Дата'] == segodna_3), 'Остаток задолженности после оплат'] = 0

    raw_data.rename(columns={
        'Сокращенное юр. наименование': 'Поставщик',
        'Ответственный': 'Лицо, ответственное за договор (счет)',
        'Дата': 'Срок оплаты',
        'Статья оборотов': 'Статья ERP'}, inplace=True)

    spravochnik = pd.read_excel('data/Справочник статей.xlsx')
    raw_data = pd.merge(raw_data, spravochnik, on='Статья ERP', how='outer')

    zakupki_data = raw_data.loc[(raw_data['ЦФО'] == 'Отдел централизованных закупок') |
                           (raw_data['ЦФО'] == 'Отдел закупок КД МСК') |
                           (raw_data['ЦФО'] == 'Отдел закупок КД СПБ') |
                           (raw_data['ЦФО'] == 'Отдел закупок КД ТГН') |
                           (raw_data['ЦФО'] == 'Отдел закупок КД СМР')]

    administation_data = raw_data.loc[(raw_data['ЦФО'] == 'УК') |
                           (raw_data['ЦФО'] == 'HR') |
                           (raw_data['ЦФО'] == 'IT') |
                           (raw_data['ЦФО'] == 'Администрация') |
                           (raw_data['ЦФО'] == 'СБ') |
                            (raw_data['ЦФО'] == 'Транспортн.отд.+ АХО')]

    marketing_data = raw_data.loc[raw_data['ЦФО'] == 'Отдел маркетинга']

    kd_msk_data = raw_data.loc[(raw_data['ЦФО'] == 'КД МСК Общие затраты') |
                           (raw_data['ЦФО'] == 'Отдел продаж №1 МСК') |
                           (raw_data['ЦФО'] == 'Отдел продаж №2 МСК') |
                           (raw_data['ЦФО'] == 'Отдел продаж №3 МСК') |
                           (raw_data['ЦФО'] == 'Отдел продаж №4 МСК')]

    kd_spb_data = raw_data.loc[(raw_data['ЦФО'] == 'СПБ Общие затраты') |
                           (raw_data['ЦФО'] == 'Отдел продаж №1 СПБ') |
                           (raw_data['ЦФО'] == 'Отдел продаж №2 СПБ') |
                           (raw_data['ЦФО'] == 'Отдел продаж №3 СПБ') |
                           (raw_data['ЦФО'] == 'Отдел продаж №4 СПБ') |
                               (raw_data['ЦФО'] == 'Склад Санкт-Петербург')]

    kd_tgn_data = raw_data.loc[(raw_data['ЦФО'] == 'Таганрог') |
                           (raw_data['ЦФО'] == 'Отдел продаж №1 ТГН')]

    kd_smr_data = raw_data.loc[raw_data['ЦФО'] == 'Отдел продаж №1 СМР']

    project_sales_data = raw_data.loc[raw_data['ЦФО'] == 'Отдел проектных продаж']
    regional_sales_data = raw_data.loc[raw_data['ЦФО'] == 'Отдел региональных продаж']


    sklad_logistics_data = raw_data.loc[(raw_data['ЦФО'] == 'Склад Подольск') |
                           (raw_data['ЦФО'] == 'Отдел складской логистики') |
                           (raw_data['ЦФО'] == 'Отдел транспортной логистики') |
                           (raw_data['ЦФО'] == 'АТП')]

    adm_fot_start_row = 17
    adm_fot_end_row = 82
    adm_arenda_start_row = 84
    adm_arenda_end_row = 123
    adm_taxes_start_row = 125
    adm_taxes_end_row = 160
    adm_VGO_start_row = 162
    adm_VGO_end_row = 189
    adm_financing_start_row = 191
    adm_financing_end_row = 213
    adm_other_start_row = 215
    adm_other_end_row = 258

    kd_msk_fot_start_row = 808
    kd_msk_fot_end_row = 846
    kd_msk_other_start_row = 848
    kd_msk_other_end_row = 903
    kd_msk_delivery_start_row = 905
    kd_msk_delivery_end_row = 985
    kd_msk_sklad_start_row = 987
    kd_msk_sklad_end_row = 1030

    kd_spb_fot_start_row = 1034
    kd_spb_fot_end_row = 1071
    kd_spb_other_start_row = 1073
    kd_spb_other_end_row = 1148
    kd_spb_delivery_start_row = 1150
    kd_spb_delivery_end_row = 1238
    kd_spb_arenda_start_row = 1240
    kd_spb_arenda_end_row = 1313
    kd_spb_sklad_start_row = 1315
    kd_spb_sklad_end_row = 1370

    kd_tgn_fot_start_row = 1374
    kd_tgn_fot_end_row = 1409
    kd_tgn_other_start_row = 1411
    kd_tgn_other_end_row = 1451
    kd_tgn_delivery_start_row = 1453
    kd_tgn_delivery_end_row = 1491
    kd_tgn_arenda_start_row = 1493
    kd_tgn_arenda_end_row = 1521
    kd_tgn_sklad_start_row = 1523
    kd_tgn_sklad_end_row = 1566

    kd_smr_fot_start_row = 1570
    kd_smr_fot_end_row = 1606
    kd_smr_other_start_row = 1608
    kd_smr_other_end_row = 1649
    kd_smr_delivery_start_row = 1651
    kd_smr_delivery_end_row = 1707
    kd_smr_arenda_start_row = 1709
    kd_smr_arenda_end_row = 1733
    kd_smr_sklad_start_row = 1735
    kd_smr_sklad_end_row = 1774

    kd_regional_other_start_row = 2217
    kd_regional_other_end_row = 2245
    kd_project_other_start_row = 2249
    kd_project_other_end_row = 2289

    sklad_fot_start_row = 2293
    sklad_fot_end_row = 2326
    sklad_sklad_start_row = 2328
    sklad_sklad_end_row = 2370
    sklad_VGO_start_row = 2372
    sklad_VGO_end_row = 2403
    sklad_other_start_row = 2405
    sklad_other_end_row = 2436
    marketing_start_row = 2440
    marketing_end_row = 2467


    paste_expenses(adm_fot_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'Фонд оплаты труда')], adm_fot_end_row)
    paste_expenses(adm_arenda_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'Аренда')], adm_arenda_end_row)
    paste_expenses(adm_taxes_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'налоги')], adm_taxes_end_row)
    paste_expenses(adm_VGO_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'Внутригрупповые обороты')], adm_VGO_end_row)
    paste_expenses(adm_financing_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'Финансирование')], adm_financing_end_row)
    paste_expenses(adm_other_start_row, administation_data.loc[(administation_data['Статья в реестре'] == 'Прочее')], adm_other_end_row)

    paste_expenses(kd_msk_fot_start_row, kd_msk_data.loc[(kd_msk_data['Статья в реестре'] == 'Фонд оплаты труда')], kd_msk_fot_end_row)
    paste_expenses(kd_msk_other_start_row, kd_msk_data.loc[(kd_msk_data['Статья в реестре'] == 'Прочее')], kd_msk_other_end_row)
    paste_expenses(kd_msk_delivery_start_row, kd_msk_data.loc[(kd_msk_data['Статья в реестре'] == 'Доставки')], kd_msk_delivery_end_row)
    paste_expenses(kd_msk_sklad_start_row, kd_msk_data.loc[(kd_msk_data['Статья в реестре'] == 'Складская логистика')], kd_msk_sklad_end_row)

    paste_expenses(kd_spb_fot_start_row, kd_spb_data.loc[(kd_spb_data['Статья в реестре'] == 'Фонд оплаты труда')], kd_spb_fot_end_row)
    paste_expenses(kd_spb_other_start_row, kd_spb_data.loc[(kd_spb_data['Статья в реестре'] == 'Прочее')], kd_spb_other_end_row)
    paste_expenses(kd_spb_delivery_start_row, kd_spb_data.loc[(kd_spb_data['Статья в реестре'] == 'Доставки')], kd_spb_delivery_end_row)
    paste_expenses(kd_spb_arenda_start_row, kd_spb_data.loc[(kd_spb_data['Статья в реестре'] == 'Аренда')], kd_spb_arenda_end_row)
    paste_expenses(kd_spb_sklad_start_row, kd_spb_data.loc[(kd_spb_data['Статья в реестре'] == 'Складская логистика')], kd_spb_sklad_end_row)

    paste_expenses(kd_tgn_fot_start_row, kd_tgn_data.loc[(kd_tgn_data['Статья в реестре'] == 'Фонд оплаты труда')], kd_tgn_fot_end_row)
    paste_expenses(kd_tgn_other_start_row, kd_tgn_data.loc[(kd_tgn_data['Статья в реестре'] == 'Прочее')], kd_tgn_other_end_row)
    paste_expenses(kd_tgn_delivery_start_row, kd_tgn_data.loc[(kd_tgn_data['Статья в реестре'] == 'Доставки')], kd_tgn_delivery_end_row)
    paste_expenses(kd_tgn_arenda_start_row, kd_tgn_data.loc[(kd_tgn_data['Статья в реестре'] == 'Аренда')], kd_tgn_arenda_end_row)
    paste_expenses(kd_tgn_sklad_start_row, kd_tgn_data.loc[(kd_tgn_data['Статья в реестре'] == 'Складская логистика')], kd_tgn_sklad_end_row)

    paste_expenses(kd_smr_fot_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Фонд оплаты труда')], kd_smr_fot_end_row)
    paste_expenses(kd_smr_other_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Прочее')], kd_smr_other_end_row)
    paste_expenses(kd_smr_delivery_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Доставки')], kd_smr_delivery_end_row)
    paste_expenses(kd_smr_arenda_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Аренда')], kd_smr_arenda_end_row)
    paste_expenses(kd_smr_sklad_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Складская логистика')], kd_smr_sklad_end_row)

    # paste_expenses(kd_smr_fot_start_row, kd_kdr_data.loc[(kd_smr_data['Статья в реестре'] == 'Фонд оплаты труда')], kd_smr_fot_end_row)
    # paste_expenses(kd_smr_other_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Прочее')], kd_smr_other_end_row)
    # paste_expenses(kd_smr_delivery_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Доставки')], kd_smr_delivery_end_row)
    # paste_expenses(kd_smr_arenda_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Аренда')], kd_smr_arenda_end_row)
    # paste_expenses(kd_smr_sklad_start_row, kd_smr_data.loc[(kd_smr_data['Статья в реестре'] == 'Складская логистика')], kd_smr_sklad_end_row)

    paste_expenses(kd_regional_other_start_row, regional_sales_data.loc[(regional_sales_data['Статья в реестре'] == 'Прочее')], kd_regional_other_end_row)
    paste_expenses(kd_project_other_start_row, project_sales_data.loc[(project_sales_data['Статья в реестре'] == 'Прочее')], kd_project_other_end_row)

    paste_expenses(sklad_fot_start_row, sklad_logistics_data.loc[(sklad_logistics_data['Статья в реестре'] == 'Фонд оплаты труда')], sklad_fot_end_row)
    paste_expenses(sklad_sklad_start_row, sklad_logistics_data.loc[(sklad_logistics_data['Статья в реестре'] == 'Складская логистика')], sklad_sklad_end_row)
    paste_expenses(sklad_VGO_start_row, sklad_logistics_data.loc[(sklad_logistics_data['Статья в реестре'] == 'Внутригрупповые обороты')], sklad_VGO_end_row)
    paste_expenses(sklad_other_start_row, sklad_logistics_data.loc[(sklad_logistics_data['Статья в реестре'] == 'Прочее')], sklad_other_end_row)
    paste_expenses(marketing_start_row, marketing_data.loc[(marketing_data['Статья в реестре'] == 'Прочее')], marketing_end_row)




"""
TO DO:

Hide rows for procurement
"""
