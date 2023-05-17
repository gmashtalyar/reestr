from aiogram import Bot, Dispatcher, executor, types
from aiogram.types.web_app_info import WebAppInfo
from aiogram.types import InputFile
import json
from reestr import format_zakupki, format_zakupki_2
import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
from perekup import perekup
from factoring import add_factoring
from novostal_reestr import get_reestr
from oper_expenses import oper_exp
"""
* Сумма оплат на сегодня
* Добавить пустые колонки
* Перекуп
* Исправить формат даты 
* Сбер

* Вставлять данные в окончательный реестр
* Опер расходы
"""

bot = Bot('5444188389:AAHdclXLU32jwy7mN_PUIYeDYZubTe6QpkA')
dp = Dispatcher(bot)



@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.answer('Инструкция реестра платежей:\n\n'
                         '1) Загрузите в чат-бот бюджет закупок на месяц\n\n'
                         '2) Загрузите в чат-бот платежи сегодняшнего дня\n\n'
                         '3) Нажмите команду /reestr , чтобы создать реестр платежей закупок\n\n'
                         '4) Сообщите финанситам о готовности реестра закупок\n\n'
                         '5) Загрузите "Заявки на расходование ДС" \n\n')


@dp.message_handler(commands=['reestr'])
async def download(message: types.Message):
    segodna = datetime.today().strftime('%Y-%m-%d')
    add_factoring()
    get_reestr()
    # document = open(f'input3/Бюджет закупок {segodna}.xlsx', 'rb')
    document = open(f'input3/Реестр Ариэль Металл {segodna}.xlsx', 'rb')
    await message.answer('Бюджет закупок готов.\n\n'
                         'Пожалуйста, сообщите финанситам, что он готов.\n\n'
                         'Ниже вы можете посмотреть на реестр платежей:')
    await bot.send_document(message.chat.id, document=document)


# перенести document download в try method?...
@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def scan_message(message: types.Message):
    segodna = datetime.today().strftime('%Y-%m-%d')
    await message.document.download(destination_file=f'input/Закупки_{segodna}.xlsx')
    try:
        format_zakupki_2()
        await message.answer('Бюджет закупок получен')
    except:
        try:
            perekup()
            await message.answer('Бюджет перекупа получен')
        except:
            try:
                await message.answer('Бот думает, дайте ему две минутки')
                oper_exp()
                document = open(f'input3/Реестр Ариэль Металл {segodna}.xlsx', 'rb')
                await message.answer('Ниже вы найдете реестр платежей:')
                await bot.send_document(message.chat.id, document=document)
            except:
                await message.answer('Вышла какая-то ошибка. Обратитесь к Геннадию.\n\n'
                                     'Или попробуйте запустить процесс снова, с самого начала :) ')


executor.start_polling(dp)
