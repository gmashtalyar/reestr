import pandas as pd
from openpyxl import load_workbook
from pandas import read_excel
from datetime import datetime
import numpy as np


model = load_workbook('data/opxl.xlsx')
sheet = model['Лист1']
# С 4 по 9
for i in range(5, 10):
    sheet.row_dimensions[i].hidden = True

# С 25 по 34
for i in range(25, 35):
    sheet.row_dimensions[i].hidden = True

# С 45 по 54
for i in range(45, 55):
    sheet.row_dimensions[i].hidden = True

model.save('data/opxl.xlsx')
