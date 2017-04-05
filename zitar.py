import openpyxl
import re
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from openpyxl.styles.fonts import Font
from openpyxl.styles import PatternFill
from time import time

tic = time()
ptt_size = r'(\d+)х(\d+)'
ptt_gost = r'(\d+)-(\d+)'
print('Загружаю Excel')
wb_zitar = openpyxl.load_workbook('Zitar.xlsx')
toc_load_excel = time()
print('Время загрузки Excel ' + str(round((toc_load_excel - tic), 2)) + ' сек')
print('Работаю с листом ' + str(wb_zitar.sheetnames))
ws_sale = wb_zitar.get_active_sheet()

for i in range(28, ws_sale.max_row - 16):
    name_nomen = ws_sale.cell(row=i, column=10).value
    size = re.search(ptt_size, name_nomen)
    # kolvo = ws_sale.cell(row=i, column=26).value
    # price = ws_sale.cell(row=i, column=33).value
    # print(name_nomen, kolvo, price)
    if size:
        print(size.group(), name_nomen)

'''Вопрос как по зитару Какая имееться в виду прочность? покрытие?'''