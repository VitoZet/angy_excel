import openpyxl
import re
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from openpyxl.styles.fonts import Font
from openpyxl.styles import PatternFill
from time import time

tic = time()
print('Загружаю Excel')
wb_zitar = openpyxl.load_workbook('Zitar.xlsx')
ws_sale = wb_zitar.get_active_sheet()
wb_smetiz = openpyxl.load_workbook('WEIGHT_Angy.xlsx')
ws_smetiz = wb_smetiz.get_sheet_by_name('TDSheet')
toc_load_excel = time()
print('Время загрузки Excel ' + str(round((toc_load_excel - tic), 2)) + ' сек')
print('Работаю с листом ' + str(wb_smetiz.sheetnames))

fill_zitar_ok = PatternFill(start_color='9ACD32', fill_type='solid')
ptt_size = r'(\d+)х(\d+)'
ptt_gost = r'(\d+)-(\d+)'

for sm_rows in range(5,  ws_smetiz.max_row + 1):
    ### name_nomen = ws_sale.cell(row=nomen_poz, column=1).value
    sm_name_metiz = ws_sale.cell(row=sm_rows, column=14).value
    sm_coating = ws_sale.cell(row=sm_rows, column=13).value
    sm_cl_pro4 = ws_sale.cell(row=sm_rows, column=12).value
    sm_gost = ws_sale.cell(row=sm_rows, column=15).value
    sm_diameter = ws_sale.cell(row=sm_rows, column=10).value  # .replace('2M' or '3M', 'М')
    sm_length = ws_sale.cell(row=sm_rows, column=11).value
    for zit_rows in range(28, ws_sale.max_row - 16):
        zit_nomen = ws_sale.cell(row=zit_rows, column=10).value
        # zit_size = re.search(ptt_size, zit_nomen)
        zit_gost = re.search(ptt_gost, zit_nomen)
        # kolvo = ws_sale.cell(row=i, column=26).value
        # price = ws_sale.cell(row=i, column=33).value
        # print(name_nomen, kolvo, price)
        # if size:
        #     print(size.group(), name_nomen)
        # if zit_gost:
        #     print(zit_gost.group(),zit_nomen)
        if 'Болт' in zit_nomen and 'Болт' == sm_name_metiz:
            print('ok')
        #     if sm_coating == 'черный' and sm_cl_pro4 == 'кл.пр.5.8':
        #         if zit_gost and zit_gost.group() == re.search(ptt_gost, sm_gost):
        #             print(zit_nomen)
            # print(zit_gost.group(), zit_nomen, zit_size.groups())