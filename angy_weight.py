import openpyxl
from openpyxl.utils import get_column_letter
from time import time

tic = time()

wb_sale = openpyxl.load_workbook('лю-АПМ БиГ Апрель 80317.xlsx')
ws_sale = wb_sale.get_active_sheet()
M10 = 0

for nomen_poz in range(5, ws_sale.max_row + 1):
    name_nomen = ws_sale.cell(row=nomen_poz, column=1).value
    # gost = ws_sale.cell(row=nomen_poz, column=15).value
    name_metiz = ws_sale.cell(row=nomen_poz, column=14).value
    coating = ws_sale.cell(row=nomen_poz, column=13).value
    cl_pro4 = ws_sale.cell(row=nomen_poz, column=12).value
    # length = ws_sale.cell(row=nomen_poz, column=11).value
    diameter = ws_sale.cell(row=nomen_poz, column=10).value  # .replace('2M' or '3M', 'М')
    kg = ws_sale.cell(row=nomen_poz, column=6).value
    sklad = ws_sale.cell(row=nomen_poz, column=9).value
    # print(kg, sklad)
    if name_metiz == 'Болт' and kg == 'кг' and coating == 'черный' and cl_pro4 == 'кл.пр.5.8' and (
            sklad == 'S' or sklad == 'SZ'):
        if diameter:
            diameter = ws_sale.cell(row=nomen_poz, column=10).value.replace('2M' or '3M', 'М')
        if diameter == 'М10':
            month1 = ws_sale.cell(row=nomen_poz, column=18).value
            if month1:
                M10 += month1
print('M10 = ' + str(M10 / 1000) + ' tonn')
# print(diameter, name_nomen, sklad, cl_pro4, coating)
