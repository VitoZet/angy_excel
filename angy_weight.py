import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from time import time

tic = time()

wb_sale = openpyxl.load_workbook('лю-АПМ БиГ Апрель 80317.xlsx')
ws_sale = wb_sale.get_active_sheet()
# head = ws_sale['4']
diameter_lst = ['2М20', '2М24', '3М12', '3М42', 'М1,6', 'М10', 'М12', 'М14', 'М16', 'М18', 'М2', 'М2,5', 'М20', 'М22',
                'М24', 'М27', 'М3', 'М30', 'М33', 'М36', 'М39', 'М4', 'М42', 'М45', 'М48', 'М5', 'М52', 'М56', 'М6',
                'М64', 'М6х', 'М72', 'М8', 'М80', 'М90']

def SearchMedium():
    for med in ws_sale['4']:
        if med.value == 'Среднее':
            return column_index_from_string(med.column)

def WeightDiametr(diam, month_col):
    # global weight
    weight = 0
    if diameter and diameter.replace('2M' or '3M', 'М') == diam:
        month = ws_sale.cell(row=nomen_poz, column=month_col).value
        if month:
            weight += month
    return weight
x=0
# weight = 0
for nomen_poz in range(5, ws_sale.max_row + 1):
    name_nomen = ws_sale.cell(row=nomen_poz, column=1).value
###    gost = ws_sale.cell(row=nomen_poz, column=15).value
    name_metiz = ws_sale.cell(row=nomen_poz, column=14).value
    coating = ws_sale.cell(row=nomen_poz, column=13).value
    cl_pro4 = ws_sale.cell(row=nomen_poz, column=12).value
###    length = ws_sale.cell(row=nomen_poz, column=11).value
    diameter = ws_sale.cell(row=nomen_poz, column=10).value  # .replace('2M' or '3M', 'М')
    kg = ws_sale.cell(row=nomen_poz, column=6).value
    sklad = ws_sale.cell(row=nomen_poz, column=9).value
    if name_metiz == 'Болт' and kg == 'кг' and coating == 'черный' and cl_pro4 == 'кл.пр.5.8' and (sklad == 'S' or sklad == 'SZ'):
        x += WeightDiametr('М10',18)
print(x)
    #     for month_col in range(18, SearchMedium()):
    #         sk = ws_sale.cell(row=4, column=month_col).value
    #         for diam in diameter_lst:
    #             print(diam)
                # print(WeightDiametr(diam, month_col))

