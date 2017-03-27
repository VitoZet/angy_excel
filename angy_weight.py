import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from time import time

tic = time()

wb_sale = openpyxl.load_workbook('лю-АПМ БиГ Апрель 80317.xlsx')
ws_sale = wb_sale.get_active_sheet()
# wb_sale.create_sheet(title='Диам в тонн')
# ws_weight = wb_sale.get_sheet_by_name('Диам в тонн')
diameter_lst = ['М1,6', 'М10', 'М12', 'М14', 'М16', 'М18', 'М2', 'М2,5', 'М20', 'М22',
                'М24', 'М27', 'М3', 'М30', 'М33', 'М36', 'М39', 'М4', 'М42', 'М45', 'М48', 'М5', 'М52', 'М56', 'М6',
                'М64', 'М72', 'М8', 'М80', 'М90']
tonnageData = {}

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
    weight = ws_sale.cell(row=nomen_poz, column=18).value
    if kg and name_metiz and coating and cl_pro4 and diameter and weight:
        tonnageData.setdefault(kg, {})  # , {name_metiz}) #,{coating,{cl_pro4,{diameter,{'weight':0}}}}})
        tonnageData[kg].setdefault(name_metiz, {})
        tonnageData[kg][name_metiz].setdefault(coating, {})
        tonnageData[kg][name_metiz][coating].setdefault(cl_pro4, {})
        tonnageData[kg][name_metiz][coating][cl_pro4].setdefault(diameter, {'weight': 0})
        tonnageData[kg][name_metiz][coating][cl_pro4][diameter]['weight'] += float(weight)
        # if (sklad == 'S' or sklad == 'SZ'):
        # if weight:
        #         tonnageData.setdefault(diameter, {'number_lenght': 0, 'weight': 0})
        #         tonnageData[diameter]['number_lenght'] +=1
        #         tonnageData[diameter]['weight'] += float(weight)
        # print(nomen_poz)
        # print(tonnageData['кг']['Болт']['черный']['кл.пр.5.8']['М10']['weight']/1000)

# for dd in diameter_lst:
#     try:
#         x = tonnageData['кг']['Болт']['черный']['кл.пр.5.8'][dd]['weight']/1000
#         if x:
#             print(dd, x)
#     except:
#         print(dd , '')
# print(tonnageData['кг']['Болт']['черный']['кл.пр.5.8'])
print(tonnageData['кг']['Болт']['черный'].keys())
# print('сохраняю...')
# wb_sale.save('WEIGHT_Angy.xlsx')
toc = time()
print(toc - tic)
print('Готово, проверяй.')
