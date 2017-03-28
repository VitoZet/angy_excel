import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from time import time

tic = time()

wb_sale = openpyxl.load_workbook('лю-АПМ БиГ Апрель 80317.xlsx')
ws_sale = wb_sale.get_active_sheet()
wb_sale.create_sheet(title='Диам в тонн')
ws_weight = wb_sale.get_sheet_by_name('Диам в тонн')

def SearchLastDate():
    for ld in ws_sale['4']:
        if ld.value == 'Итого':
            return column_index_from_string(ld.column)

tonnageData = {}
for m_s in range(18, SearchLastDate()):
    for nomen_poz in range(5, ws_sale.max_row + 1):
        ### name_nomen = ws_sale.cell(row=nomen_poz, column=1).value
        gost = ws_sale.cell(row=nomen_poz, column=15).value
        name_metiz = ws_sale.cell(row=nomen_poz, column=14).value
        coating = ws_sale.cell(row=nomen_poz, column=13).value
        cl_pro4 = ws_sale.cell(row=nomen_poz, column=12).value
        ###    length = ws_sale.cell(row=nomen_poz, column=11).value
        diameter = ws_sale.cell(row=nomen_poz, column=10).value  # .replace('2M' or '3M', 'М')
        measure_unit = ws_sale.cell(row=nomen_poz, column=6).value
        sklad = ws_sale.cell(row=nomen_poz, column=9).value
        month_sale = ws_sale.cell(row=4, column=m_s).value
        weight = ws_sale.cell(row=nomen_poz, column=m_s).value
        if measure_unit and name_metiz and coating and cl_pro4 and diameter and weight:
            tonnageData.setdefault(sklad, {})
            tonnageData[sklad].setdefault(measure_unit, {})
            tonnageData[sklad][measure_unit].setdefault(name_metiz, {})
            tonnageData[sklad][measure_unit][name_metiz].setdefault(coating, {})
            tonnageData[sklad][measure_unit][name_metiz][coating].setdefault(cl_pro4, {})
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4].setdefault(gost, {})
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4][gost].setdefault(diameter, {})
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4][gost][diameter].setdefault(month_sale, {'weight': 0})
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4][gost][diameter][month_sale]['weight'] += float(weight)
print('словарь создал')
toc_dic = time()
print('Время на словарь '+ str(round((toc_dic - tic), 2)) + ' сек')
print('-------------')
# for t in tonnageData['кг']['Болт']['черный']['кл.пр.5.8']['М10']['8.2016']:
#     print(t)
# print(tonnageData['кг']['Болт']['черный']['кл.пр.5.8']['М16']['9.2016']) # для тестов
lst_name_metiz = ['Болт', 'Гайка']
lst_coating = ['черный', 'цинк']
all_month_sale = SearchLastDate() - 18

e_cell = 1
for sk in tonnageData:
    # for m_unit in sorted(tonnageData[sk]):
    for nm in lst_name_metiz:
        for co in lst_coating:
            for kls_pro4 in sorted(tonnageData[sk]['кг'][nm][co]):
                # print(sk, nm, co, kls_pro4)
                for gst in tonnageData[sk]['кг'][nm][co][kls_pro4]:
                    for diam in sorted(tonnageData[sk]['кг'][nm][co][kls_pro4][gst]):
                        # for m_sal in tonnageData[sk]['кг'][nm][co][kls_pro4][gst][diam]:
                        # print(sk, nm, co, kls_pro4, gst, diam)
                    # for m_sal in tonnageData['кг'][nm][co][kls_pro4][diam]:
                        e_cell += 1
                        ws_weight['A'+str(e_cell)] = sk
                        # ws_weight['B'+str(e_cell)] = m_unit
                        ws_weight['B'+str(e_cell)] = nm
                        ws_weight['C'+str(e_cell)] = co
                        ws_weight['D'+str(e_cell)] = kls_pro4
                        ws_weight['E'+str(e_cell)] = gst
                        ws_weight['F'+str(e_cell)] = diam
                    # for m_s_s in (tonnageData['кг'][nm][co][kls_pro4][diam]):
                    #     for col in range(5, all_month_sale + 5):
                    #         ws_weight[get_column_letter(col)+str(e_cell)] = tonnageData['кг'][nm][co][kls_pro4][diam][m_s_s]['weight'] / 1000


toc_work = time()
print('Время обработки '+ str(round((toc_work - tic), 2)) + ' сек')
print('-------------')
print('сохраняю...')
wb_sale.save('WEIGHT_Angy.xlsx')
toc_save = time()
print('Время сохранения '+ str(round((toc_save - toc_work), 2)) + ' сек')
toc = time()
print('Полное время '+ str(round((toc - tic), 2)) + ' сек')
print('Готово, проверяй.')

# diameter_lst = ['М1,6', 'М10', 'М12', 'М14', 'М16', 'М18', 'М2', 'М2,5', 'М20', 'М22',
#                 'М24', 'М27', 'М3', 'М30', 'М33', 'М36', 'М39', 'М4', 'М42', 'М45', 'М48', 'М5', 'М52', 'М56', 'М6',
#                 'М64', 'М72', 'М8', 'М80', 'М90']

# tonnageData['кг']['Болт']['черный']['кл.пр.5.8']['М10']['weight'] / 1000
                # tonnageData['кг']['Болт']['черный']['кл.пр.5.8']['М10']['weight'] / 1000
        # print(e_cell, k)

# for k in sorted(tonnageData['кг']['Болт']['цинк']): # как взять ключ из словаря. можно так
# for k in sorted(tonnageData['кг']['Болт']['черный']): # как взять ключ из словаря. можно так
#     print(k)
#     e_cell += 1
#     ws_weight['A'+str(e_cell)] = 'цинк'
#     ws_weight['B'+str(e_cell)] = k

# for k in sorted(tonnageData['кг']['Болт']['цинк']['кл.пр.5.8']): # как взять ключ из словаря. можно так
#     e_cell += 1
#     ws_weight['A'+str(e_cell)] = 'цинк'
#     ws_weight['B'+str(e_cell)] = k

        # print(e_cell, coat, k)
# print(value)
# print(tonnageData['кг']['Болт']['черный']['кл.пр.5.8'])
# print(tonnageData['кг']['Болт']['черный'].keys())


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
# for key, value in tonnageData['кг']['Болт'].items(): # как взять ключ из словаря.
#     print(key)
# for k in tonnageData['кг']['Болт'].keys(): # как взять ключ из словаря. можно так
# print(all_month_sale)