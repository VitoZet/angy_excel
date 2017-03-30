import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
from openpyxl.styles.fonts import Font
from time import time

tic = time()

print('Загружаю Excel')
wb_sale = openpyxl.load_workbook('лю-АПМ БиГ Апрель 80317.xlsx')
toc_load_excel = time()
print('Время загрузки Excel ' + str(round((toc_load_excel - tic), 2)) + ' сек')
print('Работаю с листом ' + str(wb_sale.sheetnames))
ws_sale = wb_sale.get_active_sheet()
wb_sale.create_sheet(title='Диам в тонн')
ws_weight = wb_sale.get_sheet_by_name('Диам в тонн')
wb_sale.create_sheet(title='ТАБЛ 1')
ws_tabl1 = wb_sale.get_sheet_by_name('ТАБЛ 1')
wb_sale.create_sheet(title='ТАБЛ 2')
ws_tabl2 = wb_sale.get_sheet_by_name('ТАБЛ 2')


def HeadWrite(lst_excel):
    head_style = Font(b=True)
    lst_excel['A1'] = ws_sale['I4'].value
    lst_excel['A1'].font = head_style
    lst_excel['B1'] = ws_sale['N4'].value
    lst_excel['B1'].font = head_style
    lst_excel['C1'] = ws_sale['M4'].value
    lst_excel['C1'].font = head_style
    lst_excel['D1'] = ws_sale['L4'].value
    lst_excel['D1'].font = head_style
    lst_excel['E1'] = ws_sale['O4'].value
    lst_excel['E1'].font = head_style
    lst_excel['F1'] = ws_sale['J4'].value
    lst_excel['F1'].font = head_style
    for e, mm in enumerate(lst_sale):
        lst_excel[get_column_letter(e + 7) + '1'] = mm
        lst_excel[get_column_letter(e + 7) + '1'].font = head_style


def SearchLastDate():
    for ld in ws_sale['4']:
        if ld.value == 'Склад':
            return column_index_from_string(ld.column)


lst_sale = []

tonnageData = {}
# Сохраняем таблицу в словарь
for m_s in range(18, SearchLastDate()):
    lst_sale.append(ws_sale.cell(row=4, column=m_s).value)
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
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4][gost][diameter].setdefault(month_sale,
                                                                                                      {'weight': 0})
            tonnageData[sklad][measure_unit][name_metiz][coating][cl_pro4][gost][diameter][month_sale][
                'weight'] += float(weight)
            # print(ws_sale.cell(row=4, column=m_s).value)
print('словарь создал')
toc_dic = time()
print('Время на создание словаря ' + str(round((toc_dic - toc_load_excel), 2)) + ' сек')
print('-------------')

# Словарь с списками необходимых группировок диаметров
DictGroupDiam = {'М16-М30': ('М16', 'М18', 'М20', 'М22', 'М24', 'М27', 'М30'),
                 'М6-М16': ['М6', 'М8', 'М10', 'М12', 'М14', 'М16'],
                 'М18-М36': ['М18', 'М20', 'М22', 'М24', 'М27', 'М30', 'М36', 'М33'],
                 'М4-М16': ['М4', 'М5', 'М6', 'М8', 'М10', 'М12', 'М14', 'М16'],
                 'М3, М42-М72': ['М3', 'М42', 'М45', 'М48', 'М52', 'М56', 'М64', 'М72']}
print(DictGroupDiam['М16-М30'])
m6_m16= ['М6', 'М8', 'М10', 'М12', 'М14', 'М16']
# S	Болт	черный	кл.пр.5.8	ГОСТ 7798-70	М6-М16
# print(tonnageData['S']['кг']['Болт']['черный']['кл.пр.5.8']['ГОСТ 7798-70']['М16']['8.2016']['weight']) # для тестов
print('начало')


def SumGroupDiamMonth(groupDiam, sklad, name_metiz, coating, cl_pro4, gost, month):
    '''Сумма веса в указанном диапозоне диаметров'''
    summa = 0
    for d in groupDiam:
        summa += tonnageData[sklad]['кг'][name_metiz][coating][cl_pro4][gost][d][month]['weight'] / 1000
    return summa


# SumGroupDiamMonth(lst_m6m16, 'S','Болт','черный','кл.пр.5.8','ГОСТ 7798-70','8.2016')
def SumGroupDiamMonthForCol(groupDiam, sklad, name_metiz, coating, cl_pro4, gost):
    '''Сохранение в Excel по коллонкам'''
    for e, month_col in enumerate(lst_sale):
        ws_tabl1[get_column_letter(e + 7)+str(ws_tabl1.max_row)] = round(SumGroupDiamMonth(groupDiam, sklad, name_metiz, coating, cl_pro4, gost, month_col), 2)
        # print(e + 7, round(SumGroupDiamMonth(groupDiam, sklad, name_metiz, coating, cl_pro4, gost, month_col), 2))
        # print(e+7, SumGroupDiamMonth(lst_m6m16, 'S','Болт','черный','кл.пр.5.8','ГОСТ 7798-70',month_col))


# SumGroupDiamMonthForCol(DictGroupDiam['М6-М16'], 'S', 'Болт', 'черный', 'кл.пр.5.8', 'ГОСТ 7798-70')


def SaveInExcel(name_metiz, coating, cl_pro4, gost, groupDiam):
    lst_sklad = ['S', 'Z', 'SZ', 'ALL']
    for sk in lst_sklad:
        ws_tabl1['A' + str(ws_tabl1.max_row+1)] = sk
        ws_tabl1['B' + str(ws_tabl1.max_row)] = name_metiz
        ws_tabl1['C' + str(ws_tabl1.max_row)] = coating
        ws_tabl1['D' + str(ws_tabl1.max_row)] = cl_pro4
        ws_tabl1['E' + str(ws_tabl1.max_row)] = gost
        # ws_tabl1['F' + str(ws_tabl1.max_row)] = groupDiam
        if not sk == 'ALL':
            SumGroupDiamMonthForCol(groupDiam,sk,name_metiz,coating,cl_pro4,gost)
        else:
            pass

SaveInExcel('Болт', 'черный', 'кл.пр.5.8', 'ГОСТ 7798-70', m6_m16)



# print('33338.524000000005')
# print(tonnageData['S']['кг']['Болт']['черный']['кл.пр.5.8']['ГОСТ 7798-70'][d]['8.2016']['weight'])  # для тестов

print('конец')
lst_name_metiz = ['Болт', 'Гайка']
lst_coating = ['черный', 'цинк']
all_month_sale = SearchLastDate() - 18

HeadWrite(ws_weight)
HeadWrite(ws_tabl1)
HeadWrite(ws_tabl2)

# print('9.2016 отгружено', tonnageData['SZ']['кг']['Болт']['черный']['кл.пр.5.8']['ГОСТ 7798-70']['М10']['2.2017']['weight'] / 1000)

e_cell = 1
# Заполняем таблицу данными
for sk in tonnageData:
    for nm in lst_name_metiz:
        for co in lst_coating:
            for kls_pro4 in sorted(tonnageData[sk]['кг'][nm][co]):
                for gst in tonnageData[sk]['кг'][nm][co][kls_pro4]:
                    for diam in sorted(tonnageData[sk]['кг'][nm][co][kls_pro4][gst]):
                        e_cell += 1
                        ws_weight['A' + str(e_cell)] = sk
                        ws_weight['B' + str(e_cell)] = nm
                        ws_weight['C' + str(e_cell)] = co
                        ws_weight['D' + str(e_cell)] = kls_pro4
                        ws_weight['E' + str(e_cell)] = gst
                        ws_weight['F' + str(e_cell)] = diam
                        for e, m_s_s in enumerate(lst_sale):
                            try:
                                ws_weight[get_column_letter(e + 7) + str(e_cell)] = \
                                tonnageData[sk]['кг'][nm][co][kls_pro4][gst][diam][m_s_s]['weight'] / 1000
                            except:
                                ws_weight[get_column_letter(e + 7) + str(e_cell)] = None

toc_work = time()
print('Время обработки ' + str(round((toc_work - tic), 2)) + ' сек')
print('-------------')
print('сохраняю...')
wb_sale.save('WEIGHT_Angy.xlsx')
toc_save = time()
print('Время сохранения ' + str(round((toc_save - toc_work), 2)) + ' сек')
toc = time()
print('Полное время ' + str(round((toc - tic), 2)) + ' сек')
print('Готово, проверяй.')
