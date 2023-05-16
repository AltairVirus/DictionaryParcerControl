from database.guide_data_base import GuideDataBase as GD

import parameters as ps
import ExcelWrite


# Функция для вычленения элементов из коллекций любой степени вложенности
def extract_values_from_collection(data, unique=True, lst=[]):
    for element in data:
        if type(element) in (set, tuple, dict, list):
            extract_values_from_collection(element)
        elif element not in lst or unique:
            lst.append(element)
    return lst


GR_db = GD(ps.GR_Config)
keys = GD.take_guide_keys(GR_db)  # Парсим из genres/rule_item список КЛЮЧЕЙ
values = GD.take_guide_values(GR_db)  # Парсим из genres/rule_values список ЗНАЧЕНИЙ
attrs = GD.take_guide_attrs_codes(GR_db)  # Парсим из genres/rule_values список КОДОВ АТРИБУТОВ
GR_db.close_connection()

OPM_db = GD(ps.OPM_Config)
opm_attrs = GD.take_atr_names_and_codes(OPM_db)  # Парсим из opm/rule_values список ВСЕХ АТРИБУТОВ + КОДЫ
OPM_db.close_connection()

opm_attrs_dict = dict(opm_attrs)  # Преобразуем список из ОРМ в словарь
reference_uniq_cols_name = ExcelWrite.read_massive_from_excel()  # Cчитываем названия столбцов из Excel-файла
db_keys_list = extract_values_from_collection(values)  # Формируем список значений из БД

full_reference_keys_list = []  # Формируем список всех необходимых Header'ов
item_reference_keys_list = []  # Формируем список всех необходимых Header'ов для ЗНАЧЕНИЙ
attrs_reference_keys_list = []  # Формируем список всех необходимых Header'ов для АТРИБУТОВ
for i in reference_uniq_cols_name:
    full_reference_keys_list.append(i[0])
    if i[0] in db_keys_list:
        item_reference_keys_list.append(i[0])
    elif i[0] not in ["ObjectType", "ProcessType"]:
        attrs_reference_keys_list.append(i[0])

try:  # Сортируем в список значения согласно Header'aм
    values_new = []
    hlp_list = []
    for a in values:
        for b in a:
            hlp_list.clear()
            for m in item_reference_keys_list:
                for c in b:
                    for d in c.keys():
                        if m == d:
                            hlp_list.append(c.get(d))
        values_new.append(hlp_list)
except:
    values_new = []

keys_new = []
hlp_list = []
for a in keys:
    for b in a:
        hlp_list.clear()
        for c in b:
            hlp_list.append(c)
    keys_new.append(hlp_list)

attr_dict = {}  # Создаем словарь, ID атрибута - [последовтельность из значений атрибута]
for a, a1 in enumerate(attrs):
    if attrs[a][1] in opm_attrs_dict.keys():
        if opm_attrs_dict.get(attrs[a][1]) not in attr_dict.keys():
            attr_dict[opm_attrs_dict.get(attrs[a][1])] = [attrs[a][2]]
        else:
            attr_dict[opm_attrs_dict.get(attrs[a][1])].append(attrs[a][2])
    else:
        print(f"Атрибут с id: {attrs[a][1]} не найден в базе ОРМ ")

full_data = [item_reference_keys_list]

if values_new:
    for a_id, a in enumerate(keys_new):
        full_data.append(values_new[a_id])
else:
    full_data += keys_new

ExcelWrite.write_massive_to_excel(full_data, col=1, row=1)
ExcelWrite.write_massive_to_excel(attr_dict, col=1, row=1)
