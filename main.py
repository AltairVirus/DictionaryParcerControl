import DBExtraction
import ExcelWrite


keys = DBExtraction.db_connect('ri.str_value', db='genres', look_for=True)        # Парсим из genres/rule_item список из КЛЮЧЕЙ
values = DBExtraction.db_connect('rv.value', db='genres', look_for=True)          # Парсим из genres/rule_values список из ЗНАЧЕНИЙ
attrbs = DBExtraction.db_connect(db='genres', look_for=False)                     # Парсим из genres/rule_values список из КОДОВ АТРИБУТОВ
opm_attrbs = DBExtraction.db_connect(db='opm')                                    # Парсим из opm/rule_values список из ВСЕХ АТРИБУТОВ + КОДЫ

print(keys)
print(values)
print(attrbs)

opm_attrbs_dict = {}                                                              # Преобразуем список из ОРМ в словарь
for a in opm_attrbs:
    opm_attrbs_dict[a[0]] = a[1]

reference_uniq_cols_name = ExcelWrite.read_massive_from_excel()                   # Cчитываем названия столбцов из Excel-файла

print(reference_uniq_cols_name)

db_keys_list = []
for a in values:                                                                  # Формируем список значений из БД
    for b in a:
        for c in b:
            for key, value in c.items():
                if key not in db_keys_list:
                    db_keys_list.append(key)
print(db_keys_list)

full_reference_keys_list = []                                                     # Формируем список всех необходимых Header'ов
item_reference_keys_list = []                                                     # Формируем список всех необходимых Header'ов для ЗНАЧЕНИЙ
attrbs_reference_keys_list = []                                                   # Формируем список всех необходимых Header'ов для АТРИБУТОВ
for i in reference_uniq_cols_name:
    full_reference_keys_list.append(i[0])
    if i[0] in db_keys_list:
        item_reference_keys_list.append(i[0])
    elif i[0] not in ["ObjectType", "ProcessType"]:
        attrbs_reference_keys_list.append(i[0])

print(full_reference_keys_list)
print(item_reference_keys_list)
print(attrbs_reference_keys_list)


try:                                                                              # Сортируем в список значения согласно Header'aм
    values_new = []
    for a in values:
        for b in a:
            help_list = []
            for m in item_reference_keys_list:
                for c in b:
                    for d in c.keys():
                        if m == d:
                            help_list.append(c.get(d))
        values_new.append(help_list)
except(Exception):
    values_new = []

keys_new = []
for a in keys:
    for b in a:
        help_list = []
        for c in b:
            help_list.append(c)
    keys_new.append(help_list)

# attr_сode_dict = {}                                      # Создаем словарь, ID атрибута : [последовтельность из значений атрибута]
# for a, a1 in enumerate(attrbs):
#     if attrbs[a][1] not in attr_dict.keys():
#         attr_dict[attrbs[a][1]] = []
#         attr_dict[attrbs[a][1]].append(attrbs[a][2])
#     else:
#         attr_dict[attrbs[a][1]].append(attrbs[a][2])

attr_dict = {}                                           # Создаем словарь, ID атрибута : [последовтельность из значений атрибута]
for a, a1 in enumerate(attrbs):
    if attrbs[a][1] in opm_attrbs_dict.keys():
        if opm_attrbs_dict.get(attrbs[a][1]) not in attr_dict.keys():
            attr_dict[opm_attrbs_dict.get(attrbs[a][1])] = []                                                        #opm_attrbs_dict.get(attrbs[a][1])
            attr_dict[opm_attrbs_dict.get(attrbs[a][1])].append(attrbs[a][2])
        else:
            attr_dict[opm_attrbs_dict.get(attrbs[a][1])].append(attrbs[a][2])
    else:
        print(f"Атрибут с id: {attrbs[a][1]} не найден в базе ОРМ ")
print(attr_dict)

full_data = []
full_data.append(attrbs_reference_keys_list + item_reference_keys_list)

if values_new != []:
    for a_id, a in enumerate(keys_new):
        full_data.append(keys_new[a_id] + values_new[a_id])
else:
    full_data = keys_new

ExcelWrite.write_massive_to_excel(full_data, col=1, row=1)
ExcelWrite.write_massive_to_excel(attr_dict, col=1, row=1)

print(keys_new)
print(values_new)
print(full_data)







