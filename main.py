import DBExtraction
import ExcelWrite


def extract_values_from_collection(d, s=[]):
    for i in d:
        if type(i) in (set, tuple, dict, list):
            extract_values_from_collection(i)
        elif i not in s:
            s.append(i)
    return s


keys = DBExtraction.db_connect('ri.str_value', db='genres', look_for=True)  # Парсим из genres/rule_item список КЛЮЧЕЙ
values = DBExtraction.db_connect('rv.value', db='genres', look_for=True)  # Парсим из genres/rule_values список ЗНАЧЕНИЙ
attrbs = DBExtraction.db_connect(db='genres', look_for=False)  # Парсим из genres/rule_values список КОДОВ АТРИБУТОВ
opm_attrbs = DBExtraction.db_connect(db='opm')  # Парсим из opm/rule_values список ВСЕХ АТРИБУТОВ + КОДЫ

opm_attrbs_dict = {a[0]: a[1] for a in opm_attrbs}  # Преобразуем список из ОРМ в словарь
reference_uniq_cols_name = ExcelWrite.read_massive_from_excel()  # Cчитываем названия столбцов из Excel-файла
db_keys_list = extract_values_from_collection(values)  # Формируем список значений из БД

full_reference_keys_list = []  # Формируем список всех необходимых Header'ов
item_reference_keys_list = []  # Формируем список всех необходимых Header'ов для ЗНАЧЕНИЙ
attrbs_reference_keys_list = []  # Формируем список всех необходимых Header'ов для АТРИБУТОВ
for i in reference_uniq_cols_name:
    full_reference_keys_list.append(i[0])
    if i[0] in db_keys_list:
        item_reference_keys_list.append(i[0])
    elif i[0] not in ["ObjectType", "ProcessType"]:
        attrbs_reference_keys_list.append(i[0])

try:  # Сортируем в список значения согласно Header'aм
    values_new = []
    for a in values:
        for b in a:
            hlp_list = []
            for m in item_reference_keys_list:
                for c in b:
                    for d in c.keys():
                        if m == d:
                            hlp_list.append(c.get(d))
        values_new.append(hlp_list)
except:
    values_new = []

keys_new = []
for a in keys:
    for b in a:
        hlp_list = []
        for c in b:
            hlp_list.append(c)
    keys_new.append(hlp_list)

attr_dict = {}  # Создаем словарь, ID атрибута : [последовтельность из значений атрибута]
for a, a1 in enumerate(attrbs):
    if attrbs[a][1] in opm_attrbs_dict.keys():
        if opm_attrbs_dict.get(attrbs[a][1]) not in attr_dict.keys():
            attr_dict[opm_attrbs_dict.get(attrbs[a][1])] = [attrbs[a][2]]
        else:
            attr_dict[opm_attrbs_dict.get(attrbs[a][1])].append(attrbs[a][2])
    else:
        print(f"Атрибут с id: {attrbs[a][1]} не найден в базе ОРМ ")

full_data = [item_reference_keys_list]

if values_new:
    for a_id, a in enumerate(keys_new):
        full_data.append(values_new[a_id])
else:
    full_data += keys_new

ExcelWrite.write_massive_to_excel(full_data, col=1, row=1)
ExcelWrite.write_massive_to_excel(attr_dict, col=1, row=1)
