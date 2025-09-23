# -*- coding: utf-8 -*-

__title__ = "Сортировка систем воздуховодов"
__author__ = 'Rage'
__doc__ = "Добавляет префикс к параметру <Имя системы> систем воздуховодов по порядку от П к ДВ"

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Получаем активный документ
doc = __revit__.ActiveUIDocument.Document

# Создание транзакции
t = Transaction(doc, "Нумерация систем воздуховодов")

# Начинаем транзакцию
t.Start()

# Определение списка интересующих систем воздуховодов
all_air_systems = FilteredElementCollector(doc)\
                  .OfCategory(BuiltInCategory.OST_DuctSystem)\
                  .WhereElementIsNotElementType()\
                  .ToElements()

# Список сортировки ключевых префиксов
sort_order = ['П', 'ПЕ', 'В', 'ВЕ', 'ДП', 'ДПЕ', 'ДВ', 'ДВЕ', 'А', 'У']

def extract_sort_key(system_name):
    # Функция возвращает индекс префикса из sort_order, иначе None
    for i, prefix in enumerate(sort_order):
        if system_name.startswith(prefix):
            return i
    return len(sort_order)  # ставим элементы без указанных префиксов последними

# Создаем словарь для хранения систем с соответствующими ключами
sorted_systems = sorted(all_air_systems, key=lambda x: extract_sort_key(x.LookupParameter("Имя системы").AsString()))

# Проходим по списку и добавляем порядковые номера
for idx, air_system in enumerate(sorted_systems):
    param_name = air_system.LookupParameter("Имя системы")
    original_value = param_name.AsString()
    
    # Формируем новое значение имени системы с добавлением префикса и индекса
    new_value = "{index}-{name}".format(index=idx + 1, name=original_value)
    
    # Записываем новый порядок в параметр "ADSK_Группирование"
    grouping_param = air_system.LookupParameter("ADSK_Группирование")
    grouping_param.Set(new_value)

# Завершаем транзакцию
try:
    t.Commit()
except Exception as e:
    print("Произошла ошибка: {}".format(e))