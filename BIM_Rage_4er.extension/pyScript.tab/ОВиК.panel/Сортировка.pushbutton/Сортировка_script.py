# -*- coding: utf-8 -*-

__title__ = "Сортировка систем воздуховодов"
__author__ = 'Rage'
__doc__ = "Добавляет префикс к параметру <Имя системы> систем воздуховодов по порядку от П к ДВ"

import clr
import re
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document

def extract_prefix_and_number(system_name):
    """Извлекает префикс и число из имени системы"""
    sort_order = ['ПЕ', 'ВЕ', 'ДПЕ', 'ДВЕ', 'П', 'В', 'ДП', 'ДВ', 'А', 'У']
    system_name_upper = system_name.upper()
    
    for prefix in sorted(sort_order, key=len, reverse=True):
        if system_name_upper.startswith(prefix.upper()):
            numbers = re.findall(r'\d+', system_name[len(prefix):])
            return prefix, int(numbers[0]) if numbers else 0
    return "", 0

def main():
    """Основная функция выполнения"""
    
    # Собираем все элементы воздуховодов
    categories = [
        BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_FlexDuctCurves,
        BuiltInCategory.OST_DuctInsulations, BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_DuctFitting
    ]
    
    all_elements = {}
    for category in categories:
        for element in FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType():
            all_elements[element.Id] = element

    # Получаем и сортируем системы
    systems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType().ToElements()
    
    sort_order = ['П', 'ПЕ', 'В', 'ВЕ', 'ДП', 'ДПЕ', 'ДВ', 'ДВЕ', 'А', 'У']
    
    def get_sort_key(system):
        name_param = system.LookupParameter("Имя системы")
        if not name_param: return (len(sort_order) + 1, 0)
        prefix, number = extract_prefix_and_number(name_param.AsString() or "")
        return (sort_order.index(prefix) if prefix in sort_order else len(sort_order), number)
    
    sorted_systems = sorted(systems, key=get_sort_key)

    # Обработка в транзакции
    with Transaction(doc, "Нумерация систем воздуховодов") as t:
        t.Start()
        
        for idx, system in enumerate(sorted_systems, 1):
            try:
                original_name = system.LookupParameter("Имя системы").AsString() or "Без имени"
                new_value = "{}. {}".format(idx, original_name)
                
                # Обновляем системный параметр
                grouping_param = system.LookupParameter("ADSK_Группирование")
                if grouping_param and not grouping_param.IsReadOnly:
                    grouping_param.Set(new_value)
                
                # Обновляем элементы системы
                updated_count = 0
                for element in system.DuctNetwork:
                    if element.Id in all_elements:
                        name_param = element.LookupParameter("ИмяСистемы")
                        if name_param and not name_param.IsReadOnly:
                            name_param.Set(new_value)
                            updated_count += 1
                
                print("{:3d}. {:<30} -> {:<35} [элементов: {}]".format(idx, original_name, new_value, updated_count))
                
            except Exception as e:
                print("{:3d}. ОШИБКА: {}".format(idx, e))
        
        t.Commit()

if __name__ == "__main__":
    main()