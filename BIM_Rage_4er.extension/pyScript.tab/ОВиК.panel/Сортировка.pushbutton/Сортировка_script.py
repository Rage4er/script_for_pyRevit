# -*- coding: utf-8 -*-

__title__ = "Сортировка систем воздуховодов"
__author__ = 'Rage'
__doc__ = "Добавляет префикс к параметру <Имя системы> систем воздуховодов по порядку от П к ДВ"

import clr
import time
import re
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import *

# Получаем активный документ
doc = __revit__.ActiveUIDocument.Document

def extract_prefix_and_number(system_name, sort_order):
    """Извлекает префикс и число из имени системы"""
    system_name_upper = system_name.upper()
    
    # Сначала проверяем более длинные префиксы, потом более короткие
    # Важно: сначала "ПЕ", потом "П"; сначала "ВЕ", потом "В" и т.д.
    sorted_prefixes = sorted(sort_order, key=len, reverse=True)
    
    for prefix in sorted_prefixes:
        if system_name_upper.startswith(prefix.upper()):
            # Извлекаем числовую часть после префикса
            number_part = system_name[len(prefix):].strip()
            # Пытаемся извлечь число
            try:
                numbers = re.findall(r'\d+', number_part)
                if numbers:
                    num = int(numbers[0])
                else:
                    num = 0
            except:
                num = 0
            return prefix, num
    
    # Если префикс не найден
    return "", 0

def get_all_duct_elements():
    """Получает все элементы воздуховодов для быстрого доступа"""
    elements_dict = {}
    
    # Собираем все элементы воздуховодов один раз
    categories = [
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_FlexDuctCurves, 
        BuiltInCategory.OST_DuctInsulations,
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_DuctAccessory,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_DuctFitting
    ]
    
    for category in categories:
        collector = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType()
        for element in collector:
            elements_dict[element.Id] = element
    
    return elements_dict

def get_sorted_duct_systems():
    """Получает и сортирует системы воздуховодов по заданному порядку префиксов"""
    
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctSystem).WhereElementIsNotElementType()
    systems = collector.ToElements()
    
    # Порядок префиксов как требуется
    sort_order = ['П', 'ПЕ', 'В', 'ВЕ', 'ДП', 'ДПЕ', 'ДВ', 'ДВЕ', 'А', 'У']
    
    def get_sort_key(system):
        system_name_param = system.LookupParameter("Имя системы")
        if not system_name_param:
            return (len(sort_order) + 1, 0)
            
        system_name = system_name_param.AsString() or ""
        
        prefix, number = extract_prefix_and_number(system_name, sort_order)
        
        if prefix in sort_order:
            prefix_index = sort_order.index(prefix)
        else:
            prefix_index = len(sort_order)
        
        return (prefix_index, number)
    
    return sorted(systems, key=get_sort_key)

def get_elements_for_system(system, all_elements_dict):
    """Получает элементы для конкретной системы"""
    system_elements = []
    
    try:
        # Получаем элементы через DuctNetwork (это самый надежный способ)
        duct_network = system.DuctNetwork
        for element in duct_network:
            if element.Id in all_elements_dict:
                system_elements.append(all_elements_dict[element.Id])
                
    except Exception as e:
        print("  Предупреждение: не удалось получить элементы системы: {}".format(e))
    
    return system_elements

def update_system_elements(system, grouping_value, all_elements_dict, errors):
    """Обновляет элементы системы воздуховодов"""
    updated_elements = 0
    
    try:
        system_elements = get_elements_for_system(system, all_elements_dict)
        
        for element in system_elements:
            try:
                # Обновляем параметр "ИмяСистемы" (общий параметр проекта)
                name_param = element.LookupParameter("ИмяСистемы")
                if name_param and not name_param.IsReadOnly:
                    name_param.Set(grouping_value)
                    updated_elements += 1
                    
            except Exception as e:
                error_msg = "Элемент {}: {}".format(element.Id, e)
                errors.append(error_msg)
                
    except Exception as e:
        error_msg = "Система {}: {}".format(system.Id, e)
        errors.append(error_msg)
    
    return updated_elements

def main():
    """Основная функция выполнения"""
    
    start_time = time.time()
    errors = []
    
    try:
        print("Подготовка данных...")
        
        # Предварительно собираем все элементы для производительности
        all_elements_dict = get_all_duct_elements()
        print("Найдено элементов воздуховодов: {}".format(len(all_elements_dict)))
        
        # Получаем отсортированные системы
        sorted_systems = get_sorted_duct_systems()
        
        if not sorted_systems:
            print("Системы воздуховодов не найдены")
            return
        
        print("Найдено систем воздуховодов: {}".format(len(sorted_systems)))
        
        # Порядок сортировки для отладки
        sort_order = ['П', 'ПЕ', 'В', 'ВЕ', 'ДП', 'ДПЕ', 'ДВ', 'ДВЕ', 'А', 'У']
        print("\nПорядок сортировки: {}".format(sort_order))
        
        # Начинаем транзакцию
        with Transaction(doc, "Нумерация систем воздуховодов") as t:
            t.Start()
            
            total_systems_updated = 0
            total_elements_updated = 0
            
            print("\nОбработка систем:")
            print("-" * 50)
            
            # Сначала выводим отладочную информацию о сортировке
            print("Отладочная информация о сортировке (первые 20 систем):")
            for idx, system in enumerate(sorted_systems[:20], 1):
                system_name_param = system.LookupParameter("Имя системы")
                if system_name_param:
                    original_name = system_name_param.AsString() or "Без имени"
                    prefix, number = extract_prefix_and_number(original_name, sort_order)
                    print("  {:2d}. {} -> префикс: '{}', число: {}".format(idx, original_name, prefix, number))
            
            for idx, system in enumerate(sorted_systems, 1):
                try:
                    # Получаем исходное имя системы
                    system_name_param = system.LookupParameter("Имя системы")
                    if not system_name_param:
                        errors.append("Система {}: не найден параметр 'Имя системы'".format(system.Id))
                        continue
                        
                    original_name = system_name_param.AsString() or "Без имени"
                    
                    # Формируем новое значение для группировки в формате "1. Имя системы"
                    new_grouping_value = "{}. {}".format(idx, original_name)
                    
                    # Обновляем параметр группировки системы
                    grouping_param = system.LookupParameter("ADSK_Группирование")
                    if grouping_param and not grouping_param.IsReadOnly:
                        grouping_param.Set(new_grouping_value)
                        total_systems_updated += 1
                    else:
                        errors.append("Система {}: не удалось обновить ADSK_Группирование".format(system.Id))
                    
                    # Обновляем элементы системы
                    elements_updated = update_system_elements(system, new_grouping_value, all_elements_dict, errors)
                    total_elements_updated += elements_updated
                    
                    print("{:3d}. {:<30} -> {:<35} [элементов: {}]".format(idx, original_name, new_grouping_value, elements_updated))
                    
                except Exception as e:
                    error_msg = "Система {}: {}".format(idx, e)
                    errors.append(error_msg)
                    print("{:3d}. ОШИБКА: {}".format(idx, e))
                    continue
            
            t.Commit()
            
            # Вывод итогов
            execution_time = time.time() - start_time
            print("\n" + "=" * 50)
            print("РЕЗУЛЬТАТЫ ВЫПОЛНЕНИЯ:")
            print("Обработано систем: {}".format(len(sorted_systems)))
            print("Обновлено систем: {}".format(total_systems_updated))
            print("Обновлено элементов: {}".format(total_elements_updated))
            print("Время выполнения: {:.2f} секунд".format(execution_time))
            
            if errors:
                print("\nОШИБКИ ({}):".format(len(errors)))
                print("-" * 30)
                for i, error in enumerate(errors[:20], 1):  # Показываем первые 20 ошибок
                    print("{:2d}. {}".format(i, error))
                if len(errors) > 20:
                    print("... и еще {} ошибок".format(len(errors) - 20))
                    
    except Exception as e:
        print("КРИТИЧЕСКАЯ ОШИБКА: {}".format(e))
        import traceback
        print(traceback.format_exc())

# Запуск
if __name__ == "__main__":
    main()