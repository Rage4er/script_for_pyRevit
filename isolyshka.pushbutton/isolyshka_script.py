# -*- coding: utf-8 -*-
__title__ = 'isolyshka'
__author__ = 'Rage'
__doc__ = 'Вывод отчета наличия изоляции систем воздуховодов'

from pyrevit import forms, output, script
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
from Autodesk.Revit.DB import FilteredElementCollector, View, BuiltInParameter, BuiltInCategory, ElementId

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output.close_others(all_open_outputs=True)
output.set_title(__title__)
output.set_width(1500)
lfy = output.linkify

def extract_system_name(element):
    """Извлекает значение параметра 'Имя системы' из элемента"""
    system_name_param = element.LookupParameter(u"Имя системы")
    if system_name_param:
        return system_name_param.AsString() or u"Не указано"
    return u"Не указано"

def process_element(element):
    """Обрабатывает элемент, извлекая его ID, тип изоляции и категорию."""
    data = {}
    data['id'] = element.Id.IntegerValue
    
    # Получаем тип изоляции
    isolation_param = element.LookupParameter(u"Тип изоляции")
    if isolation_param:
        data['isolation'] = isolation_param.AsString() or u"Нет изоляции ⛔"
    else:
        data['isolation'] = u"Нет изоляции ⛔"
    
    # Получаем категорию элемента
    category = element.Category
    if category:
        data['category'] = category.Name
    else:
        data['category'] = u"Без категории"
    
    return data

def group_by_system(elements):
    """Группирует элементы по параметру 'Имя системы'"""
    grouped_elements = {}
    for element in elements:
        system_name = extract_system_name(element)
        if system_name not in grouped_elements:
            grouped_elements[system_name] = []
        grouped_elements[system_name].append(process_element(element))
    return grouped_elements

def main(categories=None):
    """
    Основная программа с возможностью выбора одной или нескольких категорий,
    и дополнительно с группировкой элементов по параметру 'Имя системы'.
    """
    # Получаем активное окно Revit
    ui_document = __revit__.ActiveUIDocument
    document = ui_document.Document
    
    # Получаем активный вид
    active_view = document.ActiveView
    
    # Создаем коллекцию элементов активного вида
    collector = FilteredElementCollector(document, active_view.Id).WhereElementIsNotElementType()
    
    # Если переданы категории, применяем фильтр по ним
    if categories is not None and len(categories) > 0:
        collectors = []
        for cat_name in categories:
            try:
                # Проверяем введенную категорию и назначаем внутреннюю категорию
                if cat_name.lower() == u"воздуховоды (сегменты)":
                    built_in_category = BuiltInCategory.OST_DuctCurves
                elif cat_name.lower() == u"соединительные детали воздуховодов":
                    built_in_category = BuiltInCategory.OST_DuctFitting
                else:
                    print(u"Ошибка: Категория '%s' не поддерживается." % cat_name)
                    continue
                
                # Создаем отдельный коллектор для каждой категории
                cat_collector = FilteredElementCollector(document, active_view.Id)\
                                    .WhereElementIsNotElementType().OfCategory(built_in_category)
                collectors.append(cat_collector.ToElements())
            except AttributeError:
                print(u"Ошибка: Категория '%s' не найдена." % cat_name)
                continue
        
        # Объединяем полученные наборы элементов вручную
        elements = set()
        for col_elements in collectors:
            elements.update(col_elements)
    else:
        elements = collector.ToElements()
    
    # Группируем элементы по параметру 'Имя системы'
    grouped_elements = group_by_system(elements)
    
    # Проходим по каждой группе и выводим таблицу
    for system_name, elements_list in grouped_elements.items():
        # Формируем таблицу с использованием pyRevit
        table_data = []
        for element_data in elements_list:
            # Преобразуем id в ссылку с помощью linkify
            linked_id = lfy(ElementId(element_data['id']))
            row = [
                linked_id,                  # Теперь это ссылка!
                element_data['category'],
                element_data['isolation']
            ]
            table_data.append(row)
        output.print_table(
            table_data=table_data,
            title='<span style="color: blue;">Система: \'%s\'</span><br/>' % system_name,
            columns=["Идентификатор (ID)", "Категория", "Тип изоляции"],
            formats=None
        )

# Основное тело программы
if __name__ == '__main__':
    # Диалоговое окно с вопросом
    result = TaskDialog.Show(
        "Подтверждение выполнения",
        "Хотели бы запустить проверку изоляции воздуховодов?",
        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    )

    # Подтверждение запуска
    if result == TaskDialogResult.Yes:
        print("Начало проверки изоляции.")
        main([u"Воздуховоды (сегменты)", u"Соединительные детали воздуховодов"])
    else:
        # Ничего не делаем и завершаем скрипт
        pass