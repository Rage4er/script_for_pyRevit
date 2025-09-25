# -*- coding: utf-8 -*-
__title__ = 'марки на схемы'
__author__ = 'Rage'
__doc__ = '_'

from pyrevit import forms, output, script
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
from Autodesk.Revit.DB import FilteredElementCollector, View, BuiltInParameter, BuiltInCategory, ElementId, ViewType

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

def get_3d_views():
    """Получает все 3D виды в проекте"""
    collector = FilteredElementCollector(doc).OfClass(View)
    views_3d = [view for view in collector if view.ViewType == ViewType.ThreeD and not view.IsTemplate]
    return views_3d

def get_elements_from_3d_views(selected_views):
    """Получает элементы из выбранных 3D видов"""
    all_elements = set()
    
    for view in selected_views:
        try:
            # Получаем элементы, видимые в данном 3D виде
            collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
            elements_in_view = collector.ToElements()
            
            # Фильтруем только элементы целевых категорий
            target_categories = [
                BuiltInCategory.OST_DuctCurves,      # Воздуховоды
                BuiltInCategory.OST_DuctFitting,     # Соединительные детали воздуховодов
                BuiltInCategory.OST_DuctTerminal,    # Воздухораспределители
                BuiltInCategory.OST_DuctAccessory,   # Арматура воздуховодов
                BuiltInCategory.OST_MechanicalEquipment,  # Оборудование
                BuiltInCategory.OST_DuctInsulations  # Материалы изоляции воздуховодов
            ]
            
            for element in elements_in_view:
                if element.Category and element.Category.Id.IntegerValue in [cat.value__ for cat in target_categories]:
                    all_elements.add(element)
                    
        except Exception as e:
            print(u"Ошибка при обработке вида '{}': {}".format(view.Name, str(e)))
    
    return list(all_elements)

def select_3d_views():
    """Позволяет пользователю выбрать 3D виды для анализа"""
    views_3d = get_3d_views()
    
    if not views_3d:
        forms.alert("В проекте не найдено 3D видов!", exitscript=True)
    
    # Создаем список для выбора
    view_options = [u"{} (ID: {})".format(view.Name, view.Id) for view in views_3d]
    
    # Диалог выбора с множественным выбором
    selected_view_indices = forms.SelectFromList.show(
        view_options,
        title="Выберите 3D виды для анализа",
        multiselect=True,
        button_name='Выбрать'
    )
    
    if not selected_view_indices:
        forms.alert("Не выбрано ни одного 3D вида!", exitscript=True)
    
    # Получаем выбранные виды
    selected_views = [views_3d[i] for i in selected_view_indices]
    return selected_views

def main():
    """
    Основная программа с выбором 3D видов и группировкой элементов по параметру 'Имя системы'.
    """
    # Выбираем 3D виды
    selected_views = select_3d_views()
    
    print(u"Выбрано 3D видов: {}".format(len(selected_views)))
    for i, view in enumerate(selected_views, 1):
        print(u"{}. {}".format(i, view.Name))
    
    # Получаем элементы из выбранных 3D видов
    elements = get_elements_from_3d_views(selected_views)
    
    if not elements:
        forms.alert("В выбранных 3D видах не найдено элементов систем воздуховодов!", exitscript=True)
    
    print(u"\nНайдено элементов в выбранных видах: {}".format(len(elements)))
    
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
        
        # Добавляем информацию о том, в каких видах отображается система
        views_info = u", ".join([view.Name for view in selected_views])
        
        output.print_table(
            table_data=table_data,
            title='<span style="color: blue;">Система: \'%s\'</span><br/>'
                  '<span style="color: green;">3D виды: %s</span><br/>' % (system_name, views_info),
            columns=["Идентификатор (ID)", "Категория", "Тип изоляции"],
            formats=None
        )

# Основное тело программы
if __name__ == '__main__':
    # Диалоговое окно с вопросом
    result = TaskDialog.Show(
        "Подтверждение выполнения",
        "Хотите запустить проверку изоляции воздуховодов в 3D видах?",
        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    )

    # Подтверждение запуска
    if result == TaskDialogResult.Yes:
        print("Начало проверки изоляции в 3D видах.")
        main()
    else:
        # Ничего не делаем и завершаем скрипт
        pass