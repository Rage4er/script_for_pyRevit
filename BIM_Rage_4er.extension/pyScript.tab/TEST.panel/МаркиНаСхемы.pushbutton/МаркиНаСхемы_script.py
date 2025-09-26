# -*- coding: utf-8 -*-
__title__ = 'Марки на схемы'
__author__ = 'Rage'
__doc__ = 'Скрипт для простановки марки элементов систем воздуховодов в 3D видах'

from pyrevit import forms, output, script
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, BuiltInParameter, BuiltInCategory, ElementId, ViewType
)

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output.close_others(all_open_outputs=True)
output.set_title(__title__)
output.set_width(1500)
lfy = output.linkify

def extract_system_name(element):
    """Извлекает значение параметра 'Имя системы' из элемента"""
    param = element.LookupParameter(u"Имя системы")
    return param.AsString() if param else u"Не указано"

def process_element(element):
    """Обрабатывает элемент, извлекая его ID, тип изоляции и категорию."""
    data = {
        'id': element.Id.IntegerValue,
        'isolation': None,
        'category': None
    }
    
    # Изоляция
    iso_param = element.LookupParameter(u"Тип изоляции")
    data['isolation'] = iso_param.AsString() if iso_param else u"Нет изоляции ⛔"
    
    # Категория
    category = element.Category
    data['category'] = category.Name if category else u"Без категории"
    
    return data

def group_by_system(elements):
    """Группирует элементы по параметру 'Имя системы'"""
    groups = {}
    for el in elements:
        sys_name = extract_system_name(el)
        if sys_name not in groups:
            groups[sys_name] = []
        groups[sys_name].append(process_element(el))
    return groups

def get_3d_views():
    """Получает все 3D виды в проекте"""
    collector = FilteredElementCollector(doc).OfClass(View)
    return [v for v in collector if v.ViewType == ViewType.ThreeD and not v.IsTemplate]

def get_elements_from_3d_views(selected_views):
    """Получает элементы из выбранных 3D видов"""
    all_elements = set()
    for view in selected_views:
        try:
            # Элементы в данном 3D виде
            collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
            elements_in_view = collector.ToElements()
            
            # Целевые категории
            categories = [
                BuiltInCategory.OST_DuctCurves,       # Воздуховоды
                BuiltInCategory.OST_DuctFitting,      # Детали соединений
                BuiltInCategory.OST_DuctTerminal,     # Распредустройства
                BuiltInCategory.OST_DuctAccessory,    # Адаптеры
                BuiltInCategory.OST_MechanicalEquipment,  # Оборудование
                BuiltInCategory.OST_DuctInsulations  # Изоляционные материалы
            ]
            
            for elem in elements_in_view:
                if elem.Category and elem.Category.Id.IntegerValue in [c.value__ for c in categories]:
                    all_elements.add(elem)
        except Exception as ex:
            print(u"Ошибка при обработке вида {}: {}".format(view.Name, ex))
    
    return list(all_elements)

def select_3d_views():
    """Позволяет пользователю выбрать 3D виды для анализа"""
    views_3d = get_3d_views()
    
    if not views_3d:
        forms.alert("В проекте отсутствуют 3D виды.", exitscript=True)
    
    # Формирование удобочитаемых вариантов
    options = ["{} ({})".format(v.Name, v.Id.IntegerValue) for v in views_3d]
    selected_options = forms.SelectFromList.show(options, title="Выберите 3D виды", multiselect=True, button_name='Выбрать')
    
    if not selected_options:
        forms.alert("Необходимо выбрать хотя бы один 3D вид.", exitscript=True)
    
    # Извлечение выбранных видов по их Id
    selected_views = []
    for opt in selected_options:
        # Парсим строку, чтобы получить ID вида
        parts = opt.split("(", 1)
        if len(parts) > 1:
            view_id_str = parts[1].rsplit(")", 1)[0].strip()
            try:
                view_id = int(view_id_str)
                for view in views_3d:
                    if view.Id.IntegerValue == view_id:
                        selected_views.append(view)
                        break
            except ValueError:
                continue
    
    return selected_views

def main():
    """Основная логика программы"""
    selected_views = select_3d_views()
    print(u"Выбрано {} 3D видов.".format(len(selected_views)))
    
    # Получить элементы из выбранных видов
    elements = get_elements_from_3d_views(selected_views)
    
    if not elements:
        forms.alert("Ни одного подходящего элемента не обнаружено в указанных видах.", exitscript=True)
    
    # Группировка элементов по системам
    grouped_elements = group_by_system(elements)
    
    # Вывод данных
    for sys_name, els in grouped_elements.items():
        rows = [[lfy(ElementId(el['id'])), el['category'], el['isolation']] for el in els]
        output.print_table(
            table_data=rows,
            title=u"Система: {}\nВиды: {}".format(sys_name, ", ".join([v.Name for v in selected_views])),
            columns=['Идентификатор (ID)', 'Категория', 'Тип изоляции']
        )

# Главное условие запуска
if __name__ == "__main__":
    result = TaskDialog.Show(
        "Подтверждение выполнения",
        "Проставить марки элементов систем воздуховодов в 3D видах?",
        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    )
    
    if result == TaskDialogResult.Yes:
        print("Начинаем проверку...")
        main()
    else:
        print("Операция отменена пользователем.")