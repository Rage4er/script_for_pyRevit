# -*- coding: utf-8 -*-
"""
Исправленная версия маркировки элементов активного вида
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

def simple_tag_from_active_view():
    """Упрощенная маркировка элементов активного вида"""
    
    # Проверяем активный вид
    if not view or view.ViewType == ViewType.ProjectBrowser:
        forms.alert("Активный вид не поддерживает маркировку!")
        return
    
    output = script.get_output()
    
    # Безопасное получение имени вида
    view_name = "Неизвестный вид"
    try:
        if hasattr(view, 'Name'):
            view_name = view.Name
        elif hasattr(view, 'ViewName'):
            view_name = view.ViewName
    except:
        view_name = "Активный вид"
    
    output.print_md("# 🏷️ Маркировка элементов активного вида")
    output.print_md("**Вид:** {}".format(view_name))
    output.print_md("---")
    
    # Получаем категории из активного вида
    categories_count = {}
    try:
        elements_collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
        
        for element in elements_collector:
            if element and element.Category:
                cat_name = element.Category.Name
                categories_count[cat_name] = categories_count.get(cat_name, 0) + 1
    except Exception as e:
        forms.alert("Ошибка при получении элементов: {}".format(str(e)))
        return
    
    if not categories_count:
        forms.alert("На активном виде не найдено элементов!")
        return
    
    # Подготовка списка категорий для выбора
    category_options = []
    for cat_name, count in sorted(categories_count.items(), key=lambda x: x[1], reverse=True):
        category_options.append("{} ({} элементов)".format(cat_name, count))
    
    # Выбор категорий
    selected_options = forms.SelectFromList.show(
        category_options,
        title="Выберите категории для маркировки",
        multiselect=True,
        button_name='Продолжить'
    )
    
    if not selected_options:
        return
    
    # Извлекаем названия категорий
    selected_categories = []
    for option in selected_options:
        category_name = option.split(" (")[0]
        selected_categories.append(category_name)
    
    # Выбор типа марки
    tag_options = {}
    try:
        tag_types_collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
        tag_types_collector = tag_types_collector.OfCategory(BuiltInCategory.OST_GenericAnnotation)
        
        for tag_type in tag_types_collector:
            if tag_type and tag_type.Family:
                key = "{} : {}".format(tag_type.Family.Name, tag_type.Name)
                tag_options[key] = tag_type
    except Exception as e:
        print("Ошибка при получении типов марок: {}".format(e))
    
    selected_tag_type = None
    if tag_options:
        selected_tag_key = forms.SelectFromList.show(
            sorted(tag_options.keys()),
            title="Выберите тип марки",
            button_name='Выбрать'
        )
        
        if not selected_tag_key:
            return
        
        selected_tag_type = tag_options[selected_tag_key]
    else:
        forms.alert("В проекте не найдено типов марок. Будет использован стандартный тип.")
    
    # Настройки выносок
    try:
        leader_option = forms.alert(
            "Добавлять выноски к маркам?",
            yes=True, no=True,
            yes_name="С выносками", 
            no_name="Без выносок"
        )
        has_leader = leader_option
    except:
        has_leader = True  # По умолчанию с выносками
    
    # Получаем существующие марки на виде
    tagged_elements = set()
    try:
        existing_tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag).ToElements()
        
        for tag in existing_tags:
            try:
                if hasattr(tag, 'TaggedLocalElementId'):
                    tagged_elements.add(tag.TaggedLocalElementId)
                elif hasattr(tag, 'GetTaggedLocalElementIds'):
                    element_ids = tag.GetTaggedLocalElementIds()
                    for elem_id in element_ids:
                        tagged_elements.add(elem_id)
            except:
                pass
    except Exception as e:
        print("Ошибка при получении существующих марок: {}".format(e))
    
    # Собираем элементы без марок
    elements_to_tag = []
    try:
        for element in elements_collector:
            if element and element.Category and element.Category.Name in selected_categories:
                if element.Id not in tagged_elements:
                    elements_to_tag.append(element)
    except Exception as e:
        forms.alert("Ошибка при фильтрации элементов: {}".format(str(e)))
        return
    
    if not elements_to_tag:
        forms.alert("Не найдено элементов без марок для выбранных категорий!")
        return
    
    output.print_md("**Найдено элементов для маркировки:** {}".format(len(elements_to_tag)))
    output.print_md("**Выбранные категории:** {}".format(", ".join(selected_categories)))
    if selected_tag_type:
        output.print_md("**Тип марки:** {} : {}".format(selected_tag_type.Family.Name, selected_tag_type.Name))
    output.print_md("**Выноски:** {}".format("Да" if has_leader else "Нет"))
    output.print_md("---")
    
    # Выполняем маркировку
    success_count = 0
    failed_count = 0
    
    with Transaction(doc, "Маркировка элементов вида") as t:
        t.Start()
        
        for i, element in enumerate(elements_to_tag):
            try:
                output.print_html("Маркируется <b>{}</b> - {} из {}".format(
                    element.Category.Name, i+1, len(elements_to_tag)))
                
                # Определяем позицию для марки
                tag_position = XYZ(0, 0, 0)
                try:
                    bbox = element.get_BoundingBox(view)
                    if bbox:
                        # Размещаем марку ниже элемента
                        tag_position = XYZ((bbox.Min.X + bbox.Max.X) / 2, bbox.Min.Y - 2.0, 0)
                    else:
                        loc = element.Location
                        if loc:
                            if isinstance(loc, LocationPoint):
                                tag_position = loc.Point
                            elif isinstance(loc, LocationCurve):
                                curve = loc.Curve
                                tag_position = curve.Evaluate(0.5, True)
                except:
                    # Используем позицию по умолчанию
                    pass
                
                # Создаем марку
                tag = IndependentTag.Create(
                    doc, 
                    view.Id, 
                    Reference(element), 
                    has_leader, 
                    TagOrientation.Horizontal, 
                    tag_position
                )
                
                if tag:
                    if selected_tag_type:
                        tag.ChangeTypeId(selected_tag_type.Id)
                    success_count += 1
                    output.print_html("✅ <span style='color:green'>Успешно</span>")
                else:
                    failed_count += 1
                    output.print_html("❌ <span style='color:red'>Ошибка создания</span>")
                    
            except Exception as e:
                failed_count += 1
                output.print_html("❌ <span style='color:red'>Ошибка: {}</span>".format(str(e)))
                continue
        
        t.Commit()
    
    # Итоговый отчет
    output.print_md("---")
    output.print_md("## 📊 Результаты маркировки:")
    output.print_md("- **Всего элементов:** {}".format(len(elements_to_tag)))
    output.print_md("- **Успешно размещено марок:** {}".format(success_count))
    output.print_md("- **Ошибок:** {}".format(failed_count))
    
    if success_count > 0:
        output.print_md("### ✅ Маркировка завершена успешно!")
        forms.alert("Маркировка завершена!\nУспешно размещено марок: {}".format(success_count))
    else:
        output.print_md("### ❌ Маркировка не выполнена")
        forms.alert("Не удалось разместить ни одной марки!")

# Запуск
try:
    simple_tag_from_active_view()
except Exception as e:
    forms.alert("Критическая ошибка: {}".format(str(e)))