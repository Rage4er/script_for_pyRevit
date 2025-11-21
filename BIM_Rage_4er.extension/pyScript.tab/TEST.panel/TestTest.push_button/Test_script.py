# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

logger = script.get_logger()

def select_and_parse_families():
    """Выбор семейств из списка всех в документе и парсинг имен"""
    
    # Получаем все семейства в документе
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
    all_families = list(collector)
    
    if not all_families:
        forms.alert("В документе не найдено семейств!")
        return
    
    # Создаем список для выбора
    family_options = []
    for family in all_families:
        if family.IsEditable and family.Name:  # Только редактируемые семейства с именем
            family_options.append(family)
    
    # Диалог выбора семейств
    selected_families = forms.SelectFromList.show(
        family_options,
        name_attr='Name',
        title="Выберите семейства для парсинга",
        multiselect=True,
        button_name='Выбрать'
    )
    
    if not selected_families:
        return
    
    # Парсим имена выбранных семейств
    results = []
    for family in selected_families:
        family_name = family.Name
        if "_" in family_name:
            prefix = family_name.split("_")[0]
            results.append("{} → {}".format(family_name, prefix))
        else:
            results.append("{} → (нет разделителя)".format(family_name))
    
    # Показываем результаты
    if results:
        forms.alert(
            "Результаты парсинга:\n\n{}".format("\n".join(results)),
            title="Парсинг имен семейств"
        )

# Версия с работой с типоразмерами семейств
def select_and_parse_family_types():
    """Выбор типоразмеров семейств и парсинг имен"""
    
    # Получаем все типоразмеры семейств в документе
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol)
    all_family_types = list(collector)
    
    if not all_family_types:
        forms.alert("В документе не найдено типоразмеров семейств!")
        return
    
    # Создаем список для выбора
    type_options = []
    for family_type in all_family_types:
        if family_type.Family and family_type.Family.Name:
            display_name = "{} : {}".format(family_type.Family.Name, family_type.Name)
            type_options.append((family_type, display_name))
    
    # Диалог выбора
    selected_types = forms.SelectFromList.show(
        type_options,
        name_attr='display_name',  # Используем кортежи с отображаемыми именами
        title="Выберите типоразмеры семейств",
        multiselect=True,
        button_name='Выбрать'
    )
    
    if not selected_types:
        return
    
    # Парсим имена
    results = []
    for family_type, display_name in selected_types:
        full_name = family_type.Family.Name
        if "_" in full_name:
            prefix = full_name.split("_")[0]
            results.append("{} → {}".format(full_name, prefix))
        else:
            results.append("{} → (нет разделителя)".format(full_name))
    
    # Результаты
    if results:
        forms.alert("Результаты:\n\n{}".format("\n".join(results)))

# Упрощенная версия
def quick_family_selection():
    """Быстрый выбор и парсинг семейств"""
    
    # Получаем все семейства
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.Family)
    all_families = list(collector)
    
    if not all_families:
        forms.alert("Нет семейств в документе!")
        return
    
    # Выбор семейств
    selected_families = forms.SelectFromList.show(
        all_families,
        name_attr='Name',
        title="Выберите семейства",
        multiselect=True
    )
    
    if not selected_families:
        return
    
    # Парсинг
    for family in selected_families:
        family_name = family.Name
        if "_" in family_name:
            prefix = family_name.split("_")[0]
            print("{} → {}".format(family_name, prefix))
        else:
            print("{} → нет '_'".format(family_name))

if __name__ == "__main__":
    select_and_parse_families()