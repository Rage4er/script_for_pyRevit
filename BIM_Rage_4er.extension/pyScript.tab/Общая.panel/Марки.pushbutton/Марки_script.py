# -*- coding: utf-8 -*-
"""
Упрощенная версия для разнесения марок
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

def simple_tag_arrangement():
    """Простое размещение марок в шахматном порядке"""
    
    # Получаем все марки на виде
    collector = FilteredElementCollector(doc, view.Id)
    tags = collector.OfClass(IndependentTag).ToElements()
    
    if not tags:
        forms.alert("Марки не найдены на активном виде!")
        return
    
    # Сортируем марки по Y координате (сверху вниз)
    sorted_tags = []
    for tag in tags:
        try:
            pos = tag.TagHeadPosition
            sorted_tags.append((pos.Y, pos.X, tag))
        except:
            continue
    
    sorted_tags.sort(reverse=True)  # Сверху вниз
    
    with Transaction(doc, "Разнесение марок") as t:
        t.Start()
        
        offset_x = 2.0  # Смещение по X в футах
        current_offset = 0.0
        
        for i, (y, x, tag) in enumerate(sorted_tags):
            try:
                # Чередуем смещение для шахматного порядка
                if i % 2 == 0:
                    new_x = x + current_offset
                else:
                    new_x = x - current_offset
                    current_offset += offset_x
                
                new_pos = XYZ(new_x, y, 0)
                tag.TagHeadPosition = new_pos
                
            except Exception as e:
                print("Ошибка с маркой {}: {}".format(tag.Id, e))
        
        t.Commit()
    
    forms.alert("Разнесение завершено!\nОбработано марок: {}".format(len(sorted_tags)))

# Запуск
simple_tag_arrangement()

