<file_path>
BIM_Rage_4er.extension\pyScript.tab\TEST.panel\Mega.pushbutton\mega_script.py
</file_path>

<edit_description>
Implement script for optimizing DuctTag positions using spiral search
</edit_description>

# -*- coding: utf-8 -*-
__title__ = "Оптимизация марок воздуховодов"
__author__ = 'Rage'
__doc__ = "Перемещает марки воздуховодов на виде в свободные позиции с помощью поиска по спирали. Обрабатывает выделенные марки или все на виде."
__version__ = "1.0"

import clr
import math

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Доступ к документу и UI
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Константы
STEP_SIZE = 1.0  # Футы, шаг спирали
INITIAL_OFFSET = 1.0  # Футы, начальное смещение
MAX_ATTEMPTS = 10  # Максимум попыток спирали
DUCT_TAG_CATEGORY = BuiltInCategory.OST_DuctTags  # Марки воздуховодов
ELEMENT_CATEGORIES = [  # Категории элементов для проверки пересечений
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_MechanicalEquipment
]

def get_tags_on_view(view):
    """Возвращает список марок DuctTags на заданном виде."""
    collector = FilteredElementCollector(doc).OfCategory(DUCT_TAG_CATEGORY).WhereElementIsNotElementType()
    tags = [tag for tag in collector if view.Id == tag.OwnerViewId]
    return tags

def get_visible_elements_bbox(view):
    """Возвращает список Outline (bbox) видимых элементов на виде."""
    bboxes = []
    for category in ELEMENT_CATEGORIES:
        collector = FilteredElementCollector(doc, view.Id).OfCategory(category).WhereElementIsNotElementType()
        for elem in collector:
            bbox = elem.get_BoundingBox(view)
            if bbox:
                bboxes.append(bbox)
    return bboxes

def get_tag_bbox(tag, view):
    """Возвращает Outline для марки на виде."""
    bbox = tag.get_BoundingBox(view)
    return bbox if bbox else None

def does_bbox_intersect(test_bbox, other_bboxes):
    """Проверяет, пересекается ли test_bbox с любым из other_bboxes."""
    for bbox in other_bboxes:
        if test_bbox.Intersects(bbox, view.Zoom):
            return True
    return False

def find_free_position(start_point, tag_bbox, other_bboxes, view):
    """Находит свободную позицию с помощью спирального поиска."""
    # Получаем ориентацию вида для преобразований
    view_orientation = view.GetOrientation()
    right_dir = view_orientation.RightDirection
    up_dir = view_orientation.UpDirection

    # Начальная точка: смещение от старта
    current_point = start_point.Add(right_dir.Multiply(INITIAL_OFFSET)).Add(up_dir.Multiply(INITIAL_OFFSET))

    directions = [(0, 1), (1, 0), (0, -1), (-1, 0)]  # Вверх, вправо, вниз, влево
    step = STEP_SIZE
    dir_index = 0
    steps_in_dir = 1
    change_dir = 0

    for attempt in range(MAX_ATTEMPTS):
        # Создаем смещенный bbox
        offset_x = (current_point.X - start_point.X) * right_dir.X + (current_point.Y - start_point.Y) * up_dir.X
        offset_y = (current_point.X - start_point.X) * right_dir.Y + (current_point.Y - start_point.Y) * up_dir.Y
        offset_z = (current_point.X - start_point.X) * right_dir.Z + (current_point.Y - start_point.Y) * up_dir.Z
        offset = XYZ(offset_x, offset_y, offset_z)

        shifted_bbox = Outline(tag_bbox.MinimumPoint.Add(offset), tag_bbox.MaximumPoint.Add(offset))

        if not does_bbox_intersect(shifted_bbox, other_bboxes):
            return current_point

        # Двигаемся по спирали
        if change_dir == 2:
            step += STEP_SIZE
            change_dir = 0
        current_point = current_point.Add(right_dir.Multiply(directions[dir_index][0] * step)).Add(up_dir.Multiply(directions[dir_index][1] * step))
        dir_index = (dir_index + 1) % 4
        change_dir += 1

    return None  # Не найдено свободное место

def optimize_tag_positions():
    """Основная функция оптимизации марок."""
    view = doc.ActiveView
    if not view or view.ViewType != ViewType.FloorPlan:
        print("Активный вид не является плоскостным видом.")
        return

    # Получаем марки для обработки
    selection = uidoc.Selection.GetElementIds()
    if selection:
        tags = [doc.GetElement(el_id) for el_id in selection if doc.GetElement(el_id).Category.Id == Category.GetCategory(doc, DUCT_TAG_CATEGORY).Id]
        if not tags:
            print("Выделенные элементы не являются марками воздуховодов. Обрабатываем все марки на виде.")
            tags = get_tags_on_view(view)
    else:
        tags = get_tags_on_view(view)

    if not tags:
        print("На виде нет марок воздуховодов.")
        return

    # Получаем bbox других элементов
    other_bboxes = get_visible_elements_bbox(view)

    moved_count = 0
    with Transaction(doc, "Оптимизация позиций марок воздуховодов") as t:
        t.Start()
        try:
            for tag in tags:
                tag_bbox = get_tag_bbox(tag, view)
                if not tag_bbox:
                    continue

                current_point = tag.Location.Point
                if does_bbox_intersect(tag_bbox, other_bboxes):
                    free_point = find_free_position(current_point, tag_bbox, other_bboxes, view)
                    if free_point:
                        tag.Location.MoveTo(free_point)
                        moved_count += 1
                        print("Марка {} перемещена.".format(tag.Id))
                    else:
                        print("Для марки {} не найдено свободное место.".format(tag.Id))
            t.Commit()
            print("Оптимизация завершена. Перемещено марок: {}.".format(moved_count))
        except Exception as e:
            t.RollBack()
            print("Ошибка: {}".format(str(e)))

if __name__ == "__main__":
    optimize_tag_positions()
