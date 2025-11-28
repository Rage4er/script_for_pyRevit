# -*- coding: utf-8 -*-
__title__ = "Экспорт вида в PNG (HQ)"
__author__ = "Rage"
__doc__ = "Экспортирует активный вид в изображение PNG с улучшенными параметрами качества (600 DPI, 4096 пикселей)"
__version__ = "1.0"

import os

from Autodesk.Revit.DB import (
    ExportRange,
    ImageExportOptions,
    ImageFileType,
    ImageResolution,
    View,
    ViewSet,
)
from Autodesk.Revit.UI import TaskDialog
from System import Enum


def export_view_to_png(doc, view, export_path):
    """
    Экспортирует указанный вид в PNG-изображение

    Args:
        doc: Документ Revit
        view: Вид для экспорта
        export_path: Путь к файлу экспорта (без расширения)
    """
    # Создаем опции экспорта изображения
    options = ImageExportOptions()

    # Настраиваем параметры экспорта для максимального качества
    options.ExportRange = ExportRange.CurrentView
    options.FilePath = export_path
    options.HLRandWFViewsFileType = ImageFileType.PNG
    options.ImageResolution = ImageResolution.DPI_600  # Очень высокое качество
    options.ZoomType = Enum.Parse(
        options.ZoomType.GetType(), "FitToPage"
    )  # Масштабировать по размеру страницы
    options.PixelSize = 4096  # Установить размер пикселей для высокого разрешения

    # Создаем набор видов для экспорта
    view_set = ViewSet()
    view_set.Insert(view)

    # Экспортируем вид
    try:
        doc.ExportImage(options)
        return True
    except Exception as e:
        TaskDialog.Show("Ошибка экспорта", "Не удалось экспортировать вид: " + str(e))
        return False


def main():
    # Получаем активный документ
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document
    active_view = doc.ActiveView

    # Проверяем, что активный вид существует
    if not active_view:
        TaskDialog.Show("Ошибка", "Активный вид не найден")
        return

    # Определяем путь для экспорта (в той же папке, что и скрипт)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    export_path = os.path.join(script_dir, "tmp")

    # Экспортируем вид в PNG
    if export_view_to_png(doc, active_view, export_path):
        TaskDialog.Show("Успех", "Вид успешно экспортирован в PNG")
    else:
        TaskDialog.Show("Ошибка", "Не удалось экспортировать вид")


# Запуск основной функции
if __name__ == "__main__":
    main()
