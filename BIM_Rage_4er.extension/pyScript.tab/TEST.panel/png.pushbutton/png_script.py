# -*- coding: utf-8 -*-
__title__ = "Экспорт вида в PNG (HQ) + Анализ OpenCV"
__author__ = "Rage"
__doc__ = "Экспортирует активный вид в изображение PNG и анализирует его с помощью OpenCV для определения свободных областей"
__version__ = "2.0"

import os
import clr
clr.AddReference('System')
from System import Enum

from Autodesk.Revit.DB import (
    UV, XYZ, ExportRange, ImageExportOptions, 
    ImageFileType, ImageResolution, View, ViewSet
)
from Autodesk.Revit.UI import TaskDialog

# Проверка доступности библиотек анализа изображений
PIL_AVAILABLE = False
CV_AVAILABLE = False

try:
    from PIL import Image, ImageDraw
    PIL_AVAILABLE = True
    print("PIL успешно импортирован")
except ImportError:
    print("PIL не найден")

try:
    import cv2
    import numpy as np
    CV_AVAILABLE = True
    print("OpenCV успешно импортирован")
except ImportError:
    print("OpenCV не найден")

if not PIL_AVAILABLE and not CV_AVAILABLE:
    msg = "Библиотеки PIL и OpenCV не найдены. Анализ изображения будет ограничен."
    print(msg)
    TaskDialog.Show("Предупреждение", msg)


def export_view_to_png(doc, view, export_path):
    """
    Экспортирует указанный вид в PNG-изображение высокого качества
    """
    print("Начало экспорта вида в PNG...")
    print("Путь экспорта: {}".format(export_path))

    options = ImageExportOptions()
    options.ExportRange = ExportRange.CurrentView
    options.FilePath = export_path
    options.HLRandWFViewsFileType = ImageFileType.PNG
    options.ImageResolution = ImageResolution.DPI_600
    options.ZoomType = Enum.Parse(options.ZoomType.GetType(), "FitToPage")
    options.PixelSize = 4096

    print("Параметры экспорта:")
    print("  Разрешение: {} DPI".format(600))
    print("  Размер пикселей: {}x{}".format(4096, 4096))
    print("  Тип файла: PNG")

    view_set = ViewSet()
    view_set.Insert(view)
    print("Вид добавлен в набор для экспорта")

    try:
        print("Выполнение экспорта...")
        doc.ExportImage(options)
        print("Экспорт успешно завершен")
        return True
    except Exception as e:
        error_msg = "Не удалось экспортировать вид: " + str(e)
        print("Ошибка экспорта: {}".format(error_msg))
        TaskDialog.Show("Ошибка экспорта", error_msg)
        return False


def analyze_with_opencv(image_path, output_path=None):
    """
    Анализ изображения с помощью OpenCV для поиска свободных областей
    с использованием Distance Transform
    """
    print("Анализ изображения с помощью OpenCV...")
    
    try:
        # Загрузка изображения
        image = cv2.imread(image_path)
        if image is None:
            print("Не удалось загрузить изображение")
            return None
            
        height, width = image.shape[:2]
        print("Размер изображения: {}x{}".format(width, height))
        
        # Конвертация в grayscale
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Бинаризация - настройте порог под ваш стиль Revit
        # THRESH_BINARY_INV: объекты черные (0), фон белый (255)
        ret, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY_INV)
        
        # Морфологические операции для очистки шума
        kernel = np.ones((5, 5), np.uint8)
        cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
        cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_OPEN, kernel)
        
        # Distance Transform - находим расстояния от каждого пикселя фона до ближайшего объекта
        dist_transform = cv2.distanceTransform(255 - cleaned, cv2.DIST_L2, 5)
        
        # Находим несколько лучших позиций
        free_areas = []
        temp_transform = dist_transform.copy()
        
        for i in range(10):  # Ищем до 10 лучших позиций
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(temp_transform)
            
            # Проверяем, что область достаточно большая
            if max_val > 50:  # Минимальный радиус свободной зоны
                x, y = max_loc
                radius = int(max_val)
                
                free_areas.append({
                    'pixels': (x, y),
                    'radius': radius,
                    'score': max_val,
                    'size': radius * 2
                })
                
                # "Замазываем" найденную область чтобы найти следующую
                cv2.circle(temp_transform, max_loc, int(max_val * 0.7), 0, -1)
            else:
                break
        
        print("Найдено свободных областей с OpenCV: {}".format(len(free_areas)))
        
        # Визуализация результатов
        if output_path and free_areas:
            result_image = image.copy()
            
            for i, area in enumerate(free_areas):
                x, y = area['pixels']
                radius = area['radius']
                
                # Рисуем зону свободного пространства
                cv2.circle(result_image, (x, y), radius, (0, 255, 0), 3)
                # Рисуем центр
                cv2.circle(result_image, (x, y), 15, (0, 0, 255), -1)
                # Добавляем текст (исправлено - без f-string)
                text = "Area {0}".format(i+1)
                cv2.putText(result_image, text, (x-40, y-25),
                           cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
            
            cv2.imwrite(output_path, result_image)
            print("Визуализация OpenCV сохранена: {}".format(output_path))
        
        # Конвертируем в формат, совместимый с остальным кодом
        result_areas = []
        for area in free_areas:
            x, y = area['pixels']
            size = area['size']
            result_areas.append((x, y, size, size))
            
        return result_areas
        
    except Exception as e:
        print("Ошибка анализа OpenCV: {}".format(str(e)))
        return None


def find_free_areas_advanced(image_path, output_path=None):
    """
    Умный анализ изображения с использованием доступных библиотек
    """
    print("Запуск расширенного анализа изображения...")
    
    # Пробуем OpenCV сначала (более мощный анализ)
    if CV_AVAILABLE:
        areas = analyze_with_opencv(image_path, output_path)
        if areas:
            return areas
        print("OpenCV анализ не дал результатов, пробуем другие методы...")
    
    # Пробуем PIL как запасной вариант
    if PIL_AVAILABLE:
        areas = find_free_areas_simple(image_path, output_path)
        if areas:
            return areas
    
    # Запасной вариант - равномерное распределение точек
    print("Используем запасной алгоритм распределения точек...")
    if PIL_AVAILABLE:
        image = Image.open(image_path)
        width, height = image.size
    else:
        width, height = 4096, 4096
    
    # Создаем сетку из 9 точек
    areas = []
    for i in range(3):
        for j in range(3):
            x = width * (i + 1) // 4
            y = height * (j + 1) // 4
            areas.append((x, y, 100, 100))
    
    print("Создано точек по сетке: {}".format(len(areas)))
    return areas


def find_free_areas_simple(image_path, output_path=None):
    """
    Упрощенный анализ изображения с использованием PIL
    """
    if not PIL_AVAILABLE:
        return None
        
    try:
        print("Анализ изображения с PIL...")
        image = Image.open(image_path)
        width, height = image.size
        print("Размер изображения: {}x{}".format(width, height))

        gray_image = image.convert("L")
        pixels = gray_image.load()

        brightness_threshold = 200
        print("Порог яркости: {}".format(brightness_threshold))

        # Простой анализ - находим самые светлые области
        bright_spots = []
        step = 50  # Шаг анализа
        
        for x in range(0, width, step):
            for y in range(0, height, step):
                # Проверяем область 100x100
                brightness_sum = 0
                count = 0
                
                for dx in range(0, min(100, width-x), 10):
                    for dy in range(0, min(100, height-y), 10):
                        if x+dx < width and y+dy < height:
                            brightness_sum += pixels[x+dx, y+dy]
                            count += 1
                
                if count > 0:
                    avg_brightness = brightness_sum / count
                    if avg_brightness > brightness_threshold:
                        bright_spots.append((x + 50, y + 50, avg_brightness))
        
        # Сортируем по яркости и берем лучшие
        bright_spots.sort(key=lambda x: x[2], reverse=True)
        free_areas = [(x, y, 100, 100) for x, y, brightness in bright_spots[:10]]
        
        print("Найдено светлых областей с PIL: {}".format(len(free_areas)))
        
        # Визуализация
        if output_path and free_areas:
            color_image = image.convert("RGB")
            draw = ImageDraw.Draw(color_image)
            
            for i, (x, y, w, h) in enumerate(free_areas):
                draw.rectangle([x-50, y-50, x+50, y+50], outline=(0, 255, 0), width=3)
                draw.ellipse([x-8, y-8, x+8, y+8], fill=(255, 0, 0))
                if i < 5:
                    print("Область {0}: ({1}, {2}) яркость: {3:.1f}".format(i+1, x, y, bright_spots[i][2]))
            
            color_image.save(output_path)
            print("Визуализация PIL сохранена")
        
        return free_areas if free_areas else None
        
    except Exception as e:
        print("Ошибка анализа PIL: {}".format(str(e)))
        return None


def pixel_to_uv(doc, view, pixel_x, pixel_y, image_width, image_height):
    """
    Преобразует координаты пикселей в UV-координаты вида Revit
    """
    print("Преобразование пикселей ({0}, {1}) в UV...".format(pixel_x, pixel_y))

    crop_box = view.CropBox
    min_pt = crop_box.Min
    max_pt = crop_box.Max
    
    print("Границы вида: Min({0:.2f}, {1:.2f}), Max({2:.2f}, {3:.2f})".format(
        min_pt.X, min_pt.Y, max_pt.X, max_pt.Y))
    
    view_width = max_pt.X - min_pt.X
    view_height = max_pt.Y - min_pt.Y
    
    print("Размеры области вида: {0:.2f} x {1:.2f}".format(view_width, view_height))

    # Преобразование с инверсией Y
    u = min_pt.X + (float(pixel_x) / image_width) * view_width
    v = max_pt.Y - (float(pixel_y) / image_height) * view_height

    print("Результат: UV({0:.2f}, {1:.2f})".format(u, v))
    return UV(u, v)


def main():
    print("=" * 60)
    print("Скрипт экспорта вида в PNG + Анализ OpenCV/PIL")
    print("=" * 60)

    # Получаем активный документ
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document
    active_view = doc.ActiveView
    
    if not active_view:
        error_msg = "Активный вид не найден"
        print(error_msg)
        TaskDialog.Show("Ошибка", error_msg)
        return

    print("Активный вид: {}".format(active_view.Name))

    # Определяем пути
    script_dir = os.path.dirname(os.path.abspath(__file__))
    export_path = os.path.join(script_dir, "tmp")
    image_path = export_path + ".png"
    
    print("Путь к скрипту: {}".format(script_dir))
    print("Путь экспорта: {}".format(export_path))

    # Экспортируем вид
    print("\n--- Экспорт вида ---")
    if not export_view_to_png(doc, active_view, export_path):
        return

    # Анализируем изображение
    print("\n--- Анализ изображения ---")
    analysis_path = os.path.join(script_dir, "analysis_result.png")
    free_areas = find_free_areas_advanced(image_path, analysis_path)

    if not free_areas:
        msg = "Не найдено свободных областей для размещения марок"
        print(msg)
        TaskDialog.Show("Анализ", msg)
        return

    # Получаем размеры изображения
    if PIL_AVAILABLE:
        with Image.open(image_path) as img:
            image_width, image_height = img.size
    else:
        image_width, image_height = 4096, 4096

    print("Размеры изображения: {}x{}".format(image_width, image_height))

    # Преобразуем координаты всех найденных областей
    print("\n--- Преобразование координат ---")
    uv_points = []
    
    for i, area in enumerate(free_areas[:5]):  # Ограничиваем первыми 5 областями
        pixel_x, pixel_y, width, height = area
        uv_point = pixel_to_uv(doc, active_view, pixel_x, pixel_y, image_width, image_height)
        uv_points.append(uv_point)
        
        print("Область {0}: пиксели ({1}, {2}) -> UV({3:.2f}, {4:.2f})".format(
            i+1, pixel_x, pixel_y, uv_point.U, uv_point.V))

    # Показываем результаты
    result_msg = "Найдено свободных областей: {0}\n\n".format(len(free_areas))
    result_msg += "Лучшие позиции для размещения марок:\n"
    
    for i, uv_point in enumerate(uv_points):
        result_msg += "{0}. UV({1:.2f}, {2:.2f})\n".format(i+1, uv_point.U, uv_point.V)
    
    result_msg += "\nВизуализация: analysis_result.png"

    print("\n--- Результаты ---")
    print(result_msg.replace("\n", " | "))
    TaskDialog.Show("Анализ завершен", result_msg)

    print("\nСкрипт успешно завершен")
    print("=" * 60)


if __name__ == "__main__":
    main()