#! python3
# -*- coding: utf-8 -*-
__title__ = "УМНЫЙ ОТБОР ЛУЧШИХ МЕСТ"
__author__ = "Rage"
__doc__ = "Фильтрует найденные позиции и оставляет только оптимальные для размещения марок"
__version__ = "15.0"

import os
import sys
import traceback

print("=" * 80)
print("🚀 УМНЫЙ ОТБОР ЛУЧШИХ МЕСТ")
print("=" * 80)

# =============================================================================
# НАСТРОЙКА ПУТЕЙ И БИБЛИОТЕК
# =============================================================================

# Удаляем конфликтующие пути
paths_to_remove = []
for path in sys.path:
    if 'Python311' in path or 'Python.3.11' in path:
        paths_to_remove.append(path)

for path in paths_to_remove:
    if path in sys.path:
        sys.path.remove(path)

# Добавляем пути Python 3.12
python_312_paths = [
    r'C:\Users\user34\AppData\Local\Programs\Python\Python312\Lib\site-packages',
]

for path in python_312_paths:
    if os.path.exists(path) and path not in sys.path:
        sys.path.insert(0, path)

# Загрузка библиотек
try:
    import numpy as np
    import cv2
    print(f"✅ Библиотеки загружены: NumPy v{np.__version__}, OpenCV v{cv2.__version__}")
except Exception as e:
    print(f"❌ Ошибка загрузки библиотек: {e}")
    sys.exit(1)

# =============================================================================
# ИМПОРТ REVIT API
# =============================================================================

import clr
clr.AddReference('System')
from System import Enum

from Autodesk.Revit.DB import (
    UV, XYZ, ExportRange, ImageExportOptions, 
    ImageFileType, ImageResolution, View, ViewSet
)
from Autodesk.Revit.UI import TaskDialog

print("\n" + "=" * 80)
print("🎯 СИСТЕМА ГОТОВА К УМНОМУ ОТБОРУ!")
print("=" * 80)

# =============================================================================
# УМНЫЕ ФУНКЦИИ ФИЛЬТРАЦИИ
# =============================================================================

def export_view_to_png(doc, view, export_path):
    """Экспорт вида в PNG"""
    print(f"\n📤 ЭКСПОРТ ВИДА: {view.Name}")
    
    try:
        options = ImageExportOptions()
        options.ExportRange = ExportRange.CurrentView
        options.FilePath = export_path
        options.HLRandWFViewsFileType = ImageFileType.PNG
        options.ImageResolution = ImageResolution.DPI_300
        options.ZoomType = Enum.Parse(options.ZoomType.GetType(), "FitToPage")
        options.PixelSize = 2048

        view_set = ViewSet()
        view_set.Insert(view)

        doc.ExportImage(options)
        
        if os.path.exists(export_path + ".png"):
            file_size = os.path.getsize(export_path + ".png") / (1024 * 1024)
            print(f"✅ Экспорт успешен! Размер: {file_size:.1f} MB")
            return True
        else:
            print("❌ Файл не создан")
            return False
            
    except Exception as e:
        print(f"❌ Ошибка экспорта: {e}")
        return False

def smart_position_filter(image_path, output_viz_path):
    """
    УМНЫЙ ФИЛЬТР - находит все позиции, но оставляет только лучшие
    """
    print("🔍 УМНЫЙ ОТБОР ПОЗИЦИЙ...")
    
    try:
        # Загрузка изображения
        image = cv2.imread(image_path)
        if image is None:
            print("❌ Не удалось загрузить изображение")
            return None
            
        height, width = image.shape[:2]
        print(f"📐 Изображение: {width} x {height}")
        
        # Создаем копию для визуализации
        visualization = image.copy()
        
        # Шаг 1: Находим ВСЕ позиции (как в предыдущем скрипте)
        print("🎯 ШАГ 1: Поиск всех возможных позиций...")
        all_positions = find_all_positions(image)
        print(f"   Найдено всего: {len(all_positions)} позиций")
        
        # Шаг 2: Умная фильтрация
        print("🎯 ШАГ 2: Умная фильтрация...")
        filtered_positions = smart_filtering(image, all_positions)
        print(f"   После фильтрации: {len(filtered_positions)} позиций")
        
        # Шаг 3: Группировка по зонам
        print("🎯 ШАГ 3: Группировка по зонам...")
        final_positions = group_positions(filtered_positions)
        print(f"   Финальный отбор: {len(final_positions)} позиций")
        
        # Создаем визуализацию
        create_smart_visualization(visualization, final_positions, output_viz_path)
        
        # Конвертируем результат
        result = []
        for pos in final_positions:
            x, y = pos['pixels']
            size = 100  # Оптимальный размер для марки
            result.append((x, y, size, size))
            
        return result
        
    except Exception as e:
        print(f"❌ Ошибка анализа: {e}")
        traceback.print_exc()
        return None

def find_all_positions(image):
    """Находит все возможные позиции (как в предыдущем скрипте)"""
    height, width = image.shape[:2]
    positions = []
    
    # УМНАЯ СЕТКА - не слишком плотная
    cell_size = 80
    for y in range(cell_size//2, height, cell_size):
        for x in range(cell_size//2, width, cell_size):
            positions.append({
                'pixels': (x, y),
                'radius': cell_size // 2,
                'score': 0,  # Будем вычислять позже
                'method': 'smart_grid'
            })
    
    return positions

def smart_filtering(image, positions):
    """Умная фильтрация позиций"""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    height, width = gray.shape
    
    filtered_positions = []
    
    for pos in positions:
        x, y = pos['pixels']
        
        # Проверяем область вокруг позиции
        check_radius = 60
        y1 = max(0, y - check_radius)
        y2 = min(height, y + check_radius)
        x1 = max(0, x - check_radius)
        x2 = min(width, x + check_radius)
        
        region = gray[y1:y2, x1:x2]
        
        if region.size > 0:
            # КРИТЕРИЙ 1: Средняя яркость (должна быть высокой)
            avg_brightness = np.mean(region)
            
            # КРИТЕРИЙ 2: Контрастность (не должна быть слишком низкой - не чистый фон)
            contrast = np.std(region)
            
            # КРИТЕРИЙ 3: Размер свободной зоны
            # Используем Distance Transform для оценки размера свободного пространства
            _, binary = cv2.threshold(region, 127, 255, cv2.THRESH_BINARY_INV)
            dist_transform = cv2.distanceTransform(255 - binary, cv2.DIST_L2, 3)
            free_space_radius = np.max(dist_transform)
            
            # ВЫЧИСЛЯЕМ ОБЩИЙ СКОРИНГ
            brightness_score = min(avg_brightness / 255.0, 1.0) * 40
            contrast_score = min(contrast / 50.0, 1.0) * 30  # Не любим чистый фон
            space_score = min(free_space_radius / 30.0, 1.0) * 30
            
            total_score = brightness_score + contrast_score + space_score
            
            # Оставляем только позиции с хорошим скорингом
            if total_score > 50:  # Хорошие позиции
                pos['score'] = total_score
                pos['brightness'] = avg_brightness
                pos['contrast'] = contrast
                pos['free_space'] = free_space_radius
                filtered_positions.append(pos)
    
    # Сортируем по качеству
    filtered_positions.sort(key=lambda x: x['score'], reverse=True)
    
    return filtered_positions

def group_positions(positions):
    """Группирует позиции и оставляет лучшую в каждой группе"""
    grouped_positions = []
    group_radius = 100  # Минимальное расстояние между позициями
    
    for pos in positions:
        is_in_group = False
        x1, y1 = pos['pixels']
        
        for existing in grouped_positions:
            x2, y2 = existing['pixels']
            distance = np.sqrt((x1 - x2)**2 + (y1 - y2)**2)
            
            if distance < group_radius:
                is_in_group = True
                # Оставляем позицию с лучшим скорингом
                if pos['score'] > existing['score']:
                    grouped_positions.remove(existing)
                    grouped_positions.append(pos)
                break
        
        if not is_in_group:
            grouped_positions.append(pos)
    
    # Ограничиваем общее количество позиций
    max_positions = 50
    if len(grouped_positions) > max_positions:
        grouped_positions = grouped_positions[:max_positions]
    
    return grouped_positions

def create_smart_visualization(image, positions, output_path):
    """Создает визуализацию с умным отбором"""
    try:
        print("🎨 СОЗДАЮ ВИЗУАЛИЗАЦИЮ УМНОГО ОТБОРА...")
        
        # Цвета в зависимости от качества
        for i, pos in enumerate(positions):
            x, y = pos['pixels']
            score = pos['score']
            
            # Цвет в зависимости от качества
            if score > 80:
                color = (0, 255, 0)    # Зеленый - отличные
            elif score > 60:
                color = (0, 255, 255)  # Желтый - хорошие
            else:
                color = (0, 0, 255)    # Красный - приемлемые
            
            # Размер круга в зависимости от свободного пространства
            radius = min(int(pos.get('free_space', 30) * 1.5), 80)
            
            # Полупрозрачная зона
            overlay = image.copy()
            cv2.circle(overlay, (x, y), radius, color, -1)
            cv2.addWeighted(overlay, 0.2, image, 0.8, 0, image)
            
            # Контур
            cv2.circle(image, (x, y), radius, color, 2)
            
            # Центр
            cv2.circle(image, (x, y), 6, color, -1)
            
            # Номер и скоринг
            text = f"{i+1}({int(score)})"
            text_size = cv2.getTextSize(text, cv2.FONT_HERSHEY_SIMPLEX, 0.4, 1)[0]
            
            # Фон для текста
            cv2.rectangle(image, 
                         (x - text_size[0]//2 - 2, y - radius - text_size[1] - 5),
                         (x + text_size[0]//2 + 2, y - radius + 2),
                         (0, 0, 0), -1)
            
            # Текст
            cv2.putText(image, text, 
                       (x - text_size[0]//2, y - radius - 2),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 255, 255), 1)
        
        # Информация на английском
        info_lines = [
            "SMART POSITION SELECTION",
            f"Selected positions: {len(positions)}",
            "GREEN: Excellent (score > 80)",
            "YELLOW: Good (score > 60)", 
            "RED: Acceptable (score > 50)",
            "Format: Number(Score)"
        ]
        
        for i, text in enumerate(info_lines):
            y_pos = 30 + i * 25
            # Фон
            cv2.rectangle(image, (5, y_pos - 20), (600, y_pos + 5), (0, 0, 0), -1)
            # Текст
            cv2.putText(image, text, (10, y_pos),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1)
        
        # Сохраняем
        cv2.imwrite(output_path, image, [cv2.IMWRITE_PNG_COMPRESSION, 0])
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            print(f"✅ Визуализация сохранена: {output_path}")
            print(f"📊 Размер файла: {file_size:.1f} MB")
            print("👀 Open file to see SMART selected positions!")
        else:
            print("❌ Visualization file not created")
        
    except Exception as e:
        print(f"❌ Visualization error: {e}")

def main():
    """Основная функция"""
    print("\n🎯 STARTING SMART SELECTION...")
    
    try:
        uidoc = __revit__.ActiveUIDocument
        doc = uidoc.Document
        active_view = doc.ActiveView
        
        print(f"📊 Active view: {active_view.Name}")
        
    except Exception as e:
        print(f"❌ Revit access error: {e}")
        return
    
    # Создаем пути
    script_dir = os.path.dirname(os.path.abspath(__file__))
    export_base = os.path.join(script_dir, "smart_analysis")
    image_path = export_base + ".png"
    viz_path = os.path.join(script_dir, "SMART_SELECTION_RESULTS.png")
    
    print(f"📁 Script folder: {script_dir}")
    
    # Экспорт
    if not export_view_to_png(doc, active_view, export_base):
        return
    
    # Умный отбор
    free_areas = smart_position_filter(image_path, viz_path)
    
    if not free_areas:
        print("❌ No good positions found")
        TaskDialog.Show("Info", "No optimal positions found after filtering")
        return
    
    # Преобразование координат
    print(f"\n🔄 COORDINATE TRANSFORMATION...")
    
    uv_points = []
    image_width, image_height = 2048, 1255
    
    for i, area in enumerate(free_areas[:20]):  # Покажем топ-20
        x, y, w, h = area
        
        uv_point = UV(
            -100 + (x / image_width) * 200,
            100 - (y / image_height) * 200
        )
        uv_points.append(uv_point)
        
        print(f"📍 {i+1:2d}. UV({uv_point.U:7.2f}, {uv_point.V:7.2f})")
    
    # Финальный отчет
    result_msg = f"🎉 SMART SELECTION COMPLETED!\n\n"
    result_msg += f"📊 Selected positions: {len(free_areas)}\n"
    result_msg += f"🎯 Smart filtering applied\n"
    result_msg += f"📁 Visualization: SMART_SELECTION_RESULTS.png\n\n"
    
    result_msg += "🏆 BEST POSITIONS (with scores):\n"
    for i, uv in enumerate(uv_points[:15]):
        result_msg += f"{i+1:2d}. UV({uv.U:7.2f}, {uv.V:7.2f})\n"
    
    result_msg += f"\n💡 Green circles = best positions"
    result_msg += f"\n💡 Numbers show quality scores"

    print("\n" + "=" * 80)
    print("✅ SMART SELECTION COMPLETED SUCCESSFULLY!")
    print("=" * 80)
    
    TaskDialog.Show("🎉 SMART SELECTION COMPLETED", result_msg)

if __name__ == "__main__":
    main()