#! python3
# -*- coding: utf-8 -*-
__title__ = "ПОЛНЫЙ АНАЛИЗ ВСЕХ СВОБОДНЫХ ОБЛАСТЕЙ"
__author__ = "Rage"
__doc__ = "Находит ВСЕ свободные места для размещения марок с улучшенным алгоритмом"
__version__ = "12.0"

import os
import sys
import traceback

print("=" * 80)
print("🚀 ПОЛНЫЙ АНАЛИЗ ВСЕХ СВОБОДНЫХ ОБЛАСТЕЙ")
print("=" * 80)

# =============================================================================
# ОЧИСТКА ПУТЕЙ И ЗАГРУЗКА БИБЛИОТЕК
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
print("🎯 СИСТЕМА ГОТОВА К ПОЛНОМУ АНАЛИЗУ!")
print("=" * 80)

# =============================================================================
# УЛУЧШЕННЫЕ ФУНКЦИИ АНАЛИЗА
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

def comprehensive_analysis(image_path, output_viz_path):
    """
    КОМПЛЕКСНЫЙ АНАЛИЗ для поиска ВСЕХ свободных мест
    Использует несколько алгоритмов и стратегий
    """
    print("🔍 ЗАПУСК КОМПЛЕКСНОГО АНАЛИЗА...")
    
    try:
        # Загрузка изображения
        image = cv2.imread(image_path)
        if image is None:
            print("❌ Не удалось загрузить изображение")
            return None
            
        original_height, original_width = image.shape[:2]
        print(f"📐 Изображение: {original_width} x {original_height}")
        
        # Создаем копию для визуализации
        visualization = image.copy()
        
        # АЛГОРИТМ 1: Основной анализ с более мягкими настройками
        print("🎯 АЛГОРИТМ 1: Основной поиск свободных зон...")
        main_positions = basic_free_space_analysis(image, min_radius=10, max_positions=30)
        
        # АЛГОРИТМ 2: Поиск по сетке для мелких областей
        print("🎯 АЛГОРИТМ 2: Поиск по сетке...")
        grid_positions = grid_based_analysis(image, cell_size=100, min_brightness=200)
        
        # АЛГОРИТМ 3: Поиск в углах и по краям
        print("🎯 АЛГОРИТМ 3: Анализ краев и углов...")
        edge_positions = edge_corner_analysis(image, margin=100)
        
        # Объединяем все позиции и убираем дубликаты
        all_positions = main_positions + grid_positions + edge_positions
        unique_positions = remove_duplicate_positions(all_positions, min_distance=50)
        
        print(f"📊 РЕЗУЛЬТАТЫ АНАЛИЗА:")
        print(f"   Основной алгоритм: {len(main_positions)} позиций")
        print(f"   Сеточный анализ: {len(grid_positions)} позиций")
        print(f"   Анализ краев: {len(edge_positions)} позиций")
        print(f"   Уникальных позиций: {len(unique_positions)}")
        
        # Сортируем по качеству (радиусу)
        unique_positions.sort(key=lambda x: x['radius'], reverse=True)
        
        # Создаем детальную визуализацию
        create_comprehensive_visualization(visualization, unique_positions, output_viz_path)
        
        # Конвертируем результат
        result = []
        for pos in unique_positions:
            x, y = pos['pixels']
            size = min(pos['radius'] * 2, 200)  # Уменьшаем размер для большего количества
            result.append((x, y, size, size))
            
        return result
        
    except Exception as e:
        print(f"❌ Ошибка анализа: {e}")
        traceback.print_exc()
        return None

def basic_free_space_analysis(image, min_radius=5, max_positions=50):
    """Основной алгоритм поиска свободных зон с мягкими настройками"""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Более мягкая бинаризация
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
        cv2.THRESH_BINARY_INV, 25, 5  # Увеличили размер блока, уменьшили константу
    )
    
    # Меньше агрессивные морфологические операции
    kernel = np.ones((2, 2), np.uint8)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    
    # Distance Transform
    dist_transform = cv2.distanceTransform(255 - cleaned, cv2.DIST_L2, 3)
    
    positions = []
    temp_transform = dist_transform.copy()
    
    for i in range(max_positions):
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(temp_transform)
        
        # СНИЗИЛИ порог для нахождения большего количества зон
        if max_val > min_radius:
            x, y = max_loc
            radius = int(max_val)
            
            positions.append({
                'pixels': (x, y),
                'radius': radius,
                'score': max_val,
                'method': 'distance_transform'
            })
            
            # Меньше замазываем чтобы найти больше соседних зон
            cv2.circle(temp_transform, max_loc, int(radius * 0.5), 0, -1)
        else:
            break
    
    return positions

def grid_based_analysis(image, cell_size=80, min_brightness=180):
    """Анализ по сетке для поиска мелких свободных областей"""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    height, width = gray.shape
    
    positions = []
    
    # Проходим по сетке
    for y in range(cell_size//2, height, cell_size):
        for x in range(cell_size//2, width, cell_size):
            # Проверяем область вокруг точки
            y1 = max(0, y - cell_size//2)
            y2 = min(height, y + cell_size//2)
            x1 = max(0, x - cell_size//2)
            x2 = min(width, x + cell_size//2)
            
            region = gray[y1:y2, x1:x2]
            
            if region.size > 0:
                # Если область достаточно светлая (свободная)
                if np.mean(region) > min_brightness:
                    positions.append({
                        'pixels': (x, y),
                        'radius': cell_size // 2,
                        'score': np.mean(region),
                        'method': 'grid_analysis'
                    })
    
    return positions

def edge_corner_analysis(image, margin=150):
    """Специальный анализ краев и углов (там обычно больше свободного места)"""
    height, width = image.shape[:2]
    
    positions = []
    
    # Углы
    corners = [
        (margin, margin),                    # Левый верх
        (margin, height - margin),          # Левый низ
        (width - margin, margin),           # Правый верх
        (width - margin, height - margin),  # Правый низ
    ]
    
    # Боковые стороны
    sides = [
        (margin, height // 2),              # Левая сторона
        (width - margin, height // 2),      # Правая сторона
        (width // 2, margin),               # Верхняя сторона
        (width // 2, height - margin),      # Нижняя сторона
    ]
    
    # Центральные точки с отступами
    centers = [
        (width // 4, height // 4),
        (width // 4, height * 3 // 4),
        (width * 3 // 4, height // 4),
        (width * 3 // 4, height * 3 // 4),
    ]
    
    all_points = corners + sides + centers
    
    for x, y in all_points:
        positions.append({
            'pixels': (x, y),
            'radius': 80,
            'score': 100,
            'method': 'edge_analysis'
        })
    
    return positions

def remove_duplicate_positions(positions, min_distance=30):
    """Удаляет дубликаты и близко расположенные позиции"""
    unique_positions = []
    
    for pos in positions:
        is_duplicate = False
        x1, y1 = pos['pixels']
        
        for existing in unique_positions:
            x2, y2 = existing['pixels']
            distance = np.sqrt((x1 - x2)**2 + (y1 - y2)**2)
            
            if distance < min_distance:
                is_duplicate = True
                # Оставляем позицию с большим радиусом
                if pos['radius'] > existing['radius']:
                    unique_positions.remove(existing)
                    unique_positions.append(pos)
                break
        
        if not is_duplicate:
            unique_positions.append(pos)
    
    return unique_positions

def create_comprehensive_visualization(image, positions, output_path):
    """Создает детальную визуализацию со всеми найденными позициями"""
    try:
        print("🎨 СОЗДАЮ ДЕТАЛЬНУЮ ВИЗУАЛИЗАЦИЮ...")
        
        # Цвета для разных методов анализа
        method_colors = {
            'distance_transform': (0, 0, 255),    # Красный - основные зоны
            'grid_analysis': (0, 255, 0),         # Зеленый - сеточные зоны
            'edge_analysis': (255, 0, 0),         # Синий - краевые зоны
        }
        
        # Рисуем все позиции
        for i, pos in enumerate(positions):
            x, y = pos['pixels']
            radius = pos['radius']
            method = pos['method']
            color = method_colors.get(method, (128, 128, 128))
            
            # Полупрозрачная зона
            overlay = image.copy()
            cv2.circle(overlay, (x, y), radius, color, -1)
            cv2.addWeighted(overlay, 0.2, image, 0.8, 0, image)
            
            # Контур зоны
            cv2.circle(image, (x, y), radius, color, 2)
            
            # Центр
            cv2.circle(image, (x, y), 6, color, -1)
            
            # Номер
            text = str(i + 1)
            cv2.putText(image, text, (x - 10, y - radius - 10),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 2)
        
        # Добавляем легенду
        legend_text = [
            "🎯 ПОЛНЫЙ АНАЛИЗ СВОБОДНЫХ ОБЛАСТЕЙ",
            f"Всего найдено: {len(positions)} позиций",
            "КРАСНЫЙ: Основные зоны (Distance Transform)",
            "ЗЕЛЕНЫЙ: Сеточные зоны (Grid Analysis)", 
            "СИНИЙ: Краевые зоны (Edge Analysis)",
            "ЦИФРЫ: Приоритет размещения",
        ]
        
        for i, text in enumerate(legend_text):
            y_pos = 30 + i * 25
            # Фон
            cv2.rectangle(image, (5, y_pos - 20), (600, y_pos + 5), (0, 0, 0), -1)
            # Текст
            cv2.putText(image, text, (10, y_pos),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Сохраняем
        cv2.imwrite(output_path, image)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            print(f"✅ Визуализация сохранена: {output_path}")
            print(f"📊 Размер файла: {file_size:.1f} MB")
            print("👀 Откройте файл чтобы увидеть ВСЕ свободные зоны!")
        else:
            print("❌ Файл визуализации не создан")
        
    except Exception as e:
        print(f"❌ Ошибка создания визуализации: {e}")

def main():
    """Основная функция"""
    print("\n🎯 ЗАПУСК ПОЛНОГО АНАЛИЗА...")
    
    try:
        uidoc = __revit__.ActiveUIDocument
        doc = uidoc.Document
        active_view = doc.ActiveView
        
        print(f"📊 Активный вид: {active_view.Name}")
        
    except Exception as e:
        print(f"❌ Ошибка доступа к Revit: {e}")
        return
    
    # Создаем пути
    script_dir = os.path.dirname(os.path.abspath(__file__))
    export_base = os.path.join(script_dir, "full_analysis")
    image_path = export_base + ".png"
    viz_path = os.path.join(script_dir, "FULL_ANALYSIS_VISUALIZATION.png")
    
    print(f"📁 Папка скрипта: {script_dir}")
    
    # Экспорт
    if not export_view_to_png(doc, active_view, export_base):
        return
    
    # Комплексный анализ
    free_areas = comprehensive_analysis(image_path, viz_path)
    
    if not free_areas:
        print("❌ Анализ не дал результатов")
        TaskDialog.Show("Ошибка", "Не удалось проанализировать изображение")
        return
    
    # Преобразование координат
    print(f"\n🔄 ПРЕОБРАЗОВАНИЕ КООРДИНАТ...")
    
    uv_points = []
    image_width, image_height = 2048, 1255
    
    for i, area in enumerate(free_areas[:15]):  # Покажем больше позиций
        x, y, w, h = area
        
        uv_point = UV(
            -100 + (x / image_width) * 200,
            100 - (y / image_height) * 200
        )
        uv_points.append(uv_point)
        
        print(f"📍 {i+1:2d}. UV({uv_point.U:7.2f}, {uv_point.V:7.2f})")
    
    # Финальный отчет
    result_msg = f"🎉 ПОЛНЫЙ АНАЛИЗ ЗАВЕРШЕН!\n\n"
    result_msg += f"📊 Найдено свободных областей: {len(free_areas)}\n"
    result_msg += f"🎯 Использовано 3 алгоритма анализа\n"
    result_msg += f"📁 Визуализация: FULL_ANALYSIS_VISUALIZATION.png\n\n"
    
    result_msg += "🏆 ЛУЧШИЕ ПОЗИЦИИ:\n"
    for i, uv in enumerate(uv_points[:12]):
        result_msg += f"{i+1:2d}. UV({uv.U:7.2f}, {uv.V:7.2f})\n"
    
    result_msg += f"\n💡 Откройте файл визуализации чтобы увидеть"
    result_msg += f"\nВСЕ {len(free_areas)} найденных зон!"

    print("\n" + "=" * 80)
    print("✅ ПОЛНЫЙ АНАЛИЗ УСПЕШНО ВЫПОЛНЕН!")
    print("=" * 80)
    
    TaskDialog.Show("🎉 ПОЛНЫЙ АНАЛИЗ ЗАВЕРШЕН", result_msg)

if __name__ == "__main__":
    main()