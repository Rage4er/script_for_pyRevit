# -*- coding: utf-8 -*-

__title__ = """Размещение
видов на листы"""
__author__ = 'Rage'
__doc__ = """Размещает 3D виды на листы по системам с оптимальным заполнением"""
__version__ = "6.0"

import clr
import re
import math
import os
import time
from collections import defaultdict, OrderedDict

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')
import System

from System.Windows.Forms import (
    Form, FormStartPosition, FormBorderStyle,
    DialogResult, MessageBox, Panel, BorderStyle,
    CheckBox, Button, Label, TextBox, GroupBox
)
from System.Drawing import (
    Font, FontStyle,
    Color, Point, Size,
    Bitmap, Imaging, Graphics, Pen, SolidBrush, Rectangle
)
from System.Drawing.Imaging import PixelFormat
from System.IO import Path, File, Directory

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ViewSheet,
    View3D,
    Viewport,
    XYZ,
    Transaction,
    FamilySymbol,
    ElementId,
    StorageType,
    ImageExportOptions,
    ImageFileType,
    ZoomFitType,
    ExportRange,
    ImageResolution,
    ViewSet,
    Line,
    DetailLine,
    SubTransaction,
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

# ============ НАСТРОЙКИ ============

GROUPING = OrderedDict([
    ("П-В",     ["П", "ПЕ", "В", "ВЕ"]),
    ("ДП-ДВ",   ["ДП", "ДПЕ", "ДВ", "ДВЕ"]),
    ("А-У",     ["А", "У"]),
])

ALL_PREFIXES = ['ПЕ', 'ВЕ', 'ДПЕ', 'ДВЕ', 'П', 'В', 'ДП', 'ДВ', 'А', 'У']

EXPORT_PIXELS = 1000
PADDING_MM = 5
STAMP_WIDTH_MM = 100  # ширина штампа (правая часть листа)
STAMP_HEIGHT_MM = 50  # высота штампа (нижняя часть листа)
CALIBRATION_SQUARE_MM = 5.0  # квадраты 5×5 мм вместо 100×100 мм (почти невидимые)
SQUARE_LINE_STEP_MM = 0.5
SQUARE_LINE_COUNT = 1
MM_PER_PX = 1.5  # 1500mm / 1000px = 1.5 мм/пиксель

# Пределы для проверки качества калибровки
MIN_SQUARE_SIZE_PX = 3     # 5мм × 0.67 px/мм ≈ 3px (фиксированный размер при 1000px)
MAX_SQUARE_SIZE_PX = 1000  # максимальный ожидаемый размер квадрата в пикселях
MIN_MM_PER_PX = 0.3        # минимальный коэффициент мм/пиксель (соответствует ~3.33 px/мм)
MAX_MM_PER_PX = 2.0        # максимальный коэффициент мм/пиксель (соответствует 0.5 px/мм)

# Кэш размеров видов для ускорения повторных измерений
VIEW_SIZE_CACHE = {}

# ============ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ============

def sanitize_sheet_name(name):
    if not name:
        return "Unnamed"
    prohibited = r'[{}[\]|;:<>?/\\*"\']'
    sanitized = re.sub(prohibited, '-', name)
    sanitized = re.sub(r'-+', '-', sanitized.strip(' -'))
    if len(sanitized) > 200:
        sanitized = sanitized[:197] + "..."
    return sanitized if sanitized else "Unnamed"


def get_view_short_name(view_name):
    prefixes_to_remove = ['Схема_Возд_', 'Схема_', 'Возд_', '3D_', '3D - ']
    short_name = view_name
    for prefix in prefixes_to_remove:
        if short_name.startswith(prefix):
            short_name = short_name[len(prefix):]
            break
    return short_name


def extract_prefix_and_numbers(view_name):
    if not view_name:
        return "", (0,)
    name_upper = view_name.upper()
    for prefix in sorted(ALL_PREFIXES, key=len, reverse=True):
        idx = name_upper.find(prefix)
        if idx >= 0:
            after = name_upper[idx + len(prefix):]
            numbers = re.findall(r'\d+', after)
            if numbers:
                return prefix, tuple(int(n) for n in numbers)
            if idx + len(prefix) == len(name_upper):
                return prefix, (0,)
    numbers = re.findall(r'\d+', view_name)
    if numbers:
        for prefix in sorted(ALL_PREFIXES, key=len, reverse=True):
            if prefix in name_upper:
                return prefix, tuple(int(n) for n in numbers)
    return "", tuple(int(n) for n in numbers) if numbers else (0,)


def get_group_for_prefix(prefix):
    for group_name, prefixes in GROUPING.items():
        if prefix in prefixes:
            return group_name
    return "Без системы"


def get_view_sort_key(view):
    prefix, numbers = extract_prefix_and_numbers(view.Name)
    all_prefixes = [p for plist in GROUPING.values() for p in plist]
    prefix_order = {p: i for i, p in enumerate(all_prefixes)}
    return (prefix_order.get(prefix, 999), numbers, view.Name)


def collect_placed_views():
    placed = set()
    try:
        sheets = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_Sheets)\
            .WhereElementIsNotElementType()\
            .ToElements()
        for sheet in sheets:
            vps = FilteredElementCollector(doc, sheet.Id)\
                .OfCategory(BuiltInCategory.OST_Viewports)\
                .WhereElementIsNotElementType()\
                .ToElements()
            for vp in vps:
                placed.add(vp.ViewId)
    except Exception as e:
        print("Ошибка при сборе размещённых видов: " + str(e))
    return placed


def get_sheet_frame_mm(sheet):
    tbs = list(FilteredElementCollector(doc, sheet.Id)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType()
               .ToElements())
    if not tbs:
        return None
    
    tb = tbs[0]
    bbox = tb.get_BoundingBox(sheet)
    if not bbox:
        return None
    
    # Стандартные размеры форматов ГОСТ (мм)
    GOST_SIZES = {
        "А4": (297, 210),
        "А3": (420, 297),
        "А2": (594, 420),
        "А1": (841, 594),
        "А0": (1189, 841),
    }
    
    # Параметры, которые могут содержать размеры
    param_height_real = tb.LookupParameter("Высота_Реальная")
    param_width_real = tb.LookupParameter("Ширина_Реальная")
    param_height = tb.LookupParameter("Высота")
    param_width = tb.LookupParameter("Ширина")
    param_format = tb.LookupParameter("Формат ГОСТ")
    
    # Размеры из bounding box (в миллиметрах)
    width_mm_from_bbox = (bbox.Max.X - bbox.Min.X) * 304.8
    height_mm_from_bbox = (bbox.Max.Y - bbox.Min.Y) * 304.8
    
    MIN_FRAME_SIZE_MM = 50.0
    MAX_FRAME_SIZE_MM = 5000.0
    
    def is_reasonable(h, w):
        """Проверяет, что размеры находятся в разумном диапазоне."""
        return (MIN_FRAME_SIZE_MM <= h <= MAX_FRAME_SIZE_MM and
                MIN_FRAME_SIZE_MM <= w <= MAX_FRAME_SIZE_MM)
    
    def try_use_params(h_param, w_param, param_names):
        if h_param and w_param:
            h_raw = h_param.AsDouble()
            w_raw = w_param.AsDouble()
            # Вывод сырых значений для отладки
            print("    🔍 Параметры рамки '{}': сырые значения ширина={}, высота={} (внутренние единицы)".format(param_names, w_raw, h_raw))
            # Внутренние единицы Revit — футы, преобразуем в миллиметры
            h_mm = h_raw * 304.8
            w_mm = w_raw * 304.8
            print("    🔍 Преобразовано в миллиметры: ширина={} мм, высота={} мм".format(int(w_mm), int(h_mm)))
            if is_reasonable(h_mm, w_mm):
                print("    📏 Используются параметры рамки ({}): Ширина={} мм, Высота={} мм".format(param_names, int(w_mm), int(h_mm)))
                return (w_mm, h_mm)
            else:
                print("    ⚠️ Параметры рамки найдены, но значения некорректны (ширина={} мм, высота={} мм).".format(int(w_mm), int(h_mm)))
        return None
    
    # Сначала пробуем определить формат по параметру "Формат ГОСТ"
    if param_format and param_format.HasValue:
        format_str = param_format.AsString()
        if format_str in GOST_SIZES:
            std_w, std_h = GOST_SIZES[format_str]
            # Проверяем ориентацию (книжная/альбомная)
            param_orientation = tb.LookupParameter("Книжная ориентация")
            if param_orientation and param_orientation.HasValue and param_orientation.AsInteger() == 0:
                # Альбомная ориентация: ширина > высоты
                width_mm = max(std_w, std_h)
                height_mm = min(std_w, std_h)
            else:
                # Книжная ориентация по умолчанию
                width_mm = min(std_w, std_h)
                height_mm = max(std_w, std_h)
            print("    📏 Используется формат ГОСТ '{}': Ширина={} мм, Высота={} мм".format(format_str, int(width_mm), int(height_mm)))
            min_x_mm = bbox.Min.X * 304.8
            min_y_mm = bbox.Min.Y * 304.8
            return (width_mm, height_mm, min_x_mm, min_y_mm)
        else:
            print("    ⚠️ Параметр 'Формат ГОСТ' содержит неизвестное значение: '{}'".format(format_str))
    
    # Пробуем параметры "Высота_Реальная"/"Ширина_Реальная"
    result = try_use_params(param_height_real, param_width_real, "Высота_Реальная/Ширина_Реальная")
    if result is None:
        # Пробуем параметры "Высота"/"Ширина"
        result = try_use_params(param_height, param_width, "Высота/Ширина")
    
    if result is not None:
        width_mm, height_mm = result
    else:
        # Если ни один параметр не подошёл, используем bounding box
        if is_reasonable(height_mm_from_bbox, width_mm_from_bbox):
            width_mm = width_mm_from_bbox
            height_mm = height_mm_from_bbox
            print("    📏 Используется bounding box рамки: Ширина={} мм, Высота={} мм".format(int(width_mm), int(height_mm)))
        else:
            # Bounding box тоже некорректен, используем стандартный размер А3 с предупреждением
            width_mm = 420.0
            height_mm = 297.0
            print("    ⚠️ Некорректные размеры рамки (bounding box: {}x{} мм). Используется стандартный формат А3 (420x297 мм).".format(
                int(width_mm_from_bbox), int(height_mm_from_bbox)))
    
    min_x_mm = bbox.Min.X * 304.8
    min_y_mm = bbox.Min.Y * 304.8
    
    return (width_mm, height_mm, min_x_mm, min_y_mm)


def get_title_block_type_id(sheet=None):
    if sheet:
        tbs = list(FilteredElementCollector(doc, sheet.Id)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsNotElementType()
                   .ToElements())
        if tbs:
            return tbs[0].GetTypeId()
    
    all_tb = list(FilteredElementCollector(doc)
                  .OfClass(FamilySymbol)
                  .OfCategory(BuiltInCategory.OST_TitleBlocks)
                  .WhereElementIsNotElementType()
                  .ToElements())
    if all_tb:
        return all_tb[0].Id
    
    raise Exception("В проекте нет загруженных титульных блоков!")


def create_sheet(template_sheet):
    tb_type_id = get_title_block_type_id(template_sheet)
    new_sheet = ViewSheet.Create(doc, tb_type_id)
    
    tbs_template = list(FilteredElementCollector(doc, template_sheet.Id)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsNotElementType()
                        .ToElements())
    tbs_new = list(FilteredElementCollector(doc, new_sheet.Id)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .WhereElementIsNotElementType()
                   .ToElements())
    
    if tbs_template and tbs_new:
        param_names = ["Формат А", "Кратность", "Книжная ориентация",
                       "Нестандартный", "Высота", "Ширина"]
        for param_name in param_names:
            p_template = tbs_template[0].LookupParameter(param_name)
            p_new = tbs_new[0].LookupParameter(param_name)
            if p_template and p_template.HasValue and p_new and not p_new.IsReadOnly:
                try:
                    if p_template.StorageType == StorageType.String:
                        p_new.Set(p_template.AsString())
                    elif p_template.StorageType == StorageType.Integer:
                        p_new.Set(p_template.AsInteger())
                    elif p_template.StorageType == StorageType.Double:
                        p_new.Set(p_template.AsDouble())
                    elif p_template.StorageType == StorageType.ElementId:
                        val = p_template.AsElementId()
                        if val and val != ElementId.InvalidElementId:
                            p_new.Set(val)
                except:
                    pass
    
    # Копирование параметра листа "ADSK_Штамп Раздел Проекта"
    param_name = "ADSK_Штамп Раздел Проекта"
    p_template = template_sheet.LookupParameter(param_name)
    p_new = new_sheet.LookupParameter(param_name)
    
    if p_template and p_template.HasValue and p_new and not p_new.IsReadOnly:
        try:
            if p_template.StorageType == StorageType.String:
                value = p_template.AsString()
                p_new.Set(value)
                print("  ✅ Скопирован параметр листа '{}': {}".format(param_name, value))
            elif p_template.StorageType == StorageType.Integer:
                p_new.Set(p_template.AsInteger())
            elif p_template.StorageType == StorageType.Double:
                p_new.Set(p_template.AsDouble())
            elif p_template.StorageType == StorageType.ElementId:
                val = p_template.AsElementId()
                if val and val != ElementId.InvalidElementId:
                    p_new.Set(val)
        except Exception as e:
            print("  ⚠️ Ошибка при копировании параметра '{}': {}".format(param_name, e))
    
    return new_sheet


def clear_sheet_viewports(sheet):
    vps = list(FilteredElementCollector(doc, sheet.Id)
               .OfCategory(BuiltInCategory.OST_Viewports)
               .WhereElementIsNotElementType()
               .ToElements())
    for vp in vps:
        doc.Delete(vp.Id)


def draw_thick_square(doc, sheet, center_ft, half_size_ft):
    """Рисует утолщённый квадрат из нескольких вложенных линий."""
    px = center_ft.X
    py = center_ft.Y
    
    step_ft = SQUARE_LINE_STEP_MM / 304.8
    
    for i in range(SQUARE_LINE_COUNT):
        offset = i * step_ft
        h = half_size_ft - offset
        if h <= 0:
            break
        
        x0, y0 = px - h, py - h
        x1, y1 = px + h, py + h
        
        doc.Create.NewDetailCurve(sheet,
            Line.CreateBound(XYZ(x0, y0, 0), XYZ(x1, y0, 0)))
        doc.Create.NewDetailCurve(sheet,
            Line.CreateBound(XYZ(x0, y1, 0), XYZ(x1, y1, 0)))
        doc.Create.NewDetailCurve(sheet,
            Line.CreateBound(XYZ(x0, y0, 0), XYZ(x0, y1, 0)))
        doc.Create.NewDetailCurve(sheet,
            Line.CreateBound(XYZ(x1, y0, 0), XYZ(x1, y1, 0)))


# ============ ПОИСК КВАДРАТОВ (EDGE DETECTION) ============

def find_squares_by_edges(bitmap):
    """Возвращает фиктивные калибровочные квадраты на основе известной геометрии."""
    w = bitmap.Width
    h = bitmap.Height
    
    print("    Используем фиксированные калибровочные квадраты (" + str(w) + "x" + str(h) + " px)")
    
    # Предполагаем, что изображение соответствует калибровочной области 1500×1500 мм
    # и экспортировано с разрешением 1000×1000 пикселей (EXPORT_PIXELS = 1000).
    # Если размер изображения отличается, масштабируем координаты пропорционально.
    ref_size = 1000  # эталонная ширина/высота (новое разрешение)
    scale_x = w / ref_size
    scale_y = h / ref_size
    
    # Размер квадрата в пикселях при эталонном размере: 5 мм * (1000 px / 1500 мм) = 3.33 ≈ 3 px
    square_size_ref = 3
    # Расстояние от центра до центра квадрата в эталонных пикселях: 700 мм * (1000/1500) = 466.67 ≈ 467 px
    offset_ref = 467
    
    # Центр изображения
    center_x = w / 2
    center_y = h / 2
    
    # Масштабированные смещение и размер
    offset_x = offset_ref * scale_x
    offset_y = offset_ref * scale_y
    square_size = square_size_ref * ((scale_x + scale_y) / 2)  # средний масштаб
    
    # Координаты центров четырёх квадратов
    squares = []
    for dx, dy in [(-1, -1), (1, -1), (-1, 1), (1, 1)]:
        cx = center_x + dx * offset_x
        cy = center_y + dy * offset_y
        squares.append({
            'x': int(cx),
            'y': int(cy),
            'size': int(square_size),
            'score': 1.0
        })
    
    print("    Создано фиктивных квадратов: " + str(len(squares)))
    return squares


def enhance_bitmap(bitmap):
    """
    Усиливает контраст и утолщает линии.
    Max pooling 5×5 + растягивание гистограммы.
    """
    w = bitmap.Width
    h = bitmap.Height
    
    # Max pooling 5×5
    enhanced = [[0] * w for _ in range(h)]
    
    for y in range(h):
        for x in range(w):
            max_val = 0
            for dy in (-2, -1, 0, 1, 2):
                ny = y + dy
                if ny < 0 or ny >= h:
                    continue
                for dx in (-2, -1, 0, 1, 2):
                    nx = x + dx
                    if nx < 0 or nx >= w:
                        continue
                    c = bitmap.GetPixel(nx, ny)
                    val = max(c.R, c.G, c.B)
                    if val > max_val:
                        max_val = val
            enhanced[y][x] = max_val
    
    # Растягиваем гистограмму (игнорируя углы калибровки ~8% от меньшей стороны)
    # Для разрешения 1000×1000 px: min(w, h) // 12 = 83, max(80, 83) = 83
    exclude_margin = max(80, min(w, h) // 12)  # примерно 8%, но не менее 80px
    # Ограничим, чтобы не превышать четверть размера
    if exclude_margin > min(w, h) // 4:
        exclude_margin = min(w, h) // 4
    min_val = 255
    max_val = 0
    for y in range(exclude_margin, h - exclude_margin):
        for x in range(exclude_margin, w - exclude_margin):
            v = enhanced[y][x]
            if v < min_val:
                min_val = v
            if v > max_val:
                max_val = v
    
    # Если диапазон слишком мал, расширяем его
    if max_val - min_val < 10:
        min_val = max(0, min_val - 5)
        max_val = min(255, max_val + 5)
    
    # Применяем растяжение ко всему изображению
    if max_val > min_val:
        scale = 255.0 / (max_val - min_val)
        for y in range(h):
            for x in range(w):
                v = enhanced[y][x]
                enhanced[y][x] = int((v - min_val) * scale)
    
    # Создаём результирующее изображение (grayscale для простоты)
    result = Bitmap(w, h, PixelFormat.Format32bppArgb)
    for y in range(h):
        for x in range(w):
            v = min(255, max(0, enhanced[y][x]))
            result.SetPixel(x, y, Color.FromArgb(255, v, v, v))
    
    return result


def detect_theme(bitmap):
    """Определяет тему по яркости краёв."""
    w = bitmap.Width
    h = bitmap.Height
    
    samples = []
    step = max(1, min(w, h) // 30)
    
    for x in range(0, w, step):
        samples.append(bitmap.GetPixel(x, 0))
        samples.append(bitmap.GetPixel(x, h - 1))
    for y in range(0, h, step):
        samples.append(bitmap.GetPixel(0, y))
        samples.append(bitmap.GetPixel(w - 1, y))
    
    if not samples:
        return 'dark'
    
    avg = sum(max(c.R, c.G, c.B) for c in samples) / len(samples)
    print("    Яркость фона: " + str(int(avg)))
    return 'dark' if avg < 100 else 'light'


def find_page_bounds(bitmap):
    """Находит границы листа A4 на изображении (исключая поля)."""
    w = bitmap.Width
    h = bitmap.Height
    
    # Определяем фон по краям изображения (предполагаем, что поля однородные)
    edge_samples = []
    step = max(1, min(w, h) // 30)
    for x in range(0, w, step):
        edge_samples.append(bitmap.GetPixel(x, 0))
        edge_samples.append(bitmap.GetPixel(x, h - 1))
    for y in range(0, h, step):
        edge_samples.append(bitmap.GetPixel(0, y))
        edge_samples.append(bitmap.GetPixel(w - 1, y))
    
    if not edge_samples:
        return 0, 0, w - 1, h - 1
    
    # Средняя яркость фона
    bg_avg = sum(max(c.R, c.G, c.B) for c in edge_samples) / len(edge_samples)
    # Порог для отделения фона от листа (эмпирически)
    threshold = bg_avg + 20 if bg_avg < 128 else bg_avg - 20
    
    # Сканируем по горизонтали для поиска левой и правой границы
    left = 0
    for x in range(0, w, 2):
        col_avg = sum(max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
                      for y in range(0, h, max(1, h // 20))) / (h // max(1, h // 20))
        if (bg_avg < 128 and col_avg > threshold) or (bg_avg >= 128 and col_avg < threshold):
            left = x
            break
    
    right = w - 1
    for x in range(w - 1, -1, -2):
        col_avg = sum(max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
                      for y in range(0, h, max(1, h // 20))) / (h // max(1, h // 20))
        if (bg_avg < 128 and col_avg > threshold) or (bg_avg >= 128 and col_avg < threshold):
            right = x
            break
    
    # Сканируем по вертикали для верхней и нижней границы
    top = 0
    for y in range(0, h, 2):
        row_avg = sum(max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
                      for x in range(0, w, max(1, w // 20))) / (w // max(1, w // 20))
        if (bg_avg < 128 and row_avg > threshold) or (bg_avg >= 128 and row_avg < threshold):
            top = y
            break
    
    bottom = h - 1
    for y in range(h - 1, -1, -2):
        row_avg = sum(max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
                      for x in range(0, w, max(1, w // 20))) / (w // max(1, w // 20))
        if (bg_avg < 128 and row_avg > threshold) or (bg_avg >= 128 and row_avg < threshold):
            bottom = y
            break
    
    # Добавляем небольшой запас (1 пиксель) чтобы гарантировать попадание внутрь листа
    left = max(0, left - 1)
    right = min(w - 1, right + 1)
    top = max(0, top - 1)
    bottom = min(h - 1, bottom + 1)
    
    print("    Границы листа: left={}, top={}, right={}, bottom={}".format(left, top, right, bottom))
    return left, top, right, bottom


def find_content_bounds(bitmap, exclusion_zones=None):
    """Находит границы содержимого, исключая зоны квадратов."""
    if exclusion_zones is None:
        exclusion_zones = []
    
    w = bitmap.Width
    h = bitmap.Height
    
    bg_pixels = []
    step = max(1, min(w, h) // 30)
    
    for x in range(0, w, step):
        excluded = any(z[0] <= x <= z[2] and z[1] <= 0 <= z[3] for z in exclusion_zones)
        if not excluded:
            c = bitmap.GetPixel(x, 0)
            bg_pixels.append(max(c.R, c.G, c.B))
        excluded = any(z[0] <= x <= z[2] and z[1] <= h-1 <= z[3] for z in exclusion_zones)
        if not excluded:
            c = bitmap.GetPixel(x, h-1)
            bg_pixels.append(max(c.R, c.G, c.B))
    
    for y in range(0, h, step):
        excluded = any(z[0] <= 0 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
        if not excluded:
            c = bitmap.GetPixel(0, y)
            bg_pixels.append(max(c.R, c.G, c.B))
        excluded = any(z[0] <= w-1 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
        if not excluded:
            c = bitmap.GetPixel(w-1, y)
            bg_pixels.append(max(c.R, c.G, c.B))
    
    if not bg_pixels:
        bg_avg = 128
    else:
        bg_avg = sum(bg_pixels) / len(bg_pixels)
    
    if bg_avg < 100:
        threshold = bg_avg + 20
        def is_content(b):
            return b > threshold
    else:
        threshold = bg_avg - 20
        def is_content(b):
            return b < threshold
    
    sample_step = max(1, min(w, h) // 80)
    
    # Отладочная информация
    print("    find_content_bounds: bg_avg={}, threshold={}, step={}, dark={}".format(
        bg_avg, threshold, sample_step, bg_avg < 100))
    
    min_x, max_x = w, 0
    min_y, max_y = h, 0
    found = False
    content_pixels = 0
    
    for y in range(0, h, sample_step):
        for x in range(0, w, sample_step):
            excluded = any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if excluded:
                continue
            c = bitmap.GetPixel(x, y)
            b = max(c.R, c.G, c.B)
            if is_content(b):
                min_x = min(min_x, x)
                max_x = max(max_x, x)
                min_y = min(min_y, y)
                max_y = max(max_y, y)
                found = True
                content_pixels += 1
    
    if not found:
        print("    find_content_bounds: не найдено ни одного пикселя содержимого")
        return None
    
    # Отладочная статистика
    print("    find_content_bounds: найдено пикселей содержимого={}, границы x=[{}, {}] y=[{}, {}]".format(
        content_pixels, min_x, max_x, min_y, max_y))
    
    return {'min_x': min_x, 'min_y': min_y, 'max_x': max_x, 'max_y': max_y}


def find_content_bounds_3pass(bitmap, exclusion_zones=None):
    """
    Трехпроходной поиск границ с адаптивным порогом для разрешения 1000×1000 px.
    Проход 1: шаг 100px по всей области (исключая exclusion_zones)
    Проход 2: шаг 25px в области, расширенной на +100px от результатов прохода 1
    Проход 3: шаг 5px в периметре (±30px от границ прохода 2)
    Возвращает словарь {'min_x', 'min_y', 'max_x', 'max_y'} или None.
    """
    if exclusion_zones is None:
        exclusion_zones = []
    
    w = bitmap.Width
    h = bitmap.Height
    
    # Вспомогательная функция для определения содержимого
    def is_content(b, bg_avg, dark_background):
        if dark_background:
            return b > bg_avg + 20
        else:
            return b < bg_avg - 20
    
    # Функция для оценки фона по краям изображения (исключая exclusion_zones)
    def estimate_background():
        bg_samples = []
        step = max(1, min(w, h) // 30)
        
        # Собираем пиксели с краев изображения
        for x in range(0, w, step):
            excluded_top = any(z[0] <= x <= z[2] and z[1] <= 0 <= z[3] for z in exclusion_zones)
            if not excluded_top:
                c = bitmap.GetPixel(x, 0)
                bg_samples.append(max(c.R, c.G, c.B))
            excluded_bottom = any(z[0] <= x <= z[2] and z[1] <= h-1 <= z[3] for z in exclusion_zones)
            if not excluded_bottom:
                c = bitmap.GetPixel(x, h-1)
                bg_samples.append(max(c.R, c.G, c.B))
        
        for y in range(0, h, step):
            excluded_left = any(z[0] <= 0 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if not excluded_left:
                c = bitmap.GetPixel(0, y)
                bg_samples.append(max(c.R, c.G, c.B))
            excluded_right = any(z[0] <= w-1 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if not excluded_right:
                c = bitmap.GetPixel(w-1, y)
                bg_samples.append(max(c.R, c.G, c.B))
        
        if not bg_samples:
            return 128, True  # по умолчанию темный фон
        
        bg_avg = sum(bg_samples) / len(bg_samples)
        dark_background = bg_avg < 100
        return bg_avg, dark_background
    
    # Оцениваем фон
    bg_avg, dark_background = estimate_background()
    
    # Игнорируем углы изображения (80px от углов для 1000px)
    corner_size = 80
    
    # Проход 1: шаг 100px по всей области
    step1 = 100
    min_x1, max_x1 = w, 0
    min_y1, max_y1 = h, 0
    found1 = False
    
    for y in range(0, h, step1):
        for x in range(0, w, step1):
            # Игнорируем exclusion_zones
            excluded = any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if excluded:
                continue
            # Игнорируем углы
            if (x < corner_size and y < corner_size) or \
               (x < corner_size and y > h - corner_size) or \
               (x > w - corner_size and y < corner_size) or \
               (x > w - corner_size and y > h - corner_size):
                continue
            
            c = bitmap.GetPixel(x, y)
            b = max(c.R, c.G, c.B)
            
            if is_content(b, bg_avg, dark_background):
                min_x1 = min(min_x1, x)
                max_x1 = max(max_x1, x)
                min_y1 = min(min_y1, y)
                max_y1 = max(max_y1, y)
                found1 = True
    
    if not found1:
        print("    find_content_bounds_3pass: проход 1 не нашел содержимого, пробуем проход 2 по всей области")
        # Проход 2 по всей области (исключая углы и exclusion_zones)
        step2 = 25
        min_x2, max_x2 = w, 0
        min_y2, max_y2 = h, 0
        found2 = False
        
        for y in range(0, h, step2):
            for x in range(0, w, step2):
                excluded = any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
                if excluded:
                    continue
                # Игнорируем углы
                if (x < corner_size and y < corner_size) or \
                   (x < corner_size and y > h - corner_size) or \
                   (x > w - corner_size and y < corner_size) or \
                   (x > w - corner_size and y > h - corner_size):
                    continue
                
                c = bitmap.GetPixel(x, y)
                b = max(c.R, c.G, c.B)
                
                if is_content(b, bg_avg, dark_background):
                    min_x2 = min(min_x2, x)
                    max_x2 = max(max_x2, x)
                    min_y2 = min(min_y2, y)
                    max_y2 = max(max_y2, y)
                    found2 = True
        
        if not found2:
            print("    find_content_bounds_3pass: проход 2 также не нашел содержимого")
            return None
        
        # Используем результаты прохода 2 как основу для прохода 3
        min_x1, max_x1 = min_x2, max_x2
        min_y1, max_y1 = min_y2, max_y2
        found1 = True
        print("    find_content_bounds_3pass: проход 2 нашел границы x=[{}, {}] y=[{}, {}]".format(
            min_x1, max_x1, min_y1, max_y1))
    
    # Расширяем область для прохода 2 (теперь это проход 2 уточнения) на +100px (шаг первого прохода)
    expand1 = step1
    search_x1 = max(0, min_x1 - expand1)
    search_x2 = min(w - 1, max_x1 + expand1)
    search_y1 = max(0, min_y1 - expand1)
    search_y2 = min(h - 1, max_y1 + expand1)
    
    # Проход 2: шаг 25px в расширенной области
    step2 = 25
    min_x2, max_x2 = w, 0
    min_y2, max_y2 = h, 0
    found2 = False
    
    for y in range(search_y1, search_y2 + 1, step2):
        for x in range(search_x1, search_x2 + 1, step2):
            excluded = any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if excluded:
                continue
            
            c = bitmap.GetPixel(x, y)
            b = max(c.R, c.G, c.B)
            
            if is_content(b, bg_avg, dark_background):
                min_x2 = min(min_x2, x)
                max_x2 = max(max_x2, x)
                min_y2 = min(min_y2, y)
                max_y2 = max(max_y2, y)
                found2 = True
    
    if not found2:
        # Используем результаты прохода 1
        min_x2, max_x2 = min_x1, max_x1
        min_y2, max_y2 = min_y1, max_y1
    else:
        # Расширяем область для прохода 3 на +25px (шаг второго прохода)
        expand2 = step2
        search_x1 = max(0, min_x2 - expand2)
        search_x2 = min(w - 1, max_x2 + expand2)
        search_y1 = max(0, min_y2 - expand2)
        search_y2 = min(h - 1, max_y2 + expand2)
    
    # Проход 3: шаг 5px в периметре (±30px от границ)
    step3 = 5
    margin = 30
    min_x3, max_x3 = w, 0
    min_y3, max_y3 = h, 0
    found3 = False
    
    # Определяем область периметра
    perim_x1 = max(0, min_x2 - margin)
    perim_x2 = min(w - 1, max_x2 + margin)
    perim_y1 = max(0, min_y2 - margin)
    perim_y2 = min(h - 1, max_y2 + margin)
    
    # Сканируем только периметр (границы области)
    for y in range(perim_y1, perim_y2 + 1, step3):
        for x in range(perim_x1, perim_x2 + 1, step3):
            # Пропускаем внутреннюю часть (только периметр)
            if x > perim_x1 + margin and x < perim_x2 - margin and \
               y > perim_y1 + margin and y < perim_y2 - margin:
                continue
            
            excluded = any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones)
            if excluded:
                continue
            
            c = bitmap.GetPixel(x, y)
            b = max(c.R, c.G, c.B)
            
            if is_content(b, bg_avg, dark_background):
                min_x3 = min(min_x3, x)
                max_x3 = max(max_x3, x)
                min_y3 = min(min_y3, y)
                max_y3 = max(max_y3, y)
                found3 = True
    
    if not found3:
        # Используем результаты прохода 2
        result = {'min_x': min_x2, 'min_y': min_y2, 'max_x': max_x2, 'max_y': max_y2}
    else:
        result = {'min_x': min_x3, 'min_y': min_y3, 'max_x': max_x3, 'max_y': max_y3}
    
    # Отладочная информация
    print("    find_content_bounds_3pass: bg_avg={}, dark={}, границы x=[{}, {}] y=[{}, {}]".format(
        int(bg_avg), dark_background, result['min_x'], result['max_x'], result['min_y'], result['max_y']))
    
    return result


# ============ ОТЛАДОЧНАЯ ОТРИСОВКА ============

def draw_debug_overlay(bitmap, squares, content_bounds, exclusion_zones, output_path, page_bounds=None, mm_per_px=None):
    """Рисует на изображении найденные границы."""
    debug_bmp = Bitmap(bitmap.Width, bitmap.Height, PixelFormat.Format32bppArgb)
    g = Graphics.FromImage(debug_bmp)
    g.DrawImage(bitmap, 0, 0)
    
    # Синий прямоугольник границ листа (если задан)
    if page_bounds:
        left, top, right, bottom = page_bounds
        blue_pen = Pen(Color.Blue, 1)
        g.DrawRectangle(blue_pen, left, top, right - left, bottom - top)
        # Подпись
        from System.Drawing import Font as GFont
        font = GFont("Arial", 8, FontStyle.Regular)
        text_brush = SolidBrush(Color.Blue)
        g.DrawString("PAGE", font, text_brush, float(left + 2), float(top + 2))
    
    # Жёлтые зоны исключения (полупрозрачные)
    yellow_brush = SolidBrush(Color.FromArgb(50, 255, 255, 0))
    yellow_pen = Pen(Color.FromArgb(100, 255, 200, 0), 1)
    for ex, ey, ex2, ey2 in exclusion_zones:
        g.FillRectangle(yellow_brush, ex, ey, ex2 - ex + 1, ey2 - ey + 1)
        g.DrawRectangle(yellow_pen, ex, ey, ex2 - ex + 1, ey2 - ey + 1)
    
    # Зелёные квадраты
    green_pen = Pen(Color.LimeGreen, 2)
    for idx, sq in enumerate(squares):
        half = sq['size'] // 2
        x = sq['x'] - half
        y = sq['y'] - half
        g.DrawRectangle(green_pen, x, y, sq['size'], sq['size'])
        # Подпись квадрата
        font = Font("Arial", 7, FontStyle.Regular)
        text_brush = SolidBrush(Color.LimeGreen)
        label = "{},{} size={}".format(sq['x'], sq['y'], sq['size'])
        if mm_per_px is not None:
            size_mm = sq['size'] * mm_per_px
            label += " ({:.1f} mm)".format(size_mm)
        g.DrawString(label, font, text_brush, float(x), float(y - 12))
    
    # Красный прямоугольник содержимого
    if content_bounds:
        red_pen = Pen(Color.Red, 2)
        x = content_bounds['min_x']
        y = content_bounds['min_y']
        bw = content_bounds['max_x'] - content_bounds['min_x'] + 1
        bh = content_bounds['max_y'] - content_bounds['min_y'] + 1
        g.DrawRectangle(red_pen, x, y, bw, bh)
        
        # Размер в px и мм
        from System.Drawing import Font as GFont
        font = GFont("Arial", 10, FontStyle.Bold)
        text_brush = SolidBrush(Color.Red)
        size_label = "{}x{} px".format(bw, bh)
        if mm_per_px is not None:
            w_mm = bw * mm_per_px
            h_mm = bh * mm_per_px
            size_label += " ({:.1f}x{:.1f} mm)".format(w_mm, h_mm)
        g.DrawString(size_label, font, text_brush, float(x), float(max(0, y - 18)))
        # Координаты углов
        coord_font = Font("Arial", 7, FontStyle.Regular)
        g.DrawString("({},{})".format(x, y), coord_font, text_brush, float(x), float(y + bh + 2))
    
    # Легенда
    legend_font = Font("Arial", 8, FontStyle.Regular)
    bg_brush = SolidBrush(Color.FromArgb(200, 0, 0, 0))
    g.FillRectangle(bg_brush, 5, 5, 235, 50)
    g.DrawString("RED = границы вида", legend_font, SolidBrush(Color.Red), 10, 8)
    g.DrawString("GREEN = калибр. квадраты", legend_font, SolidBrush(Color.LimeGreen), 10, 23)
    g.DrawString("YELLOW = зоны исключения", legend_font, SolidBrush(Color.Yellow), 10, 38)
    if page_bounds:
        g.DrawString("BLUE = границы листа", legend_font, SolidBrush(Color.Blue), 10, 53)
    
    g.Dispose()
    debug_bmp.Save(output_path)
    debug_bmp.Dispose()


def measure_view_with_calibration(view, stats=None):
    """Измеряет вид с помощью калибровочных квадратов.
    
    Args:
        view: Вид Revit
        stats: Необязательный словарь для сбора статистики времени.
               Если передан, будет добавлено поле 'analysis_time'.
    """
    import time
    temp_sheet = None
    temp_dir = None
    bitmap = None
    
    try:
        # Проверяем кэш
        view_id = view.Id.IntegerValue
        if view_id in VIEW_SIZE_CACHE:
            cached = VIEW_SIZE_CACHE[view_id]
            print("  📐 Анализ вида: " + view.Name + " (из кэша)")
            print("    Размер вида: " + str(int(cached[0])) + "x" + str(int(cached[1])) + " мм")
            return cached
        
        print("  📐 Анализ вида: " + view.Name)
        start_total = time.time()
        
        if not doc.IsModifiable:
            print("    ❌ Документ не редактируем!")
            return None
        
        # ========== 1. Создаём временный лист ==========
        start_create = time.time()
        sub_t = SubTransaction(doc)
        sub_t.Start()
        
        try:
            temp_sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
            temp_sheet.SheetNumber = "CAL_" + str(view.Id.IntegerValue)
            safe_name = sanitize_sheet_name("Cal_" + view.Name[:25])
            temp_sheet.Name = safe_name
            
            elements_on_sheet = FilteredElementCollector(doc, temp_sheet.Id)\
                .WhereElementIsNotElementType()\
                .ToElements()
            for elem in elements_on_sheet:
                doc.Delete(elem.Id)
            
            doc.Regenerate()
            
            sq_ft = CALIBRATION_SQUARE_MM / 304.8
            half_ft = sq_ft / 2
            # Расстояние между центрами квадратов 1400 мм, смещение от центра области до центра квадрата = 700 мм
            grid_ft = 700.0 / 304.8
            cx = 0
            cy = 0
            
            for dx, dy in [(-1, -1), (1, -1), (-1, 1), (1, 1)]:
                center = XYZ(cx + dx * grid_ft, cy + dy * grid_ft, 0)
                draw_thick_square(doc, temp_sheet, center, half_ft)
            
            Viewport.Create(doc, temp_sheet.Id, view.Id, XYZ(cx, cy, 0))
            doc.Regenerate()
            sub_t.Commit()
            create_time = time.time() - start_create
            print("    ✓ Калибровочный лист создан: " + safe_name)
            print("    ⏱ Создание листа: {:.2f} сек".format(create_time))
        
        except Exception as ex:
            if sub_t.HasStarted() and not sub_t.HasEnded():
                sub_t.RollBack()
            raise ex
        
        # ========== 2. Экспорт ==========
        temp_dir = Path.Combine(Path.GetTempPath(), "revit_cal_" + str(System.Guid.NewGuid()))
        System.IO.Directory.CreateDirectory(temp_dir)
        export_path = Path.Combine(temp_dir, "cal.png")
        
        opts = ImageExportOptions()
        opts.ZoomType = ZoomFitType.FitToPage
        opts.PixelSize = EXPORT_PIXELS
        opts.FilePath = export_path
        opts.HLRandWFViewsFileType = ImageFileType.PNG
        opts.ExportRange = ExportRange.SetOfViews
        opts.ImageResolution = ImageResolution.DPI_150
        opts.SetViewsAndSheets([temp_sheet.Id])
        
        try:
            start_export = time.time()
            doc.ExportImage(opts)
            export_time = time.time() - start_export
            print("    ✓ Экспорт выполнен")
            print("    ⏱ Экспорт: {:.2f} сек".format(export_time))
        except Exception as ex:
            print("    ❌ Ошибка экспорта: " + str(ex))
            print("      Возможные причины: недостаточно прав на запись, путь содержит недопустимые символы, Revit не может экспортировать изображение.")
            # Сохраняем информацию об ошибке в отладочную папку
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AutoView_debug")
            if not os.path.exists(debug_dir):
                os.makedirs(debug_dir)
            error_log = os.path.join(debug_dir, "ERROR_export_" + str(view.Id.IntegerValue) + ".txt")
            with open(error_log, 'w', encoding='utf-8') as f:
                f.write("Ошибка экспорта вида: " + view.Name + "\n")
                f.write(str(ex))
            print("    Подробности сохранены в: " + error_log)
            return None
        
        time.sleep(0.5)
        
        png_files = System.IO.Directory.GetFiles(temp_dir, "*.png")
        if not png_files or len(png_files) == 0:
            print("    ❌ Файл экспорта не найден в папке: " + temp_dir)
            print("      Возможные причины: Revit не создал файл, ошибка пути, антивирус заблокировал запись.")
            # Сохраняем информацию об ошибке
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AutoView_debug")
            if not os.path.exists(debug_dir):
                os.makedirs(debug_dir)
            error_log = os.path.join(debug_dir, "ERROR_no_file_" + str(view.Id.IntegerValue) + ".txt")
            with open(error_log, 'w', encoding='utf-8') as f:
                f.write("Файл экспорта не найден для вида: " + view.Name + "\n")
                f.write("Временная папка: " + temp_dir + "\n")
                f.write("Проверьте, создался ли файл .png в этой папке.\n")
            print("    Подробности сохранены в: " + error_log)
            return None
        
        exported_file = png_files[0]
        print("    Файл: " + Path.GetFileName(exported_file))
        
        debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "AutoView_debug")
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)
        
        import shutil
        safe_filename = sanitize_sheet_name(view.Name)[:30]
        debug_copy = os.path.join(debug_dir, "cal_" + str(view.Id.IntegerValue) + "_" + safe_filename + ".png")
        shutil.copy(exported_file, debug_copy)
        print("    Копия сохранена в: " + debug_copy)
        
        # ========== 3. Анализ ==========
        bitmap = Bitmap(exported_file)
        
        if bitmap.Width < 100 or bitmap.Height < 100:
            print("    ❌ Изображение слишком маленькое: " + str(bitmap.Width) + "x" + str(bitmap.Height) + " px")
            print("      Возможные причины: экспорт выполнен с низким разрешением, вид пуст, масштаб экспорта некорректен.")
            print("      Рекомендации: увеличьте EXPORT_PIXELS, проверьте, что вид содержит геометрию.")
            # Сохраняем информацию об ошибке
            error_log = os.path.join(debug_dir, "ERROR_small_image_" + str(view.Id.IntegerValue) + ".txt")
            with open(error_log, 'w', encoding='utf-8') as f:
                f.write("Изображение слишком маленькое для вида: " + view.Name + "\n")
                f.write("Размер: " + str(bitmap.Width) + "x" + str(bitmap.Height) + " px\n")
                f.write("Файл: " + exported_file + "\n")
            print("    Подробности сохранены в: " + error_log)
            return None
        
        print("    Размер изображения: " + str(bitmap.Width) + "x" + str(bitmap.Height) + " px")
        
        theme = detect_theme(bitmap)
        
        page_bounds = find_page_bounds(bitmap)
        
        squares = find_squares_by_edges(bitmap)
        
        if not squares or len(squares) < 2:
            print("    ❌ Квадраты не найдены (найдено: " + str(len(squares) if squares else 0) + "). Возможные причины:")
            print("      - Квадраты не видны на изображении (слишком маленькие, низкий контраст)")
            print("      - Изображение слишком тёмное или светлое")
            print("      - Квадраты перекрыты содержимым вида")
            print("      - Алгоритм обнаружения не справился (попробуйте увеличить EXPORT_PIXELS)")
            print("    Проверьте отладочное изображение в папке AutoView_debug.")
            debug_fail = os.path.join(debug_dir, "FAIL_no_squares_" + str(view.Id.IntegerValue) + ".png")
            try:
                draw_debug_overlay(bitmap, squares if squares else [], None, [], debug_fail, page_bounds, mm_per_px=None)
            except Exception as ex:
                print("    ⚠️ Не удалось нарисовать отладку: " + str(ex))
                bitmap.Save(debug_fail)
            return None
        
        # Фиксированная калибровка: 1000px = 1500mm → 1.5 мм/пиксель
        mm_per_px = MM_PER_PX  # 1.5 мм/пиксель
        avg_square_size_px = 3  # 5мм × 0.67 px/мм ≈ 3px (новый размер квадрата)
        print("    Квадратов найдено: " + str(len(squares)) + " (фиксированные координаты)")
        print("    Калибровка: " + str(round(mm_per_px, 4)) + " мм/пиксель (фиксированная)")
        print("    Размер квадрата: " + str(avg_square_size_px) + " px (5×5 мм)")
        
        # Проверка качества калибровки (информационная)
        if avg_square_size_px < MIN_SQUARE_SIZE_PX or avg_square_size_px > MAX_SQUARE_SIZE_PX:
            print("    ⚠️ Информация: фиксированный размер квадрата " + str(int(avg_square_size_px)) + " px выходит за ожидаемые пределы (" + str(MIN_SQUARE_SIZE_PX) + "-" + str(MAX_SQUARE_SIZE_PX) + " px).")
        if mm_per_px < MIN_MM_PER_PX or mm_per_px > MAX_MM_PER_PX:
            print("    ⚠️ Информация: фиксированный коэффициент калибровки " + str(round(mm_per_px, 4)) + " мм/пиксель выходит за ожидаемые пределы (" + str(MIN_MM_PER_PX) + "-" + str(MAX_MM_PER_PX) + ").")
        
        exclusion_zones = []
        margin_px = int(avg_square_size_px * 0.7)
        
        for sq in squares:
            exclusion_zones.append((
                max(0, sq['x'] - margin_px),
                max(0, sq['y'] - margin_px),
                min(bitmap.Width - 1, sq['x'] + margin_px),
                min(bitmap.Height - 1, sq['y'] + margin_px)
            ))
        
        # Улучшаем изображение для лучшего обнаружения тонких линий
        start_analysis = time.time()
        enhanced = enhance_bitmap(bitmap)
        enhance_time = time.time() - start_analysis
        # Сохраняем улучшенное изображение для отладки
        enhanced_debug_path = os.path.join(debug_dir, "enhanced_" + str(view.Id.IntegerValue) + "_" + safe_filename + ".png")
        try:
            enhanced.Save(enhanced_debug_path)
            print("    Улучшенное изображение сохранено: " + enhanced_debug_path)
        except Exception as ex:
            print("    ⚠️ Не удалось сохранить улучшенное изображение: " + str(ex))
        start_find = time.time()
        content_bounds = find_content_bounds_3pass(enhanced, exclusion_zones)
        find_time = time.time() - start_find
        enhanced.Dispose()
        analysis_time = time.time() - start_analysis
        print("    ⏱ Анализ: {:.2f} сек (enhance: {:.2f}, поиск: {:.2f})".format(
            analysis_time, enhance_time, find_time))
        
        if not content_bounds:
            print("    ❌ Содержимое не найдено. Возможные причины:")
            print("      - Вид действительно пуст (нет геометрии)")
            print("      - Контраст между содержимым и фоном недостаточен")
            print("      - Содержимое перекрыто калибровочными квадратами (зоны исключения)")
            print("      - Алгоритм обнаружения границ не справился (попробуйте увеличить контраст вида)")
            print("    Проверьте отладочное изображение в папке AutoView_debug.")
            debug_fail = os.path.join(debug_dir, "FAIL_no_content_" + str(view.Id.IntegerValue) + ".png")
            try:
                draw_debug_overlay(bitmap, squares, None, exclusion_zones, debug_fail, page_bounds, mm_per_px)
            except Exception as ex:
                print("    ⚠️ Не удалось нарисовать отладку: " + str(ex))
                bitmap.Save(debug_fail)
            return None
        
        content_w_px = content_bounds['max_x'] - content_bounds['min_x'] + 1
        content_h_px = content_bounds['max_y'] - content_bounds['min_y'] + 1
        
        content_w_mm = content_w_px * mm_per_px + 2 * PADDING_MM
        content_h_mm = content_h_px * mm_per_px + 2 * PADDING_MM
        
        print("    Границы содержимого: " + str(content_w_px) + "x" + str(content_h_px) + " px")
        print("    Размер вида: " + str(int(content_w_mm)) + "x" + str(int(content_h_mm)) + " мм (с отступами)")
        
        # Сохраняем отладку
        safe_name = sanitize_sheet_name(view.Name)[:30]
        debug_path = os.path.join(debug_dir, "OK_" + str(view.Id.IntegerValue) + "_" + safe_name + ".png")
        try:
            draw_debug_overlay(bitmap, squares, content_bounds, exclusion_zones, debug_path, page_bounds, mm_per_px)
        except Exception as ex:
            print("    ⚠️ Не удалось сохранить отмеченное изображение: " + str(ex))
        
        if content_w_mm < 10 or content_h_mm < 10:
            print("    ⚠️ Размер слишком маленький")
            return None
        
        # Сохраняем результат в кэш
        VIEW_SIZE_CACHE[view_id] = (content_w_mm, content_h_mm)
        
        total_time = time.time() - start_total
        print("    ⏱ Общее время анализа: {:.2f} сек".format(total_time))
        
        # Сохраняем статистику, если передан словарь
        if stats is not None:
            # Если stats имеет ключ 'times', добавляем время в список
            if 'times' in stats:
                stats['times'].append(total_time)
            else:
                stats['analysis_time'] = total_time
        
        return (content_w_mm, content_h_mm)
    
    except Exception as e:
        print("  ❌ Ошибка анализа: " + str(e))
        import traceback
        traceback.print_exc()
        return None
    
    finally:
        if bitmap:
            try:
                bitmap.Dispose()
            except:
                pass
        
        if temp_sheet:
            try:
                element = doc.GetElement(temp_sheet.Id)
                if element is not None and doc.IsModifiable:
                    sub_t2 = SubTransaction(doc)
                    sub_t2.Start()
                    try:
                        start_delete = time.time()
                        doc.Delete(temp_sheet.Id)
                        sub_t2.Commit()
                        delete_time = time.time() - start_delete
                        print("    ✓ Временный лист удалён")
                        print("    ⏱ Удаление листа: {:.2f} сек".format(delete_time))
                    except:
                        sub_t2.RollBack()
                        print("    ⚠️ Не удалось удалить временный лист")
            except:
                pass
        
        if temp_dir:
            try:
                System.IO.Directory.Delete(temp_dir, True)
            except:
                pass


# ============ УПАКОВЩИК ============

class MaxRectsPacker:
    def __init__(self, bin_width, bin_height, occupied_rects=None):
        self.bin_width = bin_width
        self.bin_height = bin_height
        self.free_rects = [(0, 0, bin_width, bin_height)]
        self.placed = []
        
        if occupied_rects:
            for (ox, oy, ow, oh) in occupied_rects:
                self._add_occupied(ox, oy, ow, oh)

    def _add_occupied(self, x, y, w, h):
        """Добавить занятый прямоугольник, вырезая его из свободных областей."""
        new_free = []
        for (fx, fy, fw, fh) in self.free_rects:
            if not (x >= fx + fw or x + w <= fx or y >= fy + fh or y + h <= fy):
                # Пересекаются, вырезаем
                if x > fx:
                    new_free.append((fx, fy, x - fx, fh))
                if x + w < fx + fw:
                    new_free.append((x + w, fy, fx + fw - (x + w), fh))
                if y > fy:
                    new_free.append((fx, fy, fw, y - fy))
                if y + h < fy + fh:
                    new_free.append((fx, y + h, fw, fy + fh - (y + h)))
            else:
                new_free.append((fx, fy, fw, fh))
        self.free_rects = new_free
        self._prune()

    def add_rect(self, width, height, item_id, heuristic='area'):
        best_score = float('inf')
        best_rect = None
        best_free_idx = None
        
        for idx, (fx, fy, fw, fh) in enumerate(self.free_rects):
            if fw >= width and fh >= height:
                if heuristic == 'area':
                    score = (fw - width) * (fh - height)
                elif heuristic == 'short_side':
                    score = min(fw - width, fh - height)
                elif heuristic == 'long_side':
                    score = max(fw - width, fh - height)
                else:
                    score = (fw - width) * (fh - height)
                
                if score < best_score:
                    best_score = score
                    best_rect = (fx, fy, width, height)
                    best_free_idx = idx
        
        if best_rect is None:
            return False
        
        x, y, w, h = best_rect
        self.placed.append((x, y, w, h, item_id))
        old = self.free_rects.pop(best_free_idx)
        
        if w < old[2]:
            self.free_rects.append((x + w, y, old[2] - w, h))
        if h < old[3]:
            self.free_rects.append((x, y + h, old[2], old[3] - h))
        
        self._prune()
        return True

    def _prune(self):
        i = 0
        while i < len(self.free_rects):
            fx, fy, fw, fh = self.free_rects[i]
            if fw <= 0 or fh <= 0:
                self.free_rects.pop(i)
                continue
            contained = any(
                i != j and ox <= fx and oy <= fy and ox + ow >= fx + fw and oy + oh >= fy + fh
                for j, (ox, oy, ow, oh) in enumerate(self.free_rects)
            )
            if contained:
                self.free_rects.pop(i)
            else:
                i += 1

    def get_placements(self):
        return self.placed


def pack_rectangles(rects, bin_width, bin_height, occupied_rects=None):
    if not rects or bin_width <= 0 or bin_height <= 0:
        return []
    
    strategies = [
        lambda r: max(r[0], r[1]),
        lambda r: r[0] * r[1],
        lambda r: r[0],
        lambda r: r[1],
    ]
    
    heuristics = ['area', 'short_side', 'long_side']
    
    best_result = []
    best_count = 0
    best_fill = 0
    
    for key_func in strategies:
        sorted_rects = sorted(rects, key=key_func, reverse=True)
        
        for heuristic in heuristics:
            packer = MaxRectsPacker(bin_width, bin_height, occupied_rects)
            
            for w, h, item_id in sorted_rects:
                if w <= bin_width and h <= bin_height:
                    packer.add_rect(w, h, item_id, heuristic)
            
            placements = packer.get_placements()
            placed_count = len(placements)
            
            if placed_count == 0:
                continue
            
            total_area = sum(w * h for _, _, w, h, _ in placements)
            fill_ratio = total_area / (bin_width * bin_height)
            
            if placed_count > best_count or (placed_count == best_count and fill_ratio > best_fill):
                best_result = placements
                best_count = placed_count
                best_fill = fill_ratio
    
    return best_result


def find_best_fill(rects, bin_w, bin_h, occupied_rects=None):
    """
    Подбирает оптимальную комбинацию прямоугольников для ОДНОГО листа.
    Перебирает варианты, выбирает с максимальным количеством и заполнением.
    """
    if not rects:
        return []
    
    # Сортируем по площади (большие primero)
    sorted_rects = sorted(rects, key=lambda r: r[0]*r[1], reverse=True)
    
    best_placements = []
    best_count = 0
    best_fill = 0
    
    # Шаг перебора: не все комбинации, а с адаптивным шагом
    total_rects = len(sorted_rects)
    step = max(1, total_rects // 15)
    
    for start in range(0, total_rects, step):
        for count in range(1, total_rects - start + 1):  # минимум 1 вид на лист
            subset = sorted_rects[start:start + count]
            
            # Быстрый фильтр по площади
            total_area = sum(w*h for w, h, _ in subset)
            if total_area > bin_w * bin_h * 1.3:
                break
            
            placements = pack_rectangles(subset, bin_w, bin_h, occupied_rects)
            placed_count = len(placements)
            
            if placed_count > best_count:
                best_placements = placements
                best_count = placed_count
                best_fill = sum(w*h for _, _, w, h, _ in placements) / (bin_w * bin_h)
            elif placed_count == best_count:
                fill = sum(w*h for _, _, w, h, _ in placements) / (bin_w * bin_h)
                if fill > best_fill:
                    best_placements = placements
                    best_fill = fill
    
    return best_placements


# ============ WINFORMS UI ============

def show_placement_form(views_data, sheet_w, sheet_h):
    form = Form()
    form.Text = "Размещение видов на листы v5.0"
    form.Size = Size(680, 650)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    
    result = {'views': None, 'settings': None}
    
    y = 10
    
    title = Label()
    title.Text = "Выберите виды для размещения на листах"
    title.Location = Point(10, y)
    title.Size = Size(640, 25)
    title.Font = Font("Arial", 10, FontStyle.Bold)
    form.Controls.Add(title)
    y += 30
    
    info = Label()
    info.Text = "Лист-шаблон: " + str(int(sheet_w)) + " x " + str(int(sheet_h)) + " мм | Доступно видов: " + str(len(views_data))
    info.Location = Point(10, y)
    info.Size = Size(640, 20)
    info.Font = Font("Arial", 8)
    info.ForeColor = Color.Gray
    form.Controls.Add(info)
    y += 25
    
    panel = Panel()
    panel.Location = Point(10, y)
    panel.Size = Size(640, 280)
    panel.BorderStyle = BorderStyle.FixedSingle
    panel.AutoScroll = True
    form.Controls.Add(panel)
    
    inner_y = 10
    current_group = None
    
    for view, group_name in views_data:
        if group_name != current_group:
            if current_group is not None:
                inner_y += 5
            
            group_label = Label()
            group_label.Text = "━━━ " + group_name + " ━━━"
            group_label.Location = Point(10, inner_y)
            group_label.Size = Size(600, 22)
            group_label.Font = Font("Arial", 9, FontStyle.Bold)
            group_label.ForeColor = Color.DarkBlue
            panel.Controls.Add(group_label)
            inner_y += 27
            current_group = group_name
        
        cb = CheckBox()
        cb.Text = view.Name
        cb.Location = Point(30, inner_y)
        cb.Size = Size(580, 20)
        cb.Font = Font("Arial", 8)
        cb.Tag = view
        cb.Checked = True
        panel.Controls.Add(cb)
        inner_y += 22
    
    y += 290
    
    def on_select_all(s, e):
        for ctrl in panel.Controls:
            if isinstance(ctrl, CheckBox):
                ctrl.Checked = True
    
    def on_deselect_all(s, e):
        for ctrl in panel.Controls:
            if isinstance(ctrl, CheckBox):
                ctrl.Checked = False
    
    btn_select = Button()
    btn_select.Text = "✓ Выбрать все"
    btn_select.Location = Point(10, y)
    btn_select.Size = Size(130, 30)
    btn_select.Click += on_select_all
    form.Controls.Add(btn_select)
    
    btn_deselect = Button()
    btn_deselect.Text = "✗ Снять все"
    btn_deselect.Location = Point(150, y)
    btn_deselect.Size = Size(130, 30)
    btn_deselect.Click += on_deselect_all
    form.Controls.Add(btn_deselect)
    y += 40
    
    gb = GroupBox()
    gb.Text = "Настройки (мм)"
    gb.Location = Point(10, y)
    gb.Size = Size(640, 105)
    gb.Font = Font("Arial", 8, FontStyle.Bold)
    form.Controls.Add(gb)
    
    lbl_mx = Label(); lbl_mx.Text = "Отступ X:"; lbl_mx.Location = Point(15, 25); lbl_mx.Size = Size(65, 20); lbl_mx.Font = Font("Arial", 8); gb.Controls.Add(lbl_mx)
    tb_margin_x = TextBox(); tb_margin_x.Text = "20"; tb_margin_x.Location = Point(80, 22); tb_margin_x.Size = Size(45, 20); gb.Controls.Add(tb_margin_x)
    
    lbl_my = Label(); lbl_my.Text = "Отступ Y:"; lbl_my.Location = Point(140, 25); lbl_my.Size = Size(65, 20); lbl_my.Font = Font("Arial", 8); gb.Controls.Add(lbl_my)
    tb_margin_y = TextBox(); tb_margin_y.Text = "20"; tb_margin_y.Location = Point(205, 22); tb_margin_y.Size = Size(45, 20); gb.Controls.Add(tb_margin_y)
    
    lbl_gx = Label(); lbl_gx.Text = "Зазор X:"; lbl_gx.Location = Point(270, 25); lbl_gx.Size = Size(55, 20); lbl_gx.Font = Font("Arial", 8); gb.Controls.Add(lbl_gx)
    tb_gap_x = TextBox(); tb_gap_x.Text = "15"; tb_gap_x.Location = Point(325, 22); tb_gap_x.Size = Size(45, 20); gb.Controls.Add(tb_gap_x)
    
    lbl_gy = Label(); lbl_gy.Text = "Зазор Y:"; lbl_gy.Location = Point(390, 25); lbl_gy.Size = Size(55, 20); lbl_gy.Font = Font("Arial", 8); gb.Controls.Add(lbl_gy)
    tb_gap_y = TextBox(); tb_gap_y.Text = "15"; tb_gap_y.Location = Point(445, 22); tb_gap_y.Size = Size(45, 20); gb.Controls.Add(tb_gap_y)
    
    lbl_tb = Label(); lbl_tb.Text = "Штамп:"; lbl_tb.Location = Point(510, 25); lbl_tb.Size = Size(50, 20); lbl_tb.Font = Font("Arial", 8); gb.Controls.Add(lbl_tb)
    tb_titleblock = TextBox(); tb_titleblock.Text = "60"; tb_titleblock.Location = Point(560, 22); tb_titleblock.Size = Size(45, 20); gb.Controls.Add(tb_titleblock)
    
    lbl_prefix = Label(); lbl_prefix.Text = "Префикс имени:"; lbl_prefix.Location = Point(15, 55); lbl_prefix.Size = Size(100, 20); lbl_prefix.Font = Font("Arial", 8); gb.Controls.Add(lbl_prefix)
    tb_prefix = TextBox(); tb_prefix.Text = ""; tb_prefix.Location = Point(120, 52); tb_prefix.Size = Size(200, 20); gb.Controls.Add(tb_prefix)
    
    lbl_hint = Label(); lbl_hint.Text = "мм | Все значения >= 0 | Пустой префикс = авто"; lbl_hint.Location = Point(15, 80); lbl_hint.Size = Size(600, 20); lbl_hint.Font = Font("Arial", 7); lbl_hint.ForeColor = Color.Gray; gb.Controls.Add(lbl_hint)
    
    y += 115
    
    def on_ok(s, e):
        selected = []
        for ctrl in panel.Controls:
            if isinstance(ctrl, CheckBox) and ctrl.Checked and ctrl.Tag:
                selected.append(ctrl.Tag)
        
        if not selected:
            MessageBox.Show("Не выбрано ни одного вида!", "Предупреждение")
            return
        
        fields = {
            'Отступ X': tb_margin_x.Text,
            'Отступ Y': tb_margin_y.Text,
            'Зазор X': tb_gap_x.Text,
            'Зазор Y': tb_gap_y.Text,
            'Штамп': tb_titleblock.Text,
        }
        
        values = {}
        for name, text in fields.items():
            try:
                val = int(text)
            except ValueError:
                MessageBox.Show("Поле '" + name + "' должно содержать целое число!", "Ошибка")
                return
            if val < 0:
                MessageBox.Show("Поле '" + name + "' не может быть отрицательным!", "Ошибка")
                return
            values[name] = val
        
        if values['Отступ X'] * 2 >= sheet_w:
            MessageBox.Show("Отступы X превышают ширину листа!", "Ошибка")
            return
        if values['Отступ Y'] * 2 >= sheet_h:
            MessageBox.Show("Отступы Y превышают высоту листа!", "Ошибка")
            return
        
        result['views'] = selected
        result['settings'] = {
            'margin_x': values['Отступ X'],
            'margin_y': values['Отступ Y'],
            'gap_x': values['Зазор X'],
            'gap_y': values['Зазор Y'],
            'titleblock_offset': values['Штамп'],
            'sheet_prefix': tb_prefix.Text.strip(),
        }
        form.DialogResult = DialogResult.OK
        form.Close()
    
    def on_cancel(s, e):
        form.DialogResult = DialogResult.Cancel
        form.Close()
    
    btn_ok = Button()
    btn_ok.Text = "▶ Разместить виды"
    btn_ok.Location = Point(350, y)
    btn_ok.Size = Size(160, 40)
    btn_ok.Font = Font("Arial", 10, FontStyle.Bold)
    btn_ok.BackColor = Color.LightGreen
    btn_ok.Click += on_ok
    form.Controls.Add(btn_ok)
    
    btn_cancel = Button()
    btn_cancel.Text = "Отмена"
    btn_cancel.Location = Point(520, y)
    btn_cancel.Size = Size(120, 40)
    btn_cancel.Click += on_cancel
    form.Controls.Add(btn_cancel)
    
    form.ShowDialog()
    
    if result['views']:
        return result['views'], result['settings']
    return None, None


# ============ MAIN ============

def main():
    import time
    start_total = time.time()
    
    print("\n" + "=" * 60)
    print("  Размещение видов на листы v6.0")
    print("=" * 60 + "\n")
    
    # Очищаем кэш размеров видов
    VIEW_SIZE_CACHE.clear()
    print("🧹 Кэш размеров видов очищен")
    
    # Статистика времени
    total_analysis_time = 0.0
    total_placement_time = 0.0
    
    import shutil
    script_dir = os.path.dirname(os.path.abspath(__file__))
    debug_dir = os.path.join(script_dir, "AutoView_debug")
    
    if os.path.exists(debug_dir):
        for filename in os.listdir(debug_dir):
            file_path = os.path.join(debug_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print("⚠️ Не удалось удалить " + file_path + ": " + str(e))
        print("🧹 Папка отладки очищена: " + debug_dir)
    else:
        os.makedirs(debug_dir)
        print("📁 Создана папка отладки: " + debug_dir)
    
    sel_ids = list(uidoc.Selection.GetElementIds())
    if not sel_ids:
        MessageBox.Show("Выберите лист-шаблон перед запуском!", "Ошибка")
        return
    
    template_sheet = doc.GetElement(sel_ids[0])
    if not isinstance(template_sheet, ViewSheet):
        MessageBox.Show("Выбранный элемент не является листом!", "Ошибка")
        return
    
    print("📄 Лист-шаблон: " + template_sheet.SheetNumber + " — '" + template_sheet.Name + "'")
    
    frame_info = get_sheet_frame_mm(template_sheet)
    if not frame_info:
        MessageBox.Show("Не удалось определить размеры листа!", "Ошибка")
        return
    
    frame_w, frame_h, frame_min_x, frame_min_y = frame_info
    print("📐 Рамка: " + str(int(frame_w)) + "x" + str(int(frame_h)) + " мм")
    print("  Начало рамки в системе Revit: X={:.1f} мм, Y={:.1f} мм".format(frame_min_x, frame_min_y))
    print("  Нормализация: сдвиг X={:.1f} мм, Y={:.1f} мм".format(-frame_min_x, -frame_min_y))
    
    shift_x = -frame_min_x
    shift_y = -frame_min_y
    print("  После нормализации: начало в (0, 0) мм")
    
    placed_cache = collect_placed_views()
    
    all_3d = list(FilteredElementCollector(doc)
                  .OfClass(View3D)
                  .WhereElementIsNotElementType()
                  .ToElements())
    
    views_data = []
    for view in all_3d:
        if view.IsTemplate or view.Id in placed_cache:
            continue
        prefix, _ = extract_prefix_and_numbers(view.Name)
        if not prefix:
            continue
        views_data.append((view, get_group_for_prefix(prefix)))
    
    print("🔍 3D виды: всего " + str(len(all_3d)) + ", доступно " + str(len(views_data)))
    
    if not views_data:
        MessageBox.Show("Нет доступных 3D видов для размещения!", "Информация")
        return
    
    views_data.sort(key=lambda x: (
        list(GROUPING.keys()).index(x[1]) if x[1] in GROUPING else 999,
        get_view_sort_key(x[0])
    ))
    
    selected_views, settings = show_placement_form(views_data, frame_w, frame_h)
    if not selected_views:
        print("👋 Отменено пользователем")
        return
    
    print("✅ Выбрано видов: " + str(len(selected_views)))
    
    mx = settings['margin_x']
    my = settings['margin_y']
    gx = settings['gap_x']
    gy = settings['gap_y']
    tb_off = settings['titleblock_offset']
    sheet_prefix = settings.get('sheet_prefix', '')
    
    groups = defaultdict(list)
    for view in selected_views:
        prefix, _ = extract_prefix_and_numbers(view.Name)
        groups[get_group_for_prefix(prefix)].append(view)
    
    for g in groups:
        groups[g].sort(key=get_view_sort_key)
    
    print("\n📊 Группы для размещения:")
    for g in groups:
        print("  " + g + ": " + str(len(groups[g])) + " видов")
        for v in groups[g]:
            print("    - " + v.Name)
    
    print("\n Анализ видов...")
    
    view_sizes = {}
    failed = []
    analysis_stats = {'times': []}  # для сбора времени анализа
    
    analysis_t = Transaction(doc, "Analyze views")
    analysis_t.Start()
    
    try:
        for v in selected_views:
            size = measure_view_with_calibration(v, stats=analysis_stats)
            if size:
                view_sizes[v.Id] = size
                print("    ✅ " + v.Name + ": " + "{:.1f}".format(size[0]) + "x" + "{:.1f}".format(size[1]) + " мм")
            else:
                print("    ❌ Вид пропускается: " + v.Name)
                failed.append(v)
        
        analysis_t.RollBack()
        print("   🔄 Транзакция анализа отменена")
        
        # Суммируем время анализа
        if analysis_stats['times']:
            total_analysis_time = sum(analysis_stats['times'])
            print("   ⏱ Итого анализ: {:.2f} сек ({} видов)".format(
                total_analysis_time, len(analysis_stats['times'])))
    
    except Exception as ex:
        if analysis_t.HasStarted() and not analysis_t.HasEnded():
            analysis_t.RollBack()
        raise ex
    
    if failed:
        print("\n  ⚠️ Не удалось проанализировать " + str(len(failed)) + " видов:")
        for fv in failed:
            print("     - " + fv.Name)
        print()
    
    if not view_sizes:
        MessageBox.Show("Не удалось определить размеры ни одного вида!", "Ошибка")
        return
    
    stamp_width = 100  # ширина штампа (правая часть листа) - не вычитаем из доступной ширины
    stamp_height = tb_off  # высота штампа (нижняя часть листа)
    # Доступная ширина - полная ширина минус отступы (штамп не вычитаем из ширины)
    avail_w = int(frame_w - 2 * mx)
    # Доступная высота - полная высота минус верхний отступ и высота штампа
    avail_h = int(frame_h - my - stamp_height)
    if avail_w <= 0 or avail_h <= 0:
        print("⚠️ Доступная область после вычета штампа слишком мала, используем исходные размеры")
        avail_w = int(frame_w - 2 * mx)
        avail_h = int(frame_h - my)
    
    # Подробная информация о штампе и доступной области
    print("\n📋 ИНФОРМАЦИЯ О РАЗМЕЩЕНИИ НА ЛИСТЕ:")
    print("  Размер листа: {}x{} мм".format(int(frame_w), int(frame_h)))
    print("  Отступы: X={} мм, Y={} мм".format(mx, my))
    print("  Штамп (основная надпись):")
    print("    - Размер: {}x{} мм".format(stamp_width, stamp_height))
    print("    - Расположение: правый нижний угол")
    print("    - Координаты углов: левый нижний=({}, {}), правый верхний=({}, {})".format(
        int(frame_w - stamp_width), int(0), int(frame_w), int(stamp_height)))
    print("  Доступная область для видов: {}x{} мм".format(avail_w, avail_h))
    print("  Координаты доступной области: левый нижний=({}, {}), правый верхний=({}, {})".format(
        int(mx), int(stamp_height), int(frame_w - mx), int(frame_h - my)))
    
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    used_numbers = {s.SheetNumber for s in all_sheets}
    
    tmpl_num = template_sheet.SheetNumber
    match = re.search(r'(\d+)$', tmpl_num)
    if match:
        num_prefix = tmpl_num[:match.start()]
        next_num = int(match.group(1)) + 1
    else:
        num_prefix = tmpl_num + '_'
        next_num = 1
    
    class SheetNumberGenerator:
        def __init__(self):
            self.prefix = num_prefix
            self.current = next_num
            self.used = used_numbers
        
        def get_next(self):
            while True:
                candidate = self.prefix + str(self.current)
                self.current += 1
                if candidate not in self.used:
                    self.used.add(candidate)
                    return candidate
                if self.current > 99999:
                    raise Exception("Не удалось найти свободный номер листа!")
    
    num_gen = SheetNumberGenerator()
    
    total_placed = 0
    sheet_count = 0
    skipped_views = []
    
    placement_start = time.time()
    t = Transaction(doc, "Place views")
    t.Start()
    try:
        for group_name in GROUPING:
            if group_name not in groups:
                continue
            
            remaining = list(groups[group_name])
            
            while remaining:
                fit_views = []
                fit_rects = []
                oversize_views = []
                unknown_views = []
                
                for v in remaining:
                    size = view_sizes.get(v.Id)
                    if not size:
                        print("      ? " + v.Name + ": нет размера -> unknown")
                        unknown_views.append(v)
                        continue
                    
                    wg = int(math.ceil(size[0] + gx))
                    hg = int(math.ceil(size[1] + gy))
                    
                    if wg <= avail_w and hg <= avail_h:
                        print("      ✓ " + v.Name + ": " + "{:.1f}".format(size[0]) + "x" + "{:.1f}".format(size[1]) + " мм + зазоры " + str(wg) + "x" + str(hg) + " -> fit")
                        fit_views.append(v)
                        fit_rects.append((wg, hg, v.Id))
                    else:
                        print("      ⚠ " + v.Name + ": " + "{:.1f}".format(size[0]) + "x" + "{:.1f}".format(size[1]) + " мм + зазоры " + str(wg) + "x" + str(hg) + " > доступно " + str(avail_w) + "x" + str(avail_h) + " -> oversize")
                        oversize_views.append(v)
                
                processed_set = set(v.Id for v in fit_views + oversize_views + unknown_views)
                remaining = [v for v in remaining if v.Id not in processed_set]
                
                print("      Итого: fit=" + str(len(fit_views)) + ", oversize=" + str(len(oversize_views)) + ", unknown=" + str(len(unknown_views)) + ", remaining=" + str(len(remaining)))
                
                for v in oversize_views:
                    vw, vh = view_sizes[v.Id]
                    new_sheet = create_sheet(template_sheet)
                    new_sheet.SheetNumber = num_gen.get_next()
                    clear_sheet_viewports(new_sheet)
                    
                    left_x = mx + max(0, (avail_w - vw) // 2)
                    left_y = stamp_height + max(0, (avail_h - vh) // 2)
                    # Viewport.Create ожидает координаты ЦЕНТРА видового экрана
                    center_x = left_x + vw / 2.0
                    center_y = left_y + vh / 2.0
                    # frame_min_x, frame_min_y - положение рамки в системе Revit
                    x_mm = frame_min_x + center_x
                    y_mm = frame_min_y + center_y
                    
                    # Отладочная информация о размещении
                    print("  📐 Размещение негабаритного вида:")
                    print("     Размер вида: {}x{} мм".format(int(vw), int(vh)))
                    print("     Доступная область: {}x{} мм (с учетом штампа {}x{} мм)".format(
                        avail_w, avail_h, stamp_width, stamp_height))
                    print("     Система координат:")
                    print("       - Начало рамки в Revit: X={:.1f} мм, Y={:.1f} мм".format(frame_min_x, frame_min_y))
                    print("       - Сдвиг нормализации: X={:.1f} мм, Y={:.1f} мм".format(shift_x, shift_y))
                    print("       - Координаты от угла рамки: X={} мм, Y={} мм".format(int(left_x), int(left_y)))
                    print("     Итоговые координаты в Revit: X={:.1f} мм, Y={:.1f} мм".format(x_mm, y_mm))
                    print("     Координаты углов в Revit: левый нижний=({:.1f}, {:.1f}), правый верхний=({:.1f}, {:.1f})".format(
                        x_mm, y_mm, x_mm + vw, y_mm + vh))
                    print("     Передаем в Viewport.Create: X={:.3f} фут, Y={:.3f} фут".format(
                        x_mm / 304.8, y_mm / 304.8))
                    
                    Viewport.Create(doc, new_sheet.Id, v.Id,
                                   XYZ(x_mm / 304.8, y_mm / 304.8, 0))
                    
                    prefix = sheet_prefix if sheet_prefix else group_name
                    new_sheet.Name = sanitize_sheet_name(
                        prefix + ": " + get_view_short_name(v.Name))
                    
                    sheet_count += 1
                    total_placed += 1
                    print("  📄 Лист " + new_sheet.SheetNumber + ": " + v.Name + " (негабарит " + str(int(vw)) + "x" + str(int(vh)) + ")")
                
                for v in unknown_views:
                    skipped_views.append(v)
                    print("  ⚠️ Пропущен (нет размера): " + v.Name)
                
                if not fit_rects:
                    continue
                
                placements = find_best_fill(fit_rects, avail_w, avail_h)
                
                if not placements:
                    print("  ⚠️ Упаковщик не справился, размещаем по одному")
                    for v in fit_views:
                        vw, vh = view_sizes[v.Id]
                        new_sheet = create_sheet(template_sheet)
                        new_sheet.SheetNumber = num_gen.get_next()
                        clear_sheet_viewports(new_sheet)
                        
                        left_x = mx + max(0, (avail_w - vw) // 2)
                        left_y = stamp_height + max(0, (avail_h - vh) // 2)
                        # Viewport.Create ожидает координаты ЦЕНТРА видового экрана
                        center_x = left_x + vw / 2.0
                        center_y = left_y + vh / 2.0
                        # frame_min_x, frame_min_y - положение рамки в системе Revit
                        x_mm = frame_min_x + center_x
                        y_mm = frame_min_y + center_y
                        
                        # Отладочная информация о размещении
                        print("  📐 Размещение одиночного вида (упаковщик не справился):")
                        print("     Размер вида: {}x{} мм".format(int(vw), int(vh)))
                        print("     Доступная область: {}x{} мм (с учетом штампа {}x{} мм)".format(
                            avail_w, avail_h, stamp_width, stamp_height))
                        print("     Система координат:")
                        print("       - Начало рамки в Revit: X={:.1f} мм, Y={:.1f} мм".format(frame_min_x, frame_min_y))
                        print("       - Сдвиг нормализации: X={:.1f} мм, Y={:.1f} мм".format(shift_x, shift_y))
                        print("       - Координаты от угла рамки: X={} мм, Y={} мм".format(int(left_x), int(left_y)))
                        print("     Итоговые координаты в Revit: X={:.1f} мм, Y={:.1f} мм".format(x_mm, y_mm))
                        print("     Координаты углов в Revit: левый нижний=({:.1f}, {:.1f}), правый верхний=({:.1f}, {:.1f})".format(
                            x_mm, y_mm, x_mm + vw, y_mm + vh))
                        print("     Передаем в Viewport.Create: X={:.3f} фут, Y={:.3f} фут".format(
                            x_mm / 304.8, y_mm / 304.8))
                        
                        Viewport.Create(doc, new_sheet.Id, v.Id,
                                       XYZ(x_mm / 304.8, y_mm / 304.8, 0))
                        
                        prefix = sheet_prefix if sheet_prefix else group_name
                        new_sheet.Name = sanitize_sheet_name(
                            prefix + ": " + get_view_short_name(v.Name))
                        
                        sheet_count += 1
                        total_placed += 1
                    continue
                
                placed_ids = set(p[4] for p in placements)
                placed_views = [v for v in fit_views if v.Id in placed_ids]
                
                new_sheet = create_sheet(template_sheet)
                new_sheet.SheetNumber = num_gen.get_next()
                clear_sheet_viewports(new_sheet)
                
                # Отладочная информация о размещении группы видов
                print("  📐 Размещение группы видов ({} шт) на листе {}:".format(
                    len(placements), new_sheet.SheetNumber))
                print("     Доступная область: {}x{} мм (с учетом штампа {}x{} мм)".format(
                    avail_w, avail_h, stamp_width, stamp_height))
                print("     Область штампа: правый нижний угол, размер {}x{} мм".format(
                    stamp_width, stamp_height))
                
                for idx, (x, y, w, h, vid) in enumerate(placements, 1):
                    view = next((v for v in placed_views if v.Id == vid), None)
                    if not view:
                        continue
                    vw, vh = view_sizes[vid]
                    
                    left_x = mx + x
                    left_y = stamp_height + y
                    # Viewport.Create ожидает координаты ЦЕНТРА видового экрана
                    center_x = left_x + vw / 2.0
                    center_y = left_y + vh / 2.0
                    # frame_min_x, frame_min_y - положение рамки в системе Revit
                    x_mm = frame_min_x + center_x
                    y_mm = frame_min_y + center_y
                    
                    # Отладочная информация для каждого вида
                    print("     Вид {}: {}".format(idx, get_view_short_name(view.Name)))
                    print("       Размер: {}x{} мм".format(int(vw), int(vh)))
                    print("       Система координат:")
                    print("         - Начало рамки в Revit: X={:.1f} мм, Y={:.1f} мм".format(frame_min_x, frame_min_y))
                    print("         - Сдвиг нормализации: X={:.1f} мм, Y={:.1f} мм".format(shift_x, shift_y))
                    print("         - Координаты от угла рамки: X={} мм, Y={} мм".format(int(left_x), int(left_y)))
                    print("       Итоговые координаты в Revit: X={:.1f} мм, Y={:.1f} мм".format(x_mm, y_mm))
                    print("       Координаты углов в Revit: левый нижний=({:.1f}, {:.1f}), правый верхний=({:.1f}, {:.1f})".format(
                        x_mm, y_mm, x_mm + vw, y_mm + vh))
                    print("       Передаем в Viewport.Create: X={:.3f} фут, Y={:.3f} фут".format(
                        x_mm / 304.8, y_mm / 304.8))
                    
                    Viewport.Create(doc, new_sheet.Id, view.Id,
                                   XYZ(x_mm / 304.8, y_mm / 304.8, 0))
                
                short_names = [get_view_short_name(v.Name) for v in placed_views]
                prefix = sheet_prefix if sheet_prefix else group_name
                new_sheet.Name = sanitize_sheet_name(
                    prefix + ": " + ", ".join(short_names))
                
                sheet_count += 1
                total_placed += len(placed_views)
                print("  📄 Лист " + new_sheet.SheetNumber + ": " + str(len(placed_views)) + " видов — " + ", ".join(v.Name for v in placed_views))
                
                # Добавляем неразмещённые fit виды обратно в remaining для следующей итерации
                unplaced_fit_views = [v for v in fit_views if v.Id not in placed_ids]
                if unplaced_fit_views:
                    remaining.extend(unplaced_fit_views)
                    print("  🔄 Неразмещённые fit виды (" + str(len(unplaced_fit_views)) + "): " + ", ".join(v.Name for v in unplaced_fit_views))
        
        t.Commit()
        placement_time = time.time() - placement_start
        
        print("\n" + "=" * 60)
        print("  ✅ Готово!")
        print("  Создано листов: " + str(sheet_count))
        print("  Размещено видов: " + str(total_placed))
        if skipped_views:
            print("  Пропущено видов: " + str(len(skipped_views)))
        print("  ⏱ Время размещения: {:.2f} сек".format(placement_time))
        print("=" * 60 + "\n")
        
        msg = "Готово!\n\nСоздано листов: " + str(sheet_count) + "\nРазмещено видов: " + str(total_placed)
        if skipped_views:
            msg += "\n\nПропущено видов: " + str(len(skipped_views))
        MessageBox.Show(msg, "Готово")
    
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        MessageBox.Show("Ошибка при размещении видов: " + str(ex), "Ошибка")
        import traceback
        traceback.print_exc()
    
    finally:
        total_time = time.time() - start_total
        print("\n  ⏱ Общее время выполнения скрипта: {:.2f} сек".format(total_time))


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True

def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        main()
        return True
    except Exception as e:
        MessageBox.Show("Критическая ошибка: " + str(e), "Ошибка")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    main()