# -*- coding: utf-8 -*-

__title__ = """Размещение
видов на листы"""
__author__ = 'Rage'
__doc__ = """Размещает 3D виды на листы по системам с оптимальным заполнением.
Выбираем лист для образца основной надписи, запускаем скрипт."""
__version__ = "6.1"

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
    CheckBox, Button, Label, TextBox, GroupBox, ComboBox,
    ComboBoxStyle, ProgressBar, ProgressBarStyle, FlatStyle,
    Timer, Clipboard
)
from System.Drawing import (
    Font, FontStyle,
    Color, Point, Size,
    Bitmap, Imaging, Graphics, Pen, SolidBrush, Rectangle,
    ContentAlignment, SystemColors
)
from System.Drawing.Imaging import PixelFormat
from System.IO import Path, File, Directory
from System.Diagnostics import Process

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

DEFAULT_GROUPING = OrderedDict([
    ("П-В",     ["П", "ПЕ", "В", "ВЕ"]),
    ("ДП-ДВ",   ["ДП", "ДПЕ", "ДВ", "ДВЕ"]),
    ("А-У",     ["А", "У"]),
])

GROUPING = OrderedDict(DEFAULT_GROUPING)

ALL_PREFIXES = ['ПЕ', 'ВЕ', 'ДПЕ', 'ДВЕ', 'П', 'В', 'ДП', 'ДВ', 'А', 'У']
GROUP_NAMES_DEFAULT = ["П-В", "ДП-ДВ", "А-У", "", ""]

EXPORT_PIXELS = 1000
PADDING_MM = 5
STAMP_WIDTH_MM = 185
STAMP_HEIGHT_MM = 55
CALIBRATION_SQUARE_MM = 5.0
SQUARE_LINE_STEP_MM = 0.5
SQUARE_LINE_COUNT = 1
MM_PER_PX = 1.5

MIN_SQUARE_SIZE_PX = 3
MAX_SQUARE_SIZE_PX = 1000
MIN_MM_PER_PX = 0.3
MAX_MM_PER_PX = 2.0

VIEW_SIZE_CACHE = {}

LOG_COLLECTOR = []
SHOW_LOGS = False
CREATION_WARNINGS = []

def log_message(msg):
    global LOG_COLLECTOR
    LOG_COLLECTOR.append(msg)
    if SHOW_LOGS:
        print(msg)

def warn_creation(msg):
    global CREATION_WARNINGS
    CREATION_WARNINGS.append(msg)
    log_message("  ⚠️ " + msg)

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
    tokens = re.split(r'[ _\-]+', name_upper)
    sorted_prefixes = sorted(ALL_PREFIXES, key=len, reverse=True)
    
    for token in tokens:
        token = token.strip()
        if not token:
            continue
        
        for prefix in sorted_prefixes:
            if len(token) <= len(prefix):
                continue
            if not token.startswith(prefix):
                continue
            
            after = token[len(prefix):]
            if after and after[0].isdigit():
                numbers = re.findall(r'\d+', after)
                if numbers:
                    return prefix, tuple(int(n) for n in numbers)
    
    for prefix in sorted_prefixes:
        idx = name_upper.find(prefix)
        if idx >= 0:
            after = name_upper[idx + len(prefix):]
            if after and after[0].isdigit():
                numbers = re.findall(r'\d+', after)
                if numbers:
                    return prefix, tuple(int(n) for n in numbers)
    
    numbers = re.findall(r'\d+', view_name)
    if numbers:
        return "", tuple(int(n) for n in numbers)
    
    return "", (0,)


def get_group_for_prefix(prefix):
    if not prefix:
        return "Без системы"
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
        log_message("Ошибка при сборе размещённых видов: " + str(e))
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
    
    GOST_SIZES = {
        "А4": (297, 210), "А3": (420, 297), "А2": (594, 420),
        "А1": (841, 594), "А0": (1189, 841),
    }
    
    param_format = tb.LookupParameter("Формат ГОСТ")
    param_height_real = tb.LookupParameter("Высота_Реальная")
    param_width_real = tb.LookupParameter("Ширина_Реальная")
    param_height = tb.LookupParameter("Высота")
    param_width = tb.LookupParameter("Ширина")
    
    width_mm_from_bbox = (bbox.Max.X - bbox.Min.X) * 304.8
    height_mm_from_bbox = (bbox.Max.Y - bbox.Min.Y) * 304.8
    
    MIN_FRAME_SIZE_MM = 50.0
    MAX_FRAME_SIZE_MM = 5000.0
    
    def is_reasonable(h, w):
        return (MIN_FRAME_SIZE_MM <= h <= MAX_FRAME_SIZE_MM and
                MIN_FRAME_SIZE_MM <= w <= MAX_FRAME_SIZE_MM)
    
    if param_format and param_format.HasValue:
        format_str = param_format.AsString()
        if format_str in GOST_SIZES:
            std_w, std_h = GOST_SIZES[format_str]
            param_orientation = tb.LookupParameter("Книжная ориентация")
            if param_orientation and param_orientation.HasValue and param_orientation.AsInteger() == 0:
                width_mm = max(std_w, std_h)
                height_mm = min(std_w, std_h)
            else:
                width_mm = min(std_w, std_h)
                height_mm = max(std_w, std_h)
            return (width_mm, height_mm, bbox.Min.X * 304.8, bbox.Min.Y * 304.8)
    
    for h_param, w_param in [(param_height_real, param_width_real), (param_height, param_width)]:
        if h_param and w_param:
            try:
                h_mm = h_param.AsDouble() * 304.8
                w_mm = w_param.AsDouble() * 304.8
                if is_reasonable(h_mm, w_mm):
                    return (w_mm, h_mm, bbox.Min.X * 304.8, bbox.Min.Y * 304.8)
            except Exception:
                pass
    
    if is_reasonable(height_mm_from_bbox, width_mm_from_bbox):
        return (width_mm_from_bbox, height_mm_from_bbox, bbox.Min.X * 304.8, bbox.Min.Y * 304.8)
    
    return (420.0, 297.0, bbox.Min.X * 304.8, bbox.Min.Y * 304.8)


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
                except Exception:
                    pass
    
    # Копируем параметры листа
    sheet_params = ["ADSK_Штамп Раздел проекта", "ADSK_Номер проекта", "ADSK_Название проекта"]
    for param_name in sheet_params:
        try:
            p_template = template_sheet.LookupParameter(param_name)
            p_new = new_sheet.LookupParameter(param_name)
            if p_template and p_template.HasValue and p_new and not p_new.IsReadOnly:
                if p_template.StorageType == StorageType.String:
                    p_new.Set(p_template.AsString())
        except Exception:
            pass
    
    return new_sheet


def clear_sheet_viewports(sheet):
    vps = list(FilteredElementCollector(doc, sheet.Id)
               .OfCategory(BuiltInCategory.OST_Viewports)
               .WhereElementIsNotElementType()
               .ToElements())
    for vp in vps:
        doc.Delete(vp.Id)


def draw_thick_square(doc, sheet, center_ft, half_size_ft):
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
        
        doc.Create.NewDetailCurve(sheet, Line.CreateBound(XYZ(x0, y0, 0), XYZ(x1, y0, 0)))
        doc.Create.NewDetailCurve(sheet, Line.CreateBound(XYZ(x0, y1, 0), XYZ(x1, y1, 0)))
        doc.Create.NewDetailCurve(sheet, Line.CreateBound(XYZ(x0, y0, 0), XYZ(x0, y1, 0)))
        doc.Create.NewDetailCurve(sheet, Line.CreateBound(XYZ(x1, y0, 0), XYZ(x1, y1, 0)))


# ============ ПОИСК КВАДРАТОВ ============

def find_squares_by_edges(bitmap):
    w = bitmap.Width
    h = bitmap.Height
    
    ref_size = 1000
    scale_x = w / ref_size
    scale_y = h / ref_size
    
    square_size_ref = 3
    offset_ref = 467
    
    center_x = w / 2
    center_y = h / 2
    
    offset_x = offset_ref * scale_x
    offset_y = offset_ref * scale_y
    square_size = square_size_ref * ((scale_x + scale_y) / 2)
    
    squares = []
    for dx, dy in [(-1, -1), (1, -1), (-1, 1), (1, 1)]:
        cx = center_x + dx * offset_x
        cy = center_y + dy * offset_y
        squares.append({'x': int(cx), 'y': int(cy), 'size': int(square_size), 'score': 1.0})
    
    return squares


def enhance_bitmap(bitmap):
    w = bitmap.Width
    h = bitmap.Height
    
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
    
    exclude_margin = max(80, min(w, h) // 12)
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
    
    if max_val - min_val < 10:
        min_val = max(0, min_val - 5)
        max_val = min(255, max_val + 5)
    
    if max_val > min_val:
        scale = 255.0 / (max_val - min_val)
        for y in range(h):
            for x in range(w):
                v = enhanced[y][x]
                enhanced[y][x] = int((v - min_val) * scale)
    
    result = Bitmap(w, h, PixelFormat.Format32bppArgb)
    for y in range(h):
        for x in range(w):
            v = min(255, max(0, enhanced[y][x]))
            result.SetPixel(x, y, Color.FromArgb(255, v, v, v))
    
    return result


def find_content_bounds_3pass(bitmap, exclusion_zones=None):
    if exclusion_zones is None:
        exclusion_zones = []
    
    w = bitmap.Width
    h = bitmap.Height
    
    def is_content(b, bg_avg, dark_background):
        if dark_background:
            return b > bg_avg + 20
        else:
            return b < bg_avg - 20
    
    def estimate_background():
        bg_samples = []
        step = max(1, min(w, h) // 30)
        
        for x in range(0, w, step):
            if not any(z[0] <= x <= z[2] and z[1] <= 0 <= z[3] for z in exclusion_zones):
                c = bitmap.GetPixel(x, 0)
                bg_samples.append(max(c.R, c.G, c.B))
            if not any(z[0] <= x <= z[2] and z[1] <= h-1 <= z[3] for z in exclusion_zones):
                c = bitmap.GetPixel(x, h-1)
                bg_samples.append(max(c.R, c.G, c.B))
        
        for y in range(0, h, step):
            if not any(z[0] <= 0 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                c = bitmap.GetPixel(0, y)
                bg_samples.append(max(c.R, c.G, c.B))
            if not any(z[0] <= w-1 <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                c = bitmap.GetPixel(w-1, y)
                bg_samples.append(max(c.R, c.G, c.B))
        
        if not bg_samples:
            return 128, True
        
        bg_avg = sum(bg_samples) / len(bg_samples)
        return bg_avg, bg_avg < 100
    
    bg_avg, dark_background = estimate_background()
    corner_size = 80
    
    # Проход 1
    step1 = 100
    min_x1, max_x1 = w, 0
    min_y1, max_y1 = h, 0
    found1 = False
    
    for y in range(0, h, step1):
        for x in range(0, w, step1):
            if any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                continue
            if (x < corner_size and y < corner_size) or \
               (x > w - corner_size and y < corner_size) or \
               (x < corner_size and y > h - corner_size) or \
               (x > w - corner_size and y > h - corner_size):
                continue
            
            b = max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
            if is_content(b, bg_avg, dark_background):
                min_x1, max_x1 = min(min_x1, x), max(max_x1, x)
                min_y1, max_y1 = min(min_y1, y), max(max_y1, y)
                found1 = True
    
    if not found1:
        step2 = 25
        for y in range(0, h, step2):
            for x in range(0, w, step2):
                if any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                    continue
                if (x < corner_size and y < corner_size) or \
                   (x > w - corner_size and y < corner_size) or \
                   (x < corner_size and y > h - corner_size) or \
                   (x > w - corner_size and y > h - corner_size):
                    continue
                
                b = max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
                if is_content(b, bg_avg, dark_background):
                    min_x1, max_x1 = min(min_x1, x), max(max_x1, x)
                    min_y1, max_y1 = min(min_y1, y), max(max_y1, y)
                    found1 = True
        
        if not found1:
            return None
    
    # Проход 2
    expand1 = step1
    sx1, sx2 = max(0, min_x1 - expand1), min(w-1, max_x1 + expand1)
    sy1, sy2 = max(0, min_y1 - expand1), min(h-1, max_y1 + expand1)
    
    step2 = 25
    min_x2, max_x2 = w, 0
    min_y2, max_y2 = h, 0
    found2 = False
    
    for y in range(sy1, sy2 + 1, step2):
        for x in range(sx1, sx2 + 1, step2):
            if any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                continue
            b = max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
            if is_content(b, bg_avg, dark_background):
                min_x2, max_x2 = min(min_x2, x), max(max_x2, x)
                min_y2, max_y2 = min(min_y2, y), max(max_y2, y)
                found2 = True
    
    if not found2:
        min_x2, max_x2 = min_x1, max_x1
        min_y2, max_y2 = min_y1, max_y1
    
    # Проход 3
    step3 = 5
    margin = 30
    min_x3, max_x3 = w, 0
    min_y3, max_y3 = h, 0
    found3 = False
    
    px1, px2 = max(0, min_x2 - margin), min(w-1, max_x2 + margin)
    py1, py2 = max(0, min_y2 - margin), min(h-1, max_y2 + margin)
    
    for y in range(py1, py2 + 1, step3):
        for x in range(px1, px2 + 1, step3):
            if x > px1 + margin and x < px2 - margin and y > py1 + margin and y < py2 - margin:
                continue
            if any(z[0] <= x <= z[2] and z[1] <= y <= z[3] for z in exclusion_zones):
                continue
            b = max(bitmap.GetPixel(x, y).R, bitmap.GetPixel(x, y).G, bitmap.GetPixel(x, y).B)
            if is_content(b, bg_avg, dark_background):
                min_x3, max_x3 = min(min_x3, x), max(max_x3, x)
                min_y3, max_y3 = min(min_y3, y), max(max_y3, y)
                found3 = True
    
    if not found3:
        return {'min_x': min_x2, 'min_y': min_y2, 'max_x': max_x2, 'max_y': max_y2}
    else:
        return {'min_x': min_x3, 'min_y': min_y3, 'max_x': max_x3, 'max_y': max_y3}


# ============ ИЗМЕРЕНИЕ ВИДА ============

def measure_view_with_calibration(view, stats=None):
    temp_sheet = None
    temp_dir = None
    bitmap = None
    
    try:
        view_id = view.Id.IntegerValue
        if view_id in VIEW_SIZE_CACHE:
            return VIEW_SIZE_CACHE[view_id]
        
        if not doc.IsModifiable:
            return None
        
        start_total = time.time()
        
        sub_t = SubTransaction(doc)
        sub_t.Start()
        
        try:
            temp_sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
            temp_sheet.SheetNumber = "CAL_" + str(view.Id.IntegerValue)
            temp_sheet.Name = sanitize_sheet_name("Cal_" + view.Name[:25])
            
            for elem in FilteredElementCollector(doc, temp_sheet.Id).WhereElementIsNotElementType().ToElements():
                doc.Delete(elem.Id)
            
            doc.Regenerate()
            
            sq_ft = CALIBRATION_SQUARE_MM / 304.8
            half_ft = sq_ft / 2
            grid_ft = 700.0 / 304.8
            
            for dx, dy in [(-1, -1), (1, -1), (-1, 1), (1, 1)]:
                center = XYZ(dx * grid_ft, dy * grid_ft, 0)
                draw_thick_square(doc, temp_sheet, center, half_ft)
            
            Viewport.Create(doc, temp_sheet.Id, view.Id, XYZ(0, 0, 0))
            doc.Regenerate()
            sub_t.Commit()
        
        except Exception as ex:
            if sub_t.HasStarted() and not sub_t.HasEnded():
                sub_t.RollBack()
            raise ex
        
        temp_dir = Path.Combine(Path.GetTempPath(), "revit_cal_" + str(System.Guid.NewGuid()))
        System.IO.Directory.CreateDirectory(temp_dir)
        
        opts = ImageExportOptions()
        opts.ZoomType = ZoomFitType.FitToPage
        opts.PixelSize = EXPORT_PIXELS
        opts.FilePath = Path.Combine(temp_dir, "cal.png")
        opts.HLRandWFViewsFileType = ImageFileType.PNG
        opts.ExportRange = ExportRange.SetOfViews
        opts.ImageResolution = ImageResolution.DPI_150
        opts.SetViewsAndSheets([temp_sheet.Id])
        
        try:
            doc.ExportImage(opts)
        except Exception:
            return None
        
        time.sleep(0.3)
        
        png_files = System.IO.Directory.GetFiles(temp_dir, "*.png")
        if not png_files:
            return None
        
        bitmap = Bitmap(png_files[0])
        
        if bitmap.Width < 100 or bitmap.Height < 100:
            return None
        
        squares = find_squares_by_edges(bitmap)
        
        if not squares or len(squares) < 2:
            return None
        
        mm_per_px = MM_PER_PX
        
        exclusion_zones = []
        margin_px = 2
        
        for sq in squares:
            exclusion_zones.append((
                max(0, sq['x'] - margin_px),
                max(0, sq['y'] - margin_px),
                min(bitmap.Width - 1, sq['x'] + margin_px),
                min(bitmap.Height - 1, sq['y'] + margin_px)
            ))
        
        enhanced = enhance_bitmap(bitmap)
        content_bounds = find_content_bounds_3pass(enhanced, exclusion_zones)
        enhanced.Dispose()
        
        if not content_bounds:
            return None
        
        content_w_px = content_bounds['max_x'] - content_bounds['min_x'] + 1
        content_h_px = content_bounds['max_y'] - content_bounds['min_y'] + 1
        
        content_w_mm = content_w_px * mm_per_px + 2 * PADDING_MM
        content_h_mm = content_h_px * mm_per_px + 2 * PADDING_MM
        
        if content_w_mm < 10 or content_h_mm < 10:
            return None
        
        VIEW_SIZE_CACHE[view_id] = (content_w_mm, content_h_mm)
        
        if stats is not None and 'times' in stats:
            stats['times'].append(time.time() - start_total)
        
        return (content_w_mm, content_h_mm)
    
    except Exception:
        return None
    
    finally:
        if bitmap:
            try:
                bitmap.Dispose()
            except Exception:
                pass
        
        if temp_sheet:
            try:
                el = doc.GetElement(temp_sheet.Id)
                if el is not None and doc.IsModifiable:
                    st = SubTransaction(doc)
                    st.Start()
                    try:
                        doc.Delete(temp_sheet.Id)
                        st.Commit()
                    except Exception:
                        st.RollBack()
            except Exception:
                pass
        
        if temp_dir:
            try:
                System.IO.Directory.Delete(temp_dir, True)
            except Exception:
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
        new_free = []
        for (fx, fy, fw, fh) in self.free_rects:
            if not (x >= fx + fw or x + w <= fx or y >= fy + fh or y + h <= fy):
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
    
    strategies = [lambda r: max(r[0], r[1]), lambda r: r[0]*r[1], lambda r: r[0], lambda r: r[1]]
    heuristics = ['area', 'short_side', 'long_side']
    best_result, best_count, best_fill = [], 0, 0
    
    for key_func in strategies:
        sorted_rects = sorted(rects, key=key_func, reverse=True)
        for heuristic in heuristics:
            packer = MaxRectsPacker(bin_width, bin_height, occupied_rects)
            for w, h, item_id in sorted_rects:
                if w <= bin_width and h <= bin_height:
                    packer.add_rect(w, h, item_id, heuristic)
            
            placements = packer.get_placements()
            if len(placements) == 0:
                continue
            
            fill = sum(w*h for _, _, w, h, _ in placements) / (bin_width * bin_height)
            if len(placements) > best_count or (len(placements) == best_count and fill > best_fill):
                best_result, best_count, best_fill = placements, len(placements), fill
    
    return best_result


def find_best_fill(rects, bin_w, bin_h, occupied_rects=None):
    if not rects:
        return []
    
    sorted_rects = sorted(rects, key=lambda r: r[0]*r[1], reverse=True)
    best_placements, best_count, best_fill = [], 0, 0
    
    total_rects = len(sorted_rects)
    step = max(1, total_rects // 15)
    
    for start in range(0, total_rects, step):
        for count in range(1, total_rects - start + 1):
            subset = sorted_rects[start:start + count]
            if sum(w*h for w, h, _ in subset) > bin_w * bin_h * 1.3:
                break
            
            placements = pack_rectangles(subset, bin_w, bin_h, occupied_rects)
            if len(placements) == 0:
                continue
            
            fill = sum(w*h for _, _, w, h, _ in placements) / (bin_w * bin_h)
            if len(placements) > best_count or (len(placements) == best_count and fill > best_fill):
                best_placements, best_count, best_fill = placements, len(placements), fill
    
    return best_placements


# ============ АНИМАЦИИ ============

def get_revit_theme_color():
    try:
        bg_color = app.ApplicationTheme.BackgroundColor
        return Color.FromArgb(bg_color.R, bg_color.G, bg_color.B)
    except Exception:
        return Color.FromArgb(30, 30, 35)


def is_dark_theme(bg_color):
    return (bg_color.R + bg_color.G + bg_color.B) / 3 < 128


def get_text_color(bg_color):
    return Color.White if is_dark_theme(bg_color) else Color.Black


def get_subtext_color(bg_color):
    return Color.FromArgb(180, 180, 180) if is_dark_theme(bg_color) else Color.FromArgb(60, 60, 60)


def show_search_animation(message="Поиск видов в проекте..."):
    bg_color = get_revit_theme_color()
    dark_bg = Color.FromArgb(
        max(0, bg_color.R - 20), max(0, bg_color.G - 20), max(0, bg_color.B - 20)
    )
    
    form = Form()
    form.Text = "Размещение видов"
    form.Size = Size(350, 220)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.BackColor = dark_bg
    form.ControlBox = False
    
    anim_panel = Panel()
    anim_panel.Location = Point(0, 15)
    anim_panel.Size = Size(350, 70)
    anim_panel.BackColor = Color.Transparent
    form.Controls.Add(anim_panel)
    
    # Состояние пульсации через список (IronPython не поддерживает nonlocal)
    pulse_state = [0]
    pulse_dir = [1]
    
    def draw_pulse(state):
        if form.IsDisposed:
            return
        try:
            g = anim_panel.CreateGraphics()
            g.Clear(dark_bg)
            cx, cy = 175, 35
            r = 12 + state * 4
            alpha = 200 - state * 40
            color = Color.FromArgb(alpha, 100, 180, 255)
            brush = SolidBrush(color)
            g.FillEllipse(brush, cx - r, cy - r, r*2, r*2)
            g.Dispose()
        except Exception:
            pass
    
    def pulse_tick(s, e):
        pulse_state[0] += pulse_dir[0]
        if pulse_state[0] >= 3:
            pulse_dir[0] = -1
        elif pulse_state[0] <= 0:
            pulse_dir[0] = 1
        draw_pulse(pulse_state[0])
    
    timer = Timer()
    timer.Interval = 150
    timer.Tick += pulse_tick
    timer.Start()
    
    label = Label()
    label.Text = message
    label.Location = Point(0, 95)
    label.Size = Size(350, 30)
    label.Font = Font("Arial", 10, FontStyle.Regular)
    label.ForeColor = get_text_color(bg_color)
    label.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(label)
    
    sublabel = Label()
    sublabel.Text = ""
    sublabel.Location = Point(0, 130)
    sublabel.Size = Size(350, 25)
    sublabel.Font = Font("Arial", 8, FontStyle.Regular)
    sublabel.ForeColor = get_subtext_color(bg_color)
    sublabel.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(sublabel)
    
    form.Show()
    form.Refresh()
    draw_pulse(0)
    
    return form, label, sublabel, timer


def update_search_text(form, label, sublabel, text):
    if form and not form.IsDisposed and label:
        try:
            label.Text = text
            form.Refresh()
        except Exception:
            pass


def close_search_animation(form, timer=None):
    if timer:
        try:
            timer.Stop()
        except Exception:
            pass
    if form and not form.IsDisposed:
        try:
            form.Close()
        except Exception:
            pass


def show_progress_panel(message="Размещение видов"):
    bg_color = get_revit_theme_color()
    dark_bg = Color.FromArgb(
        max(0, bg_color.R - 20), max(0, bg_color.G - 20), max(0, bg_color.B - 20)
    )
    panel_bg = Color.FromArgb(
        max(0, bg_color.R - 10), max(0, bg_color.G - 10), max(0, bg_color.B - 10)
    )
    
    form = Form()
    form.Text = message
    form.Size = Size(500, 300)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.BackColor = dark_bg
    form.ControlBox = False
    
    title = Label()
    title.Text = message
    title.Location = Point(0, 15)
    title.Size = Size(500, 30)
    title.Font = Font("Arial", 12, FontStyle.Bold)
    title.ForeColor = get_text_color(bg_color)
    title.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(title)
    
    grid = Panel()
    grid.Location = Point(30, 55)
    grid.Size = Size(440, 140)
    grid.BackColor = panel_bg
    form.Controls.Add(grid)
    
    stage = Label()
    stage.Text = "Подготовка..."
    stage.Location = Point(0, 205)
    stage.Size = Size(500, 20)
    stage.Font = Font("Arial", 9)
    stage.ForeColor = get_subtext_color(bg_color)
    stage.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(stage)
    
    progress = ProgressBar()
    progress.Location = Point(30, 235)
    progress.Size = Size(440, 15)
    progress.Style = ProgressBarStyle.Marquee
    progress.MarqueeAnimationSpeed = 25
    form.Controls.Add(progress)
    
    pct = Label()
    pct.Text = ""
    pct.Location = Point(0, 255)
    pct.Size = Size(500, 20)
    pct.Font = Font("Arial", 8, FontStyle.Bold)
    pct.ForeColor = Color.FromArgb(100, 200, 100)
    pct.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(pct)
    
    form.Show()
    form.Refresh()
    
    return form, grid, stage, progress, pct


def update_grid(form, grid, views_data, current_idx, completed):
    if form.IsDisposed:
        return
    grid.Controls.Clear()
    cols, size, gap = 6, 55, 8
    
    for i, (view, group) in enumerate(views_data):
        col, row = i % cols, i // cols
        x, y = 10 + col * (size + gap), 10 + row * (size + gap)
        
        box = Panel()
        box.Location = Point(x, y)
        box.Size = Size(size, size)
        box.BorderStyle = BorderStyle.FixedSingle
        
        if i < completed:
            box.BackColor = Color.FromArgb(60, 180, 75)
        elif i == current_idx:
            box.BackColor = Color.FromArgb(255, 180, 30)
        else:
            box.BackColor = Color.FromArgb(60, 60, 68)
        
        lbl = Label()
        lbl.Text = str(i + 1)
        lbl.Location = Point(0, 0)
        lbl.Size = Size(size, size)
        lbl.Font = Font("Arial", 9, FontStyle.Bold)
        lbl.ForeColor = Color.White
        lbl.TextAlign = ContentAlignment.MiddleCenter
        box.Controls.Add(lbl)
        grid.Controls.Add(box)
    
    form.Refresh()


def update_progress_info(form, stage, progress, pct, stage_text, percent_text=""):
    if form.IsDisposed:
        return
    stage.Text = stage_text
    if percent_text:
        pct.Text = percent_text
        progress.Style = ProgressBarStyle.Blocks
        try:
            val = int(percent_text.replace("%", ""))
            progress.Value = min(100, max(0, val))
        except Exception:
            pass
    form.Refresh()


def show_results_form(sheet_count, total_placed, skipped_views, failed, total_time):
    bg_color = get_revit_theme_color()
    dark_bg = Color.FromArgb(
        max(0, bg_color.R - 20), max(0, bg_color.G - 20), max(0, bg_color.B - 20)
    )
    txt_color = get_text_color(bg_color)
    sub_color = get_subtext_color(bg_color)
    
    form = Form()
    form.Text = "Результаты размещения"
    form.Size = Size(450, 400)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.BackColor = dark_bg
    
    y = 20
    
    title = Label()
    title.Text = "✅ Размещение завершено!"
    title.Location = Point(0, y)
    title.Size = Size(450, 30)
    title.Font = Font("Arial", 12, FontStyle.Bold)
    title.ForeColor = Color.FromArgb(100, 200, 100)
    title.TextAlign = ContentAlignment.MiddleCenter
    form.Controls.Add(title)
    y += 45
    
    stats = [
        ("Создано листов:", str(sheet_count)),
        ("Размещено видов:", str(total_placed)),
    ]
    if skipped_views:
        stats.append(("Пропущено видов:", str(len(skipped_views))))
    if failed:
        stats.append(("Не проанализировано:", str(len(failed))))
    if CREATION_WARNINGS:
        stats.append(("Предупреждений:", str(len(CREATION_WARNINGS))))
    stats.append(("Общее время:", str(round(total_time, 1)) + " сек"))
    
    for label_text, value in stats:
        lbl = Label()
        lbl.Text = label_text
        lbl.Location = Point(50, y)
        lbl.Size = Size(200, 25)
        lbl.Font = Font("Arial", 10)
        lbl.ForeColor = sub_color
        lbl.TextAlign = ContentAlignment.MiddleRight
        form.Controls.Add(lbl)
        
        val = Label()
        val.Text = value
        val.Location = Point(255, y)
        val.Size = Size(120, 25)
        val.Font = Font("Arial", 10, FontStyle.Bold)
        val.ForeColor = txt_color
        val.TextAlign = ContentAlignment.MiddleLeft
        form.Controls.Add(val)
        y += 30
    
    y += 15
    
    # Кнопка копирования логов
    def copy_logs(s, e):
        if LOG_COLLECTOR:
            log_text = "\n".join(LOG_COLLECTOR)
            try:
                Clipboard.SetText(log_text)
                MessageBox.Show("Логи скопированы в буфер обмена!", "Логи")
            except Exception:
                log_form = Form()
                log_form.Text = "Логи выполнения"
                log_form.Size = Size(600, 400)
                log_form.StartPosition = FormStartPosition.CenterScreen
                
                log_box = TextBox()
                log_box.Multiline = True
                log_box.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
                log_box.Text = log_text
                log_box.Dock = System.Windows.Forms.DockStyle.Fill
                log_box.Font = Font("Consolas", 9)
                log_box.ReadOnly = True
                log_form.Controls.Add(log_box)
                
                log_form.ShowDialog()
        else:
            MessageBox.Show("Логи пусты", "Логи")
    
    btn_logs = Button()
    btn_logs.Text = "📋 Копировать логи"
    btn_logs.Location = Point(60, y)
    btn_logs.Size = Size(140, 35)
    btn_logs.Font = Font("Arial", 9, FontStyle.Bold)
    btn_logs.BackColor = Color.FromArgb(60, 80, 120)
    btn_logs.ForeColor = Color.White
    btn_logs.FlatStyle = FlatStyle.Flat
    btn_logs.Click += copy_logs
    form.Controls.Add(btn_logs)
    
    btn_ok = Button()
    btn_ok.Text = "OK"
    btn_ok.Location = Point(240, y)
    btn_ok.Size = Size(120, 35)
    btn_ok.Font = Font("Arial", 10, FontStyle.Bold)
    btn_ok.BackColor = Color.FromArgb(60, 180, 75)
    btn_ok.ForeColor = Color.White
    btn_ok.FlatStyle = FlatStyle.Flat
    btn_ok.Click += lambda s, e: form.Close()
    form.Controls.Add(btn_ok)
    
    form.ShowDialog()


# ============ ФОРМА НАСТРОЙКИ ГРУПП ============

def show_group_config_form():
    bg_color = get_revit_theme_color()
    txt_color = get_text_color(bg_color)
    
    form = Form()
    form.Text = "Настройка групп систем"
    form.Size = Size(860, 410)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.BackColor = bg_color
    
    result = {'grouping': None, 'show_logs': False}
    y = 10
    
    title = Label()
    title.Text = "Настройка групп систем (каждая строка = одна группа листов)"
    title.Location = Point(10, y)
    title.Size = Size(830, 25)
    title.Font = Font("Arial", 10, FontStyle.Bold)
    title.ForeColor = txt_color
    form.Controls.Add(title)
    y += 30
    
    panel = Panel()
    panel.Location = Point(10, y)
    panel.Size = Size(830, 220)
    panel.BorderStyle = BorderStyle.FixedSingle
    panel.AutoScroll = True
    panel.BackColor = Color.FromArgb(
        max(0, bg_color.R - 15), max(0, bg_color.G - 15), max(0, bg_color.B - 15)
    )
    form.Controls.Add(panel)
    
    group_combos = []
    group_name_boxes = []
    
    inner_y = 10
    for group_idx in range(5):
        group_label = Label()
        group_label.Text = "Группа " + str(group_idx + 1) + ":"
        group_label.Location = Point(10, inner_y)
        group_label.Size = Size(60, 20)
        group_label.Font = Font("Arial", 8, FontStyle.Regular)
        group_label.ForeColor = txt_color
        panel.Controls.Add(group_label)
        
        row_combos = []
        combo_x = 75
        for i in range(10):
            cb = ComboBox()
            cb.Location = Point(combo_x, inner_y)
            cb.Size = Size(50, 20)
            cb.Font = Font("Arial", 8)
            cb.DropDownStyle = ComboBoxStyle.DropDownList
            cb.Items.Add("")
            for prefix in ALL_PREFIXES:
                cb.Items.Add(prefix)
            
            if group_idx < len(DEFAULT_GROUPING):
                default_prefixes = list(DEFAULT_GROUPING.values())[group_idx]
                if i < len(default_prefixes):
                    cb.SelectedItem = default_prefixes[i]
                else:
                    cb.SelectedIndex = 0
            else:
                cb.SelectedIndex = 0
            
            panel.Controls.Add(cb)
            row_combos.append(cb)
            combo_x += 50
            if i < 9:
                plus_label = Label()
                plus_label.Text = "+"
                plus_label.Location = Point(combo_x + 2, inner_y)
                plus_label.Size = Size(10, 20)
                plus_label.Font = Font("Arial", 8)
                plus_label.ForeColor = txt_color
                panel.Controls.Add(plus_label)
                combo_x += 8
        
        group_combos.append(row_combos)
        
        name_label = Label()
        name_label.Text = "Название:"
        name_label.Location = Point(combo_x + 8, inner_y)
        name_label.Size = Size(60, 20)
        name_label.Font = Font("Arial", 8)
        name_label.ForeColor = txt_color
        panel.Controls.Add(name_label)
        
        name_box = TextBox()
        name_box.Location = Point(combo_x + 68, inner_y)
        name_box.Size = Size(100, 20)
        name_box.Font = Font("Arial", 8)
        if group_idx < len(GROUP_NAMES_DEFAULT):
            name_box.Text = GROUP_NAMES_DEFAULT[group_idx]
        panel.Controls.Add(name_box)
        group_name_boxes.append(name_box)
        inner_y += 25
    
    y += 230
    
    presets_label = Label()
    presets_label.Text = "Быстрые пресеты:"
    presets_label.Location = Point(10, y)
    presets_label.Size = Size(120, 20)
    presets_label.Font = Font("Arial", 8, FontStyle.Bold)
    presets_label.ForeColor = txt_color
    form.Controls.Add(presets_label)
    
    def apply_preset(preset_name):
        presets = {
            "П+В, ДП+ДВ, А+У": [
                (["П", "ПЕ", "В", "ВЕ"], "П-В"),
                (["ДП", "ДПЕ", "ДВ", "ДВЕ"], "ДП-ДВ"),
                (["А", "У"], "А-У"), ([], ""), ([], ""),
            ],
            "П, В, ДП+ДВ, А+У": [
                (["П", "ПЕ"], "П"), (["В", "ВЕ"], "В"),
                (["ДП", "ДПЕ", "ДВ", "ДВЕ"], "ДП-ДВ"),
                (["А", "У"], "А-У"), ([], ""),
            ],
            "П, В, ДП, ДВ, А, У": [
                (["П", "ПЕ"], "П"), (["В", "ВЕ"], "В"),
                (["ДП", "ДПЕ"], "ДП"), (["ДВ", "ДВЕ"], "ДВ"),
                (["А", "У"], "А-У"),
            ],
        }
        if preset_name in presets:
            for g_idx, (prefixes, name) in enumerate(presets[preset_name]):
                for i in range(10):
                    if i < len(prefixes):
                        group_combos[g_idx][i].SelectedItem = prefixes[i]
                    else:
                        group_combos[g_idx][i].SelectedIndex = 0
                group_name_boxes[g_idx].Text = name
    
    preset_x = 140
    for preset_name in ["П+В, ДП+ДВ, А+У", "П, В, ДП+ДВ, А+У", "П, В, ДП, ДВ, А, У"]:
        btn = Button()
        btn.Text = preset_name
        btn.Location = Point(preset_x, y - 2)
        btn.Size = Size(180, 25)
        btn.Font = Font("Arial", 8)
        btn.Click += lambda s, e, pn=preset_name: apply_preset(pn)
        form.Controls.Add(btn)
        preset_x += 190
    
    y += 30
    
    chk_logs = CheckBox()
    chk_logs.Text = "Показывать логи в терминале"
    chk_logs.Location = Point(10, y)
    chk_logs.Size = Size(250, 20)
    chk_logs.Font = Font("Arial", 8)
    chk_logs.ForeColor = txt_color
    chk_logs.Checked = False
    form.Controls.Add(chk_logs)
    
    def on_ok(s, e):
        new_grouping = OrderedDict()
        for g_idx in range(5):
            prefixes = []
            for cb in group_combos[g_idx]:
                if cb.SelectedItem and cb.SelectedItem.ToString():
                    prefixes.append(cb.SelectedItem.ToString())
            name = group_name_boxes[g_idx].Text.strip()
            if prefixes and name:
                new_grouping[name] = prefixes
        if new_grouping:
            result['grouping'] = new_grouping
            result['show_logs'] = chk_logs.Checked
            form.DialogResult = DialogResult.OK
            form.Close()
        else:
            MessageBox.Show("Настройте хотя бы одну группу!", "Предупреждение")
    
    def on_cancel(s, e):
        form.DialogResult = DialogResult.Cancel
        form.Close()
    
    btn_ok = Button()
    btn_ok.Text = "▶ Продолжить"
    btn_ok.Location = Point(350, y + 5)
    btn_ok.Size = Size(140, 35)
    btn_ok.Font = Font("Arial", 10, FontStyle.Bold)
    btn_ok.BackColor = Color.LightGreen
    btn_ok.Click += on_ok
    form.Controls.Add(btn_ok)
    
    btn_cancel = Button()
    btn_cancel.Text = "По умолчанию"
    btn_cancel.Location = Point(500, y + 5)
    btn_cancel.Size = Size(140, 35)
    btn_cancel.Font = Font("Arial", 9)
    btn_cancel.Click += on_cancel
    form.Controls.Add(btn_cancel)
    
    form.ShowDialog()
    
    if result['grouping'] is not None:
        return result['grouping'], result['show_logs']
    return None, False


# ============ WINFORMS UI ============

def show_placement_form(views_data, sheet_w, sheet_h):
    bg_color = get_revit_theme_color()
    txt_color = get_text_color(bg_color)
    sub_color = get_subtext_color(bg_color)
    
    form = Form()
    form.Text = "Размещение видов на листы v6.1"
    form.Size = Size(680, 650)
    form.StartPosition = FormStartPosition.CenterScreen
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.BackColor = bg_color
    
    result = {'views': None, 'settings': None}
    y = 10
    
    title = Label()
    title.Text = "Выберите виды для размещения на листах"
    title.Location = Point(10, y)
    title.Size = Size(640, 25)
    title.Font = Font("Arial", 10, FontStyle.Bold)
    title.ForeColor = txt_color
    form.Controls.Add(title)
    y += 30
    
    info = Label()
    info.Text = "Лист: " + str(int(sheet_w)) + "x" + str(int(sheet_h)) + " мм | Видов: " + str(len(views_data))
    info.Location = Point(10, y)
    info.Size = Size(640, 20)
    info.Font = Font("Arial", 8)
    info.ForeColor = sub_color
    form.Controls.Add(info)
    y += 25
    
    panel = Panel()
    panel.Location = Point(10, y)
    panel.Size = Size(640, 280)
    panel.BorderStyle = BorderStyle.FixedSingle
    panel.AutoScroll = True
    panel.BackColor = Color.FromArgb(
        max(0, bg_color.R - 10), max(0, bg_color.G - 10), max(0, bg_color.B - 10)
    )
    form.Controls.Add(panel)
    
    inner_y = 10
    current_group = None
    
    for view, group_name in views_data:
        if group_name != current_group:
            if current_group is not None:
                inner_y += 5
            gl = Label()
            gl.Text = "━━━ " + group_name + " ━━━"
            gl.Location = Point(10, inner_y)
            gl.Size = Size(600, 22)
            gl.Font = Font("Arial", 9, FontStyle.Bold)
            gl.ForeColor = Color.DarkBlue if not is_dark_theme(bg_color) else Color.FromArgb(100, 180, 255)
            panel.Controls.Add(gl)
            inner_y += 27
            current_group = group_name
        
        cb = CheckBox()
        cb.Text = view.Name
        cb.Location = Point(30, inner_y)
        cb.Size = Size(580, 20)
        cb.Font = Font("Arial", 8)
        cb.ForeColor = txt_color
        cb.Tag = view
        cb.Checked = True
        panel.Controls.Add(cb)
        inner_y += 22
    
    y += 290
    
    def select_all(s, e):
        for c in panel.Controls:
            if isinstance(c, CheckBox):
                c.Checked = True
    
    def deselect_all(s, e):
        for c in panel.Controls:
            if isinstance(c, CheckBox):
                c.Checked = False
    
    b1 = Button()
    b1.Text = "✓ Все"
    b1.Location = Point(10, y)
    b1.Size = Size(100, 30)
    b1.Click += select_all
    form.Controls.Add(b1)
    
    b2 = Button()
    b2.Text = "✗ Снять"
    b2.Location = Point(120, y)
    b2.Size = Size(100, 30)
    b2.Click += deselect_all
    form.Controls.Add(b2)
    y += 40
    
    gb = GroupBox()
    gb.Text = "Настройки (мм)"
    gb.Location = Point(10, y)
    gb.Size = Size(640, 105)
    gb.Font = Font("Arial", 8, FontStyle.Bold)
    gb.ForeColor = txt_color
    form.Controls.Add(gb)
    
    # Исправленные координаты для полей
    field_defs = [
        ("Отступ X:", 15, "20", 80),
        ("Отступ Y:", 140, "20", 205),
        ("Зазор X:", 270, "15", 325),
        ("Зазор Y:", 390, "15", 445),
        ("Штамп:", 510, "60", 560),
    ]
    
    textboxes = {}
    for lbl_text, lbl_x, default_val, tb_x in field_defs:
        lbl = Label()
        lbl.Text = lbl_text
        lbl.Location = Point(lbl_x, 25)
        lbl.Size = Size(65, 20)
        lbl.Font = Font("Arial", 8)
        lbl.ForeColor = txt_color
        gb.Controls.Add(lbl)
        
        tb = TextBox()
        tb.Text = default_val
        tb.Location = Point(tb_x, 22)
        tb.Size = Size(45, 20)
        gb.Controls.Add(tb)
        textboxes[lbl_text.replace(":", "")] = tb
    
    lpf = Label()
    lpf.Text = "Префикс:"
    lpf.Location = Point(15, 55)
    lpf.Size = Size(65, 20)
    lpf.Font = Font("Arial", 8)
    lpf.ForeColor = txt_color
    gb.Controls.Add(lpf)
    
    tpf = TextBox()
    tpf.Text = ""
    tpf.Location = Point(80, 52)
    tpf.Size = Size(200, 20)
    gb.Controls.Add(tpf)
    
    y += 115
    
    def ok(s, e):
        sel = [c.Tag for c in panel.Controls if isinstance(c, CheckBox) and c.Checked and c.Tag]
        if not sel:
            MessageBox.Show("Не выбрано!")
            return
        
        vals = {}
        for name in ['Отступ X', 'Отступ Y', 'Зазор X', 'Зазор Y', 'Штамп']:
            try:
                v = int(textboxes[name].Text)
            except Exception:
                MessageBox.Show("Поле '" + name + "' — число!")
                return
            if v < 0:
                MessageBox.Show("Поле '" + name + "' >= 0!")
                return
            vals[name] = v
        
        if vals['Отступ X']*2 >= sheet_w or vals['Отступ Y']*2 >= sheet_h:
            MessageBox.Show("Отступы > лист!")
            return
        
        result['views'] = sel
        result['settings'] = {
            'margin_x': vals['Отступ X'], 'margin_y': vals['Отступ Y'],
            'gap_x': vals['Зазор X'], 'gap_y': vals['Зазор Y'],
            'titleblock_offset': vals['Штамп'], 'sheet_prefix': tpf.Text.strip()
        }
        form.DialogResult = DialogResult.OK
        form.Close()
    
    def cancel(s, e):
        form.DialogResult = DialogResult.Cancel
        form.Close()
    
    bok = Button()
    bok.Text = "▶ Разместить"
    bok.Location = Point(350, y)
    bok.Size = Size(140, 40)
    bok.Font = Font("Arial", 10, FontStyle.Bold)
    bok.BackColor = Color.LightGreen
    bok.Click += ok
    form.Controls.Add(bok)
    
    bc = Button()
    bc.Text = "Отмена"
    bc.Location = Point(500, y)
    bc.Size = Size(120, 40)
    bc.Click += cancel
    form.Controls.Add(bc)
    
    form.ShowDialog()
    return (result['views'], result['settings']) if result['views'] else (None, None)


# ============ MAIN ============

def main():
    global SHOW_LOGS, LOG_COLLECTOR, CREATION_WARNINGS
    LOG_COLLECTOR = []
    CREATION_WARNINGS = []
    
    start_total = time.time()
    
    # Настройка групп
    custom_grouping, show_logs_flag = show_group_config_form()
    SHOW_LOGS = show_logs_flag
    
    if custom_grouping is not None:
        global GROUPING
        GROUPING = custom_grouping
        log_message("📋 Пользовательская группировка:")
        for name, prefixes in GROUPING.items():
            log_message("  " + name + ": " + ", ".join(prefixes))
    else:
        GROUPING = OrderedDict(DEFAULT_GROUPING)
        log_message("📋 Группировка по умолчанию")
    
    # Сбор видов с анимацией
    search_form, search_label, search_sublabel, search_timer = show_search_animation("Сбор видов из проекта...")
    update_search_text(search_form, search_label, search_sublabel, "Поиск 3D видов...")
    
    import shutil
    script_dir = os.path.dirname(os.path.abspath(__file__))
    debug_dir = os.path.join(script_dir, "AutoView_debug")
    
    if os.path.exists(debug_dir):
        for f in os.listdir(debug_dir):
            try:
                fp = os.path.join(debug_dir, f)
                if os.path.isfile(fp):
                    os.unlink(fp)
            except Exception:
                pass
    
    sel_ids = list(uidoc.Selection.GetElementIds())
    if not sel_ids:
        close_search_animation(search_form, search_timer)
        MessageBox.Show("Выберите лист-шаблон!", "Ошибка")
        return
    
    template_sheet = doc.GetElement(sel_ids[0])
    if not isinstance(template_sheet, ViewSheet):
        close_search_animation(search_form, search_timer)
        MessageBox.Show("Не лист!", "Ошибка")
        return
    
    update_search_text(search_form, search_label, search_sublabel, "Чтение параметров листа...")
    frame_info = get_sheet_frame_mm(template_sheet)
    if not frame_info:
        close_search_animation(search_form, search_timer)
        MessageBox.Show("Нет рамки!", "Ошибка")
        return
    
    frame_w, frame_h, frame_min_x, frame_min_y = frame_info
    log_message("📐 Рамка: " + str(int(frame_w)) + "x" + str(int(frame_h)) + " мм")
    
    update_search_text(search_form, search_label, search_sublabel, "Поиск 3D видов...")
    placed_cache = collect_placed_views()
    all_3d = list(FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements())
    
    update_search_text(search_form, search_label, search_sublabel, "Группировка видов...")
    views_data = []
    for view in all_3d:
        if view.IsTemplate or view.Id in placed_cache:
            continue
        prefix, _ = extract_prefix_and_numbers(view.Name)
        views_data.append((view, get_group_for_prefix(prefix)))
    
    log_message("🔍 3D виды: всего " + str(len(all_3d)) + ", доступно " + str(len(views_data)))
    
    update_search_text(search_form, search_label, search_sublabel, "Найдено: " + str(len(views_data)) + " видов")
    time.sleep(0.5)
    close_search_animation(search_form, search_timer)
    
    if not views_data:
        MessageBox.Show("Нет видов!", "Информация")
        return
    
    views_data.sort(key=lambda x: (
        list(GROUPING.keys()).index(x[1]) if x[1] in GROUPING else 999,
        get_view_sort_key(x[0])
    ))
    
    # Выбор видов
    selected_views, settings = show_placement_form(views_data, frame_w, frame_h)
    if not selected_views:
        return
    
    log_message("✅ Выбрано видов: " + str(len(selected_views)))
    
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
    
    # Анализ с прогресс-панелью
    progress_form, grid, stage, progress, pct = show_progress_panel("Анализ видов")
    all_selected = [(v, get_group_for_prefix(extract_prefix_and_numbers(v.Name)[0])) for v in selected_views]
    update_grid(progress_form, grid, all_selected, 0, 0)
    update_progress_info(progress_form, stage, progress, pct,
                         "Начинаю анализ " + str(len(selected_views)) + " видов...", "")
    
    view_sizes = {}
    failed = []
    analysis_stats = {'times': []}
    
    analysis_t = Transaction(doc, "Analyze views")
    analysis_t.Start()
    
    try:
        for idx, v in enumerate(selected_views):
            update_grid(progress_form, grid, all_selected, idx, idx)
            pct_text = str(int((idx + 1) / len(selected_views) * 90)) + "%"
            update_progress_info(progress_form, stage, progress, pct,
                                 "Анализ: " + get_view_short_name(v.Name), pct_text)
            
            log_message("  📐 Анализ: " + v.Name)
            size = measure_view_with_calibration(v, stats=analysis_stats)
            
            if size:
                view_sizes[v.Id] = size
                update_grid(progress_form, grid, all_selected, idx, idx + 1)
                log_message("    ✅ " + str(int(size[0])) + "x" + str(int(size[1])) + " мм")
            else:
                failed.append(v)
                update_grid(progress_form, grid, all_selected, idx, idx + 1)
                log_message("    ❌ Не удалось определить размер")
        
        analysis_t.RollBack()
    
    except Exception as ex:
        if analysis_t.HasStarted() and not analysis_t.HasEnded():
            analysis_t.RollBack()
        log_message("  ❌ Ошибка анализа: " + str(ex))
        raise ex
    
    if not view_sizes:
        progress_form.Close()
        MessageBox.Show("Нет размеров!", "Ошибка")
        return
    
    # Размещение
    update_progress_info(progress_form, stage, progress, pct, "Размещение на листах...", "92%")
    
    stamp_width = STAMP_WIDTH_MM
    stamp_height = tb_off
    avail_w = int(frame_w - 2 * mx)
    avail_h = int(frame_h - my - stamp_height)
    if avail_w <= 0 or avail_h <= 0:
        avail_w = int(frame_w - 2 * mx)
        avail_h = int(frame_h - my)
    
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    used_numbers = {s.SheetNumber for s in all_sheets}
    
    tmpl_num = template_sheet.SheetNumber
    match = re.search(r'(\d+)$', tmpl_num)
    num_prefix = tmpl_num[:match.start()] if match else tmpl_num + '_'
    next_num = int(match.group(1)) + 1 if match else 1
    
    class SheetNumberGenerator:
        def __init__(self):
            self.p = num_prefix
            self.c = next_num
            self.u = used_numbers
        def get_next(self):
            while True:
                cand = self.p + str(self.c)
                self.c += 1
                if cand not in self.u:
                    self.u.add(cand)
                    return cand
    
    num_gen = SheetNumberGenerator()
    
    total_placed = 0
    sheet_count = 0
    skipped_views = []
    
    t = Transaction(doc, "Place views")
    t.Start()
    try:
        for group_name in GROUPING:
            if group_name not in groups:
                continue
            
            remaining = list(groups[group_name])
            log_message("📋 Группа: " + group_name + " (" + str(len(remaining)) + " видов)")
            
            while remaining:
                fit_views, fit_rects, oversize_views, unknown_views = [], [], [], []
                
                for v in remaining:
                    size = view_sizes.get(v.Id)
                    if not size:
                        unknown_views.append(v)
                        continue
                    
                    wg = int(math.ceil(size[0] + gx))
                    hg = int(math.ceil(size[1] + gy))
                    
                    if wg <= avail_w and hg <= avail_h:
                        fit_views.append(v)
                        fit_rects.append((wg, hg, v.Id))
                    else:
                        oversize_views.append(v)
                
                done = set(v.Id for v in fit_views + oversize_views + unknown_views)
                remaining = [v for v in remaining if v.Id not in done]
                
                for v in oversize_views:
                    vw, vh = view_sizes[v.Id]
                    ns = create_sheet(template_sheet)
                    ns.SheetNumber = num_gen.get_next()
                    clear_sheet_viewports(ns)
                    
                    left_x = mx + max(0, (avail_w - vw) // 2)
                    left_y = stamp_height + max(0, (avail_h - vh) // 2)
                    cx = left_x + vw / 2.0
                    cy = left_y + vh / 2.0
                    
                    try:
                        Viewport.Create(doc, ns.Id, v.Id, XYZ((frame_min_x + cx) / 304.8, (frame_min_y + cy) / 304.8, 0))
                        
                        p = sheet_prefix if sheet_prefix else group_name
                        # Проверка имени на запрещённые символы
                        safe_name = sanitize_sheet_name(p + ": " + get_view_short_name(v.Name))
                        ns.Name = safe_name
                        sheet_count += 1
                        total_placed += 1
                        log_message("  📄 " + ns.SheetNumber + ": " + v.Name)
                    except Exception as ex:
                        warn_creation("Не удалось создать лист для вида '" + v.Name + "': " + str(ex))
                
                for v in unknown_views:
                    skipped_views.append(v)
                    log_message("  ⚠️ Пропущен: " + v.Name)
                if not fit_rects:
                    continue
                
                placements = find_best_fill(fit_rects, avail_w, avail_h)
                
                if not placements:
                    for v in fit_views:
                        vw, vh = view_sizes[v.Id]
                        ns = create_sheet(template_sheet)
                        ns.SheetNumber = num_gen.get_next()
                        clear_sheet_viewports(ns)
                        
                        left_x = mx + max(0, (avail_w - vw) // 2)
                        left_y = stamp_height + max(0, (avail_h - vh) // 2)
                        cx = left_x + vw / 2.0
                        cy = left_y + vh / 2.0
                        
                        try:
                            Viewport.Create(doc, ns.Id, v.Id, XYZ((frame_min_x + cx) / 304.8, (frame_min_y + cy) / 304.8, 0))
                            safe_name = sanitize_sheet_name((sheet_prefix if sheet_prefix else group_name) + ": " + get_view_short_name(v.Name))
                            ns.Name = safe_name
                            sheet_count += 1
                            total_placed += 1
                        except Exception as ex:
                            warn_creation("Не удалось создать лист: " + str(ex))
                    continue
                
                pids = set(p[4] for p in placements)
                pv = [v for v in fit_views if v.Id in pids]
                
                ns = create_sheet(template_sheet)
                ns.SheetNumber = num_gen.get_next()
                clear_sheet_viewports(ns)
                
                for x, y, w, h, vid in placements:
                    v = next((vv for vv in pv if vv.Id == vid), None)
                    if not v:
                        continue
                    vw, vh = view_sizes[vid]
                    
                    left_x = mx + x
                    left_y = stamp_height + y
                    cx = left_x + vw / 2.0
                    cy = left_y + vh / 2.0
                    
                    Viewport.Create(doc, ns.Id, v.Id, XYZ((frame_min_x + cx) / 304.8, (frame_min_y + cy) / 304.8, 0))
                
                sn = [get_view_short_name(v.Name) for v in pv]
                p = sheet_prefix if sheet_prefix else group_name
                try:
                    ns.Name = sanitize_sheet_name(p + ": " + ", ".join(sn))
                    sheet_count += 1
                    total_placed += len(pv)
                    log_message("  📄 " + ns.SheetNumber + ": " + str(len(pv)) + " видов")
                except Exception as ex:
                    warn_creation("Не удалось задать имя листа: " + str(ex))
                
                unplaced = [v for v in fit_views if v.Id not in pids]
                if unplaced:
                    remaining.extend(unplaced)
        
        t.Commit()
    
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        log_message("❌ Ошибка размещения: " + str(ex))
        MessageBox.Show("Ошибка: " + str(ex))
    
    # Результаты
    update_progress_info(progress_form, stage, progress, pct, "Готово!", "100%")
    time.sleep(0.5)
    progress_form.Close()
    
    total_time = time.time() - start_total
    log_message("\n⏱ Общее время: " + str(round(total_time, 1)) + " сек")
    log_message("✅ Листов: " + str(sheet_count) + ", видов: " + str(total_placed))
    
    show_results_form(sheet_count, total_placed, skipped_views, failed, total_time)


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True


def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        main()
        return True
    except Exception as e:
        MessageBox.Show("Критическая ошибка: " + str(e))
        return False


if __name__ == "__main__":
    main()