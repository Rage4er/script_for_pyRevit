# -*- coding: utf-8 -*-
__title__ = """Марки на видах"""
__author__ = "Rage"
__doc__ = "Автоматическое размещение марок на 3D-видах, планах и разрезах"
__version__ = "2.0"

import clr
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import datetime
import json
import os
import traceback

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Drawing import *
from System.Windows.Forms import *

# Константы
MM_TO_FEET = 304.8
DEFAULT_OFFSET_X = 60.0
DEFAULT_OFFSET_Y = 30.0
VIEW_TYPES = ["3D виды", "Планы", "Разрезы", "Фасады"]


# Логирование
class Logger:
    def __init__(self, enabled=False):
        self.messages = []
        self.enabled = enabled

    def add(self, message):
        if self.enabled:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.messages.append("[{0}] {1}".format(timestamp, message))

    def show(self):
        if not self.enabled:
            MessageBox.Show("Логирование отключено", "Информация")
            return
        if not self.messages:
            MessageBox.Show("Нет записей в логе", "Информация")
            return
        log_text = "\n".join(self.messages)
        form = Form()
        form.Text = "Логи выполнения"
        form.Size = Size(600, 400)
        textbox = TextBox()
        textbox.Multiline = True
        textbox.Dock = DockStyle.Fill
        textbox.ScrollBars = ScrollBars.Vertical
        textbox.Text = log_text
        form.Controls.Add(textbox)
        form.ShowDialog()


# Настройки
class TagSettings(object):
    def __init__(self):
        self.selected_view_types = []
        self.selected_views = []
        self.selected_categories = []
        self.category_tag_families = {}
        self.category_tag_types = {}
        self.offset_x = DEFAULT_OFFSET_X
        self.offset_y = DEFAULT_OFFSET_Y
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True
        self.enable_logging = False
        self.view_filter_rule = "Все"


# Главная форма
class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}
        self.all_views_dict = {}
        self.lstViewsChecked = {}
        self.category_mapping = {}
        self.logger = Logger(self.settings.enable_logging)
        self.tag_defaults = self.LoadTagDefaults()
        self.view_types_collection = {}

        self.InitializeComponent()
        self.LoadAllViews()

    def InitializeComponent(self):
        self.Text = "Расстановка марок на видах v{}".format(__version__)
        self.Size = Size(850, 650)
        self.StartPosition = FormStartPosition.CenterScreen
        self.clbViewTypes.Font = Font("Segoe UI", 9)
        
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = [
            "1. Типы видов",
            "2. Выбор видов",
            "3. Категории",
            "4. Марки",
            "5. Настройки",
            "6. Выполнение",
        ]
        
        for i, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            getattr(self, "SetupTab" + str(i + 1))(tab)
            self.tabControl.TabPages.Add(tab)

        self.Controls.Add(self.tabControl)

    def CreateControl(self, control_type, **kwargs):
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control

    def CreateButton(self, text, location, size=None, click_handler=None):
        size = size or Size(90, 28)
        btn = self.CreateControl(Button, Text=text, Location=location, Size=size)
        if click_handler:
            btn.Click += click_handler
        return btn

    # Вкладка 1: Типы видов
    def SetupTab1(self, tab):
        # Заголовок
        lblTitle = self.CreateControl(
            Label,
            Text="Выберите типы видов для обработки:",
            Location=Point(15, 15),
            Size=Size(400, 25)
        )
        lblTitle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        tab.Controls.Add(lblTitle)
    
        # Список типов видов
        self.clbViewTypes = CheckedListBox()
        self.clbViewTypes.Location = Point(15, 45)
        self.clbViewTypes.Size = Size(350, 180)
        self.clbViewTypes.CheckOnClick = True
        # Убрали проблемную строку с Font
        tab.Controls.Add(self.clbViewTypes)
    
        # Фильтр видов
        lblFilter = self.CreateControl(
            Label,
            Text="Дополнительный фильтр:",
            Location=Point(15, 240),
            Size=Size(180, 20)
        )
        tab.Controls.Add(lblFilter)
    
        self.cmbViewFilter = ComboBox()
        self.cmbViewFilter.Location = Point(200, 240)
        self.cmbViewFilter.Size = Size(200, 22)
        # Исправленный AddRange
        self.cmbViewFilter.Items.Add("Все")
        self.cmbViewFilter.Items.Add("Только незаблокированные")
        self.cmbViewFilter.Items.Add("Без шаблонов")
        self.cmbViewFilter.SelectedIndex = 0
        tab.Controls.Add(self.cmbViewFilter)
    
        # Кнопка далее
        self.btnNextTab1 = self.CreateButton(
            "Далее →",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnNextTab1Click
        )
        tab.Controls.Add(self.btnNextTab1)
    
        # Информация
        lblInfo = self.CreateControl(
            Label,
            Text="Совет: Выберите один или несколько типов видов.\nЗатем на следующей вкладке выберите конкретные   виды.",
            Location=Point(15, 280),
            Size=Size(500, 40),
            ForeColor=Color.DarkSlateGray
        )
        tab.Controls.Add(lblInfo)

    # Вкладка 2: Выбор видов
    def SetupTab2(self, tab):
        # Заголовок
        self.lblViews = self.CreateControl(
            Label,
            Text="Выберите конкретные виды:",
            Location=Point(15, 15),
            Size=Size(400, 25),
            self.lblViews.Font = System.Drawing.Font(self.Font.FontFamily, self.Font.Size, System.Drawing.FontStyle.Bold)
        )
        tab.Controls.Add(self.lblViews)
        
        # Поиск
        self.lblSearch = self.CreateControl(
            Label,
            Text="Поиск:",
            Location=Point(15, 50),
            Size=Size(50, 20)
        )
        tab.Controls.Add(self.lblSearch)
        
        self.txtSearchViews = self.CreateControl(
            TextBox,
            Location=Point(70, 50),
            Size=Size(200, 23)
        )
        tab.Controls.Add(self.txtSearchViews)
        
        # Кнопки выбора
        self.btnSelectAll = self.CreateButton(
            "Выбрать все",
            Point(280, 50),
            Size(100, 25),
            click_handler=self.OnSelectAllViews
        )
        tab.Controls.Add(self.btnSelectAll)
        
        self.btnDeselectAll = self.CreateButton(
            "Снять выбор",
            Point(390, 50),
            Size(100, 25),
            click_handler=self.OnDeselectAllViews
        )
        tab.Controls.Add(self.btnDeselectAll)
        
        self.btnRefreshViews = self.CreateButton(
            "Обновить список",
            Point(500, 50),
            Size(120, 25),
            click_handler=self.OnRefreshViewsClick
        )
        tab.Controls.Add(self.btnRefreshViews)
        
        # Список видов
        self.lstViews = self.CreateControl(
            CheckedListBox,
            Location=Point(15, 85),
            Size=Size(800, 380),
            CheckOnClick=True
        )
        tab.Controls.Add(self.lstViews)
        
        # Логирование
        self.chkLogging = self.CreateControl(
            CheckBox,
            Text="Включить логирование",
            Location=Point(15, 480),
            Size=Size(180, 20)
        )
        tab.Controls.Add(self.chkLogging)
        
        self.btnShowLogs = self.CreateButton(
            "Показать логи",
            Point(200, 480),
            Size(100, 25),
            click_handler=self.OnShowLogsClick
        )
        tab.Controls.Add(self.btnShowLogs)
        
        # Кнопки навигации
        self.btnBack1 = self.CreateButton(
            "← Назад",
            Point(550, 550),
            click_handler=self.OnBack1Click
        )
        tab.Controls.Add(self.btnBack1)
        
        self.btnNext1 = self.CreateButton(
            "Далее →",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnNext1Click
        )
        tab.Controls.Add(self.btnNext1)
        
        # Обработчики событий
        self.txtSearchViews.TextChanged += self.OnSearchViewsTextChanged
        self.chkLogging.CheckedChanged += self.OnLoggingCheckedChanged

    # Вкладка 3: Категории
    def SetupTab3(self, tab):
        self.lblCategories = self.CreateControl(
            Label,
            Text="Выберите категории элементов:",
            Location=Point(15, 15),
            Size=Size(400, 25),
            Font=System.Drawing.Font(self.Font.FontFamily, self.Font.Size, System.Drawing.FontStyle.Bold)
        )
        tab.Controls.Add(self.lblCategories)
        
        self.lstCategories = self.CreateControl(
            CheckedListBox,
            Location=Point(15, 45),
            Size=Size(800, 420),
            CheckOnClick=True
        )
        tab.Controls.Add(self.lstCategories)
        
        # Кнопки навигации
        self.btnBack2 = self.CreateButton(
            "← Назад",
            Point(550, 550),
            click_handler=self.OnBack2Click
        )
        tab.Controls.Add(self.btnBack2)
        
        self.btnNext2 = self.CreateButton(
            "Далее →",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnNext2Click
        )
        tab.Controls.Add(self.btnNext2)

    # Вкладка 4: Марки
    def SetupTab4(self, tab):
        self.lblTagFamilies = self.CreateControl(
            Label,
            Text="Выберите марки для категорий:",
            Location=Point(15, 15),
            Size=Size(400, 25),
            Font=System.Drawing.Font(self.Font.FontFamily, self.Font.Size, System.Drawing.FontStyle.Bold)
        )
        tab.Controls.Add(self.lblTagFamilies)
        
        self.lstTagFamilies = ListView()
        self.lstTagFamilies.Location = Point(15, 45)
        self.lstTagFamilies.Size = Size(800, 420)
        self.lstTagFamilies.View = View.Details
        self.lstTagFamilies.FullRowSelect = True
        self.lstTagFamilies.GridLines = True
        self.lstTagFamilies.Font = Font(self.Font, 9)
        
        # Колонки
        self.lstTagFamilies.Columns.Add("Категория", 180)
        self.lstTagFamilies.Columns.Add("Семейство марки", 250)
        self.lstTagFamilies.Columns.Add("Типоразмер марки", 350)
        
        self.lstTagFamilies.DoubleClick += self.OnTagFamilyDoubleClick
        tab.Controls.Add(self.lstTagFamilies)
        
        # Инструкция
        lblInstruction = self.CreateControl(
            Label,
            Text="Подсказка: Дважды щелкните по строке для выбора другой марки",
            Location=Point(15, 475),
            Size=Size(500, 20),
            ForeColor=Color.DarkSlateGray
        )
        tab.Controls.Add(lblInstruction)
        
        # Кнопки навигации
        self.btnBack3 = self.CreateButton(
            "← Назад",
            Point(550, 550),
            click_handler=self.OnBack3Click
        )
        tab.Controls.Add(self.btnBack3)
        
        self.btnNext3 = self.CreateButton(
            "Далее →",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnNext3Click
        )
        tab.Controls.Add(self.btnNext3)

    # Вкладка 5: Настройки
    def SetupTab5(self, tab):
        lblTitle = self.CreateControl(
            Label,
            Text="Настройки размещения марок:",
            Location=Point(15, 15),
            Size=Size(400, 25),
            Font=System.Drawing.Font(self.Font.FontFamily, self.Font.Size, System.Drawing.FontStyle.Bold)
        )
        tab.Controls.Add(lblTitle)
        
        # Смещение
        lblOffsetX = self.CreateControl(
            Label,
            Text="Смещение по X (мм):",
            Location=Point(15, 60),
            Size=Size(150, 20)
        )
        tab.Controls.Add(lblOffsetX)
        
        self.txtOffsetX = self.CreateControl(
            TextBox,
            Location=Point(170, 60),
            Size=Size(100, 23),
            Text=str(DEFAULT_OFFSET_X)
        )
        tab.Controls.Add(self.txtOffsetX)
        
        lblOffsetY = self.CreateControl(
            Label,
            Text="Смещение по Y (мм):",
            Location=Point(15, 90),
            Size=Size(150, 20)
        )
        tab.Controls.Add(lblOffsetY)
        
        self.txtOffsetY = self.CreateControl(
            TextBox,
            Location=Point(170, 90),
            Size=Size(100, 23),
            Text=str(DEFAULT_OFFSET_Y)
        )
        tab.Controls.Add(self.txtOffsetY)
        
        # Ориентация
        lblOrientation = self.CreateControl(
            Label,
            Text="Ориентация марки:",
            Location=Point(15, 130),
            Size=Size(150, 20)
        )
        tab.Controls.Add(lblOrientation)
        
        self.cmbOrientation = ComboBox()
        self.cmbOrientation.Location = Point(170, 130)
        self.cmbOrientation.Size = Size(150, 22)
        self.cmbOrientation.Items.Add("Горизонтальная")
        self.cmbOrientation.Items.Add("Вертикальная")
        self.cmbOrientation.SelectedIndex = 0
        tab.Controls.Add(self.cmbOrientation)
        
        # Выноска
        self.chkUseLeader = CheckBox()
        self.chkUseLeader.Text = "Использовать выноску"
        self.chkUseLeader.Location = Point(15, 170)
        self.chkUseLeader.Size = Size(200, 20)
        self.chkUseLeader.Checked = True
        tab.Controls.Add(self.chkUseLeader)
        
        # Информация о смещении
        lblOffsetInfo = self.CreateControl(
            Label,
            Text="Смещение применяется с учетом масштаба вида.\nДля 3D видов: случайное направление\nДля планов/разрезов: вправо и вверх",
            Location=Point(15, 210),
            Size=Size(400, 40),
            ForeColor=Color.DarkSlateGray
        )
        tab.Controls.Add(lblOffsetInfo)
        
        # Кнопки навигации
        self.btnBack4 = self.CreateButton(
            "← Назад",
            Point(550, 550),
            click_handler=self.OnBack4Click
        )
        tab.Controls.Add(self.btnBack4)
        
        self.btnNext4 = self.CreateButton(
            "Далее →",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnNext4Click
        )
        tab.Controls.Add(self.btnNext4)

    # Вкладка 6: Выполнение
    def SetupTab6(self, tab):
        lblTitle = self.CreateControl(
            Label,
            Text="Готово к выполнению:",
            Location=Point(15, 15),
            Size=Size(400, 25),
            Font=System.Drawing.Font(self.Font.FontFamily, self.Font.Size, System.Drawing.FontStyle.Bold)
        )
        tab.Controls.Add(lblTitle)
        
        self.txtSummary = TextBox()
        self.txtSummary.Location = Point(15, 45)
        self.txtSummary.Size = Size(800, 400)
        self.txtSummary.Multiline = True
        self.txtSummary.ScrollBars = ScrollBars.Vertical
        self.txtSummary.ReadOnly = True
        self.txtSummary.Font = Font("Consolas", 9)
        tab.Controls.Add(self.txtSummary)
        
        self.progressBar = ProgressBar()
        self.progressBar.Location = Point(15, 460)
        self.progressBar.Size = Size(800, 25)
        self.progressBar.Minimum = 0
        self.progressBar.Maximum = 100
        tab.Controls.Add(self.progressBar)
        
        # Кнопки навигации
        self.btnBack5 = self.CreateButton(
            "← Назад",
            Point(550, 550),
            click_handler=self.OnBack5Click
        )
        tab.Controls.Add(self.btnBack5)
        
        self.btnExecute = self.CreateButton(
            "Выполнить",
            Point(650, 550),
            Size(120, 35),
            click_handler=self.OnExecuteClick
        )
        self.btnExecute.BackColor = Color.LightGreen
        tab.Controls.Add(self.btnExecute)

    # ============ МЕТОДЫ ЗАГРУЗКИ ДАННЫХ ============
    
    def LoadAllViews(self):
        try:
            self.all_views_dict.clear()
            self.view_types_collection.clear()
            
            # Собираем все подходящие виды
            views = FilteredElementCollector(self.doc)\
                .OfClass(View)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            for view in views:
                if self.IsViewSuitable(view):
                    self.CategorizeView(view)
            
            # Обновляем информацию о количестве
            self.UpdateViewTypesSummary()
            
        except Exception as e:
            self.logger.add("Ошибка загрузки видов: " + str(e))
            MessageBox.Show("Ошибка загрузки видов: " + str(e), "Ошибка")

    def IsViewSuitable(self, view):
        try:
            # Пропускаем шаблоны
            if view.IsTemplate:
                return False
                
            # Проверяем, можно ли печатать вид (опционально)
            if hasattr(view, 'CanBePrinted') and not view.CanBePrinted:
                return False
            
            # Определяем тип вида
            if isinstance(view, View3D):
                # Только 3D виды, не являющиеся шаблонами
                return not view.IsTemplate
            elif isinstance(view, ViewPlan):
                # Пропускаем чертежные листы и шаблоны видов
                if view.ViewType == ViewType.DrawingSheet:
                    return False
                # Включаем планы этажей и потолков
                return view.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan]
            elif isinstance(view, ViewSection):
                # Разрезы и фасады
                return view.ViewType in [ViewType.Section, ViewType.Elevation, ViewType.Detail]
            
            return False
            
        except Exception as e:
            self.logger.add("Ошибка проверки вида {}: {}".format(view.Id, e))
            return False

    def CategorizeView(self, view):
        try:
            view_name = self.GetViewDisplayName(view)
            
            # Определяем тип вида для группировки
            if isinstance(view, View3D):
                view_type = "3D виды"
            elif isinstance(view, ViewPlan):
                if view.ViewType == ViewType.FloorPlan:
                    view_type = "Планы"
                elif view.ViewType == ViewType.CeilingPlan:
                    view_type = "Планы"
                else:
                    view_type = "Планы"
            elif isinstance(view, ViewSection):
                if view.ViewType == ViewType.Section:
                    view_type = "Разрезы"
                elif view.ViewType == ViewType.Elevation:
                    view_type = "Фасады"
                else:
                    view_type = "Разрезы"
            else:
                view_type = "Другие виды"
            
            # Сохраняем
            if view_type not in self.view_types_collection:
                self.view_types_collection[view_type] = []
            
            self.view_types_collection[view_type].append(view)
            self.all_views_dict[view_name] = view
            
        except Exception as e:
            self.logger.add("Ошибка категоризации вида {}: {}".format(view.Id, e))

    def GetViewDisplayName(self, view):
        try:
            base_name = view.Name
            view_id = view.Id.IntegerValue
            
            # Добавляем информацию о типе
            if isinstance(view, View3D):
                type_info = " [3D]"
            elif isinstance(view, ViewPlan):
                if view.ViewType == ViewType.FloorPlan:
                    type_info = " [План этажа]"
                elif view.ViewType == ViewType.CeilingPlan:
                    type_info = " [План потолка]"
                else:
                    type_info = " [План]"
            elif isinstance(view, ViewSection):
                if view.ViewType == ViewType.Section:
                    type_info = " [Разрез]"
                elif view.ViewType == ViewType.Elevation:
                    type_info = " [Фасад]"
                else:
                    type_info = " [Сечение]"
            else:
                type_info = " [Вид]"
            
            return "{}{} (ID: {})".format(base_name, type_info, view_id)
            
        except Exception as e:
            self.logger.add("Ошибка получения имени вида: {}".format(e))
            return "Вид {}".format(view.Id.IntegerValue)

    def UpdateViewTypesSummary(self):
        if hasattr(self, 'clbViewTypes'):
            for i in range(self.clbViewTypes.Items.Count):
                item_text = self.clbViewTypes.Items[i]
                # Извлекаем чистый тип вида (без количества)
                if '(' in item_text:
                    view_type = item_text.split(' (')[0]
                else:
                    view_type = item_text
                
                # Обновляем текст с количеством
                count = len(self.view_types_collection.get(view_type, []))
                self.clbViewTypes.Items[i] = "{} ({})".format(view_type, count)

    def LoadSelectedViewTypes(self):
        try:
            self.lstViews.Items.Clear()
            self.views_dict.clear()
            self.lstViewsChecked.clear()
            
            # Собираем виды выбранных типов
            selected_views = []
            for view_type in self.settings.selected_view_types:
                if view_type in self.view_types_collection:
                    selected_views.extend(self.view_types_collection[view_type])
            
            # Применяем фильтры
            filtered_views = self.ApplyViewFilters(selected_views)
            
            # Добавляем в список
            for view in filtered_views:
                name = self.GetViewDisplayName(view)
                self.views_dict[name] = view
                self.lstViewsChecked[name] = False
                self.lstViews.Items.Add(name, False)
            
            if self.lstViews.Items.Count > 0:
                self.lstViews.SetItemChecked(0, True)
                self.lstViewsChecked[self.lstViews.Items[0]] = True
                
        except Exception as e:
            MessageBox.Show("Ошибка загрузки видов: " + str(e), "Ошибка")

    def ApplyViewFilters(self, views):
        filtered = []
        
        for view in views:
            include = True
            
            # Фильтр по заблокированности
            if self.settings.view_filter_rule == "Только незаблокированные":
                if not self.IsViewUnlocked(view):
                    include = False
            
            # Фильтр "Без шаблонов" уже применен в IsViewSuitable
            
            if include:
                filtered.append(view)
        
        return filtered

    def IsViewUnlocked(self, view):
        try:
            # Проверяем параметр блокировки вида
            lock_param = view.get_Parameter(BuiltInParameter.VIEW_LOCK)
            if lock_param and lock_param.HasValue:
                return lock_param.AsInteger() == 0
            return True
        except:
            return True

    # ============ МЕТОДЫ НАВИГАЦИИ ============
    
    # Вкладка 1 -> Вкладка 2
    def OnNextTab1Click(self, sender, args):
        # Сохраняем выбранные типы видов
        self.settings.selected_view_types = []
        for i in range(self.clbViewTypes.Items.Count):
            if self.clbViewTypes.GetItemChecked(i):
                item_text = self.clbViewTypes.Items[i]
                # Извлекаем тип вида без количества в скобках
                view_type = item_text.split(' (')[0]
                self.settings.selected_view_types.append(view_type)
        
        # Сохраняем фильтр
        self.settings.view_filter_rule = self.cmbViewFilter.SelectedItem
        
        if not self.settings.selected_view_types:
            MessageBox.Show("Выберите хотя бы один тип видов!", "Внимание")
            return
        
        # Загружаем виды выбранных типов
        self.LoadSelectedViewTypes()
        
        # Переходим на следующую вкладку
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    # Вкладка 2 -> Вкладка 3
    def OnNext1Click(self, sender, args):
        self.settings.enable_logging = self.chkLogging.Checked
        self.logger.enabled = self.settings.enable_logging

        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            name = self.lstViews.Items[i]
            checked = self.lstViews.GetItemChecked(i)
            if checked and name in self.views_dict:
                self.settings.selected_views.append(self.views_dict[name])
            self.lstViewsChecked[name] = checked

        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!", "Внимание")
            return

        self.CollectCategories()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    # Вкладка 3 -> Вкладка 4
    def OnNext2Click(self, sender, args):
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            name = self.lstCategories.Items[i]
            if self.lstCategories.GetItemChecked(i) and name in self.category_mapping:
                self.settings.selected_categories.append(self.category_mapping[name])

        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!", "Внимание")
            return

        self.PopulateTagFamilies()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 3
        self.tabControl.Selecting += self.OnTabSelecting

    # Вкладка 4 -> Вкладка 5
    def OnNext3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 4
        self.tabControl.Selecting += self.OnTabSelecting

    # Вкладка 5 -> Вкладка 6
    def OnNext4Click(self, sender, args):
        try:
            # Сохраняем настройки смещения
            self.settings.offset_x = float(self.txtOffsetX.Text)
            self.settings.offset_y = float(self.txtOffsetY.Text)
        except:
            MessageBox.Show("Некорректное значение смещения. Использованы значения по умолчанию.", "Внимание")
            self.settings.offset_x = DEFAULT_OFFSET_X
            self.settings.offset_y = DEFAULT_OFFSET_Y
        
        self.settings.orientation = (
            TagOrientation.Horizontal 
            if self.cmbOrientation.SelectedIndex == 0 
            else TagOrientation.Vertical
        )
        self.settings.use_leader = self.chkUseLeader.Checked
        
        # Генерируем сводку
        self.txtSummary.Text = self.GenerateSummary()
        
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 5
        self.tabControl.Selecting += self.OnTabSelecting

    # Назад: Вкладка 2 -> Вкладка 1
    def OnBack1Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 0
        self.tabControl.Selecting += self.OnTabSelecting

    # Назад: Вкладка 3 -> Вкладка 2
    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    # Назад: Вкладка 4 -> Вкладка 3
    def OnBack3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    # Назад: Вкладка 5 -> Вкладка 4
    def OnBack4Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 3
        self.tabControl.Selecting += self.OnTabSelecting

    # Назад: Вкладка 6 -> Вкладка 5
    def OnBack5Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 4
        self.tabControl.Selecting += self.OnTabSelecting

    def OnTabSelecting(self, sender, args):
        args.Cancel = True

    # ============ ОБРАБОТЧИКИ СОБЫТИЙ ============
    
    def OnSelectAllViews(self, sender, args):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, True)
            name = self.lstViews.Items[i]
            self.lstViewsChecked[name] = True

    def OnDeselectAllViews(self, sender, args):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, False)
            name = self.lstViews.Items[i]
            self.lstViewsChecked[name] = False

    def OnSearchViewsTextChanged(self, sender, args):
        filter_text = sender.Text
        self.UpdateViewsList(filter_text)

    def UpdateViewsList(self, filter_text):
        self.lstViews.Items.Clear()
        filter_lower = filter_text.lower()
        
        for name, view in self.all_views_dict.items():
            # Проверяем, соответствует ли вид выбранным типам
            view_type = None
            if isinstance(view, View3D):
                view_type = "3D виды"
            elif isinstance(view, ViewPlan):
                view_type = "Планы"
            elif isinstance(view, ViewSection):
                if hasattr(view, 'ViewType') and view.ViewType == ViewType.Elevation:
                    view_type = "Фасады"
                else:
                    view_type = "Разрезы"
            
            # Фильтруем по типу и тексту
            if view_type in self.settings.selected_view_types and filter_lower in name.lower():
                self.lstViews.Items.Add(name, self.lstViewsChecked.get(name, False))

    def OnLoggingCheckedChanged(self, sender, args):
        self.logger.enabled = sender.Checked

    def OnRefreshViewsClick(self, sender, args):
        self.LoadAllViews()
        self.LoadSelectedViewTypes()
        MessageBox.Show("Список видов обновлен", "Информация")

    def OnShowLogsClick(self, sender, args):
        self.logger.show()

    # ============ КАТЕГОРИИ И МАРКИ ============
    
    def CollectCategories(self):
        categories = [
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_MechanicalEquipment,
        ]

        unique_cats = set()
        for cat in categories:
            try:
                cat_obj = Category.GetCategory(self.doc, cat)
                if cat_obj:
                    unique_cats.add(cat_obj)
            except Exception as e:
                self.logger.add("Ошибка получения категории {}: {}".format(cat, e))

        self.settings.selected_categories = list(unique_cats)
        self.lstCategories.Items.Clear()
        self.category_mapping.clear()

        for cat in sorted(
            self.settings.selected_categories, 
            key=lambda x: self.GetCategoryName(x)
        ):
            name = self.GetCategoryName(cat)
            self.lstCategories.Items.Add(name, True)
            self.category_mapping[name] = cat

    def GetCategoryName(self, category):
        if not category:
            return "Неизвестная категория"
        try:
            if hasattr(category, "Id") and category.Id.IntegerValue < 0:
                return LabelUtils.GetLabelFor(BuiltInCategory(category.Id.IntegerValue))
        except Exception as e:
            self.logger.add("Ошибка получения имени категории: {}".format(e))
        return getattr(category, "Name", "Неизвестная категория")

    def GetDuctsToTag(self, duct_elements):
        groups = {}

        for duct in duct_elements:
            try:
                system_name = ""
                section = ""
                length = 0.0

                # Имя системы
                system_name_param = duct.LookupParameter("Имя системы")
                if not system_name_param:
                    system_name_param = duct.get_Parameter(
                        BuiltInParameter.RBS_SYSTEM_NAME_PARAM
                    )
                
                if system_name_param and system_name_param.HasValue:
                    if system_name_param.StorageType == StorageType.String:
                        system_name = system_name_param.AsString()
                    else:
                        system_name = system_name_param.AsValueString()

                # Сечение
                section = self.GetDuctSection(duct)

                # Длина
                length_param = duct.LookupParameter("Длина")
                if not length_param:
                    length_param = duct.get_Parameter(
                        BuiltInParameter.CURVE_ELEM_LENGTH
                    )
                
                if length_param and length_param.HasValue:
                    if length_param.StorageType == StorageType.Double:
                        length = length_param.AsDouble()
                    else:
                        try:
                            length = float(length_param.AsValueString().replace(",", "."))
                        except:
                            length = 0.0

                # Группируем
                if system_name not in groups:
                    groups[system_name] = {}

                if section not in groups[system_name]:
                    groups[system_name][section] = []

                groups[system_name][section].append({"element": duct, "length": length})

            except Exception as e:
                self.logger.add("Ошибка обработки воздуховода {}: {}".format(duct.Id, e))
                continue

        # Выбираем самые длинные
        selected_ducts = []
        for system_name, sections in groups.items():
            for section, ducts in sections.items():
                if ducts:
                    longest_duct = max(ducts, key=lambda x: x["length"])
                    selected_ducts.append(longest_duct["element"])
                    self.logger.add(
                        "Выбран воздуховод ID {} для системы '{}', сечения '{}', длина {:.2f}".format(
                            longest_duct["element"].Id,
                            system_name,
                            section,
                            longest_duct["length"],
                        )
                    )

        return selected_ducts

    def GetDuctSection(self, duct):
        try:
            # Круглые воздуховоды
            diameter_param = duct.LookupParameter("Диаметр")
            if not diameter_param:
                diameter_param = duct.get_Parameter(
                    BuiltInParameter.RBS_CURVE_DIAMETER_PARAM
                )

            if diameter_param and diameter_param.HasValue:
                if diameter_param.StorageType == StorageType.Double:
                    diameter = diameter_param.AsDouble() * 304.8
                    return "Ø{:.0f}".format(diameter)

            # Прямоугольные воздуховоды
            width_param = duct.LookupParameter("Ширина")
            if not width_param:
                width_param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)

            height_param = duct.LookupParameter("Высота")
            if not height_param:
                height_param = duct.get_Parameter(
                    BuiltInParameter.RBS_CURVE_HEIGHT_PARAM
                )

            if (
                width_param
                and height_param
                and width_param.HasValue
                and height_param.HasValue
            ):
                width = 0.0
                height = 0.0

                if width_param.StorageType == StorageType.Double:
                    width = width_param.AsDouble() * 304.8
                else:
                    try:
                        width = float(width_param.AsValueString().replace(",", "."))
                    except:
                        width = 0.0

                if height_param.StorageType == StorageType.Double:
                    height = height_param.AsDouble() * 304.8
                else:
                    try:
                        height = float(height_param.AsValueString().replace(",", "."))
                    except:
                        height = 0.0

                return "{:.0f}x{:.0f}".format(width, height)

            return "Не определено"
        except Exception as e:
            self.logger.add("Ошибка получения сечения воздуховода {}: {}".format(duct.Id, e))
            return "Ошибка"

    def PopulateTagFamilies(self):
        self.lstTagFamilies.Items.Clear()
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            item.Tag = category

            # Ищем сохраненную марку
            cat_name = self.GetCategoryName(category)
            default = self.tag_defaults.get(cat_name, {})
            family_name = default.get("family")
            type_name = default.get("type")

            if family_name and type_name:
                tag_family, tag_type = self.FindSavedTag(family_name, type_name)
                if tag_family and tag_type:
                    item.SubItems.Add(self.GetElementName(tag_family))
                    item.SubItems.Add(self.GetElementName(tag_type))
                    self.settings.category_tag_families[category] = tag_family
                    self.settings.category_tag_types[category] = tag_type
                    self.lstTagFamilies.Items.Add(item)
                    continue

            # Ищем подходящую марку
            tag_family, tag_type = self.FindTagForCategory(category)

            if tag_family and tag_type:
                item.SubItems.Add(self.GetElementName(tag_family))
                item.SubItems.Add(self.GetElementName(tag_type))
                self.settings.category_tag_families[category] = tag_family
                self.settings.category_tag_types[category] = tag_type
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")

            self.lstTagFamilies.Items.Add(item)

    def FindTagForCategory(self, category):
        tag_category_id = self.GetTagCategoryId(category)
        if not tag_category_id:
            return None, None

        collector = FilteredElementCollector(self.doc)
        tag_families = (
            collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        )

        for family in tag_families:
            if not family or not hasattr(family, "FamilyCategory"):
                continue

            if family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                symbol_ids = family.GetFamilySymbolIds()
                if symbol_ids and symbol_ids.Count > 0:
                    # Ищем активный типоразмер
                    for symbol_id in list(symbol_ids):
                        tag_type = self.doc.GetElement(symbol_id)
                        if tag_type and tag_type.IsActive:
                            return family, tag_type
                    # Берем первый доступный
                    tag_type = self.doc.GetElement(list(symbol_ids)[0])
                    return family, tag_type

        return None, None

    def GetTagCategoryId(self, element_category):
        mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipment: BuiltInCategory.OST_MechanicalEquipmentTags,
        }

        try:
            if hasattr(element_category, "Id") and element_category.Id.IntegerValue < 0:
                element_cat = BuiltInCategory(element_category.Id.IntegerValue)
                if element_cat in mapping:
                    tag_cat = Category.GetCategory(self.doc, mapping[element_cat])
                    if tag_cat:
                        return tag_cat.Id
        except Exception as e:
            self.logger.add("Ошибка получения категории марки: {}".format(e))
        return None

    def GetElementName(self, element):
        if not element:
            return "Без имени"
        try:
            if isinstance(element, FamilySymbol):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name

                if hasattr(element, "Family") and element.Family:
                    family_name = (
                        element.Family.Name
                        if hasattr(element.Family, "Name") and element.Family.Name
                        else ""
                    )

                    type_name = ""
                    for param_name in ["Тип", "Type Name", "Имя типа"]:
                        param = element.LookupParameter(param_name)
                        if param and param.HasValue:
                            type_name = param.AsString()
                            break

                    if family_name and type_name:
                        return "{} - {}".format(family_name, type_name)
                    elif type_name:
                        return type_name
                    elif family_name:
                        return family_name

                return "Типоразмер " + str(element.Id.IntegerValue)

            elif isinstance(element, Family):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            if hasattr(element, "Name") and element.Name:
                name = element.Name
                if name and not name.startswith("IronPython"):
                    return name

        except Exception as e:
            self.logger.add("GetElementName: Ошибка при получении имени элемента: {}".format(e))

        return "Элемент " + str(element.Id.IntegerValue)

    def OnTagFamilyDoubleClick(self, sender, args):
        if self.lstTagFamilies.SelectedItems.Count == 0:
            return
        selected = self.lstTagFamilies.SelectedItems[0]
        category = selected.Tag

        available_families = self.GetAvailableTagFamiliesForCategory(category)

        if available_families:
            current_family = self.settings.category_tag_families.get(category)
            current_type = self.settings.category_tag_types.get(category)

            form = TagFamilySelectionForm(
                self.doc, available_families, current_family, current_type
            )
            if (
                form.ShowDialog() == DialogResult.OK
                and form.SelectedFamily
                and form.SelectedType
            ):
                selected.SubItems[1].Text = self.GetElementName(form.SelectedFamily)
                selected.SubItems[2].Text = self.GetElementName(form.SelectedType)
                self.settings.category_tag_families[category] = form.SelectedFamily
                self.settings.category_tag_types[category] = form.SelectedType
        else:
            MessageBox.Show("Нет доступных семейств марок для этой категории", "Внимание")

    def GetAvailableTagFamiliesForCategory(self, category):
        tag_category_id = self.GetTagCategoryId(category)
        if not tag_category_id:
            return []

        collector = FilteredElementCollector(self.doc)
        tag_families = (
            collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        )

        available_families = []
        for family in tag_families:
            if (
                family
                and hasattr(family, "FamilyCategory")
                and family.FamilyCategory
                and family.FamilyCategory.Id == tag_category_id
            ):
                available_families.append(family)

        return available_families

    def GenerateSummary(self):
        summary = "=" * 60 + "\n"
        summary += "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ\n"
        summary += "=" * 60 + "\n\n"
        
        # Типы видов
        summary += "ТИПЫ ВИДОВ:\n"
        summary += "- " + ", ".join(self.settings.selected_view_types) + "\n"
        summary += "Фильтр: " + self.settings.view_filter_rule + "\n\n"
        
        # Виды
        summary += "ВЫБРАНО ВИДОВ: {}\n".format(len(self.settings.selected_views))
        for i, view in enumerate(self.settings.selected_views[:5]):  # Показываем первые 5
            summary += "  {}. {}\n".format(i+1, self.GetViewDisplayName(view))
        if len(self.settings.selected_views) > 5:
            summary += "  ... и еще {} видов\n".format(len(self.settings.selected_views) - 5)
        summary += "\n"
        
        # Категории
        summary += "КАТЕГОРИИ ЭЛЕМЕНТОВ: {}\n".format(len(self.settings.selected_categories))
        for category in self.settings.selected_categories:
            summary += "- {}\n".format(self.GetCategoryName(category))
        summary += "\n"
        
        # Настройки
        summary += "НАСТРОЙКИ РАЗМЕЩЕНИЯ:\n"
        summary += "- Смещение: X={}мм, Y={}мм\n".format(self.settings.offset_x, self.settings.offset_y)
        orientation_text = "Горизонтальная" if self.settings.orientation == TagOrientation.Horizontal else "Вертикальная"
        summary += "- Ориентация: {}\n".format(orientation_text)
        summary += "- Выноска: {}\n".format("Да" if self.settings.use_leader else "Нет")
        summary += "- Логирование: {}\n".format("Включено" if self.settings.enable_logging else "Отключено")
        summary += "\n"
        
        # Марки
        summary += "НАСТРОЙКИ МАРОК:\n"
        for category in self.settings.selected_categories:
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            if tag_family and tag_type:
                status = "{} ({})".format(
                    self.GetElementName(tag_family),
                    self.GetElementName(tag_type)
                )
            else:
                status = "НЕ НАЗНАЧЕНА"
            summary += "- {}: {}\n".format(self.GetCategoryName(category), status)
        
        summary += "\n" + "=" * 60 + "\n"
        summary += "Готово к выполнению. Нажмите 'Выполнить' для начала.\n"
        summary += "=" * 60
        
        return summary

    def OnExecuteClick(self, sender, args):
        success_count = 0
        errors = []
        
        self.logger.add("=" * 50)
        self.logger.add("Начало выполнения расстановки марок")
        self.logger.add("=" * 50)
        
        # Блокируем кнопку выполнения
        self.btnExecute.Enabled = False
        self.btnExecute.Text = "Выполняется..."
        
        try:
            # Предварительный сбор элементов для прогресс-бара
            self.logger.add("Предварительный подсчет элементов...")
            all_elements_for_progress = []
            
            for view in self.settings.selected_views:
                self.logger.add("Подготовка вида: {}".format(self.GetViewDisplayName(view)))
                
                for category in self.settings.selected_categories:
                    try:
                        elements = list(
                            FilteredElementCollector(self.doc, view.Id)
                            .OfCategoryId(category.Id)
                            .WhereElementIsNotElementType()
                            .ToElements()
                        )
                        
                        # Фильтрация воздуховодов
                        if category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
                            elements = self.GetDuctsToTag(elements)
                            self.logger.add("  Категория '{}': {} элементов после фильтрации".format(
                                self.GetCategoryName(category), len(elements)
                            ))
                        
                        all_elements_for_progress.extend(elements)
                        
                    except Exception as e:
                        error_msg = "Ошибка сбора элементов для категории {}: {}".format(
                            self.GetCategoryName(category), str(e)
                        )
                        errors.append(error_msg)
                        self.logger.add(error_msg)
            
            # Настройка прогресс-бара
            total_operations = len(all_elements_for_progress)
            self.logger.add("Всего операций для выполнения: {}".format(total_operations))
            
            if total_operations == 0:
                self.progressBar.Maximum = 1
                self.progressBar.Value = 0
                MessageBox.Show(
                    "Нет элементов для маркировки на выбранных видах.\nПроверьте выбор категорий и фильтры.", 
                    "Информация"
                )
                self.btnExecute.Enabled = True
                self.btnExecute.Text = "Выполнить"
                return
            
            self.progressBar.Maximum = total_operations
            self.progressBar.Value = 0
            
            # Основная транзакция
            trans = Transaction(self.doc, "Расстановка марок на видах")
            trans.Start()
            
            try:
                processed_count = 0
                
                for view in self.settings.selected_views:
                    view_name = self.GetViewDisplayName(view)
                    self.logger.add("\n--- Обработка вида: {} ---".format(view_name))
                    
                    for category in self.settings.selected_categories:
                        category_name = self.GetCategoryName(category)
                        tag_type = self.settings.category_tag_types.get(category)
                        
                        if not tag_type:
                            error_msg = "Нет марки для категории '{}'".format(category_name)
                            errors.append(error_msg)
                            self.logger.add(error_msg)
                            continue
                        
                        # Собираем элементы для этого вида и категории
                        try:
                            elements = list(
                                FilteredElementCollector(self.doc, view.Id)
                                .OfCategoryId(category.Id)
                                .WhereElementIsNotElementType()
                                .ToElements()
                            )
                            
                            # Фильтрация воздуховодов
                            if category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
                                elements = self.GetDuctsToTag(elements)
                            
                            self.logger.add("Категория '{}': найдено {} элементов".format(category_name, len(elements)))
                            
                            for element in elements:
                                # Проверка существующей марки
                                if self.HasExistingTag(element, view):
                                    self.logger.add("  Элемент {} уже имеет марку, пропущен".format(element.Id))
                                    processed_count += 1
                                    self.progressBar.Value = processed_count
                                    continue
                                
                                # Создание марки
                                if self.CreateTag(element, view, tag_type):
                                    success_count += 1
                                    self.logger.add("  ✓ Марка создана для элемента {}".format(element.Id))
                                else:
                                    error_msg = "  ✗ Не удалось создать марку для элемента {}".format(element.Id)
                                    errors.append(error_msg)
                                    self.logger.add(error_msg)
                                
                                # Обновление прогресса
                                processed_count += 1
                                self.progressBar.Value = processed_count
                                
                        except Exception as e:
                            error_msg = "Ошибка обработки категории '{}': {}".format(category_name, str(e))
                            errors.append(error_msg)
                            self.logger.add(error_msg)
                            self.logger.add(traceback.format_exc())
                
                trans.Commit()
                self.logger.add("\nТранзакция подтверждена успешно")
                
            except Exception as e:
                trans.RollBack()
                error_msg = "Критическая ошибка при выполнении транзакции: {}".format(str(e))
                errors.append(error_msg)
                self.logger.add(error_msg)
                self.logger.add(traceback.format_exc())
            finally:
                trans.Dispose()
                self.logger.add("Транзакция завершена")
            
        except Exception as e:
            error_msg = "Критическая ошибка: {}".format(str(e))
            errors.append(error_msg)
            self.logger.add(error_msg)
            self.logger.add(traceback.format_exc())
        finally:
            # Восстанавливаем кнопку
            self.btnExecute.Enabled = True
            self.btnExecute.Text = "Выполнить"
            
            # Финальное обновление прогресс-бара
            self.progressBar.Value = self.progressBar.Maximum
            
            # Сохранение настроек
            self.SaveTagDefaults()
            
            # Вывод результатов
            result_msg = "=" * 50 + "\n"
            result_msg += "РЕЗУЛЬТАТ ВЫПОЛНЕНИЯ\n"
            result_msg += "=" * 50 + "\n\n"
            result_msg += "Успешно расставлено марок: {}\n".format(success_count)
            result_msg += "Всего операций: {}\n".format(total_operations)
            
            if errors:
                result_msg += "\nОшибки ({}):\n".format(len(errors))
                for i, error in enumerate(errors[:10]):
                    result_msg += "{}. {}\n".format(i+1, error)
                if len(errors) > 10:
                    result_msg += "... и еще {} ошибок\n".format(len(errors) - 10)
            
            result_msg += "\n" + "=" * 50
            
            # Обновляем сводку
            self.txtSummary.Text = result_msg
            
            # Показываем результат
            MessageBox.Show(
                "Выполнение завершено.\nУспешно: {} марок\nОшибок: {}".format(success_count, len(errors)),
                "Результат",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            
            if self.settings.enable_logging and self.logger.messages:
                self.logger.show()

    def CreateTag(self, element, view, tag_type):
        try:
            # Получаем bounding box элемента
            bbox = element.get_BoundingBox(view)
            if not bbox or not bbox.Min or not bbox.Max:
                self.logger.add("  Элемент {} не видим на виде {}, пропущен".format(element.Id, view.Name))
                return False
            
            # Центр элемента
            center = (bbox.Min + bbox.Max) / 2
            
            # Коэффициент масштаба
            scale_factor = 100.0 / view.Scale
            
            # Смещение в футах (Revit использует футы)
            offset_x = (self.settings.offset_x * scale_factor) / MM_TO_FEET
            offset_y = (self.settings.offset_y * scale_factor) / MM_TO_FEET
            
            # Определяем точку размещения марки в зависимости от типа вида
            if isinstance(view, View3D):
                # Для 3D видов - случайное смещение для лучшего распределения
                direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
                direction_y = 1 if element.Id.IntegerValue % 3 == 0 else -1
                
                tag_point = XYZ(
                    center.X + offset_x * direction_x,
                    center.Y + offset_y * direction_y,
                    center.Z,
                )
            else:
                # Для 2D видов (планы, разрезы) - смещение вправо и вверх
                tag_point = XYZ(
                    center.X + offset_x,
                    center.Y + offset_y,
                    center.Z,
                )
            
            # Создаем марку
            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                Reference(element),
                self.settings.use_leader,
                TagMode.TM_ADDBY_CATEGORY,
                self.settings.orientation,
                tag_point,
            )
            
            if tag and tag_type:
                try:
                    tag.ChangeTypeId(tag_type.Id)
                    
                    # Дополнительная настройка для 2D видов
                    if not isinstance(view, View3D):
                        self.AdjustTagFor2DView(tag, element, view)
                    
                    return True
                except Exception as e:
                    self.logger.add("  Ошибка изменения типа марки: {}".format(e))
                    return False
            return False
            
        except Exception as e:
            self.logger.add("  Ошибка создания марки для элемента {}: {}".format(element.Id, e))
            return False

    def AdjustTagFor2DView(self, tag, element, view):
        """Дополнительная настройка марки для 2D видов"""
        try:
            if self.settings.use_leader and hasattr(tag, 'LeaderEndCondition'):
                # Устанавливаем свободный конец выноски
                tag.LeaderEndCondition = LeaderEndCondition.Free
                
                # Получаем точку на элементе для выноски
                bbox = element.get_BoundingBox(view)
                if bbox and bbox.Min and bbox.Max:
                    element_center = (bbox.Min + bbox.Max) / 2
                    if hasattr(tag, 'LeaderEnd'):
                        tag.LeaderEnd = element_center
                        
        except Exception as e:
            self.logger.add("  Ошибка настройки марки для 2D вида: {}".format(e))

    def HasExistingTag(self, element, view):
        try:
            collector = FilteredElementCollector(self.doc, view.Id)
            existing_tags = collector.OfClass(IndependentTag).ToElements()
            
            for tag in existing_tags:
                try:
                    # Проверяем, ссылается ли марка на этот элемент
                    tagged_refs = tag.GetTaggedReferences()
                    for ref in tagged_refs:
                        if ref.ElementId == element.Id:
                            return True
                except:
                    # Альтернативный способ проверки для старых версий Revit
                    if hasattr(tag, 'TaggedLocalElementId') and tag.TaggedLocalElementId == element.Id:
                        return True
            return False
        except Exception as e:
            self.logger.add("Ошибка проверки существующей марки: {}".format(e))
            return False

    def LoadTagDefaults(self):
        config_path = self.GetConfigPath()
        if os.path.exists(config_path):
            try:
                with open(config_path, "r") as f:
                    return json.load(f)
            except Exception as e:
                self.logger.add("Ошибка загрузки настроек марок: {}".format(e))
        return {}

    def SaveTagDefaults(self):
        defaults = {}
        for cat, family in self.settings.category_tag_families.items():
            cat_name = self.GetCategoryName(cat)
            family_name = self.GetElementName(family)
            type_elem = self.settings.category_tag_types.get(cat)
            type_name = self.GetElementName(type_elem) if type_elem else ""
            if family_name and type_name:
                defaults[cat_name] = {"family": family_name, "type": type_name}

        config_path = self.GetConfigPath()
        try:
            with open(config_path, "w") as f:
                json.dump(defaults, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.logger.add("Ошибка сохранения настроек марок: {}".format(e))

    def FindSavedTag(self, family_name, type_name):
        collector = FilteredElementCollector(self.doc).OfClass(Family)
        for family in collector:
            if self.GetElementName(family) == family_name:
                symbol_ids = family.GetFamilySymbolIds()
                for symbol_id in symbol_ids:
                    symbol = self.doc.GetElement(symbol_id)
                    if self.GetElementName(symbol) == type_name:
                        return family, symbol
        return None, None

    def GetConfigPath(self):
        script_dir = os.path.dirname(__file__)
        return os.path.join(script_dir, "tag_defaults.json")


# Форма выбора семейства марки
class TagFamilySelectionForm(Form):
    def __init__(self, doc, available_families, current_family, current_type):
        self.doc = doc
        self.available_families = available_families
        self.selected_family = None
        self.selected_type = None
        self.family_dict = {}
        self.type_dict = {}
        
        self.InitializeComponent()
        self.PopulateFamilies()

    def InitializeComponent(self):
        self.Text = "Выбор семейства и типоразмера марки"
        self.Size = Size(850, 550)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        
        # Семейства
        self.lblFamilies = Label(
            Text="Выберите семейство марки:",
            Location=Point(15, 15),
            Size=Size(250, 20),
            Font=Font("Segoe UI", 9, FontStyle.Bold)
        )
        self.Controls.Add(self.lblFamilies)
        
        self.lstFamilies = ListBox()
        self.lstFamilies.Location = Point(15, 40)
        self.lstFamilies.Size = Size(300, 400)
        self.lstFamilies.SelectedIndexChanged += self.OnFamilySelected
        self.Controls.Add(self.lstFamilies)
        
        # Типоразмеры
        self.lblTypes = Label(
            Text="Выберите типоразмер:",
            Location=Point(330, 15),
            Size=Size(250, 20),
            Font=Font("Segoe UI", 9, FontStyle.Bold)
        )
        self.Controls.Add(self.lblTypes)
        
        self.lstTypes = ListBox()
        self.lstTypes.Location = Point(330, 40)
        self.lstTypes.Size = Size(450, 400)
        self.Controls.Add(self.lstTypes)
        
        # Кнопки
        self.btnOK = Button(Text="OK", Location=Point(580, 460), Size=Size(100, 30))
        self.btnOK.Click += self.OnOKClick
        self.Controls.Add(self.btnOK)
        
        self.btnCancel = Button(Text="Отмена", Location=Point(690, 460), Size=Size(100, 30))
        self.btnCancel.Click += self.OnCancelClick
        self.Controls.Add(self.btnCancel)
        
        # Подсказка
        lblHint = Label(
            Text="Подсказка: Активные типоразмеры помечены [Активный]",
            Location=Point(15, 460),
            Size=Size(400, 20),
            ForeColor=Color.DarkSlateGray
        )
        self.Controls.Add(lblHint)

    def PopulateFamilies(self):
        self.lstFamilies.Items.Clear()
        self.family_dict.clear()
        
        for family in self.available_families:
            if family:
                family_name = self.GetElementName(family)
                self.lstFamilies.Items.Add(family_name)
                self.family_dict[family_name] = family
        
        if self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0

    def GetElementName(self, element):
        if not element:
            return "Без имени"
        try:
            if isinstance(element, FamilySymbol):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name

                if hasattr(element, "Family") and element.Family:
                    family_name = (
                        element.Family.Name
                        if hasattr(element.Family, "Name") and element.Family.Name
                        else ""
                    )

                    type_name = ""
                    for param_name in ["Тип", "Type Name", "Имя типа"]:
                        param = element.LookupParameter(param_name)
                        if param and param.HasValue:
                            type_name = param.AsString()
                            break

                    if family_name and type_name:
                        return "{} - {}".format(family_name, type_name)
                    elif type_name:
                        return type_name
                    elif family_name:
                        return family_name

                return "Типоразмер " + str(element.Id.IntegerValue)

            elif isinstance(element, Family):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            if hasattr(element, "Name") and element.Name:
                name = element.Name
                if name and not name.startswith("IronPython"):
                    return name

        except Exception as e:
            print("Ошибка получения имени элемента: {}".format(e))

        return "Элемент " + str(element.Id.IntegerValue)

    def OnFamilySelected(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0:
            selected_name = self.lstFamilies.SelectedItem
            selected_family = self.family_dict.get(selected_name)
            if selected_family:
                self.PopulateTypesList(selected_family)

    def PopulateTypesList(self, family):
        self.lstTypes.Items.Clear()
        self.type_dict.clear()

        try:
            symbol_ids = family.GetFamilySymbolIds()
            if symbol_ids and symbol_ids.Count > 0:
                for symbol_id in list(symbol_ids):
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = self.GetElementName(symbol)
                        status = " [Активный]" if symbol.IsActive else " [Не активный]"
                        display_name = "{}{} [ID:{}]".format(symbol_name, status, symbol.Id.IntegerValue)
                        
                        self.lstTypes.Items.Add(display_name)
                        self.type_dict[display_name] = symbol

                # Выбираем первый активный типоразмер
                for i in range(self.lstTypes.Items.Count):
                    display_name = self.lstTypes.Items[i]
                    symbol = self.type_dict.get(display_name)
                    if symbol and symbol.IsActive:
                        self.lstTypes.SelectedIndex = i
                        return

                # Если нет активных, выбираем первый
                if self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
            else:
                MessageBox.Show("В выбранном семействе нет типоразмеров", "Информация")

        except Exception as e:
            MessageBox.Show("Ошибка при загрузке типоразмеров: {}".format(e), "Ошибка")

    def OnOKClick(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0 and self.lstTypes.SelectedIndex >= 0:
            selected_family_name = self.lstFamilies.SelectedItem
            selected_type_name = self.lstTypes.SelectedItem

            self.selected_family = self.family_dict.get(selected_family_name)
            self.selected_type = self.type_dict.get(selected_type_name)

            if self.selected_family and self.selected_type:
                self.DialogResult = DialogResult.OK
                self.Close()
            else:
                MessageBox.Show("Ошибка при получении выбранных объектов!", "Ошибка")
        else:
            MessageBox.Show("Выберите семейство и типоразмер марки!", "Внимание")

    def OnCancelClick(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    @property
    def SelectedFamily(self):
        return self.selected_family

    @property
    def SelectedType(self):
        return self.selected_type


# Запуск
def main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        if doc and uidoc:
            Application.Run(MainForm(doc, uidoc))
        else:
            MessageBox.Show("Нет доступа к документу Revit", "Ошибка")
    except Exception as e:
        MessageBox.Show("Ошибка запуска приложения: {}\n\n{}".format(str(e), traceback.format_exc()), "Критическая ошибка")


if __name__ == "__main__":
    main()