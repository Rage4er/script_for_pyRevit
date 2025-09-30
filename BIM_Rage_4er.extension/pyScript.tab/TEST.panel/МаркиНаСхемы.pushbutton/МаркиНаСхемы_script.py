# -*- coding: utf-8 -*-
__title__ = 'Автоматическое размещение маркировочных меток на 3D-видах'
__author__ = 'г'
__doc__ = ' '

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from System.Windows.Forms import *
from System.Drawing import *
from System import EventHandler
import sys
import os
import datetime

# Глобальная переменная для логов
log_messages = []

def add_log(message):
    """Добавление сообщения в лог"""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_message = "[{0}] {1}".format(timestamp, message)
    log_messages.append(log_message)
    print(log_message)

def show_logs():
    """Показать все логи в MessageBox"""
    log_text = "\n".join(log_messages)
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

# Класс для хранения настроек пользователя
class TagSettings(object):
    def __init__(self):
        self.selected_views = []
        self.selected_categories = []
        self.category_tag_families = {}   # {category: selected_tag_family}
        self.category_tag_types = {}      # {category: selected_tag_type}
        self.offset_x = 60  # мм
        self.offset_y = 30  # мм
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True

# Вспомогательный класс для отображения в ListBox
class FamilyListItem(object):
    def __init__(self, display_name, tag):
        self.display_name = display_name
        self.Tag = tag
    
    def __str__(self):
        return self.display_name

# Форма для выбора семейства и типоразмера марки
class TagFamilySelectionForm(Form):
    def __init__(self, doc, available_families, current_family, current_type):
        add_log("TagFamilySelectionForm: Инициализация формы выбора марки")
        self.doc = doc
        self.available_families = available_families
        self.current_family = current_family
        self.current_type = current_type
        self.selected_family = None
        self.selected_type = None
        
        self.InitializeComponent()
        self.PopulateFamiliesList()
    
    def InitializeComponent(self):
        add_log("TagFamilySelectionForm: Создание компонентов формы")
        self.Text = "Выбор семейства и типоразмера марки"
        self.Size = Size(500, 400)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        
        # Метка для семейств
        self.lblFamilies = Label()
        self.lblFamilies.Text = "Выберите семейство марки:"
        self.lblFamilies.Location = Point(10, 10)
        self.lblFamilies.Size = Size(300, 20)
        
        # Список семейств
        self.lstFamilies = ListBox()
        self.lstFamilies.Location = Point(10, 40)
        self.lstFamilies.Size = Size(200, 200)
        self.lstFamilies.SelectedIndexChanged += self.OnFamilySelected
        
        # Метка для типоразмеров
        self.lblTypes = Label()
        self.lblTypes.Text = "Выберите типоразмер:"
        self.lblTypes.Location = Point(220, 10)
        self.lblTypes.Size = Size(300, 20)
        
        # Список типоразмеров
        self.lstTypes = ListBox()
        self.lstTypes.Location = Point(220, 40)
        self.lstTypes.Size = Size(200, 200)
        
        # Кнопка OK
        self.btnOK = Button()
        self.btnOK.Text = "OK"
        self.btnOK.Location = Point(300, 250)
        self.btnOK.Size = Size(75, 25)
        self.btnOK.Click += self.OnOKClick
        
        # Кнопка Отмена
        self.btnCancel = Button()
        self.btnCancel.Text = "Отмена"
        self.btnCancel.Location = Point(380, 250)
        self.btnCancel.Size = Size(75, 25)
        self.btnCancel.Click += self.OnCancelClick
        
        controls = [self.lblFamilies, self.lstFamilies, self.lblTypes, self.lstTypes, self.btnOK, self.btnCancel]
        for control in controls:
            self.Controls.Add(control)
    
    def PopulateFamiliesList(self):
        """Заполнение списка семейств"""
        add_log("TagFamilySelectionForm: Заполнение списка семейств. Доступно: {0}".format(len(self.available_families)))
        self.lstFamilies.Items.Clear()
        
        for i, family in enumerate(self.available_families):
            if family:
                # Получаем имя семейства
                family_name = self.GetElementName(family)
                add_log("TagFamilySelectionForm: Семейство {0}: {1}".format(i, family_name))
                # Создаем объект для хранения и семейства и его имени
                item = FamilyListItem(family_name, family)
                self.lstFamilies.Items.Add(item)
        
        # Выбираем текущее семейство если есть
        if self.current_family:
            for i in range(self.lstFamilies.Items.Count):
                if self.lstFamilies.Items[i].Tag == self.current_family:
                    self.lstFamilies.SelectedIndex = i
                    add_log("TagFamilySelectionForm: Выбрано текущее семейство с индексом {0}".format(i))
                    break
        # Иначе выбираем первое семейство
        elif self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0
            add_log("TagFamilySelectionForm: Выбрано первое семейство по умолчанию")
    
    def OnFamilySelected(self, sender, args):
        """Обработка выбора семейства"""
        if self.lstFamilies.SelectedItem:
            selected_item = self.lstFamilies.SelectedItem
            selected_family = selected_item.Tag
            add_log("TagFamilySelectionForm: Выбрано семейство: {0}".format(self.GetElementName(selected_family)))
            self.PopulateTypesList(selected_family)
    
    def PopulateTypesList(self, family):
        """Заполнение списка типоразмеров для выбранного семейства"""
        add_log("TagFamilySelectionForm: Загрузка типоразмеров для семейства {0}".format(self.GetElementName(family)))
        self.lstTypes.Items.Clear()
        
        try:
            # Получаем все типоразмеры выбранного семейства
            symbol_ids = family.GetFamilySymbolIds()
            add_log("TagFamilySelectionForm: Найдено типоразмеров: {0}".format(symbol_ids.Count if symbol_ids else 0))
            
            if symbol_ids and symbol_ids.Count > 0:
                symbol_id_list = list(symbol_ids)
                for i, symbol_id in enumerate(symbol_id_list):
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        # Получаем имя типоразмера
                        symbol_name = self.GetElementName(symbol)
                        add_log("TagFamilySelectionForm: Типоразмер {0}: {1}".format(i, symbol_name))
                        # Создаем объект для хранения и типоразмера и его имени
                        type_item = FamilyListItem(symbol_name, symbol)
                        self.lstTypes.Items.Add(type_item)
                
                # Выбираем текущий типоразмер если есть
                if self.current_type:
                    for i in range(self.lstTypes.Items.Count):
                        if self.lstTypes.Items[i].Tag == self.current_type:
                            self.lstTypes.SelectedIndex = i
                            add_log("TagFamilySelectionForm: Выбран текущий типоразмер с индексом {0}".format(i))
                            break
                # Иначе выбираем первый типоразмер
                elif self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
                    add_log("TagFamilySelectionForm: Выбран первый типоразмер по умолчанию")
            else:
                add_log("TagFamilySelectionForm: В выбранном семействе нет типоразмеров")
                MessageBox.Show("В выбранном семействе нет типоразмеров")
                    
        except Exception as e:
            add_log("TagFamilySelectionForm: Ошибка при загрузке типоразмеров: {0}".format(str(e)))
            MessageBox.Show("Ошибка при загрузке типоразмеров")
    
    def GetElementName(self, element):
        """Получение корректного имени элемента"""
        if not element:
            return "Без имени"
        
        try:
            # Для FamilySymbol (типоразмеров)
            if isinstance(element, FamilySymbol):
                # Пробуем получить имя типа
                if hasattr(element, 'Name') and element.Name:
                    return element.Name
                
                # Или через параметры
                param = element.LookupParameter("Тип")
                if param and param.HasValue:
                    return param.AsString()
                
                param = element.LookupParameter("Type")
                if param and param.HasValue:
                    return param.AsString()
            
            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, 'Name') and element.Name:
                    return element.Name
            
            # Общий случай
            if hasattr(element, 'Name') and element.Name:
                return element.Name
                
        except Exception as e:
            add_log("TagFamilySelectionForm: Ошибка при получении имени элемента: {0}".format(str(e)))
        
        return "Без имени"
    
    def OnOKClick(self, sender, args):
        """Обработка нажатия OK"""
        if self.lstFamilies.SelectedItem and self.lstTypes.SelectedItem:
            self.selected_family = self.lstFamilies.SelectedItem.Tag
            self.selected_type = self.lstTypes.SelectedItem.Tag
            add_log("TagFamilySelectionForm: Выбрано - Семейство: {0}, Тип: {1}".format(
                self.GetElementName(self.selected_family), 
                self.GetElementName(self.selected_type)))
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            add_log("TagFamilySelectionForm: Не выбрано семейство или типоразмер")
            MessageBox.Show("Выберите семейство и типоразмер марки!")
    
    def OnCancelClick(self, sender, args):
        """Обработка нажатия Отмена"""
        add_log("TagFamilySelectionForm: Отмена выбора")
        self.DialogResult = DialogResult.Cancel
        self.Close()
    
    @property
    def SelectedFamily(self):
        return self.selected_family
    
    @property
    def SelectedType(self):
        return self.selected_type

# Главное окно приложения
class MainForm(Form):
    def __init__(self, doc, uidoc):
        add_log("MainForm: Инициализация главного окна")
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}  # Словарь для хранения соответствия названий и объектов View
        self.category_mapping = {}  # Словарь для хранения соответствия названий категорий и объектов Category
        
        self.InitializeComponent()
        self.Load3DViewsWithDuctSystems()
    
    def InitializeComponent(self):
        add_log("MainForm: Создание компонентов интерфейса")
        self.Text = "Расстановка марок на 3D видах"
        self.Size = Size(800, 600)
        self.StartPosition = FormStartPosition.CenterScreen
        
        # Создание вкладок
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        
        # Вкладка 1: Выбор видов
        self.tabViews = TabPage()
        self.tabViews.Text = "1. Выбор видов"
        
        # Вкладка 2: Выбор категорий
        self.tabCategories = TabPage()
        self.tabCategories.Text = "2. Категории"
        
        # Вкладка 3: Выбор семейств марок
        self.tabTags = TabPage()
        self.tabTags.Text = "3. Марки"
        
        # Вкладка 4: Настройки
        self.tabSettings = TabPage()
        self.tabSettings.Text = "4. Настройки"
        
        # Вкладка 5: Выполнение
        self.tabExecute = TabPage()
        self.tabExecute.Text = "5. Выполнение"
        
        self.tabControl.TabPages.Add(self.tabViews)
        self.tabControl.TabPages.Add(self.tabCategories)
        self.tabControl.TabPages.Add(self.tabTags)
        self.tabControl.TabPages.Add(self.tabSettings)
        self.tabControl.TabPages.Add(self.tabExecute)
        
        self.Controls.Add(self.tabControl)
        self.SetupViewsTab()
        self.SetupCategoriesTab()
        self.SetupTagsTab()
        self.SetupSettingsTab()
        self.SetupExecuteTab()
    
    def SetupViewsTab(self):
        add_log("MainForm: Настройка вкладки выбора видов")
        # Список 3D видов
        self.lblViews = Label()
        self.lblViews.Text = "Выберите 3D виды с системами воздуховодов:"
        self.lblViews.Location = Point(10, 10)
        self.lblViews.Size = Size(300, 20)
        
        self.lstViews = CheckedListBox()
        self.lstViews.Location = Point(10, 40)
        self.lstViews.Size = Size(500, 300)
        self.lstViews.CheckOnClick = True
        
        # Кнопка продолжить
        self.btnNext1 = Button()
        self.btnNext1.Text = "Далее →"
        self.btnNext1.Location = Point(400, 350)
        self.btnNext1.Click += self.OnNext1Click
        
        # Кнопка показа логов
        self.btnShowLogs = Button()
        self.btnShowLogs.Text = "Показать логи"
        self.btnShowLogs.Location = Point(10, 350)
        self.btnShowLogs.Click += self.OnShowLogsClick
        
        controls = [self.lblViews, self.lstViews, self.btnNext1, self.btnShowLogs]
        for control in controls:
            self.tabViews.Controls.Add(control)
    
    def SetupCategoriesTab(self):
        add_log("MainForm: Настройка вкладки выбора категорий")
        self.lblCategories = Label()
        self.lblCategories.Text = "Выберите категории элементов:"
        self.lblCategories.Location = Point(10, 10)
        self.lblCategories.Size = Size(300, 20)
        
        # ListBox для категорий
        self.lstCategories = CheckedListBox()
        self.lstCategories.Location = Point(10, 40)
        self.lstCategories.Size = Size(500, 300)
        self.lstCategories.CheckOnClick = True
        
        self.btnNext2 = Button()
        self.btnNext2.Text = "Далее →"
        self.btnNext2.Location = Point(400, 350)
        self.btnNext2.Click += self.OnNext2Click
        
        self.btnBack1 = Button()
        self.btnBack1.Text = "← Назад"
        self.btnBack1.Location = Point(300, 350)
        self.btnBack1.Click += self.OnBack1Click
        
        controls = [
            self.lblCategories, self.lstCategories, 
            self.btnNext2, self.btnBack1
        ]
        for control in controls:
            self.tabCategories.Controls.Add(control)
    
    def SetupTagsTab(self):
        add_log("MainForm: Настройка вкладки выбора марок")
        self.lblTags = Label()
        self.lblTags.Text = "Выберите семейства и типоразмеры марок для категорий:"
        self.lblTags.Location = Point(10, 10)
        self.lblTags.Size = Size(300, 20)
        
        # ListView для выбора семейств и типоразмеров
        self.lstTagFamilies = ListView()
        self.lstTagFamilies.Location = Point(10, 40)
        self.lstTagFamilies.Size = Size(500, 300)
        self.lstTagFamilies.View = View.Details
        self.lstTagFamilies.FullRowSelect = True
        self.lstTagFamilies.Columns.Add("Категория", 150)
        self.lstTagFamilies.Columns.Add("Семейство марки", 200)
        self.lstTagFamilies.Columns.Add("Типоразмер марки", 150)
        
        # Добавляем обработчик двойного клика для изменения марки
        self.lstTagFamilies.DoubleClick += self.OnTagFamilyDoubleClick
        
        self.btnNext3 = Button()
        self.btnNext3.Text = "Далее →"
        self.btnNext3.Location = Point(400, 350)
        self.btnNext3.Click += self.OnNext3Click
        
        self.btnBack2 = Button()
        self.btnBack2.Text = "← Назад"
        self.btnBack2.Location = Point(300, 350)
        self.btnBack2.Click += self.OnBack2Click
        
        controls = [
            self.lblTags, self.lstTagFamilies, 
            self.btnNext3, self.btnBack2
        ]
        for control in controls:
            self.tabTags.Controls.Add(control)
    
    def SetupSettingsTab(self):
        add_log("MainForm: Настройка вкладки настроек")
        self.lblSettings = Label()
        self.lblSettings.Text = "Настройки размещения марок:"
        self.lblSettings.Location = Point(10, 10)
        self.lblSettings.Size = Size(300, 20)
        
        # Смещение X
        self.lblOffsetX = Label()
        self.lblOffsetX.Text = "Смещение по X (мм):"
        self.lblOffsetX.Location = Point(10, 50)
        self.lblOffsetX.Size = Size(150, 20)
        
        self.txtOffsetX = TextBox()
        self.txtOffsetX.Text = "60"
        self.txtOffsetX.Location = Point(170, 50)
        self.txtOffsetX.Size = Size(100, 20)
        
        # Смещение Y
        self.lblOffsetY = Label()
        self.lblOffsetY.Text = "Смещение по Y (мм):"
        self.lblOffsetY.Location = Point(10, 80)
        self.lblOffsetY.Size = Size(150, 20)
        
        self.txtOffsetY = TextBox()
        self.txtOffsetY.Text = "30"
        self.txtOffsetY.Location = Point(170, 80)
        self.txtOffsetY.Size = Size(100, 20)
        
        # Ориентация
        self.lblOrientation = Label()
        self.lblOrientation.Text = "Ориентация марки:"
        self.lblOrientation.Location = Point(10, 110)
        self.lblOrientation.Size = Size(150, 20)
        
        self.cmbOrientation = ComboBox()
        self.cmbOrientation.Location = Point(170, 110)
        self.cmbOrientation.Size = Size(100, 20)
        self.cmbOrientation.Items.Add("Горизонтальная")
        self.cmbOrientation.Items.Add("Вертикальная")
        self.cmbOrientation.SelectedIndex = 0
        
        # Использование выноски
        self.chkUseLeader = CheckBox()
        self.chkUseLeader.Text = "Использовать выноску"
        self.chkUseLeader.Location = Point(10, 140)
        self.chkUseLeader.Size = Size(200, 20)
        self.chkUseLeader.Checked = True
        
        self.btnNext4 = Button()
        self.btnNext4.Text = "Далее →"
        self.btnNext4.Location = Point(400, 350)
        self.btnNext4.Click += self.OnNext4Click
        
        self.btnBack3 = Button()
        self.btnBack3.Text = "← Назад"
        self.btnBack3.Location = Point(300, 350)
        self.btnBack3.Click += self.OnBack3Click
        
        controls = [
            self.lblSettings, self.lblOffsetX, self.txtOffsetX,
            self.lblOffsetY, self.txtOffsetY, self.lblOrientation,
            self.cmbOrientation, self.chkUseLeader,
            self.btnNext4, self.btnBack3
        ]
        for control in controls:
            self.tabSettings.Controls.Add(control)
    
    def SetupExecuteTab(self):
        add_log("MainForm: Настройка вкладки выполнения")
        self.lblExecute = Label()
        self.lblExecute.Text = "Готово к выполнению:"
        self.lblExecute.Location = Point(10, 10)
        self.lblExecute.Size = Size(300, 20)
        
        self.txtSummary = TextBox()
        self.txtSummary.Location = Point(10, 40)
        self.txtSummary.Size = Size(500, 200)
        self.txtSummary.Multiline = True
        self.txtSummary.ScrollBars = ScrollBars.Vertical
        self.txtSummary.ReadOnly = True
        
        self.btnExecute = Button()
        self.btnExecute.Text = "Выполнить расстановку"
        self.btnExecute.Location = Point(300, 250)
        self.btnExecute.Size = Size(150, 30)
        self.btnExecute.Click += self.OnExecuteClick
        
        self.btnBack4 = Button()
        self.btnBack4.Text = "← Назад"
        self.btnBack4.Location = Point(150, 250)
        self.btnBack4.Click += self.OnBack4Click
        
        controls = [
            self.lblExecute, self.txtSummary, 
            self.btnExecute, self.btnBack4
        ]
        for control in controls:
            self.tabExecute.Controls.Add(control)
    
    def Load3DViewsWithDuctSystems(self):
        """Загрузка 3D видов, на которых отображены системы воздуховодов"""
        try:
            add_log("Load3DViewsWithDuctSystems: Начало загрузки 3D видов")
            collector = FilteredElementCollector(self.doc)
            views = collector.OfClass(View3D).WhereElementIsNotElementType().ToElements()
            
            self.lstViews.Items.Clear()
            self.views_dict = {}
            
            view_count = 0
            for view in views:
                # Проверяем, что вид не является шаблоном и может быть напечатан
                if not view.IsTemplate and view.CanBePrinted:
                    display_name = "{0} (ID: {1})".format(view.Name, view.Id.IntegerValue)
                    self.lstViews.Items.Add(display_name, False)
                    self.views_dict[display_name] = view
                    view_count += 1
            
            add_log("Load3DViewsWithDuctSystems: Загружено 3D видов: {0}".format(view_count))
            
            # Если не найдено видов, показываем сообщение
            if self.lstViews.Items.Count == 0:
                add_log("Load3DViewsWithDuctSystems: Не найдено 3D видов в проекте")
                MessageBox.Show("Не найдено 3D видов в проекте")
            else:
                # Автоматически выбираем первый вид для удобства
                if self.lstViews.Items.Count > 0:
                    self.lstViews.SetItemChecked(0, True)
                    add_log("Load3DViewsWithDuctSystems: Автоматически выбран первый вид")
                
        except Exception as e:
            add_log("Load3DViewsWithDuctSystems: Ошибка при загрузке видов: {0}".format(str(e)))
            MessageBox.Show("Ошибка при загрузке видов: " + str(e))
    
    def OnNext1Click(self, sender, args):
        """Переход от выбора видов к выбору категорий"""
        add_log("OnNext1Click: Переход к выбору категорий")
        # Сохраняем выбранные виды
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            if self.lstViews.GetItemChecked(i):
                display_name = self.lstViews.Items[i]
                if display_name in self.views_dict:
                    self.settings.selected_views.append(self.views_dict[display_name])
        
        add_log("OnNext1Click: Выбрано видов: {0}".format(len(self.settings.selected_views)))
        
        if not self.settings.selected_views:
            add_log("OnNext1Click: Не выбрано ни одного вида")
            MessageBox.Show("Выберите хотя бы один вид!")
            return
        
        # Собираем уникальные категории из выбранных видов
        self.CollectUniqueCategories()
        self.PopulateCategoriesList()
        
        self.tabControl.SelectedTab = self.tabCategories

    def CollectUniqueCategories(self):
        """Сбор уникальных категорий из выбранных видов"""
        add_log("CollectUniqueCategories: Сбор категорий")
        unique_categories = set()
        
        # Целевые категории для систем воздуховодов
        target_categories = [
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory
        ]
        
        add_log("CollectUniqueCategories: Целевые категории: {0}".format(len(target_categories)))
        
        # Просто добавляем все целевые категории
        for category in target_categories:
            try:
                cat_obj = Category.GetCategory(self.doc, category)
                if cat_obj:
                    unique_categories.add(cat_obj)
                    add_log("CollectUniqueCategories: Добавлена категория: {0}".format(self.GetCategoryDisplayName(cat_obj)))
            except Exception as e:
                add_log("CollectUniqueCategories: Ошибка при получении категории {0}: {1}".format(category, str(e)))
        
        # Сохраняем уникальные категории
        self.settings.selected_categories = list(unique_categories)
        add_log("CollectUniqueCategories: Всего собрано категорий: {0}".format(len(self.settings.selected_categories)))
    
    def PopulateCategoriesList(self):
        """Заполнение списка категорий"""
        add_log("PopulateCategoriesList: Заполнение списка категорий")
        self.lstCategories.Items.Clear()
        self.category_mapping = {}  # Очищаем словарь соответствий
        
        # Сортируем категории по имени для удобства
        sorted_categories = sorted(self.settings.selected_categories, key=lambda x: self.GetCategoryDisplayName(x))
        
        for category in sorted_categories:
            # Отображаем корректное имя категории
            display_name = self.GetCategoryDisplayName(category)
            self.lstCategories.Items.Add(display_name, True)  # По умолчанию выбираем все
            # Сохраняем соответствие в словаре
            self.category_mapping[display_name] = category
            add_log("PopulateCategoriesList: Добавлена категория в список: {0}".format(display_name))
    
    def GetCategoryDisplayName(self, category):
        """Получение корректного отображаемого имени категории"""
        if not category:
            return "Неизвестная категория"
        
        # Для встроенных категорий используем BuiltInCategory
        try:
            if hasattr(category, 'Id') and category.Id.IntegerValue < 0:
                built_in_cat = BuiltInCategory(category.Id.IntegerValue)
                return LabelUtils.GetLabelFor(built_in_cat)
        except:
            pass
        
        # Для остальных категорий используем Name
        return getattr(category, 'Name', 'Неизвестная категория')
    
    def OnNext2Click(self, sender, args):
        """Переход от категорий к выбору марок"""
        add_log("OnNext2Click: Переход к выбору марок")
        # Сохраняем выбранные категории
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            if self.lstCategories.GetItemChecked(i):
                display_name = self.lstCategories.Items[i]
                if display_name in self.category_mapping:
                    self.settings.selected_categories.append(self.category_mapping[display_name])
        
        add_log("OnNext2Click: Выбрано категорий: {0}".format(len(self.settings.selected_categories)))
        
        if not self.settings.selected_categories:
            add_log("OnNext2Click: Не выбрано ни одной категории")
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return
        
        # Загружаем доступные семейства марок
        self.PopulateTagFamilies()
        self.tabControl.SelectedTab = self.tabTags
    
    def PopulateTagFamilies(self):
        """Заполнение списка доступных семейств марок"""
        add_log("PopulateTagFamilies: Заполнение списка марок")
        self.lstTagFamilies.Items.Clear()
        
        # Заполняем список для выбранных категорий
        for category in self.settings.selected_categories:
            cat_name = self.GetCategoryDisplayName(category)
            item = ListViewItem(cat_name)
            item.Tag = category  # Сохраняем категорию в Tag
            
            # Ищем подходящие семейства марок
            add_log("PopulateTagFamilies: Поиск марки для категории: {0}".format(cat_name))
            tag_family, tag_type = self.FindSuitableTagForCategory(category)
            
            if tag_family and tag_type:
                family_name = self.GetElementName(tag_family)
                type_name = self.GetElementName(tag_type)
                item.SubItems.Add(family_name)
                item.SubItems.Add(type_name)
                self.settings.category_tag_families[category] = tag_family
                self.settings.category_tag_types[category] = tag_type
                add_log("PopulateTagFamilies: Найдена марка для {0}: {1} ({2})".format(cat_name, family_name, type_name))
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")
                self.settings.category_tag_families[category] = None
                self.settings.category_tag_types[category] = None
                add_log("PopulateTagFamilies: Не найдено марок для категории: {0}".format(cat_name))
            
            self.lstTagFamilies.Items.Add(item)
    
    def FindSuitableTagForCategory(self, category):
        """Поиск подходящего семейства и типоразмера марки для категории"""
        add_log("FindSuitableTagForCategory: Поиск марки для категории: {0}".format(self.GetCategoryDisplayName(category)))
        
        # Определяем категорию марки в зависимости от категории элемента
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            add_log("FindSuitableTagForCategory: Не найдена категория марки для {0}".format(self.GetCategoryDisplayName(category)))
            return None, None
        
        add_log("FindSuitableTagForCategory: Категория марки ID: {0}".format(tag_category_id.IntegerValue))
        
        # Получаем все семейства марок нужной категории в проекте
        collector = FilteredElementCollector(self.doc)
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        
        add_log("FindSuitableTagForCategory: Всего семейств в проекте: {0}".format(len(list(tag_families))))
        
        # Ищем семейства марки для данной категории
        for family in tag_families:
            if not family or not hasattr(family, 'FamilyCategory'):
                continue
                
            add_log("FindSuitableTagForCategory: Проверяем семейство: {0}".format(self.GetElementName(family)))
                
            if family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                add_log("FindSuitableTagForCategory: Нашли подходящее семейство: {0}".format(self.GetElementName(family)))
                
                # Получаем все доступные типоразмеры
                symbol_ids = family.GetFamilySymbolIds()
                add_log("FindSuitableTagForCategory: Типоразмеров в семействе: {0}".format(symbol_ids.Count if symbol_ids else 0))
                
                if symbol_ids and symbol_ids.Count > 0:
                    # Преобразуем HashSet в список для безопасного доступа
                    symbol_id_list = list(symbol_ids)
                    if symbol_id_list:
                        # Берем первый активный типоразмер
                        for symbol_id in symbol_id_list:
                            tag_type = self.doc.GetElement(symbol_id)
                            if tag_type and tag_type.IsActive:
                                add_log("FindSuitableTagForCategory: Найден активный типоразмер: {0}".format(self.GetElementName(tag_type)))
                                return family, tag_type
                        # Если нет активных, берем первый доступный
                        tag_type = self.doc.GetElement(symbol_id_list[0])
                        add_log("FindSuitableTagForCategory: Используем первый типоразмер: {0}".format(self.GetElementName(tag_type)))
                        return family, tag_type
        
        add_log("FindSuitableTagForCategory: Не найдено подходящих марок")
        return None, None
    
    def GetTagCategoryForElementCategory(self, element_category):
        """Определяет категорию марки для категории элемента"""
        if not element_category:
            return None
        
        # Сопоставление категорий элементов с категориями марок
        category_mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags
        }
        
        # Получаем BuiltInCategory для элемента
        try:
            if hasattr(element_category, 'Id') and element_category.Id.IntegerValue < 0:
                element_builtin_cat = BuiltInCategory(element_category.Id.IntegerValue)
                add_log("GetTagCategoryForElementCategory: Категория элемента: {0}".format(element_builtin_cat))
                if element_builtin_cat in category_mapping:
                    tag_builtin_cat = category_mapping[element_builtin_cat]
                    tag_category = Category.GetCategory(self.doc, tag_builtin_cat)
                    if tag_category:
                        add_log("GetTagCategoryForElementCategory: Найдена категория марки: {0}".format(tag_builtin_cat))
                        return tag_category.Id
        except Exception as e:
            add_log("GetTagCategoryForElementCategory: Ошибка при получении категории марки: {0}".format(str(e)))
        
        add_log("GetTagCategoryForElementCategory: Категория марки не найдена")
        return None
    
    def GetElementName(self, element):
        """Получение корректного имени элемента"""
        if not element:
            return "Без имени"
        
        try:
            # Для FamilySymbol (типоразмеров)
            if isinstance(element, FamilySymbol):
                # Пробуем получить имя типа
                if hasattr(element, 'Name') and element.Name:
                    return element.Name
                
                # Или через параметры
                param = element.LookupParameter("Тип")
                if param and param.HasValue:
                    return param.AsString()
                
                param = element.LookupParameter("Type")
                if param and param.HasValue:
                    return param.AsString()
            
            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, 'Name') and element.Name:
                    return element.Name
            
            # Общий случай
            if hasattr(element, 'Name') and element.Name:
                return element.Name
                
        except Exception as e:
            add_log("GetElementName: Ошибка при получении имени элемента: {0}".format(str(e)))
        
        return "Без имени"
    
    def OnTagFamilyDoubleClick(self, sender, args):
        """Обработка двойного клика для изменения семейства марки"""
        if self.lstTagFamilies.SelectedItems.Count > 0:
            selected_item = self.lstTagFamilies.SelectedItems[0]
            category = selected_item.Tag
            add_log("OnTagFamilyDoubleClick: Двойной клик по категории: {0}".format(self.GetCategoryDisplayName(category)))
            
            # Получаем все доступные семейства марок для этой категории
            available_families = self.GetAvailableTagFamiliesForCategory(category)
            add_log("OnTagFamilyDoubleClick: Доступно семейств марок: {0}".format(len(available_families)))
            
            if available_families:
                # Создаем диалог для выбора семейства и типоразмера
                current_family = self.settings.category_tag_families.get(category)
                current_type = self.settings.category_tag_types.get(category)
                form = TagFamilySelectionForm(self.doc, available_families, current_family, current_type)
                result = form.ShowDialog()
                
                if result == DialogResult.OK and form.SelectedFamily and form.SelectedType:
                    # Обновляем выбранное семейство и типоразмер
                    family_name = self.GetElementName(form.SelectedFamily)
                    type_name = self.GetElementName(form.SelectedType)
                    selected_item.SubItems[1].Text = family_name
                    selected_item.SubItems[2].Text = type_name
                    self.settings.category_tag_families[category] = form.SelectedFamily
                    self.settings.category_tag_types[category] = form.SelectedType
                    add_log("OnTagFamilyDoubleClick: Обновлена марка для {0}: {1} ({2})".format(
                        self.GetCategoryDisplayName(category), family_name, type_name))
            else:
                add_log("OnTagFamilyDoubleClick: Нет доступных семейств марок для {0}".format(self.GetCategoryDisplayName(category)))
                MessageBox.Show("Нет доступных семейств марок для этой категории")

    def GetAvailableTagFamiliesForCategory(self, category):
        """Получение всех доступных семейств марок для категории"""
        add_log("GetAvailableTagFamiliesForCategory: Поиск марок для {0}".format(self.GetCategoryDisplayName(category)))
        # Определяем категорию марки в зависимости от категории элемента
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            add_log("GetAvailableTagFamiliesForCategory: Не найдена категория марки для: {0}".format(self.GetCategoryDisplayName(category)))
            return []
        
        try:
            tag_category = Category.GetCategory(self.doc, BuiltInCategory(tag_category_id.IntegerValue))
            add_log("GetAvailableTagFamiliesForCategory: Ищем марки категории: {0}".format(tag_category.Name if tag_category else 'Unknown'))
        except:
            add_log("GetAvailableTagFamiliesForCategory: Ищем марки категории ID: {0}".format(tag_category_id.IntegerValue))
        
        collector = FilteredElementCollector(self.doc)
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        
        available_families = []
        
        for family in tag_families:
            if not family or not hasattr(family, 'FamilyCategory'):
                continue
                
            if family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                available_families.append(family)
                # Проверим сразу есть ли типоразмеры
                symbol_ids = family.GetFamilySymbolIds()
                add_log("GetAvailableTagFamiliesForCategory: Семейство: {0} (типоразмеров: {1})".format(
                    self.GetElementName(family), 
                    symbol_ids.Count if symbol_ids else 0))
        
        add_log("GetAvailableTagFamiliesForCategory: Итого найдено семейств марок для категории {0}: {1}".format(
            self.GetCategoryDisplayName(category), len(available_families)))
        
        return available_families
    
    def OnNext3Click(self, sender, args):
        """Переход к настройкам"""
        add_log("OnNext3Click: Переход к настройкам")
        self.tabControl.SelectedTab = self.tabSettings
    
    def OnNext4Click(self, sender, args):
        """Переход к выполнению"""
        add_log("OnNext4Click: Переход к выполнению")
        # Сохраняем настройки
        try:
            self.settings.offset_x = float(self.txtOffsetX.Text)
            self.settings.offset_y = float(self.txtOffsetY.Text)
            if self.cmbOrientation.SelectedIndex == 0:
                self.settings.orientation = TagOrientation.Horizontal
            else:
                self.settings.orientation = TagOrientation.Vertical
            self.settings.use_leader = self.chkUseLeader.Checked
            add_log("OnNext4Click: Настройки сохранены - X:{0}, Y:{1}, ориентация:{2}, выноска:{3}".format(
                self.settings.offset_x, self.settings.offset_y, self.settings.orientation, self.settings.use_leader))
        except Exception as e:
            add_log("OnNext4Click: Ошибка сохранения настроек: {0}".format(str(e)))
            MessageBox.Show("Проверьте правильность введенных значений! Ошибка: " + str(e))
            return
        
        # Формируем сводку
        summary = self.GenerateSummary()
        self.txtSummary.Text = summary
        
        self.tabControl.SelectedTab = self.tabExecute
    
    def GenerateSummary(self):
        """Генерация сводки перед выполнением"""
        add_log("GenerateSummary: Формирование сводки")
        summary = "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ:\n\n"
        summary += "Выбрано видов: {0}\n".format(len(self.settings.selected_views))
        summary += "Выбрано категорий: {0}\n".format(len(self.settings.selected_categories))
        summary += "Смещение: X={0} мм, Y={1} мм\n".format(self.settings.offset_x, self.settings.offset_y)
        
        if self.settings.orientation == TagOrientation.Horizontal:
            orientation_text = "Горизонтальная"
        else:
            orientation_text = "Вертикальная"
        summary += "Ориентация: {0}\n".format(orientation_text)
        summary += "Выноска: {0}\n\n".format("Да" if self.settings.use_leader else "Нет")
        
        summary += "Детали по категориям:\n"
        for category in self.settings.selected_categories:
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            cat_name = self.GetCategoryDisplayName(category)
            
            if tag_family and tag_type:
                family_name = self.GetElementName(tag_family)
                type_name = self.GetElementName(tag_type)
                summary += "- {0}: {1} ({2})\n".format(cat_name, family_name, type_name)
            else:
                summary += "- {0}: НЕТ МАРКИ\n".format(cat_name)
        
        add_log("GenerateSummary: Сводка сформирована")
        return summary
    
    def OnExecuteClick(self, sender, args):
        """Выполнение расстановки марок"""
        add_log("OnExecuteClick: Начало выполнения расстановки марок")
        errors = []
        success_count = 0
        
        add_log("OnExecuteClick: Выбрано видов: {0}".format(len(self.settings.selected_views)))
        add_log("OnExecuteClick: Выбрано категорий: {0}".format(len(self.settings.selected_categories)))
        
        # Начинаем транзакцию
        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        add_log("OnExecuteClick: Транзакция начата")
        
        try:
            for view in self.settings.selected_views:
                add_log("OnExecuteClick: Обработка вида: {0} (ID: {1})".format(view.Name, view.Id))
                
                if not isinstance(view, View3D):
                    error_msg = "Вид '{0}' не является 3D видом".format(view.Name)
                    add_log("OnExecuteClick: {0}".format(error_msg))
                    errors.append(error_msg)
                    continue
                
                # Получаем элементы по категориям
                for category in self.settings.selected_categories:
                    cat_name = self.GetCategoryDisplayName(category)
                    add_log("OnExecuteClick: Обработка категории: {0}".format(cat_name))
                    
                    collector = FilteredElementCollector(self.doc, view.Id)
                    elements = collector.OfCategoryId(category.Id).WhereElementIsNotElementType().ToElements()
                    
                    element_list = list(elements)
                    add_log("OnExecuteClick: Найдено элементов категории {0}: {1}".format(cat_name, len(element_list)))
                    
                    for element in element_list:
                        if not element:
                            continue
                        
                        add_log("OnExecuteClick: Обработка элемента ID: {0}, Категория: {1}".format(element.Id, element.Category.Name))
                        
                        # Проверяем, есть ли уже марка у элемента на этом виде
                        if self.HasExistingTag(element, view):
                            add_log("OnExecuteClick: Марка уже существует для элемента {0}".format(element.Id))
                            continue
                        
                        # Получаем семейство и типоразмер марки для категории
                        tag_family = self.settings.category_tag_families.get(category)
                        tag_type = self.settings.category_tag_types.get(category)
                        
                        if not tag_family or not tag_type:
                            error_msg = "Нет марки для категории {0}".format(cat_name)
                            add_log("OnExecuteClick: {0}".format(error_msg))
                            errors.append(error_msg)
                            continue
                        
                        add_log("OnExecuteClick: Создание марки для элемента {0}".format(element.Id))
                        # Создаем марку - ИСПРАВЛЕННЫЙ ВАРИАНТ
                        if self.CreateTagImproved(element, view, tag_type):
                            success_count += 1
                            add_log("OnExecuteClick: Успешно создана марка для элемента {0}".format(element.Id))
                        else:
                            error_msg = "Не удалось создать марку для {0}".format(element.Id)
                            add_log("OnExecuteClick: {0}".format(error_msg))
                            errors.append(error_msg)
            
            trans.Commit()
            add_log("OnExecuteClick: Транзакция завершена успешно. Создано марок: {0}".format(success_count))
            
        except Exception as e:
            trans.RollBack()
            error_msg = "Ошибка выполнения: {0}".format(str(e))
            add_log("OnExecuteClick: {0}".format(error_msg))
            errors.append(error_msg)
        finally:
            trans.Dispose()
            add_log("OnExecuteClick: Транзакция disposed")
        
        # Показываем результаты
        result_msg = "Успешно расставлено марок: {0}\n".format(success_count)
        if errors:
            result_msg += "\nОшибки ({0}):\n".format(len(errors)) + "\n".join(errors[:10])
            if len(errors) > 10:
                result_msg += "\n... и еще {0} ошибок".format(len(errors) - 10)
        
        add_log("OnExecuteClick: Результат: {0}".format(result_msg))
        MessageBox.Show(result_msg, "Результат выполнения")
        
        # Показываем логи после выполнения
        show_logs()
        
        self.Close()
    
    def CreateTagImproved(self, element, view, tag_type):
        """Улучшенное создание марки на основе кода из Dynamo"""
        try:
            add_log("CreateTagImproved: Создание марки для элемента {0}".format(element.Id))
            
            # Получаем точку для размещения марки
            bbox = element.get_BoundingBox(view)
            if not bbox:
                add_log("CreateTagImproved: Нет bounding box у элемента {0}".format(element.Id))
                return False
            
            center = (bbox.Min + bbox.Max) / 2
            
            # Конвертация мм в футы (1 фут = 304.8 мм)
            offset_x_feet = self.settings.offset_x / 304.8
            offset_y_feet = self.settings.offset_y / 304.8
            
            # Создаем точку смещения от центра элемента
            tag_point = XYZ(
                center.X + offset_x_feet,
                center.Y + offset_y_feet,
                center.Z
            )
            
            add_log("CreateTagImproved: Точка размещения: {0}".format(tag_point))
            
            # Создаем ссылку на элемент - КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ
            element_ref = Reference(element)
            add_log("CreateTagImproved: Ссылка создана")
            
            # Используем правильный TagMode как в Dynamo
            tag_mode = TagMode.TM_ADDBY_CATEGORY
            
            # Создаем марку - ИСПРАВЛЕННЫЙ ВЫЗОВ
            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                element_ref,
                self.settings.use_leader,
                tag_mode,
                self.settings.orientation,
                tag_point
            )
            
            if tag:
                add_log("CreateTagImproved: Марка создана успешно, ID: {0}".format(tag.Id))
                
                # Устанавливаем нужный типоразмер
                if tag_type:
                    try:
                        tag.ChangeTypeId(tag_type.Id)
                        add_log("CreateTagImproved: Типоразмер установлен: {0}".format(self.GetElementName(tag_type)))
                    except Exception as e:
                        add_log("CreateTagImproved: Ошибка при установке типоразмера: {0}".format(str(e)))
                
                return True
            else:
                add_log("CreateTagImproved: Не удалось создать марку")
                return False
                
        except Exception as e:
            add_log("CreateTagImproved: Ошибка создания марки: {0}".format(str(e)))
            return False
    
    def HasExistingTag(self, element, view):
        """Проверка существующей марки у элемента на виде"""
        try:
            collector = FilteredElementCollector(self.doc, view.Id)
            tags = collector.OfClass(IndependentTag).ToElements()
            
            tag_count = 0
            for tag in tags:
                tag_count += 1
                try:
                    # Проверяем все привязанные элементы
                    tagged_elements = tag.GetTaggedLocalElements()
                    for tagged_elem in tagged_elements:
                        if tagged_elem.Id == element.Id:
                            add_log("HasExistingTag: Найдена существующая марка для элемента {0}".format(element.Id))
                            return True
                except:
                    # Альтернативный способ для старых версий Revit
                    if hasattr(tag, 'TaggedLocalElementId'):
                        if tag.TaggedLocalElementId == element.Id:
                            add_log("HasExistingTag: Найдена существующая марка для элемента {0} (старый метод)".format(element.Id))
                            return True
            add_log("HasExistingTag: Проверено {0} марок, существующих марок не найдено для элемента {1}".format(tag_count, element.Id))
            return False
        except Exception as e:
            add_log("HasExistingTag: Ошибка при проверке существующих марок: {0}".format(str(e)))
            return False
    
    def OnShowLogsClick(self, sender, args):
        """Показать логи"""
        show_logs()
    
    # Методы навигации назад
    def OnBack1Click(self, sender, args):
        add_log("OnBack1Click: Возврат к выбору видов")
        self.tabControl.SelectedTab = self.tabViews
    
    def OnBack2Click(self, sender, args):
        add_log("OnBack2Click: Возврат к выбору категорий")
        self.tabControl.SelectedTab = self.tabCategories
    
    def OnBack3Click(self, sender, args):
        add_log("OnBack3Click: Возврат к выбору марок")
        self.tabControl.SelectedTab = self.tabTags
    
    def OnBack4Click(self, sender, args):
        add_log("OnBack4Click: Возврат к настройкам")
        self.tabControl.SelectedTab = self.tabSettings

# Основная функция
def main():
    try:
        add_log("=== ЗАПУСК ПРИЛОЖЕНИЯ ===")
        # Получаем документ Revit
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        
        add_log("main: Документ получен: {0}".format(doc.Title if doc else 'None'))
        
        # Проверяем, что документ доступен
        if doc and uidoc:
            add_log("main: Создание главной формы")
            form = MainForm(doc, uidoc)
            add_log("main: Запуск приложения")
            Application.Run(form)
            add_log("main: Приложение завершено")
        else:
            error_msg = "Не удалось получить доступ к документу Revit"
            add_log("main: {0}".format(error_msg))
            MessageBox.Show(error_msg)
    except Exception as e:
        error_msg = "Критическая ошибка при запуске: {0}".format(str(e))
        add_log("main: {0}".format(error_msg))
        MessageBox.Show(error_msg)
        show_logs()

if __name__ == "__main__":
    main()