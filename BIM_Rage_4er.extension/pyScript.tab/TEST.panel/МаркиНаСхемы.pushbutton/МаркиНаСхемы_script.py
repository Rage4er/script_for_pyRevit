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
        self.category_tag_families = {}
        self.category_tag_types = {}
        self.offset_x = 60  # мм - фиксированное значение
        self.offset_y = 30  # мм - фиксированное значение
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True

# Базовый класс для форм с общей функциональностью
class BaseForm(Form):
    def GetElementNameImproved(self, element):
        """Улучшенное получение корректного имени элемента"""
        if not element:
            return "Без имени"
        
        try:
            # Для FamilySymbol (типоразмеров)
            if isinstance(element, FamilySymbol):
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name
                
                if hasattr(element, 'Family') and element.Family:
                    family_name = element.Family.Name if hasattr(element.Family, 'Name') and element.Family.Name else ""
                    
                    # Пробуем получить имя типа через параметры
                    type_name = ""
                    for param_name in ["Тип", "Type Name", "Имя типа"]:
                        param = element.LookupParameter(param_name)
                        if param and param.HasValue:
                            type_name = param.AsString()
                            break
                    
                    if family_name and type_name:
                        return "{0} - {1}".format(family_name, type_name)
                    elif type_name:
                        return type_name
                    elif family_name:
                        return family_name
                
                return "Типоразмер {0}".format(element.Id.IntegerValue)
            
            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name
                return "Семейство {0}".format(element.Id.IntegerValue)
            
            # Общий случай
            if hasattr(element, 'Name') and element.Name:
                name = element.Name
                if name and not name.startswith('IronPython'):
                    return name
                
        except Exception as e:
            add_log("GetElementNameImproved: Ошибка при получении имени элемента: {0}".format(str(e)))
        
        return "Элемент {0}".format(element.Id.IntegerValue)

# Форма для выбора семейства и типоразмера марки
class TagFamilySelectionForm(BaseForm):
    def __init__(self, doc, available_families, current_family, current_type):
        add_log("TagFamilySelectionForm: Инициализация формы выбора марки")
        self.doc = doc
        self.available_families = available_families
        self.current_family = current_family
        self.current_type = current_type
        self.selected_family = None
        self.selected_type = None
        
        self.family_dict = {}
        self.type_dict = {}
        
        self.InitializeComponent()
        self.PopulateFamiliesList()
    
    def InitializeComponent(self):
        self.Text = "Выбор семейства и типоразмера марки"
        self.Size = Size(800, 500)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        
        # Создание элементов управления
        controls = [
            self.CreateLabel("Выберите семейство марки:", 10, 10, 250, 20),
            self.CreateListBox(10, 40, 250, 350, self.OnFamilySelected),
            self.CreateLabel("Выберите типоразмер:", 270, 10, 250, 20),
            self.CreateListBox(270, 40, 500, 350),
            self.CreateButton("OK", 400, 400, 75, 25, self.OnOKClick),
            self.CreateButton("Отмена", 485, 400, 75, 25, self.OnCancelClick)
        ]
        
        self.lblFamilies, self.lstFamilies, self.lblTypes, self.lstTypes, self.btnOK, self.btnCancel = controls
        
        for control in controls:
            self.Controls.Add(control)
    
    def CreateLabel(self, text, x, y, width, height):
        label = Label()
        label.Text = text
        label.Location = Point(x, y)
        label.Size = Size(width, height)
        return label
    
    def CreateListBox(self, x, y, width, height, event_handler=None):
        listbox = ListBox()
        listbox.Location = Point(x, y)
        listbox.Size = Size(width, height)
        if event_handler:
            listbox.SelectedIndexChanged += event_handler
        return listbox
    
    def CreateButton(self, text, x, y, width, height, click_handler):
        button = Button()
        button.Text = text
        button.Location = Point(x, y)
        button.Size = Size(width, height)
        button.Click += click_handler
        return button
    
    def PopulateFamiliesList(self):
        add_log("TagFamilySelectionForm: Заполнение списка семейств. Доступно: {0}".format(len(self.available_families)))
        self.lstFamilies.Items.Clear()
        self.family_dict.clear()
        
        for family in self.available_families:
            if family:
                family_name = self.GetElementNameImproved(family)
                self.lstFamilies.Items.Add(family_name)
                self.family_dict[family_name] = family
        
        # Выбор текущего или первого семейства
        if self.current_family:
            current_name = self.GetElementNameImproved(self.current_family)
            self.SelectItemByName(self.lstFamilies, current_name, "текущее семейство")
        elif self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0
    
    def SelectItemByName(self, listbox, name, item_type):
        for i in range(listbox.Items.Count):
            if listbox.Items[i] == name:
                listbox.SelectedIndex = i
                add_log("TagFamilySelectionForm: Выбрано {0} с индексом {1}".format(item_type, i))
                return True
        return False
    
    def OnFamilySelected(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0:
            selected_name = self.lstFamilies.SelectedItem
            selected_family = self.family_dict.get(selected_name)
            if selected_family:
                add_log("TagFamilySelectionForm: Выбрано семейство: {0}".format(selected_name))
                self.PopulateTypesList(selected_family)
    
    def PopulateTypesList(self, family):
        add_log("TagFamilySelectionForm: Загрузка типоразмеров для семейства {0}".format(self.GetElementNameImproved(family)))
        self.lstTypes.Items.Clear()
        self.type_dict.clear()

        try:
            symbol_ids = family.GetFamilySymbolIds()
            add_log("TagFamilySelectionForm: Найдено типоразмеров: {0}".format(symbol_ids.Count if symbol_ids else 0))
        
            if symbol_ids and symbol_ids.Count > 0:
                for symbol_id in list(symbol_ids):
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = self.GetElementNameImproved(symbol)
                        status = " (активный)" if symbol.IsActive else " (не активный)"
                        display_name = "{0}{1} [ID:{2}]".format(symbol_name, status, symbol.Id.IntegerValue)
                        
                        self.lstTypes.Items.Add(display_name)
                        self.type_dict[display_name] = symbol
            
                # Выбор типоразмера
                if self.current_type:
                    current_name = self.GetElementNameImproved(self.current_type)
                    self.SelectItemByName(self.lstTypes, current_name, "текущий типоразмер")
                else:
                    self.SelectFirstActiveType()
            else:
                add_log("TagFamilySelectionForm: В выбранном семействе нет типоразмеров")
                MessageBox.Show("В выбранном семействе нет типоразмеров")
                
        except Exception as e:
            add_log("TagFamilySelectionForm: Ошибка при загрузке типоразмеров: {0}".format(str(e)))
            MessageBox.Show("Ошибка при загрузке типоразмеров")

    def SelectFirstActiveType(self):
        for i in range(self.lstTypes.Items.Count):
            display_name = self.lstTypes.Items[i]
            symbol = self.type_dict.get(display_name)
            if symbol and symbol.IsActive:
                self.lstTypes.SelectedIndex = i
                add_log("TagFamilySelectionForm: Выбран первый активный типоразмер с индексом {0}".format(i))
                return
        
        if self.lstTypes.Items.Count > 0:
            self.lstTypes.SelectedIndex = 0
            add_log("TagFamilySelectionForm: Выбран первый типоразмер по умолчанию")
    
    def OnOKClick(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0 and self.lstTypes.SelectedIndex >= 0:
            selected_family_name = self.lstFamilies.SelectedItem
            selected_type_name = self.lstTypes.SelectedItem
            
            self.selected_family = self.family_dict.get(selected_family_name)
            self.selected_type = self.type_dict.get(selected_type_name)
            
            if self.selected_family and self.selected_type:
                add_log("TagFamilySelectionForm: Выбрано - Семейство: {0}, Тип: {1}".format(selected_family_name, selected_type_name))
                self.DialogResult = DialogResult.OK
                self.Close()
            else:
                MessageBox.Show("Ошибка при получении выбранных объектов!")
        else:
            MessageBox.Show("Выберите семейство и типоразмер марки!")
    
    def OnCancelClick(self, sender, args):
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
class MainForm(BaseForm):
    def __init__(self, doc, uidoc):
        add_log("MainForm: Инициализация главного окна")
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}
        self.category_mapping = {}
        
        self.InitializeComponent()
        self.Load3DViewsWithDuctSystems()
    
    def InitializeComponent(self):
        self.Text = "Расстановка марок на 3D видах"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        
        # Создание вкладок
        tabs = [
            ("1. Выбор видов", self.SetupViewsTab),
            ("2. Категории", self.SetupCategoriesTab),
            ("3. Марки", self.SetupTagsTab),
            ("4. Настройки", self.SetupSettingsTab),
            ("5. Выполнение", self.SetupExecuteTab)
        ]
        
        for tab_text, setup_method in tabs:
            tab = TabPage()
            tab.Text = tab_text
            self.tabControl.TabPages.Add(tab)
            setup_method(tab)
        
        self.Controls.Add(self.tabControl)
    
    def CreateTabControl(self, tab, controls):
        """Универсальный метод создания элементов управления на вкладке"""
        for control in controls:
            tab.Controls.Add(control)
    
    def SetupViewsTab(self, tab):
        controls = [
            self.CreateLabel("Выберите 3D виды с системами воздуховодов:", 10, 10, 300, 20),
            self.CreateCheckedListBox(10, 40, 600, 400),
            self.CreateButton("Показать логи", 10, 450, 100, 25, self.OnShowLogsClick),
            self.CreateButton("Далее →", 500, 450, 80, 25, self.OnNext1Click)
        ]
        
        self.lblViews, self.lstViews, self.btnShowLogs, self.btnNext1 = controls
        self.CreateTabControl(tab, controls)
    
    def SetupCategoriesTab(self, tab):
        controls = [
            self.CreateLabel("Выберите категории элементов:", 10, 10, 300, 20),
            self.CreateCheckedListBox(10, 40, 600, 400),
            self.CreateButton("← Назад", 400, 450, 80, 25, self.OnBack1Click),
            self.CreateButton("Далее →", 500, 450, 80, 25, self.OnNext2Click)
        ]
        
        self.lblCategories, self.lstCategories, self.btnBack1, self.btnNext2 = controls
        self.CreateTabControl(tab, controls)
    
    def SetupTagsTab(self, tab):
        list_view = ListView()
        list_view.Location = Point(10, 40)
        list_view.Size = Size(700, 400)
        list_view.View = View.Details
        list_view.FullRowSelect = True
        list_view.GridLines = True
        list_view.Columns.Add("Категория", 180)
        list_view.Columns.Add("Семейство марки", 200)
        list_view.Columns.Add("Типоразмер марки", 300)
        list_view.DoubleClick += self.OnTagFamilyDoubleClick
        
        controls = [
            self.CreateLabel("Выберите семейства и типоразмеры марок для категорий:", 10, 10, 400, 20),
            list_view,
            self.CreateButton("← Назад", 500, 450, 80, 25, self.OnBack2Click),
            self.CreateButton("Далее →", 600, 450, 80, 25, self.OnNext3Click)
        ]
        
        self.lblTags, self.lstTagFamilies, self.btnBack2, self.btnNext3 = controls
        self.CreateTabControl(tab, controls)
    
    def SetupSettingsTab(self, tab):
        cmb_orientation = ComboBox()
        cmb_orientation.Location = Point(170, 50)
        cmb_orientation.Size = Size(100, 20)
        cmb_orientation.Items.Add("Горизонтальная")
        cmb_orientation.Items.Add("Вертикальная")
        cmb_orientation.SelectedIndex = 0
        
        chk_leader = CheckBox()
        chk_leader.Text = "Использовать выноску"
        chk_leader.Location = Point(10, 80)
        chk_leader.Size = Size(200, 20)
        chk_leader.Checked = True
        
        controls = [
            self.CreateLabel("Настройки размещения марок:", 10, 10, 300, 20),
            self.CreateLabel("Ориентация марки:", 10, 50, 150, 20),
            cmb_orientation,
            chk_leader,
            self.CreateButton("← Назад", 500, 450, 80, 25, self.OnBack3Click),
            self.CreateButton("Далее →", 600, 450, 80, 25, self.OnNext4Click)
        ]
        
        self.lblSettings, self.lblOrientation, self.cmbOrientation, self.chkUseLeader, self.btnBack3, self.btnNext4 = controls
        self.CreateTabControl(tab, controls)
    
    def SetupExecuteTab(self, tab):
        textbox = TextBox()
        textbox.Location = Point(10, 40)
        textbox.Size = Size(700, 400)
        textbox.Multiline = True
        textbox.ScrollBars = ScrollBars.Vertical
        textbox.ReadOnly = True
        
        controls = [
            self.CreateLabel("Готово к выполнению:", 10, 10, 300, 20),
            textbox,
            self.CreateButton("← Назад", 340, 450, 80, 25, self.OnBack4Click),
            self.CreateButton("Выполнить расстановку", 500, 450, 150, 30, self.OnExecuteClick)
        ]
        
        self.lblExecute, self.txtSummary, self.btnBack4, self.btnExecute = controls
        self.CreateTabControl(tab, controls)
    
    def CreateLabel(self, text, x, y, width, height):
        label = Label()
        label.Text = text
        label.Location = Point(x, y)
        label.Size = Size(width, height)
        return label
    
    def CreateCheckedListBox(self, x, y, width, height):
        listbox = CheckedListBox()
        listbox.Location = Point(x, y)
        listbox.Size = Size(width, height)
        listbox.CheckOnClick = True
        return listbox
    
    def CreateButton(self, text, x, y, width, height, click_handler):
        button = Button()
        button.Text = text
        button.Location = Point(x, y)
        button.Size = Size(width, height)
        button.Click += click_handler
        return button
    
    def Load3DViewsWithDuctSystems(self):
        try:
            add_log("Load3DViewsWithDuctSystems: Начало загрузки 3D видов")
            collector = FilteredElementCollector(self.doc)
            views = collector.OfClass(View3D).WhereElementIsNotElementType().ToElements()
            
            self.lstViews.Items.Clear()
            self.views_dict = {}
            
            for view in views:
                if not view.IsTemplate and view.CanBePrinted:
                    display_name = "{0} (ID: {1})".format(view.Name, view.Id.IntegerValue)
                    self.lstViews.Items.Add(display_name, False)
                    self.views_dict[display_name] = view
            
            add_log("Load3DViewsWithDuctSystems: Загружено 3D видов: {0}".format(self.lstViews.Items.Count))
            
            if self.lstViews.Items.Count == 0:
                MessageBox.Show("Не найдено 3D видов в проекте")
            else:
                self.lstViews.SetItemChecked(0, True)
                
        except Exception as e:
            add_log("Load3DViewsWithDuctSystems: Ошибка при загрузке видов: {0}".format(str(e)))
            MessageBox.Show("Ошибка при загрузке видов: " + str(e))
    
    # Методы навигации
    def OnNext1Click(self, sender, args):
        self.settings.selected_views = [self.views_dict[name] for i, name in enumerate(self.lstViews.Items) 
                                      if self.lstViews.GetItemChecked(i) and name in self.views_dict]
        
        add_log("OnNext1Click: Выбрано видов: {0}".format(len(self.settings.selected_views)))
        
        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return
        
        self.CollectUniqueCategories()
        self.PopulateCategoriesList()
        self.tabControl.SelectedTab = self.tabControl.TabPages[1]
    
    def OnNext2Click(self, sender, args):
        self.settings.selected_categories = [self.category_mapping[name] for i, name in enumerate(self.lstCategories.Items)
                                           if self.lstCategories.GetItemChecked(i) and name in self.category_mapping]
        
        add_log("OnNext2Click: Выбрано категорий: {0}".format(len(self.settings.selected_categories)))
        
        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return
        
        self.PopulateTagFamilies()
        self.tabControl.SelectedTab = self.tabControl.TabPages[2]
    
    def OnNext3Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabControl.TabPages[3]
    
    def OnNext4Click(self, sender, args):
        try:
            self.settings.orientation = TagOrientation.Horizontal if self.cmbOrientation.SelectedIndex == 0 else TagOrientation.Vertical
            self.settings.use_leader = self.chkUseLeader.Checked
            
            add_log("OnNext4Click: Настройки сохранены - ориентация:{0}, выноска:{1}".format(
                self.settings.orientation, self.settings.use_leader))
                
        except Exception as e:
            MessageBox.Show("Проверьте правильность введенных значений! Ошибка: " + str(e))
            return
        
        self.txtSummary.Text = self.GenerateSummary()
        self.tabControl.SelectedTab = self.tabControl.TabPages[4]
    
    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabControl.TabPages[0]
    
    def OnBack2Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabControl.TabPages[1]
    
    def OnBack3Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabControl.TabPages[2]
    
    def OnBack4Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabControl.TabPages[3]
    
    def CollectUniqueCategories(self):
        target_categories = [
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_MechanicalEquipment  # ДОБАВЛЕНА НОВАЯ КАТЕГОРИЯ
        ]
        
        unique_categories = set()
        for category in target_categories:
            try:
                cat_obj = Category.GetCategory(self.doc, category)
                if cat_obj:
                    unique_categories.add(cat_obj)
            except Exception as e:
                add_log("CollectUniqueCategories: Ошибка при получении категории {0}: {1}".format(category, str(e)))
        
        self.settings.selected_categories = list(unique_categories)
        add_log("CollectUniqueCategories: Всего собрано категорий: {0}".format(len(self.settings.selected_categories)))
    
    def PopulateCategoriesList(self):
        self.lstCategories.Items.Clear()
        self.category_mapping = {}
        
        sorted_categories = sorted(self.settings.selected_categories, key=lambda x: self.GetCategoryDisplayName(x))
        
        for category in sorted_categories:
            display_name = self.GetCategoryDisplayName(category)
            self.lstCategories.Items.Add(display_name, True)
            self.category_mapping[display_name] = category
    
    def GetCategoryDisplayName(self, category):
        if not category:
            return "Неизвестная категория"
        
        try:
            if hasattr(category, 'Id') and category.Id.IntegerValue < 0:
                built_in_cat = BuiltInCategory(category.Id.IntegerValue)
                return LabelUtils.GetLabelFor(built_in_cat)
        except:
            pass
        
        return getattr(category, 'Name', 'Неизвестная категория')
    
    def PopulateTagFamilies(self):
        self.lstTagFamilies.Items.Clear()
        
        for category in self.settings.selected_categories:
            cat_name = self.GetCategoryDisplayName(category)
            item = ListViewItem(cat_name)
            item.Tag = category
            
            tag_family, tag_type = self.FindSuitableTagForCategory(category)
            
            if tag_family and tag_type:
                family_name = self.GetElementNameImproved(tag_family)
                type_name = self.GetElementNameImproved(tag_type)
                item.SubItems.Add(family_name)
                item.SubItems.Add(type_name)
                self.settings.category_tag_families[category] = tag_family
                self.settings.category_tag_types[category] = tag_type
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")
                self.settings.category_tag_families[category] = None
                self.settings.category_tag_types[category] = None
            
            self.lstTagFamilies.Items.Add(item)
    
    def FindSuitableTagForCategory(self, category):
        add_log("FindSuitableTagForCategory: Поиск марки для категории: {0}".format(self.GetCategoryDisplayName(category)))
        
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            return None, None
        
        collector = FilteredElementCollector(self.doc)
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        
        for family in tag_families:
            if not family or not hasattr(family, 'FamilyCategory'):
                continue
                
            if family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                symbol_ids = family.GetFamilySymbolIds()
                if symbol_ids and symbol_ids.Count > 0:
                    for symbol_id in list(symbol_ids):
                        tag_type = self.doc.GetElement(symbol_id)
                        if tag_type and tag_type.IsActive:
                            return family, tag_type
                    # Если нет активных, берем первый доступный
                    tag_type = self.doc.GetElement(list(symbol_ids)[0])
                    return family, tag_type
        
        return None, None
    
    def GetTagCategoryForElementCategory(self, element_category):
        if not element_category:
            return None
        
        category_mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipment: BuiltInCategory.OST_MechanicalEquipmentTags  # ДОБАВЛЕНА НОВАЯ КАТЕГОРИЯ МАРОК
        }
        
        try:
            if hasattr(element_category, 'Id') and element_category.Id.IntegerValue < 0:
                element_builtin_cat = BuiltInCategory(element_category.Id.IntegerValue)
                if element_builtin_cat in category_mapping:
                    tag_builtin_cat = category_mapping[element_builtin_cat]
                    tag_category = Category.GetCategory(self.doc, tag_builtin_cat)
                    if tag_category:
                        return tag_category.Id
        except Exception as e:
            add_log("GetTagCategoryForElementCategory: Ошибка: {0}".format(str(e)))
        
        return None
    
    def OnTagFamilyDoubleClick(self, sender, args):
        if self.lstTagFamilies.SelectedItems.Count > 0:
            selected_item = self.lstTagFamilies.SelectedItems[0]
            category = selected_item.Tag
            
            available_families = self.GetAvailableTagFamiliesForCategory(category)
            
            if available_families:
                current_family = self.settings.category_tag_families.get(category)
                current_type = self.settings.category_tag_types.get(category)
                form = TagFamilySelectionForm(self.doc, available_families, current_family, current_type)
                
                if form.ShowDialog() == DialogResult.OK and form.SelectedFamily and form.SelectedType:
                    family_name = self.GetElementNameImproved(form.SelectedFamily)
                    type_name = self.GetElementNameImproved(form.SelectedType)
                    selected_item.SubItems[1].Text = family_name
                    selected_item.SubItems[2].Text = type_name
                    self.settings.category_tag_families[category] = form.SelectedFamily
                    self.settings.category_tag_types[category] = form.SelectedType
            else:
                MessageBox.Show("Нет доступных семейств марок для этой категории")

    def GetAvailableTagFamiliesForCategory(self, category):
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            return []
        
        collector = FilteredElementCollector(self.doc)
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        
        return [family for family in tag_families 
                if family and hasattr(family, 'FamilyCategory') 
                and family.FamilyCategory and family.FamilyCategory.Id == tag_category_id]
    
    def GenerateSummary(self):
        summary = "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ:\n\n"
        summary += "Выбрано видов: {0}\n".format(len(self.settings.selected_views))
        summary += "Выбрано категорий: {0}\n".format(len(self.settings.selected_categories))
        
        orientation_text = "Горизонтальная" if self.settings.orientation == TagOrientation.Horizontal else "Вертикальная"
        summary += "Ориентация: {0}\n".format(orientation_text)
        summary += "Выноска: {0}\n\n".format("Да" if self.settings.use_leader else "Нет")
        
        summary += "Детали по категориям:\n"
        for category in self.settings.selected_categories:
            cat_name = self.GetCategoryDisplayName(category)
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            
            if tag_family and tag_type:
                family_name = self.GetElementNameImproved(tag_family)
                type_name = self.GetElementNameImproved(tag_type)
                summary += "- {0}: {1} ({2})\n".format(cat_name, family_name, type_name)
            else:
                summary += "- {0}: НЕТ МАРКИ\n".format(cat_name)
        
        return summary
    
    def OnExecuteClick(self, sender, args):
        add_log("OnExecuteClick: Начало выполнения расстановки марок")
        errors = []
        success_count = 0
        
        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        
        try:
            for view in self.settings.selected_views:
                if not isinstance(view, View3D):
                    errors.append("Вид '{0}' не является 3D видом".format(view.Name))
                    continue
                
                for category in self.settings.selected_categories:
                    collector = FilteredElementCollector(self.doc, view.Id)
                    elements = collector.OfCategoryId(category.Id).WhereElementIsNotElementType().ToElements()
                    
                    for element in elements:
                        if not element:
                            continue
                        
                        if self.HasExistingTag(element, view):
                            continue
                        
                        tag_family = self.settings.category_tag_families.get(category)
                        tag_type = self.settings.category_tag_types.get(category)
                        
                        if not tag_family or not tag_type:
                            errors.append("Нет марки для категории {0}".format(self.GetCategoryDisplayName(category)))
                            continue
                        
                        if self.CreateTagImproved(element, view, tag_type):
                            success_count += 1
            
            trans.Commit()
            add_log("OnExecuteClick: Создано марок: {0}".format(success_count))
            
        except Exception as e:
            trans.RollBack()
            errors.append("Ошибка выполнения: {0}".format(str(e)))
        finally:
            trans.Dispose()
        
        # Показ результатов
        result_msg = "Успешно расставлено марок: {0}\n".format(success_count)
        if errors:
            result_msg += "\nОшибки ({0}):\n".format(len(errors)) + "\n".join(errors[:10])
            if len(errors) > 10:
                result_msg += "\n... и еще {0} ошибок".format(len(errors) - 10)
        
        MessageBox.Show(result_msg, "Результат выполнения")
        show_logs()
        self.Close()
    
    def CreateTagImproved(self, element, view, tag_type):
        try:
            bbox = element.get_BoundingBox(view)
            if not bbox:
                return False
            
            center = (bbox.Min + bbox.Max) / 2
            scale = view.Scale
            scale_factor = 100.0 / scale
            
            # Конвертация мм в футы с учетом масштаба
            offset_x_feet = (self.settings.offset_x * scale_factor) / 304.8
            offset_y_feet = (self.settings.offset_y * scale_factor) / 304.8
            
            # Случайное направление для разнообразия
            direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
            direction_y = 1 if (element.Id.IntegerValue % 3 == 0) else -1
            
            tag_point = XYZ(
                center.X + (offset_x_feet * direction_x),
                center.Y + (offset_y_feet * direction_y),
                center.Z
            )
            
            element_ref = Reference(element)
            tag_mode = TagMode.TM_ADDBY_CATEGORY
            
            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                element_ref,
                self.settings.use_leader,
                tag_mode,
                self.settings.orientation,
                tag_point
            )
            
            if tag and tag_type:
                try:
                    tag.ChangeTypeId(tag_type.Id)
                except Exception as e:
                    add_log("CreateTagImproved: Ошибка при установке типоразмера: {0}".format(str(e)))
                return True
            
            return False
            
        except Exception as e:
            add_log("CreateTagImproved: Ошибка создания марки: {0}".format(str(e)))
            return False
    
    def HasExistingTag(self, element, view):
        try:
            collector = FilteredElementCollector(self.doc, view.Id)
            tags = collector.OfClass(IndependentTag).ToElements()
            
            for tag in tags:
                try:
                    for tagged_elem in tag.GetTaggedLocalElements():
                        if tagged_elem.Id == element.Id:
                            return True
                except:
                    if hasattr(tag, 'TaggedLocalElementId') and tag.TaggedLocalElementId == element.Id:
                        return True
            return False
        except Exception as e:
            add_log("HasExistingTag: Ошибка при проверке существующих марок: {0}".format(str(e)))
            return False
    
    def OnShowLogsClick(self, sender, args):
        show_logs()

# Основная функция
def main():
    try:
        add_log("=== ЗАПУСК ПРИЛОЖЕНИЯ ===")
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        
        if doc and uidoc:
            form = MainForm(doc, uidoc)
            Application.Run(form)
        else:
            MessageBox.Show("Не удалось получить доступ к документу Revit")
    except Exception as e:
        error_msg = "Критическая ошибка при запуске: {0}".format(str(e))
        add_log(error_msg)
        MessageBox.Show(error_msg)
        show_logs()

if __name__ == "__main__":
    main()