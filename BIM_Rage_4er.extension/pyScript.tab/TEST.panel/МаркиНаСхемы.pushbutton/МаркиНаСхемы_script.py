# -*- coding: utf-8 -*-
__title__ = 'Автоматическое размещение маркировочных меток на 3D-видах'
__author__ = 'г'
__doc__ = ' '

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from System.Windows.Forms import *
from System.Drawing import *
import sys
import os

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

# Главное окно приложения
class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}  # Словарь для хранения соответствия названий и объектов View
        self.category_mapping = {}  # Словарь для хранения соответствия названий категорий и объектов Category
        self.InitializeComponent()
        self.Load3DViewsWithDuctSystems()
    
    def InitializeComponent(self):
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
        
        controls = [self.lblViews, self.lstViews, self.btnNext1]
        for control in controls:
            self.tabViews.Controls.Add(control)
    
    def SetupCategoriesTab(self):
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
        collector = FilteredElementCollector(self.doc)
        views = collector.OfClass(View3D).WhereElementIsNotElementType().ToElements()
        
        self.lstViews.Items.Clear()
        self.views_dict = {}
        
        target_categories = [
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_PlaceHolderDucts,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctTerminal
        ]
        
        for view in views:
            if not view.IsTemplate and view.CanBePrinted:
                # Проверяем, есть ли в виде элементы систем воздуховодов
                has_duct_elements = False
                
                # Проверяем наличие элементов каждой категории
                for category in target_categories:
                    try:
                        collector = FilteredElementCollector(self.doc, view.Id)
                        elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                        if elements and any(elements):
                            has_duct_elements = True
                            break
                    except:
                        continue
                
                if has_duct_elements:
                    # Создаем отображаемое имя
                    display_name = "{} (ID: {})".format(view.Name, view.Id.IntegerValue)
                    # Добавляем в список
                    self.lstViews.Items.Add(display_name, False)
                    # Сохраняем соответствие в словаре
                    self.views_dict[display_name] = view
        
        # Если не найдено видов с воздуховодами, показываем все 3D виды
        if self.lstViews.Items.Count == 0:
            for view in views:
                if not view.IsTemplate and view.CanBePrinted:
                    display_name = "{} (ID: {})".format(view.Name, view.Id.IntegerValue)
                    self.lstViews.Items.Add(display_name, False)
                    self.views_dict[display_name] = view
    
    def OnNext1Click(self, sender, args):
        """Переход от выбора видов к выбору категорий"""
        # Сохраняем выбранные виды
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            if self.lstViews.GetItemChecked(i):
                display_name = self.lstViews.Items[i]
                if display_name in self.views_dict:
                    self.settings.selected_views.append(self.views_dict[display_name])
        
        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return
        
        # Собираем уникальные категории из выбранных видов
        self.CollectUniqueCategories()
        self.PopulateCategoriesList()
        
        self.tabControl.SelectedTab = self.tabCategories

    def CollectUniqueCategories(self):
        """Сбор уникальных категорий из выбранных видов"""
        unique_categories = set()
        
        # Целевые категории для систем воздуховодов
        target_categories = [
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_PlaceHolderDucts,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctTerminal
        ]
        
        # Просто добавляем все целевые категории
        for category in target_categories:
            try:
                cat_obj = Category.GetCategory(self.doc, category)
                if cat_obj:
                    unique_categories.add(cat_obj)
            except Exception as e:
                print("Ошибка при получении категории {}: {}".format(category, str(e)))
        
        # Сохраняем уникальные категории
        self.settings.selected_categories = list(unique_categories)
    
    def PopulateCategoriesList(self):
        """Заполнение списка категорий"""
        self.lstCategories.Items.Clear()
        self.category_mapping = {}  # Очищаем словарь соответствий
        
        # Сортируем категории по имени для удобства
        sorted_categories = sorted(self.settings.selected_categories, key=lambda x: self.GetCategoryDisplayName(x))
        
        for category in sorted_categories:
            # Отображаем корректное имя категории
            display_name = self.GetCategoryDisplayName(category)
            self.lstCategories.Items.Add(display_name, False)
            # Сохраняем соответствие в словаре
            self.category_mapping[display_name] = category
    
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
        # Сохраняем выбранные категории
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            if self.lstCategories.GetItemChecked(i):
                display_name = self.lstCategories.Items[i]
                if display_name in self.category_mapping:
                    self.settings.selected_categories.append(self.category_mapping[display_name])
        
        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return
        
        # Загружаем доступные семейства марок
        self.PopulateTagFamilies()
        self.tabControl.SelectedTab = self.tabTags
    
    def PopulateTagFamilies(self):
        """Заполнение списка доступных семейств марок"""
        self.lstTagFamilies.Items.Clear()
        
        # Заполняем список для выбранных категорий
        for category in self.settings.selected_categories:
            cat_name = self.GetCategoryDisplayName(category)
            item = ListViewItem(cat_name)
            item.Tag = category  # Сохраняем категорию в Tag
            
            # Ищем подходящие семейства марок
            tag_family, tag_type = self.FindSuitableTagForCategory(category)
            
            if tag_family and tag_type:
                family_name = getattr(tag_family, 'Name', 'Без имени')
                type_name = getattr(tag_type, 'Name', 'Без имени')
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
        """Поиск подходящего семейства и типоразмера марки для категории"""
        print("Поиск марки для категории: {}".format(self.GetCategoryDisplayName(category)))
        
        # Определяем категорию марки в зависимости от категории элемента
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            print("Не найдена категория марки")
            return None, None
        
        # Получаем все семейства марок нужной категории в проекте
        collector = FilteredElementCollector(self.doc)
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()
        
        print("Всего семейств в проекте: {}".format(len(list(tag_families))))
        
        # Ищем семейства марки для данной категории
        for family in tag_families:
            if not family or not hasattr(family, 'FamilyCategory'):
                continue
                
            print("Проверяем семейство: {} (категория: {})".format(
                getattr(family, 'Name', 'Без имени'),
                getattr(family.FamilyCategory, 'Name', 'Без категории') if family.FamilyCategory else 'Нет категории'
            ))
                
            if family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                print("Нашли подходящее семейство: {}".format(getattr(family, 'Name', 'Без имени')))
                
                # Получаем все доступные типоразмеры
                symbol_ids = family.GetFamilySymbolIds()
                print("Типоразмеров в семействе: {}".format(symbol_ids.Count if symbol_ids else 0))
                
                if symbol_ids and symbol_ids.Count > 0:
                    # Преобразуем HashSet в список для безопасного доступа
                    symbol_id_list = list(symbol_ids)
                    if symbol_id_list:
                        # Берем первый активный типоразмер
                        for symbol_id in symbol_id_list:
                            tag_type = self.doc.GetElement(symbol_id)
                            if tag_type and tag_type.IsActive:
                                print("Найден активный типоразмер: {}".format(getattr(tag_type, 'Name', 'Без имени')))
                                return family, tag_type
                        # Если нет активных, берем первый доступный
                        tag_type = self.doc.GetElement(symbol_id_list[0])
                        print("Используем первый типоразмер: {}".format(getattr(tag_type, 'Name', 'Без имени')))
                        return family, tag_type
        
        print("Не найдено подходящих марок")
        return None, None
    
    def GetTagCategoryForElementCategory(self, element_category):
        """Определяет категорию марки для категории элемента"""
        if not element_category:
            return None
        
        # Сопоставление категорий элементов с категориями марок
        category_mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_PlaceHolderDucts: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags
        }
        
        # Получаем BuiltInCategory для элемента
        try:
            if hasattr(element_category, 'Id') and element_category.Id.IntegerValue < 0:
                element_builtin_cat = BuiltInCategory(element_category.Id.IntegerValue)
                if element_builtin_cat in category_mapping:
                    tag_builtin_cat = category_mapping[element_builtin_cat]
                    tag_category = Category.GetCategory(self.doc, tag_builtin_cat)
                    if tag_category:
                        return tag_category.Id
        except Exception as e:
            print("Ошибка при получении категории марки: {}".format(str(e)))
        
        return None
    
    def OnTagFamilyDoubleClick(self, sender, args):
        """Обработка двойного клика для изменения семейства марки"""
        if self.lstTagFamilies.SelectedItems.Count > 0:
            selected_item = self.lstTagFamilies.SelectedItems[0]
            category = selected_item.Tag
            
            # Получаем все доступные семейства марок для этой категории
            available_families = self.GetAvailableTagFamiliesForCategory(category)
            
            if available_families:
                # Создаем диалог для выбора семейства и типоразмера
                current_family = self.settings.category_tag_families.get(category)
                current_type = self.settings.category_tag_types.get(category)
                form = TagFamilySelectionForm(self.doc, available_families, current_family, current_type)
                result = form.ShowDialog()
                
                if result == DialogResult.OK and form.SelectedFamily and form.SelectedType:
                    # Обновляем выбранное семейство и типоразмер
                    family_name = getattr(form.SelectedFamily, 'Name', 'Без имени')
                    type_name = getattr(form.SelectedType, 'Name', 'Без имени')
                    selected_item.SubItems[1].Text = family_name
                    selected_item.SubItems[2].Text = type_name
                    self.settings.category_tag_families[category] = form.SelectedFamily
                    self.settings.category_tag_types[category] = form.SelectedType
            else:
                MessageBox.Show("Нет доступных семейств марок для этой категории")
    
    def GetAvailableTagFamiliesForCategory(self, category):
        """Получение всех доступных семейств марок для категории"""
        # Определяем категорию марки в зависимости от категории элемента
        tag_category_id = self.GetTagCategoryForElementCategory(category)
        if not tag_category_id:
            print("Не найдена категория марки для: {}".format(self.GetCategoryDisplayName(category)))
            return []
        
        print("Ищем марки категории: {}".format(Category.GetCategory(self.doc, BuiltInCategory(tag_category_id.IntegerValue)).Name))
        
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
                print("Семейство: {} (типоразмеров: {})".format(
                    getattr(family, 'Name', 'Без имени'),
                    symbol_ids.Count if symbol_ids else 0
                ))
        
        print("Итого найдено семейств марок для категории {}: {}".format(
            self.GetCategoryDisplayName(category), len(available_families)))
        
        return available_families
    
    def OnNext3Click(self, sender, args):
        """Переход к настройкам"""
        self.tabControl.SelectedTab = self.tabSettings
    
    def OnNext4Click(self, sender, args):
        """Переход к выполнению"""
        # Сохраняем настройки
        try:
            self.settings.offset_x = float(self.txtOffsetX.Text)
            self.settings.offset_y = float(self.txtOffsetY.Text)
            if self.cmbOrientation.SelectedIndex == 0:
                self.settings.orientation = TagOrientation.Horizontal
            else:
                self.settings.orientation = TagOrientation.Vertical
            self.settings.use_leader = self.chkUseLeader.Checked
        except Exception as e:
            MessageBox.Show("Проверьте правильность введенных значений! Ошибка: " + str(e))
            return
        
        # Формируем сводку
        summary = self.GenerateSummary()
        self.txtSummary.Text = summary
        
        self.tabControl.SelectedTab = self.tabExecute
    
    def GenerateSummary(self):
        """Генерация сводки перед выполнением"""
        summary = "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ:\n\n"
        summary += "Выбрано видов: {}\n".format(len(self.settings.selected_views))
        summary += "Выбрано категорий: {}\n".format(len(self.settings.selected_categories))
        summary += "Смещение: X={} мм, Y={} мм\n".format(self.settings.offset_x, self.settings.offset_y)
        
        if self.settings.orientation == TagOrientation.Horizontal:
            orientation_text = "Горизонтальная"
        else:
            orientation_text = "Вертикальная"
        summary += "Ориентация: {}\n".format(orientation_text)
        summary += "Выноска: {}\n\n".format("Да" if self.settings.use_leader else "Нет")
        
        summary += "Детали по категориям:\n"
        for category in self.settings.selected_categories:
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            cat_name = self.GetCategoryDisplayName(category)
            
            if tag_family and tag_type:
                family_name = getattr(tag_family, 'Name', 'Без имени')
                type_name = getattr(tag_type, 'Name', 'Без имени')
                summary += "- {}: {} ({})\n".format(cat_name, family_name, type_name)
            else:
                summary += "- {}: НЕТ МАРКИ\n".format(cat_name)
        
        return summary
    
    def OnExecuteClick(self, sender, args):
        """Выполнение расстановки марок"""
        errors = []
        success_count = 0
        
        # Начинаем транзакцию
        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        
        try:
            for view in self.settings.selected_views:
                if not isinstance(view, View3D):
                    errors.append("Вид '{}' не является 3D видом".format(view.Name))
                    continue
                
                # Получаем элементы вида через FilteredElementCollector
                collector = FilteredElementCollector(self.doc, view.Id)
                elements = collector.WhereElementIsNotElementType().ToElements()
                
                for element in elements:
                    if not element or not element.Category:
                        continue
                    
                    # Проверяем, подходит ли элемент под выбранные критерии
                    category = element.Category
                    if category not in self.settings.selected_categories:
                        continue
                    
                    # Проверяем, есть ли уже марка у элемента на этом виде
                    if self.HasExistingTag(element, view):
                        continue
                    
                    # Получаем семейство и типоразмер марки для категории
                    tag_family = self.settings.category_tag_families.get(category)
                    tag_type = self.settings.category_tag_types.get(category)
                    
                    if not tag_family or not tag_type:
                        cat_name = self.GetCategoryDisplayName(category)
                        errors.append("Нет марки для категории {}".format(cat_name))
                        continue
                    
                    # Создаем марку
                    if self.CreateTag(element, view, tag_type):
                        success_count += 1
                    else:
                        errors.append("Не удалось создать марку для {}".format(element.Id))
            
            trans.Commit()
            
        except Exception as e:
            trans.RollBack()
            errors.append("Ошибка выполнения: {}".format(str(e)))
        finally:
            trans.Dispose()
        
        # Показываем результаты
        result_msg = "Успешно расставлено марок: {}\n".format(success_count)
        if errors:
            result_msg += "\nОшибки ({}):\n".format(len(errors)) + "\n".join(errors[:10])
            if len(errors) > 10:
                result_msg += "\n... и еще {} ошибок".format(len(errors) - 10)
        
        MessageBox.Show(result_msg, "Результат выполнения")
        self.Close()
    
    def HasExistingTag(self, element, view):
        """Проверка существующей марки у элемента на виде"""
        collector = FilteredElementCollector(self.doc, view.Id)
        tags = collector.OfClass(IndependentTag).ToElements()
        
        for tag in tags:
            if tag.TaggedLocalElementId == element.Id:
                return True
        return False
    
    def CreateTag(self, element, view, tag_type):
        """Создание марки для элемента"""
        try:
            # Получаем точку для размещения марки
            bbox = element.get_BoundingBox(view)
            if not bbox:
                return False
            
            center = (bbox.Min + bbox.Max) / 2
            # конвертация мм в футы
            offset_x_feet = self.settings.offset_x / 304.8
            offset_y_feet = self.settings.offset_y / 304.8
            
            tag_point = XYZ(
                center.X + offset_x_feet,
                center.Y + offset_y_feet,
                center.Z
            )
            
            # Создаем марку с указанным типоразмером
            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                Reference(element),
                self.settings.use_leader,
                self.settings.orientation,
                tag_point
            )
            
            if tag:
                # Устанавливаем нужный типоразмер
                tag.ChangeTypeId(tag_type.Id)
                return True
            
            return False
            
        except Exception as e:
            print("Ошибка создания марки: {}".format(str(e)))
            return False
    
    # Методы навигации назад
    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabViews
    
    def OnBack2Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabCategories
    
    def OnBack3Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabTags
    
    def OnBack4Click(self, sender, args):
        self.tabControl.SelectedTab = self.tabSettings

# Форма для выбора семейства и типоразмера марки
class TagFamilySelectionForm(Form):
    def __init__(self, doc, available_families, current_family, current_type):
        self.doc = doc
        self.available_families = available_families
        self.current_family = current_family
        self.current_type = current_type
        self.selected_family = None
        self.selected_type = None
        
        self.InitializeComponent()
        self.PopulateFamiliesList()
    
    def InitializeComponent(self):
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
        self.lstFamilies.Items.Clear()
        
        for family in self.available_families:
            if family and hasattr(family, 'Name'):
                family_name = getattr(family, 'Name', 'Без имени')
                # Создаем объект для хранения и семейства и его имени
                item = FamilyListItem(family_name, family)
                self.lstFamilies.Items.Add(item)
        
        # Выбираем первое семейство по умолчанию
        if self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0
    
    def OnFamilySelected(self, sender, args):
        """Обработка выбора семейства"""
        if self.lstFamilies.SelectedItem:
            selected_item = self.lstFamilies.SelectedItem
            selected_family = selected_item.Tag
            self.PopulateTypesList(selected_family)
    
    def PopulateTypesList(self, family):
        """Заполнение списка типоразмеров для выбранного семейства"""
        self.lstTypes.Items.Clear()
        
        try:
            # Получаем все типоразмеры выбранного семейства
            symbol_ids = family.GetFamilySymbolIds()
            
            print("Найдено типоразмеров в семействе {}: {}".format(
                getattr(family, 'Name', 'Без имени'), 
                symbol_ids.Count if symbol_ids else 0))
            
            if symbol_ids and symbol_ids.Count > 0:
                symbol_id_list = list(symbol_ids)
                for symbol_id in symbol_id_list:
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = getattr(symbol, 'Name', 'Без имени')
                        print(" - Типоразмер: {}".format(symbol_name))
                        # Создаем объект для хранения и типоразмера и его имени
                        type_item = FamilyListItem(symbol_name, symbol)
                        self.lstTypes.Items.Add(type_item)
                
                # Выбираем первый типоразмер
                if self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
                else:
                    print("Типоразмеры найдены, но не удалось загрузить их имена")
            else:
                print("Семейство не содержит типоразмеров")
                
                # Альтернативный способ - попробовать получить типоразмеры через FamilySymbol
                collector = FilteredElementCollector(self.doc)
                symbols = collector.OfClass(FamilySymbol).WhereElementIsNotElementType().ToElements()
                
                family_symbols = []
                for symbol in symbols:
                    if symbol.Family and symbol.Family.Id == family.Id:
                        family_symbols.append(symbol)
                
                print("Альтернативный поиск: найдено типоразмеров: {}".format(len(family_symbols)))
                
                for symbol in family_symbols:
                    symbol_name = getattr(symbol, 'Name', 'Без имени')
                    print(" - Типоразмер (альт.): {}".format(symbol_name))
                    type_item = FamilyListItem(symbol_name, symbol)
                    self.lstTypes.Items.Add(type_item)
                
                if self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
                    
        except Exception as e:
            print("Ошибка при загрузке типоразмеров: " + str(e))
            # Показываем сообщение только если действительно нет типоразмеров
            if self.lstTypes.Items.Count == 0:
                MessageBox.Show("В выбранном семействе нет доступных типоразмеров!")
    
    def OnOKClick(self, sender, args):
        """Обработка нажатия OK"""
        if self.lstFamilies.SelectedItem and self.lstTypes.SelectedItem:
            self.selected_family = self.lstFamilies.SelectedItem.Tag
            self.selected_type = self.lstTypes.SelectedItem.Tag
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            MessageBox.Show("Выберите семейство и типоразмер марки!")
    
    def OnCancelClick(self, sender, args):
        """Обработка нажатия Отмена"""
        self.DialogResult = DialogResult.Cancel
        self.Close()
    
    @property
    def SelectedFamily(self):
        return self.selected_family
    
    @property
    def SelectedType(self):
        return self.selected_type

# Вспомогательный класс для отображения в ListBox
class FamilyListItem(object):
    def __init__(self, display_name, tag):
        self.display_name = display_name
        self.Tag = tag
    
    def __str__(self):
        return self.display_name

# Основная функция
def main():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    
    form = MainForm(doc, uidoc)
    Application.Run(form)

if __name__ == "__main__":
    main()