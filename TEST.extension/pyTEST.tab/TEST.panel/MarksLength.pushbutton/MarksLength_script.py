# -*- coding: utf-8 -*-
__title__ = """Автодлина
марок"""
__author__ = "Rage"
__doc__ = (
    """Автоматическа корректировка длины полки марок 
    на виде в зависимости от содержимого"""
)
__ver__ = "1.1"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import json
import math
import os
import sys
import traceback

from Autodesk.Revit.DB import *
from System import Array, Environment
from System.Drawing import *
from System.Windows.Forms import *

# Логирование для pyRevit
try:
    from pyrevit import script
    from pyrevit import forms
    from pyrevit import output
    logger = script.get_logger()
    output = script.get_output()
    PYREVIT_AVAILABLE = True
except:
    PYREVIT_AVAILABLE = False

MM_TO_FEET = 304.8


class Settings(object):
    """Класс для хранения настроек скрипта: выбранные виды, категории и параметры."""

    def __init__(self):
        self.selected_views = []
        self.selected_categories = []
        self.available_parameters = {}
        self.selected_parameters = {}


class MainForm(Form):
    """Главная форма скрипта для анализа и корректировки марок."""

    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = Settings()
        self.all_views_dict = {}
        self.views_dict = {}
        self.category_mapping = {}

        self.LoadTagDefaults()

        self.InitializeComponent()
        self.Load3DViews()
        
        # Инициализируем логирование
        if PYREVIT_AVAILABLE:
            logger.info("Инициализация формы MainForm")

    def InitializeComponent(self):
        self.Text = "Объединенный анализ и корректировка марок"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = ["1. Виды", "2. Категории и Параметры", "3. Результаты"]
        for i, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            if i == 0:
                self.SetupTab1(tab)
            elif i == 1:
                self.SetupTab2(tab)
            elif i == 2:
                self.SetupTab3(tab)
            self.tabControl.TabPages.Add(tab)

        self.Controls.Add(self.tabControl)
        
        if PYREVIT_AVAILABLE:
            logger.debug("Форма MainForm инициализирована")

    def CreateControl(self, control_type, **kwargs):
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control

    def SetupTab1(self, tab):  # Виды
        self.txtSearchViews = self.CreateControl(
            TextBox, Location=Point(120, 35), Size=Size(140, 20)
        )
        self.btnSelectAll = self.CreateControl(
            Button, Text="Выбрать все", Location=Point(270, 35), Size=Size(100, 25)
        )
        self.btnSelectAll.Click += lambda s, e: self.OnToggleViews(True)
        self.btnDeselectAll = self.CreateControl(
            Button, Text="Снять выбор", Location=Point(380, 35), Size=Size(100, 25)
        )
        self.btnDeselectAll.Click += lambda s, e: self.OnToggleViews(False)
        controls = [
            self.CreateControl(
                Label,
                Text="Выберите 3D виды:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.CreateControl(
                Label, Text="Поиск:", Location=Point(70, 35), Size=Size(50, 20)
            ),
            self.txtSearchViews,
            self.btnSelectAll,
            self.btnDeselectAll,
            self.CreateControl(
                CheckedListBox,
                Location=Point(10, 65),
                Size=Size(600, 360),
                CheckOnClick=True,
            ),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 440)),
        ]
        self.lblViews = controls[0]
        self.lblSearch = controls[1]
        self.lstViews = controls[5]
        self.btnNext1 = controls[6]
        self.btnNext1.Click += self.OnNext1Click
        for c in controls:
            tab.Controls.Add(c)
        self.txtSearchViews.TextChanged += self.OnSearchViewsTextChanged
        
        if PYREVIT_AVAILABLE:
            logger.debug("Вкладка 1 (Виды) настроена")

    def SetupTab2(self, tab):  # Категории и Параметры
        self.lstParameterSelection = ListView()
        self.lstParameterSelection.Location = Point(10, 40)
        self.lstParameterSelection.Size = Size(700, 300)
        self.lstParameterSelection.View = View.Details
        self.lstParameterSelection.FullRowSelect = True
        self.lstParameterSelection.GridLines = True
        self.lstParameterSelection.Columns.Add("Категория", 200)
        self.lstParameterSelection.Columns.Add("Параметр", 400)
        self.lstParameterSelection.DoubleClick += self.OnParameterSelectionDoubleClick

        controls = [
            self.CreateControl(
                Label,
                Text="Выберите параметр для каждой категории:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.lstParameterSelection,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 350), Size=Size(80, 25)
            ),
            self.CreateControl(
                Button, Text="Далее →", Location=Point(600, 350), Size=Size(80, 25)
            ),
        ]
        self.btnBack1 = controls[2]
        self.btnNext2 = controls[3]
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click
        for c in controls:
            tab.Controls.Add(c)
        self.CollectCategories()
        self.CollectParameters()
        self.PopulateParameterSelection()
        
        if PYREVIT_AVAILABLE:
            logger.debug("Вкладка 2 (Категории и Параметры) настроена")

    def SetupTab3(self, tab):
        """Настраивает вкладку 3: результаты обработки."""
        self.txtResults = self.CreateControl(
            TextBox,
            Location=Point(10, 60),
            Size=Size(700, 350),
            Multiline=True,
            ReadOnly=True,
            ScrollBars=ScrollBars.Vertical,
        )
        self.progressBar = self.CreateControl(
            ProgressBar,
            Location=Point(10, 420),
            Size=Size(700, 20),
            Minimum=0,
            Maximum=100,
            Value=0,
        )
        controls = [
            self.CreateControl(
                Label,
                Text="Результаты анализа и корректировки:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.txtResults,
            self.progressBar,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(260, 450), Size=Size(80, 30)
            ),
            self.CreateControl(
                Button, Text="Завершить", Location=Point(360, 450), Size=Size(100, 30)
            ),
            self.CreateControl(
                Button,
                Text="Выполнить корректировку",
                Location=Point(480, 450),
                Size=Size(170, 30),
            ),
        ]
        self.btnBack3, self.btnFinish, self.btnExecute = (
            controls[3],
            controls[4],
            controls[5],
        )
        self.btnBack3.Click += self.OnBack3Click
        self.btnFinish.Click += self.OnFinishClick
        self.btnExecute.Click += self.OnExecuteClick
        for c in controls:
            tab.Controls.Add(c)
        
        if PYREVIT_AVAILABLE:
            logger.debug("Вкладка 3 (Результаты) настроена")

    def Load3DViews(self):
        try:
            views = (
                FilteredElementCollector(self.doc)
                .OfClass(View3D)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            self.lstViews.Items.Clear()
            self.all_views_dict.clear()
            for view in views:
                if (
                    view.CanBePrinted
                    and not view.IsTemplate
                    and not view.Name.startswith("{")
                ):
                    self.lstViews.Items.Add(view.Name, False)
                    self.all_views_dict[view.Name] = view
            self.UpdateViewsList("")
            
            if PYREVIT_AVAILABLE:
                logger.info("Загружено {} 3D видов".format(len(self.all_views_dict)))
        except Exception as e:
            if PYREVIT_AVAILABLE:
                logger.error("Ошибка при загрузке 3D видов: {}".format(str(e)))
            else:
                print("Ошибка при загрузке 3D видов: {}".format(str(e)))

    def UpdateViewsList(self, filter_text):
        self.lstViews.Items.Clear()
        self.views_dict.clear()
        for name, view in self.all_views_dict.items():
            if not filter_text or filter_text.lower() in name.lower():
                checked = name in [v.Name for v in self.settings.selected_views]
                self.lstViews.Items.Add(name, checked)
                self.views_dict[name] = view

    def OnToggleViews(self, checked):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, checked)

    def OnSearchViewsTextChanged(self, sender, args):
        self.UpdateViewsList(sender.Text)

    def OnNext1Click(self, sender, args):
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            if self.lstViews.GetItemChecked(i):
                self.settings.selected_views.append(
                    self.all_views_dict[self.lstViews.Items[i]]
                )
        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting
        
        if PYREVIT_AVAILABLE:
            logger.info("Выбрано {} видов для обработки".format(len(self.settings.selected_views)))

    def OnBack1Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 0
        self.tabControl.Selecting += self.OnTabSelecting

    def OnNext2Click(self, sender, args):
        selected_cats = [item.Tag for item in self.lstParameterSelection.Items]
        for item in self.lstParameterSelection.Items:
            if item.SubItems[1].Text != "Не выбран":
                selected_cats.append(item.Tag)
        if not any(p for p in self.settings.selected_parameters.values() if p):
            MessageBox.Show("Выберите параметры для категорий!")
            return
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def CollectCategories(self):
        try:
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
                    if PYREVIT_AVAILABLE:
                        logger.warning("Не удалось получить категорию {}: {}".format(cat, str(e)))

            self.settings.selected_categories = list(unique_cats)
            self.CollectParameters()
            
            if PYREVIT_AVAILABLE:
                logger.info("Собрано {} категорий".format(len(self.settings.selected_categories)))
        except Exception as e:
            if PYREVIT_AVAILABLE:
                logger.error("Ошибка при сборе категорий: {}".format(str(e)))
            else:
                print("Ошибка при сборе категорий: {}".format(str(e)))

    def CollectParameters(self):
        self.settings.parameters_by_category = {}

        for category in self.settings.selected_categories:
            try:
                # Использовать WhereElementIsNotElementType для исключения типов
                collector = (
                    FilteredElementCollector(self.doc)
                    .OfCategoryId(category.Id)
                    .WhereElementIsNotElementType()
                )

                elements = collector.ToElements()
                if not elements:
                    if PYREVIT_AVAILABLE:
                        logger.debug("Нет элементов в категории {}".format(self.GetCategoryName(category)))
                    continue

                # Собрать параметры из первых 5 элементов
                params_set = set()
                num_elements = min(len(elements), 5)
                for i in range(num_elements):
                    el = elements[i]
                    # Параметры instance
                    for param in el.Parameters:
                        if (
                            param
                            and param.Definition
                            and param.Definition.Name
                            and param.Definition.Name not in ["Имя типа", "Тип"]
                        ):
                            params_set.add(param.Definition.Name)

                    # Параметры type
                    type_elem = self.doc.GetElement(el.GetTypeId())
                    if type_elem:
                        for param in type_elem.Parameters:
                            if (
                                param
                                and param.Definition
                                and param.Definition.Name
                                and param.Definition.Name not in ["Имя типа", "Тип"]
                            ):
                                params_set.add(param.Definition.Name)

                sorted_params = sorted(list(params_set))
                self.settings.parameters_by_category[category.Id.IntegerValue] = (
                    sorted_params
                )
                # Использовать сохранённый по умолчанию или первый
                default_param = self.tag_defaults.get(
                    self.GetCategoryName(category), {}
                ).get("param")
                if default_param and default_param in sorted_params:
                    self.settings.selected_parameters[category.Id.IntegerValue] = (
                        default_param
                    )
                else:
                    self.settings.selected_parameters[category.Id.IntegerValue] = (
                        sorted_params[0] if sorted_params else None
                    )
                    
                if PYREVIT_AVAILABLE:
                    logger.debug("Категория {}: найдено {} параметров".format(self.GetCategoryName(category), len(sorted_params)))

            except Exception as e:
                if PYREVIT_AVAILABLE:
                    logger.error("Ошибка при сборе параметров для категории {}: {}".format(self.GetCategoryName(category), str(e)))
                else:
                    print("Ошибка при сборе параметров для категории {}: {}".format(self.GetCategoryName(category), str(e)))

    def PopulateParameterSelection(self):
        self.lstParameterSelection.Items.Clear()
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            param = self.settings.selected_parameters.get(
                category.Id.IntegerValue, "Не выбран"
            )
            item.SubItems.Add(param)
            item.Tag = category.Id.IntegerValue
            self.lstParameterSelection.Items.Add(item)

    def OnParameterSelectionDoubleClick(self, sender, args):
        if self.lstParameterSelection.SelectedItems.Count == 0:
            return
        selected = self.lstParameterSelection.SelectedItems[0]
        cat_id = selected.Tag
        category_obj = next(
            (
                c
                for c in self.settings.selected_categories
                if c.Id.IntegerValue == cat_id
            ),
            None,
        )
        if not category_obj:
            return
        available_params = self.settings.parameters_by_category.get(cat_id, [])
        current_param = self.settings.selected_parameters.get(cat_id)
        form = ParameterSelectionForm(
            self.doc, category_obj, available_params, current_param
        )
        try:
            if form.ShowDialog() == DialogResult.OK:
                self.settings.selected_parameters[cat_id] = form.SelectedParameter
                selected.SubItems[1].Text = form.SelectedParameter or "Не выбран"
                # Сохранить выбор в defaults
                cat_name = self.GetCategoryName(category_obj)
                if cat_name not in self.tag_defaults:
                    self.tag_defaults[cat_name] = {}
                self.tag_defaults[cat_name]["param"] = form.SelectedParameter
                self.SaveTagDefaults()
                if PYREVIT_AVAILABLE:
                    logger.info("Выбран параметр '{}' для категории '{}'".format(form.SelectedParameter, cat_name))
        except Exception as e:
            error_msg = "Ошибка в диалоге выбора параметров: {}".format(str(e))
            if PYREVIT_AVAILABLE:
                logger.error(error_msg)
            MessageBox.Show(error_msg)

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnFinishClick(self, sender, args):
        if PYREVIT_AVAILABLE:
            logger.info("Завершение работы скрипта")
        self.Close()

    def LoadTagDefaults(self):
        self.tag_defaults = {}
        script_dir = os.path.dirname(__file__)
        defaults_path = os.path.join(script_dir, "tag_defaults.json")
        if os.path.exists(defaults_path):
            try:
                with open(defaults_path, "r") as f:
                    self.tag_defaults = json.load(f)
                if PYREVIT_AVAILABLE:
                    logger.info("Загружены настройки из {}".format(defaults_path))
            except Exception as e:
                if PYREVIT_AVAILABLE:
                    logger.warning("Не удалось загрузить настройки: {}".format(str(e)))
        else:
            if PYREVIT_AVAILABLE:
                logger.info("Файл настроек не найден, будут использованы значения по умолчанию")

    def SaveTagDefaults(self):
        script_dir = os.path.dirname(__file__)
        defaults_path = os.path.join(script_dir, "tag_defaults.json")
        try:
            with open(defaults_path, "w") as f:
                json.dump(self.tag_defaults, f, indent=4, ensure_ascii=False)
            if PYREVIT_AVAILABLE:
                logger.debug("Настройки сохранены в {}".format(defaults_path))
        except Exception as e:
            if PYREVIT_AVAILABLE:
                logger.error("Ошибка при сохранении настроек: {}".format(str(e)))

    def generate_new_name_and_num(self, tag_type, required_length):
        name_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param and name_param.HasValue:
            symbol_name = name_param.AsString()
        else:
            symbol_name = tag_type.Name

        parts = symbol_name.split("_")
        if len(parts) > 1:
            try:
                base_num = int(parts[-1])
                prefix = "_".join(parts[:-1])
            except ValueError:
                prefix = symbol_name
        else:
            prefix = symbol_name

        family = tag_type.Family
        existing_names = set()
        symbol_ids = family.GetFamilySymbolIds()
        for symbol_id in symbol_ids:
            symbol = self.doc.GetElement(symbol_id)
            if symbol:
                name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param and name_param.HasValue:
                    name = name_param.AsString()
                else:
                    name = symbol.Name
                existing_names.add(name)

        target_name = "{}_{}".format(prefix, int(required_length))
        # Проверить, существует ли символ с target_name и подходящей длиной
        for symbol_id in symbol_ids:
            symbol = self.doc.GetElement(symbol_id)
            if (
                symbol
                and symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                == target_name
            ):
                # Проверить длину полки
                shelf_param_names = [
                    "Длина полки",
                    "Shelf Length",
                    "Length",
                    "Длина",
                ]
                for param_name in shelf_param_names:
                    shelf_param = symbol.LookupParameter(param_name)
                    if (
                        shelf_param
                        and shelf_param.HasValue
                        and shelf_param.StorageType == StorageType.Double
                    ):
                        current_length = shelf_param.AsDouble() * 304.8  # в мм
                        if abs(current_length - required_length) < 1:
                            return target_name, int(required_length), symbol.Id

        # Если не найден подходящий, создать новый уникальный
        num = int(required_length)
        while True:
            new_name = "{}_{}".format(prefix, num)
            if new_name not in existing_names:
                return new_name, num, None
            num += 1  # Инкремент в случае конфликта

    # Анализ и корректировка
    def OnExecuteClick(self, sender, args):
        # Убираем подавление вывода и включаем полное логирование
        if PYREVIT_AVAILABLE:
            logger.info("=" * 80)
            logger.info("НАЧАЛО ВЫПОЛНЕНИЯ КОРРЕКТИРОВКИ МАРОК")
            logger.info("=" * 80)
        else:
            print("=" * 80)
            print("НАЧАЛО ВЫПОЛНЕНИЯ КОРРЕКТИРОВКИ МАРОК")
            print("=" * 80)
            
        tag_categories = [
            BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
        ]

        results = []
        total_tags_all = 0
        tags_with_category = 0
        tags_with_elements = 0
        tags_with_param = 0
        tags_with_value = 0
        changed_tags = 0
        created_symbols = {}
        actual_category_ids = set()

        # Кэширование маркеров для производительности
        all_tags = []
        for view in self.settings.selected_views:
            if PYREVIT_AVAILABLE:
                logger.info("Сбор марок для вида: {}".format(view.Name))
            collector = FilteredElementCollector(self.doc, view.Id).OfClass(
                IndependentTag
            )
            tags_in_view = collector.ToElements()
            all_tags.extend(tags_in_view)
            if PYREVIT_AVAILABLE:
                logger.debug("Вид {}: найдено {} марок".format(view.Name, len(tags_in_view)))

        # Инициализация прогресса
        self.progressBar.Value = 0
        processed_tags = 0
        total_tags_estimated = len(all_tags) or 1
        
        if PYREVIT_AVAILABLE:
            logger.info("Всего марок для обработки: {}".format(len(all_tags)))

        trans = Transaction(self.doc, "Корректировка типоразмеров марок")
        try:
            trans.Start()
            
            for tag in all_tags:
                try:
                    total_tags_all += 1
                    
                    # Обновление прогресса
                    processed_tags += 1
                    progress = int(processed_tags * 100 / total_tags_estimated)
                    if progress > 100:
                        progress = 100
                    self.progressBar.Value = progress
                    
                    # Периодическое логирование прогресса
                    if processed_tags % 50 == 0 and PYREVIT_AVAILABLE:
                        logger.debug("Обработано {}/{} марок ({}%)".format(processed_tags, len(all_tags), progress))

                    if tag.Category:
                        actual_category_ids.add(tag.Category.Id.IntegerValue)
                    else:
                        if PYREVIT_AVAILABLE:
                            logger.debug("Марка ID {}: нет категории, пропускаем".format(tag.Id))
                        continue
                        
                    if tag.Category and tag.Category.Id.IntegerValue in [
                        c.value__ for c in tag_categories
                    ]:
                        tags_with_category += 1
                        tagged_elements = tag.GetTaggedLocalElements()
                        if tagged_elements:
                            tags_with_elements += 1
                            element = tagged_elements[0]
                            category = element.Category
                            param_name = self.settings.selected_parameters.get(
                                category.Id.IntegerValue
                            )
                            if param_name:
                                tags_with_param += 1
                                # Проверка instance или type параметра
                                param = element.LookupParameter(param_name)
                                if not param:
                                    # Проверить type parameter
                                    type_elem = self.doc.GetElement(element.GetTypeId())
                                    if type_elem:
                                        param = type_elem.LookupParameter(param_name)
                                if param:
                                    if param and param.HasValue:
                                        tags_with_value += 1
                                        value = param.AsValueString()
                                        if value is None:
                                            value = ""
                                        char_count = len(value)
                                        required_length = math.ceil(char_count * 1.6 + 1)
                                        
                                        if PYREVIT_AVAILABLE:
                                            logger.debug("Марка ID {}: значение '{}', символов: {}, требуемая длина: {}мм".format(tag.Id, value, char_count, required_length))

                                        # Получить базовый символ марки
                                        base_symbol = self.doc.GetElement(tag.GetTypeId())
                                        if not isinstance(base_symbol, FamilySymbol):
                                            if PYREVIT_AVAILABLE:
                                                logger.warning("Марка ID {}: не является семейством, пропускаем".format(tag.Id))
                                            continue
                                        try:
                                            if not base_symbol.Family:
                                                if PYREVIT_AVAILABLE:
                                                    logger.warning("Марка ID {}: нет семейства, пропускаем".format(tag.Id))
                                                continue
                                        except Exception as fam_e:
                                            if PYREVIT_AVAILABLE:
                                                logger.warning("Марка ID {}: ошибка при получении семейства: {}".format(tag.Id, str(fam_e)))
                                            continue

                                        family_id = base_symbol.Family.Id.IntegerValue

                                        # Инициализировать словарь для семейства
                                        if family_id not in created_symbols:
                                            created_symbols[family_id] = {}

                                        # Создать новый символ, если нужен
                                        if (
                                            required_length
                                            not in created_symbols[family_id]
                                        ):
                                            new_name, shelf_num, existing_id = (
                                                self.generate_new_name_and_num(
                                                    base_symbol, required_length
                                                )
                                            )

                                            if existing_id is not None:
                                                # Использовать существующий символ
                                                new_symbol = self.doc.GetElement(
                                                    existing_id
                                                )
                                                created_symbols[family_id][
                                                    required_length
                                                ] = new_symbol
                                                if PYREVIT_AVAILABLE:
                                                    logger.info("Создан новый символ: {} (длина {}мм) для семейства ID {}".format(new_name, required_length, family_id))
                                            else:
                                                # Создать новый символ
                                                new_symbol = base_symbol.Duplicate(new_name)
                                                # Установить "Длина полки"
                                                shelf_param_names = [
                                                    "Длина полки",
                                                    "Shelf Length",
                                                    "Length",
                                                    "Длина",
                                                ]
                                                shelf_set = False
                                                for param_name_shelf in shelf_param_names:
                                                    shelf_param = (
                                                        new_symbol.LookupParameter(
                                                            param_name_shelf
                                                        )
                                                    )
                                                    if (
                                                        shelf_param
                                                        and not shelf_param.IsReadOnly
                                                        and shelf_param.StorageType
                                                        == StorageType.Double
                                                    ):
                                                        shelf_param.Set(
                                                            required_length / 304.8
                                                        )
                                                        shelf_set = True
                                                        if PYREVIT_AVAILABLE:
                                                            logger.debug("Установлена длина полки {}мм для символа {}".format(required_length, new_name))
                                                        break
                                                if not shelf_set:
                                                    if PYREVIT_AVAILABLE:
                                                        logger.warning("Не найдено параметра для установки длины полки в символе {}".format(new_name))
                                                created_symbols[family_id][
                                                    required_length
                                                ] = new_symbol

                                        # Назначить новый символ марке
                                        new_symbol = created_symbols[family_id][
                                            required_length
                                        ]
                                        if tag.GetTypeId() != new_symbol.Id:
                                            tag.ChangeTypeId(new_symbol.Id)
                                            changed_tags += 1
                                            if PYREVIT_AVAILABLE:
                                                logger.debug("Марка ID {}: изменен тип на {}".format(tag.Id, new_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))

                                        results.append(
                                            "Марка ID {}: присвоен тип '{}' с длиной полки {}мм".format(
                                                tag.Id, new_name, required_length
                                            )
                                        )
                                    else:
                                        if PYREVIT_AVAILABLE:
                                            logger.debug("Марка ID {}: параметр '{}' не имеет значения".format(tag.Id, param_name))
                                else:
                                    if PYREVIT_AVAILABLE:
                                        logger.debug("Марка ID {}: параметр '{}' не найден".format(tag.Id, param_name))
                            else:
                                if PYREVIT_AVAILABLE:
                                    logger.debug("Марка ID {}: нет выбранного параметра для категории {}".format(tag.Id, category.Name if category else 'Unknown'))
                        else:
                            if PYREVIT_AVAILABLE:
                                logger.debug("Марка ID {}: нет привязанных элементов".format(tag.Id))
                    else:
                        if PYREVIT_AVAILABLE:
                            logger.debug("Марка ID {}: категория {} не в списке целевых".format(tag.Id, tag.Category.Name if tag.Category else 'Unknown'))

                except Exception as tag_error:
                    error_msg = "Ошибка при обработке марки ID {}: {}".format(tag.Id, str(tag_error))
                    if PYREVIT_AVAILABLE:
                        logger.error(error_msg)
                        logger.error(traceback.format_exc())
                    else:
                        print(error_msg)
                    results.append(error_msg)

            trans.Commit()
            if PYREVIT_AVAILABLE:
                logger.info("Транзакция успешно завершена")
                
        except Exception as e:
            error_msg = "Критическая ошибка: {}".format(str(e))
            if PYREVIT_AVAILABLE:
                logger.error(error_msg)
                logger.error(traceback.format_exc())
            else:
                print(error_msg)
            results.append(error_msg)
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
                if PYREVIT_AVAILABLE:
                    logger.warning("Транзакция отменена")
        finally:
            trans.Dispose()

        # Формирование итогового отчета
        lines = [
            "=" * 60,
            "ИТОГОВЫЙ ОТЧЕТ ОБ ОБРАБОТКЕ",
            "=" * 60,
            "Всего марок: {}".format(total_tags_all),
            "С целевой категорией: {}".format(tags_with_category),
            "С привязанными элементами: {}".format(tags_with_elements),
            "С выбранным параметром: {}".format(tags_with_param),
            "С заполненным значением: {}".format(tags_with_value),
            "Изменено марок: {}".format(changed_tags),
            "Создано символов: {}".format(sum(len(symbols) for symbols in created_symbols.values())),
            "=" * 60,
            "ДЕТАЛЬНЫЙ ЛОГ:"
        ]
        
        # Добавляем детальные результаты
        for result in results[-20:]:  # Последние 20 записей
            lines.append(result)
            
        log_text = Environment.NewLine.join(lines)

        self.txtResults.Text = log_text
        
        # Выводим итоги в консоль pyRevit
        if PYREVIT_AVAILABLE:
            logger.info("=" * 60)
            logger.info("ИТОГИ ОБРАБОТКИ")
            logger.info("=" * 60)
            logger.info("Всего марок: {}".format(total_tags_all))
            logger.info("С целевой категорией: {}".format(tags_with_category))
            logger.info("С привязанными элементами: {}".format(tags_with_elements))
            logger.info("С выбранным параметром: {}".format(tags_with_param))
            logger.info("С заполненным значением: {}".format(tags_with_value))
            logger.info("Изменено марок: {}".format(changed_tags))
            logger.info("Создано символов: {}".format(sum(len(symbols) for symbols in created_symbols.values())))
            logger.info("=" * 60)
            logger.info("ВЫПОЛНЕНИЕ ЗАВЕРШЕНО")
            logger.info("=" * 60)
        else:
            print("=" * 60)
            print("ИТОГИ ОБРАБОТКИ")
            print("=" * 60)
            print("Всего марок: {}".format(total_tags_all))
            print("С целевой категорией: {}".format(tags_with_category))
            print("С привязанными элементами: {}".format(tags_with_elements))
            print("С выбранным параметром: {}".format(tags_with_param))
            print("С заполненным значением: {}".format(tags_with_value))
            print("Изменено марок: {}".format(changed_tags))
            print("Создано символов: {}".format(sum(len(symbols) for symbols in created_symbols.values())))
            print("=" * 60)
            print("ВЫПОЛНЕНИЕ ЗАВЕРШЕНО")
            print("=" * 60)

    def OnTabSelecting(self, sender, args):
        args.Cancel = True

    def GetCategoryName(self, category):
        if not category:
            return "Неизвестная категория"
        try:
            if hasattr(category, "Id") and category.Id.IntegerValue < 0:
                return LabelUtils.GetLabelFor(BuiltInCategory(category.Id.IntegerValue))
        except:
            pass
        return getattr(category, "Name", "Неизвестная категория")


class ParameterSelectionForm(Form):
    def __init__(self, doc, category, available_params, current_param):
        self.doc = doc
        self.category = category
        self.available_params = available_params
        self.selected_parameter = current_param
        self.filtered_params = list(available_params)

        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "Выбор параметра для категории '{}'".format(
            __revit__.Application.ActiveUIDocument.Document.GetCategoryName(
                self.category
            )
            if self.category
            else "Неизвестная"
        )
        self.Size = Size(450, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.lblSearch = Label()
        self.lblSearch.Text = "Поиск:"
        self.lblSearch.Location = Point(10, 10)
        self.lblSearch.Size = Size(50, 20)

        self.txtSearch = TextBox()
        self.txtSearch.Location = Point(70, 10)
        self.txtSearch.Size = Size(200, 20)
        self.txtSearch.TextChanged += self.OnSearchTextChanged

        self.lstParams = ListBox()
        self.lstParams.Location = Point(10, 40)
        self.lstParams.Size = Size(410, 250)
        self.lstParams.SelectionMode = SelectionMode.One
        self.lstParams.Items.AddRange(Array[Object](self.filtered_params))
        if self.selected_parameter and self.selected_parameter in self.filtered_params:
            self.lstParams.SelectedItem = self.selected_parameter

        self.btnOK = Button()
        self.btnOK.Text = "OK"
        self.btnOK.Location = Point(120, 310)
        self.btnOK.Size = Size(80, 30)
        self.btnOK.Click += self.OnOKClick

        self.btnCancel = Button()
        self.btnCancel.Text = "Отмена"
        self.btnCancel.Location = Point(220, 310)
        self.btnCancel.Size = Size(80, 30)
        self.btnCancel.Click += self.OnCancelClick

        self.Controls.AddRange(
            [self.lblSearch, self.txtSearch, self.lstParams, self.btnOK, self.btnCancel]
        )

    def OnSearchTextChanged(self, sender, args):
        search_text = sender.Text.lower()
        self.filtered_params = [
            p for p in self.available_params if search_text in p.lower()
        ]
        self.lstParams.Items.Clear()
        self.lstParams.Items.AddRange(Array[Object](self.filtered_params))

    def OnOKClick(self, sender, args):
        if self.lstParams.SelectedItem:
            self.selected_parameter = self.lstParams.SelectedItem
        self.DialogResult = DialogResult.OK
        self.Close()

    def OnCancelClick(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    @property
    def SelectedParameter(self):
        return self.selected_parameter


def main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        if doc and uidoc:
            if PYREVIT_AVAILABLE:
                logger.info("Запуск скрипта Автодлина марок")
                logger.info("Версия: {}".format(__ver__))
            Application.Run(MainForm(doc, uidoc))
    except Exception as e:
        error_msg = "Ошибка при запуске скрипта: {}".format(str(e))
        if PYREVIT_AVAILABLE:
            logger.error(error_msg)
            logger.error(traceback.format_exc())
        else:
            print(error_msg)
            traceback.print_exc()
        MessageBox.Show(error_msg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)


if __name__ == "__main__":
    main()