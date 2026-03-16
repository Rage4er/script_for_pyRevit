# -*- coding: utf-8 -*-
__title__ = """Автодлина
марок"""
__author__ = "Rage"
__doc__ = (
    """Автоматическа корректировка длины полки марок 
    на виде в зависимости от содержимого"""
)
__ver__ = "1.0"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import json
import math
import os

from Autodesk.Revit.DB import *
from System import Array, Environment
from System.Drawing import *
from System.Windows.Forms import *

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

    def Load3DViews(self):
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
            except:
                pass

        self.settings.selected_categories = list(unique_cats)
        self.CollectParameters()

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

            except Exception as e:
                pass

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
        except Exception, e:
            MessageBox.Show("Ошибка в диалоге выбора параметров: {}".format(str(e)))

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnFinishClick(self, sender, args):
        self.Close()

    def LoadTagDefaults(self):
        self.tag_defaults = {}
        script_dir = os.path.dirname(__file__)
        defaults_path = os.path.join(script_dir, "tag_defaults.json")
        if os.path.exists(defaults_path):
            try:
                with open(defaults_path, "r") as f:
                    self.tag_defaults = json.load(f)
            except Exception, e:
                pass

    def SaveTagDefaults(self):
        script_dir = os.path.dirname(__file__)
        defaults_path = os.path.join(script_dir, "tag_defaults.json")
        try:
            with open(defaults_path, "w") as f:
                json.dump(self.tag_defaults, f, indent=4)
        except Exception, e:
            pass

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
        # Перенаправляем вывод в null для предотвращения отображения консоли
        import os
        import sys

        sys.stdout = open(os.devnull, "w")
        sys.stderr = open(os.devnull, "w")
        # Скрываем консоль pyRevit для предотвращения открытия при выполнении
        try:
            import pyrevit

            pyrevit.console.Hide()
        except:
            pass
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
            collector = FilteredElementCollector(self.doc, view.Id).OfClass(
                IndependentTag
            )
            all_tags.extend(collector.ToElements())

        # Инициализация прогресса
        self.progressBar.Value = 0
        processed_tags = 0
        total_tags_estimated = len(all_tags) or 1

        trans = Transaction(self.doc, "Корректировка типоразмеров марок")
        try:
            trans.Start()
            for tag in all_tags:
                total_tags_all += 1
                print("Обработка марки ID {}".format(tag.Id))

                # Обновление прогресса
                processed_tags += 1
                progress = int(processed_tags * 100 / total_tags_estimated)
                if progress > 100:
                    progress = 100
                self.progressBar.Value = progress

                if tag.Category:
                    actual_category_ids.add(tag.Category.Id.IntegerValue)
                else:
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

                                    # Получить базовый символ марки
                                    base_symbol = self.doc.GetElement(tag.GetTypeId())
                                    if not isinstance(base_symbol, FamilySymbol):
                                        continue
                                    try:
                                        if not base_symbol.Family:
                                            continue
                                    except Exception as fam_e:
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
                                                    break
                                            if not shelf_set:
                                                pass
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

                                    results.append(
                                        "Марка ID {}: присвоен тип '{}' с длиной полки {}мм".format(
                                            tag.Id, new_name, required_length
                                        )
                                    )
                                else:
                                    pass
                            else:
                                pass
                        else:
                            pass
                    else:
                        pass

            trans.Commit()
        except Exception as e:
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
            results.append("Ошибка: {}".format(str(e)))
        finally:
            trans.Dispose()

        # Повторно скрываем консоль на случай, если она открылась
        try:
            import pyrevit

            pyrevit.console.Hide()
        except:
            pass

        # Восстанавливаем stdout и stderr
        try:
            sys.stdout.close()
            sys.stderr.close()
        except:
            pass

        lines = [
            "Полный лог обработки:",
            "Всего марок: {}".format(total_tags_all),
            "С категорией: {}".format(tags_with_category),
            "С привязанными элементами: {}".format(tags_with_elements),
            "С выбранным параметром: {}".format(tags_with_param),
            "С заполненным значением: {}".format(tags_with_value),
            "Изменено: {}".format(changed_tags),
        ]
        log_text = Environment.NewLine.join(lines)

        self.txtResults.Text = log_text

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


class ParameterSelectionForm(Form):
    def __init__(self, doc, category, available_params, current_param):
        self.doc = doc
        self.category = category
        self.available_params = available_params
        self.selected_parameter = current_param

        self.InitializeComponent()

    def InitializeComponent(self):
        self.Text = "Выбор параметра для " + LabelUtils.GetLabelFor(
            BuiltInCategory(self.category.Id.IntegerValue)
        )
        self.Size = Size(400, 300)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True
        self.lstParams = ListBox()
        self.lstParams.Location = Point(10, 40)
        self.lstParams.Size = Size(360, 200)
        self.lstParams.DoubleClick += self.OnParamSelected

        controls = [
            Label(
                Text="Выберите параметр:", Location=Point(10, 10), Size=Size(200, 20)
            ),
            self.lstParams,
            Button(Text="OK", Location=Point(200, 250), Size=Size(75, 25)),
            Button(Text="Отмена", Location=Point(285, 250), Size=Size(75, 25)),
        ]
        self.btnOK, self.btnCancel = controls[2], controls[3]
        self.btnOK.Click += self.OnOKClick
        self.btnCancel.Click += self.OnCancelClick

        self.PopulateParams()

        for c in controls:
            self.Controls.Add(c)

    def PopulateParams(self):
        self.lstParams.Items.Clear()
        for param in self.available_params:
            self.lstParams.Items.Add(param)
        if self.selected_parameter and self.selected_parameter in self.available_params:
            self.lstParams.SelectedItem = self.selected_parameter

    def OnParamSelected(self, sender, args):
        if self.lstParams.SelectedItem:
            self.selected_parameter = self.lstParams.SelectedItem
            self.DialogResult = DialogResult.OK
            self.Close()

    def OnOKClick(self, sender, args):
        if self.lstParams.SelectedItem:
            self.selected_parameter = self.lstParams.SelectedItem
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            print("Выберите параметр!")

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
            Application.Run(MainForm(doc, uidoc))
    except Exception, e:
        pass


if __name__ == "__main__":
    main()
