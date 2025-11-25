# -*- coding: utf-8 -*-
__title__ = "Длина полки марок"
__author__ = "Rage"
__doc__ = "Получение длины полки из выбранных марок"
__ver__ = "1"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import datetime

from Autodesk.Revit.DB import *
from System.Drawing import *
from System.Windows.Forms import *

MM_TO_FEET = 304.8


class TagSettings(object):
    def __init__(self):
        self.selected_categories = []
        self.selected_tag_types = []


class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.category_mapping = {}
        self.allow_tab_change = False
        self.available_types = {}  # family_name -> list of type_names

        self.InitializeComponent()
        self.CollectCategories()

    def InitializeComponent(self):
        self.Text = "Длина полки марок"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = ["1. Категории", "2. Марки", "3. Результат", "4. Марки на виде"]
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

    def SetupTab1(self, tab):
        controls = [
            self.CreateControl(
                Label,
                Text="Выберите категории элементов:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.CreateControl(
                CheckedListBox,
                Location=Point(10, 40),
                Size=Size(600, 400),
                CheckOnClick=True,
            ),
            self.CreateControl(
                Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.lblCategories, self.lstCategories, self.btnNext1 = controls
        self.btnNext1.Click += self.OnNext1Click
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab2(self, tab):
        self.lstTagFamilies = ListView()
        self.lstTagFamilies.Location = Point(10, 40)
        self.lstTagFamilies.Size = Size(700, 400)
        self.lstTagFamilies.View = View.Details
        self.lstTagFamilies.FullRowSelect = True
        self.lstTagFamilies.GridLines = True
        self.lstTagFamilies.CheckBoxes = True
        self.lstTagFamilies.Columns.Add("Категория", 150)
        self.lstTagFamilies.Columns.Add("Семейство марки", 200)
        self.lstTagFamilies.Columns.Add("Типоразмер марки", 250)
        self.lstTagFamilies.Columns.Add("Длина полки (мм)", 100)

        self.lblStatusTab2 = self.CreateControl(
            Label,
            Text="",
            Location=Point(10, 450),
            Size=Size(300, 20),
        )

        self.btnDuplicate = self.CreateControl(
            Button,
            Text="Дублировать марку",
            Location=Point(350, 450),
            Size=Size(120, 25),
        )
        self.btnDuplicate.Click += self.OnDuplicateClick

        controls = [
            self.CreateControl(
                Label,
                Text="Выберите типоразмеры марок:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.lstTagFamilies,
            self.lblStatusTab2,
            self.btnDuplicate,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)
            ),
            self.CreateControl(
                Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.btnBack1, self.btnNext2 = controls[4], controls[5]
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click

        for c in controls:
            tab.Controls.Add(c)

    def SetupTab3(self, tab):
        self.txtSummary = TextBox()
        self.txtSummary.Location = Point(10, 40)
        self.txtSummary.Size = Size(700, 400)
        self.txtSummary.Multiline = True
        self.txtSummary.ScrollBars = ScrollBars.Vertical
        self.txtSummary.ReadOnly = True

        self.btnViewTags = self.CreateControl(
            Button,
            Text="Далее →",
            Location=Point(600, 450),
            Size=Size(80, 25),
        )
        self.btnViewTags.Click += self.OnViewTagsClick

        controls = [
            self.CreateControl(
                Label, Text="Результат:", Location=Point(10, 10), Size=Size(300, 20)
            ),
            self.txtSummary,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)
            ),
            self.btnViewTags,
        ]
        self.btnBack2 = controls[2]
        self.btnBack2.Click += self.OnBack2Click
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab4(self, tab):
        # Создаем DataGridView для отображения марок на виде
        self.dgTags = DataGridView()
        self.dgTags.Location = Point(10, 40)
        self.dgTags.Size = Size(700, 400)
        self.dgTags.AllowUserToAddRows = False
        self.dgTags.AllowUserToDeleteRows = False
        self.dgTags.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.dgTags.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        # Создаем колонки
        colFamily = DataGridViewTextBoxColumn()
        colFamily.HeaderText = "Семейство"
        colFamily.ReadOnly = True
        colFamily.Name = "Family"
        colFamily.Width = 200
        self.dgTags.Columns.Add(colFamily)

        # Создаем колонку с комбобоксом
        colType = DataGridViewComboBoxColumn()
        colType.HeaderText = "Типоразмер"
        colType.Name = "Type"
        colType.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        colType.FlatStyle = FlatStyle.Flat
        colType.Width = 250
        self.dgTags.Columns.Add(colType)

        colCategory = DataGridViewTextBoxColumn()
        colCategory.HeaderText = "Категория"
        colCategory.ReadOnly = True
        colCategory.Name = "Category"
        colCategory.Width = 150
        self.dgTags.Columns.Add(colCategory)

        # Обработчики событий
        self.dgTags.EditingControlShowing += self.OnEditingControlShowing
        self.dgTags.CellValueChanged += self.OnCellValueChanged
        self.dgTags.DataError += self.OnDataError

        # Кнопка обновления
        self.btnRefresh = self.CreateControl(
            Button,
            Text="Обновить",
            Location=Point(600, 450),
            Size=Size(80, 25),
        )
        self.btnRefresh.Click += self.OnRefreshClick

        # Статус
        self.lblStatusTab4 = self.CreateControl(
            Label,
            Text="Нажмите 'Обновить' для загрузки марок",
            Location=Point(10, 450),
            Size=Size(500, 20),
        )

        controls = [
            self.CreateControl(
                Label,
                Text="Марки на активном виде:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.dgTags,
            self.btnRefresh,
            self.lblStatusTab4,
        ]
        for c in controls:
            tab.Controls.Add(c)

    def OnDataError(self, sender, args):
        # Игнорируем ошибки данных в комбобоксах
        args.ThrowException = False

    def OnNext1Click(self, sender, args):
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            name = self.lstCategories.Items[i]
            if self.lstCategories.GetItemChecked(i) and name in self.category_mapping:
                self.settings.selected_categories.append(self.category_mapping[name])

        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return

        self.PopulateTagFamilies()
        self.allow_tab_change = True
        self.tabControl.SelectedIndex = 1

    def OnNext2Click(self, sender, args):
        self.settings.selected_tag_types = []
        for item in self.lstTagFamilies.Items:
            if item.Checked:
                self.settings.selected_tag_types.append(item.Tag)
        if not self.settings.selected_tag_types:
            MessageBox.Show("Выберите хотя бы один типоразмер марки!")
            return
        self.txtSummary.Text = self.GenerateSummary()
        self.allow_tab_change = True
        self.tabControl.SelectedIndex = 2

    def OnBack1Click(self, sender, args):
        self.allow_tab_change = True
        self.tabControl.SelectedIndex = 0

    def OnBack2Click(self, sender, args):
        self.allow_tab_change = True
        self.tabControl.SelectedIndex = 1

    def OnViewTagsClick(self, sender, args):
        print("OnViewTagsClick: Переход на вкладку 4")
        self.allow_tab_change = True
        self.tabControl.SelectedIndex = 3
        # Автоматически собираем типоразмеры, если еще не собраны
        if not self.available_types:
            print(
                "OnViewTagsClick: Сбор типоразмеров не выполнен, запускаем CollectAllTagTypes"
            )
            self.CollectAllTagTypes()
        else:
            print("OnViewTagsClick: Типоразмеры уже собраны, пропускаем")
        # Автоматически загружаем марки на виде
        print("OnViewTagsClick: Загрузка марок на виде")
        self.PopulateTagsOnView()

    def OnRefreshClick(self, sender, args):
        print("OnRefreshClick: Принудительная перезагрузка типоразмеров и марок")
        # Принудительно перезагружаем типоразмеры
        self.CollectAllTagTypes()
        self.PopulateTagsOnView()

    def CollectAllTagTypes(self):
        """Собираем ВСЕ доступные типоразмеры маркеров в проекте"""
        print("CollectAllTagTypes: Начало сбора типоразмеров")
        self.available_types = {}

        self.lblStatusTab4.Text = "Сбор всех типоразмеров маркеров..."

        try:
            # Получаем ВСЕ семейства в проекте
            collector = FilteredElementCollector(self.doc).OfClass(Family)
            all_families = list(collector)
            print("CollectAllTagTypes: Найдено семейств: {}".format(len(all_families)))

            # Категории маркеров
            tag_category_ids = [
                ElementId(BuiltInCategory.OST_DuctTags),
                ElementId(BuiltInCategory.OST_FlexDuctTags),
                ElementId(BuiltInCategory.OST_DuctInsulationsTags),
                ElementId(BuiltInCategory.OST_DuctTerminalTags),
                ElementId(BuiltInCategory.OST_DuctAccessoryTags),
                ElementId(BuiltInCategory.OST_MechanicalEquipmentTags),
            ]
            print(
                "CollectAllTagTypes: Категории маркеров: {}".format(
                    len(tag_category_ids)
                )
            )

            for family in all_families:
                try:
                    # Проверяем, является ли семейство маркером
                    if (
                        family.FamilyCategory
                        and family.FamilyCategory.Id in tag_category_ids
                    ):
                        family_name = family.Name
                        symbol_ids = family.GetFamilySymbolIds()
                        print(
                            "CollectAllTagTypes: Обработка семейства '{}' с {} символами".format(
                                family_name, len(symbol_ids)
                            )
                        )

                        if family_name not in self.available_types:
                            self.available_types[family_name] = []

                        # Собираем все типоразмеры этого семейства
                        for symbol_id in symbol_ids:
                            symbol = self.doc.GetElement(symbol_id)
                            if symbol and symbol.IsActive:
                                type_name = self.GetSymbolTypeName(symbol)
                                print(
                                    "CollectAllTagTypes: Символ '{}' имеет тип '{}'".format(
                                        symbol.Name, type_name
                                    )
                                )
                                if (
                                    type_name
                                    and type_name
                                    not in self.available_types[family_name]
                                ):
                                    self.available_types[family_name].append(type_name)

                except Exception as e:
                    print(
                        "CollectAllTagTypes: Ошибка при обработке семейства: {}".format(
                            str(e)
                        )
                    )
                    continue

            # Диагностика
            total_types = sum(len(types) for types in self.available_types.values())
            print(
                "CollectAllTagTypes: Собрано семейств: {}, типоразмеров: {}".format(
                    len(self.available_types), total_types
                )
            )
            self.lblStatusTab4.Text = (
                "Загружено {} семейств маркеров, {} типоразмеров".format(
                    len(self.available_types), total_types
                )
            )

            # Выводим отладочную информацию только при первом сборе
            if self.available_types and not hasattr(self, "_debug_shown"):
                debug_info = "Доступные семейства маркеров:\n"
                for family_name, types in self.available_types.items():
                    debug_info += "{}: {} типов\n".format(family_name, len(types))
                MessageBox.Show(debug_info, "Отладка - доступные типоразмеры")
                self._debug_shown = True

        except Exception as e:
            print("CollectAllTagTypes: Критическая ошибка: {}".format(str(e)))
            MessageBox.Show("Ошибка при сборе типоразмеров: " + str(e), "Ошибка")

    def GetSymbolTypeName(self, symbol):
        """Получаем имя типоразмера символа"""
        if not symbol:
            return "Без имени"

        # Сначала пробуем получить имя из параметра "Type Name"
        type_name_param = symbol.LookupParameter("Type Name")
        if type_name_param and type_name_param.HasValue:
            name = type_name_param.AsString()
            if name and name.strip():
                return name.strip()

        # Пробуем параметр "Тип"
        type_param = symbol.LookupParameter("Тип")
        if type_param and type_param.HasValue:
            name = type_param.AsString()
            if name and name.strip():
                return name.strip()

        # Пробуем параметр имени символа
        name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if name_param and name_param.HasValue:
            name = name_param.AsString()
            if name and name.strip():
                return name.strip()

        # Используем свойство Name
        if hasattr(symbol, "Name") and symbol.Name:
            name = symbol.Name
            if name and name.strip():
                return name.strip()

        return "Типоразмер_" + str(symbol.Id.IntegerValue)

    def OnDuplicateClick(self, sender, args):
        selected_items = []
        for item in self.lstTagFamilies.Items:
            if item.Checked:
                selected_items.append(item)

        if not selected_items:
            self.lblStatusTab2.Text = "Выберите марку для дублирования!"
            return

        trans = Transaction(self.doc, "Дублирование марок")
        duplicated_count = 0
        try:
            trans.Start()
            for selected_item in selected_items:
                tag_type = selected_item.Tag
                if not isinstance(tag_type, FamilySymbol):
                    continue

                new_name, shelf_num = self.generate_new_name_and_num(tag_type)
                new_symbol = tag_type.Duplicate(new_name)

                # Установить длину полки в новый символ
                param_names = ["Длина полки", "Shelf Length"]
                for param_name in param_names:
                    shelf_param = new_symbol.LookupParameter(param_name)
                    if shelf_param and not shelf_param.IsReadOnly:
                        shelf_param.Set(shelf_num / MM_TO_FEET)
                        break

                duplicated_count += 1
            trans.Commit()
            self.lblStatusTab2.Text = "Дублировано {0} марок.".format(duplicated_count)
            self.PopulateTagFamilies()
        except Exception as e:
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
            self.lblStatusTab2.Text = "Ошибка дублирования: {0}".format(e)
        finally:
            trans.Dispose()

    def generate_new_name_and_num(self, tag_type):
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
                return symbol_name + "_копия_" + datetime.datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                ), 0
        else:
            return symbol_name + "_копия_" + datetime.datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            ), 0

        family = tag_type.Family
        existing_names = set()
        symbol_ids = family.GetFamilySymbolIds()
        for symbol_id in symbol_ids:
            symbol = self.doc.GetElement(symbol_id)
            if symbol and symbol != tag_type:
                name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param and name_param.HasValue:
                    name = name_param.AsString()
                else:
                    name = symbol.Name
                existing_names.add(name)

        num = base_num + 1
        while True:
            new_name = "{}_{}".format(prefix, num)
            if new_name not in existing_names:
                break
            num += 1

        return new_name, num

    def PopulateTagsOnView(self):
        print("PopulateTagsOnView: Начало загрузки марок на виде")
        # Очищаем DataGridView
        self.dgTags.Rows.Clear()
        print("PopulateTagsOnView: DataGridView очищен")

        view = self.uidoc.ActiveView
        if not view:
            print("PopulateTagsOnView: Активный вид не найден")
            MessageBox.Show("Нет активного вида!")
            return
        print("PopulateTagsOnView: Активный вид: {}".format(view.Name))

        # Категории марок
        tag_categories = [
            BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctTags,
            BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
        ]
        print(
            "PopulateTagsOnView: Категории для поиска: {}".format(len(tag_categories))
        )

        element_count = 0
        tag_count = 0

        # Собираем марки типа FamilyInstance
        collector = FilteredElementCollector(self.doc, view.Id).OfClass(FamilyInstance)
        for element in collector:
            if element.Category and element.Category.Id.IntegerValue < 0:
                builtin_cat = BuiltInCategory(element.Category.Id.IntegerValue)
                if builtin_cat in tag_categories:
                    # Получаем имя семейства
                    if element.Symbol and element.Symbol.Family:
                        family_name = element.Symbol.Family.Name
                    else:
                        family_name = "Неизвестно"

                    # Получаем имя типоразмера из параметра "Тип"
                    type_name = self.GetElementTypeName(element)
                    print(
                        "PopulateTagsOnView: FamilyInstance - семейство: {}, тип: {}".format(
                            family_name, type_name
                        )
                    )
                    category_name = (
                        element.Category.Name if element.Category else "Без категории"
                    )

                    # Добавляем строку в DataGridView
                    row_index = self.dgTags.Rows.Add()
                    self.dgTags.Rows[row_index].Cells["Family"].Value = family_name
                    self.dgTags.Rows[row_index].Cells["Category"].Value = category_name

                    # Для комбобокса устанавливаем значение и заполняем варианты
                    type_cell = self.dgTags.Rows[row_index].Cells["Type"]
                    if isinstance(type_cell, DataGridViewComboBoxCell):
                        # Устанавливаем текущее значение
                        type_cell.Value = type_name
                        print(
                            "PopulateTagsOnView: Установлено значение комбо: {}".format(
                                type_name
                            )
                        )

                        # Заполняем варианты для этого семейства
                        type_cell.Items.Clear()

                        # ВСЕГДА добавляем текущее значение
                        if type_name and type_name not in type_cell.Items:
                            type_cell.Items.Add(type_name)

                        # Добавляем все доступные типы для этого семейства
                        if family_name in self.available_types:
                            print(
                                "PopulateTagsOnView: Добавление {} типов для семейства {}".format(
                                    len(self.available_types[family_name]), family_name
                                )
                            )
                            for available_type in self.available_types[family_name]:
                                if available_type != type_name:
                                    type_cell.Items.Add(available_type)
                        else:
                            # Если семейства нет в available_types, добавляем только текущее значение
                            if type_name and type_name not in type_cell.Items:
                                type_cell.Items.Add(type_name)
                            print(
                                "PopulateTagsOnView: Семейство {} не найдено в available_types".format(
                                    family_name
                                )
                            )

                    # Сохраняем элемент для последующего использования
                    self.dgTags.Rows[row_index].Tag = element
                    element_count += 1

        print(
            "PopulateTagsOnView: Найдено {} FamilyInstance марок".format(element_count)
        )

        # Собираем независимые марки
        collector_tags = FilteredElementCollector(self.doc, view.Id).OfClass(
            IndependentTag
        )
        for tag in collector_tags:
            if tag.Category and tag.Category.Id.IntegerValue < 0:
                builtin_cat = BuiltInCategory(tag.Category.Id.IntegerValue)
                if builtin_cat in tag_categories:
                    category_name = (
                        tag.Category.Name if tag.Category else "Без категории"
                    )

                    # Для независимых марок используем категорию как имя семейства
                    family_name = category_name

                    # Получаем имя типоразмера из параметра "Тип"
                    type_value = self.GetElementTypeName(tag)
                    print(
                        "PopulateTagsOnView: IndependentTag - категория: {}, тип: {}".format(
                            family_name, type_value
                        )
                    )

                    # Добавляем строку в DataGridView
                    row_index = self.dgTags.Rows.Add()
                    self.dgTags.Rows[row_index].Cells["Family"].Value = family_name
                    self.dgTags.Rows[row_index].Cells["Category"].Value = category_name

                    # Для комбобокса устанавливаем значение и заполняем варианты
                    type_cell = self.dgTags.Rows[row_index].Cells["Type"]
                    if isinstance(type_cell, DataGridViewComboBoxCell):
                        # Устанавливаем текущее значение
                        type_cell.Value = type_value

                        # Заполняем варианты для этого семейства
                        type_cell.Items.Clear()

                        # ВСЕГДА добавляем текущее значение
                        if type_value and type_value not in type_cell.Items:
                            type_cell.Items.Add(type_value)

                        # Добавляем все доступные типы для этого семейства
                        if family_name in self.available_types:
                            for available_type in self.available_types[family_name]:
                                if available_type != type_value:
                                    type_cell.Items.Add(available_type)
                        else:
                            # Если семейства нет в available_types, добавляем только текущее значение
                            if type_value and type_value not in type_cell.Items:
                                type_cell.Items.Add(type_value)

                    # Сохраняем элемент для последующего использования
                    self.dgTags.Rows[row_index].Tag = tag
                    tag_count += 1

        print("PopulateTagsOnView: Найдено {} IndependentTag марок".format(tag_count))

        self.lblStatusTab4.Text = "Найдено маркеров: {} ({} элементов, {} тегов). Доступно {} семейств для замены".format(
            element_count + tag_count,
            element_count,
            tag_count,
            len(self.available_types),
        )
        print("PopulateTagsOnView: Итого марок: {}".format(element_count + tag_count))

    def GetElementTypeName(self, element):
        """Получаем имя типоразмера из параметра 'Тип' элемента"""
        if not element:
            return "Без типа"

        # Пробуем параметр "Тип"
        type_param = element.LookupParameter("Тип")
        if type_param and type_param.HasValue:
            name = type_param.AsString()
            if name and name.strip():
                return name.strip()

        # Пробуем параметр "Type"
        type_param_en = element.LookupParameter("Type")
        if type_param_en and type_param_en.HasValue:
            name = type_param_en.AsString()
            if name and name.strip():
                return name.strip()

        # Пробуем параметр "Type Name"
        type_name_param = element.LookupParameter("Type Name")
        if type_name_param and type_name_param.HasValue:
            name = type_name_param.AsString()
            if name and name.strip():
                return name.strip()

        return "Нет типа"

    def OnEditingControlShowing(self, sender, args):
        if (
            isinstance(args.Control, ComboBox)
            and self.dgTags.CurrentCell.ColumnIndex == 1
        ):
            combo = args.Control
            combo.DropDownStyle = ComboBoxStyle.DropDownList
            row_index = self.dgTags.CurrentCell.RowIndex
            family_name = self.dgTags.Rows[row_index].Cells["Family"].Value

            if family_name:
                # Сохраняем текущее значение
                current_value = self.dgTags.Rows[row_index].Cells["Type"].Value

                # Очищаем и заполняем комбобокс
                combo.Items.Clear()

                # ВСЕГДА добавляем текущее значение
                if current_value and current_value not in combo.Items:
                    combo.Items.Add(current_value)

                # Добавляем все доступные типы для этого семейства
                if family_name in self.available_types:
                    for type_name in self.available_types[family_name]:
                        if type_name != current_value and type_name not in combo.Items:
                            combo.Items.Add(type_name)

    def OnCellValueChanged(self, sender, args):
        print(
            "OnCellValueChanged: Изменение в строке {}, колонке {}".format(
                args.RowIndex, args.ColumnIndex
            )
        )
        if args.ColumnIndex == 1 and args.RowIndex >= 0:
            row = self.dgTags.Rows[args.RowIndex]
            tag_obj = row.Tag
            selected_type = row.Cells["Type"].Value
            print(
                "OnCellValueChanged: Выбран тип: {}, объект: {}".format(
                    selected_type, type(tag_obj)
                )
            )

            if selected_type and tag_obj and selected_type != "":
                trans = Transaction(self.doc, "Изменение типоразмера марки")
                print("OnCellValueChanged: Транзакция начата")
                try:
                    trans.Start()

                    # Для ВСЕХ элементов меняем параметр "Тип"
                    type_param = tag_obj.LookupParameter("Тип")
                    if not type_param:
                        # Пробуем английское имя
                        type_param = tag_obj.LookupParameter("Type")
                        print(
                            "OnCellValueChanged: Используем параметр 'Type' вместо 'Тип'"
                        )

                    if type_param and not type_param.IsReadOnly:
                        print(
                            "OnCellValueChanged: Установка параметра в '{}'".format(
                                selected_type
                            )
                        )
                        type_param.Set(selected_type)
                        trans.Commit()
                        print("OnCellValueChanged: Транзакция завершена успешно")
                        try:
                            self.doc.Regenerate()
                            self.uidoc.RefreshActiveView()
                            print("OnCellValueChanged: Вид обновлен")
                        except:
                            print("OnCellValueChanged: Ошибка обновления вида")
                        MessageBox.Show("Типоразмер изменен")
                        # Обновляем отображаемое значение в колонке "Текущий тип"
                        row.Cells[2].Value = selected_type
                        print(
                            "OnCellValueChanged: Строка обновлена, повторная загрузка"
                        )
                        self.PopulateTagsOnView()
                    else:
                        trans.RollBack()
                        print("OnCellValueChanged: Параметр не найден или readonly")
                        MessageBox.Show("Не удалось найти параметр 'Тип' для изменения")

                except Exception as e:
                    if trans.GetStatus() == TransactionStatus.Started:
                        trans.RollBack()
                    print("OnCellValueChanged: Ошибка: {}".format(str(e)))
                    MessageBox.Show("Ошибка: " + str(e))
                    MessageBox.Show("Ошибка при изменении типоразмера: " + str(e))

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
        self.lstCategories.Items.Clear()
        self.category_mapping.clear()

        for cat in sorted(
            self.settings.selected_categories, key=lambda x: self.GetCategoryName(x)
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
        except:
            pass
        return getattr(category, "Name", "Неизвестная категория")

    def PopulateTagFamilies(self):
        self.lstTagFamilies.Items.Clear()

        for category in self.settings.selected_categories:
            available_families = self.GetAvailableTagFamiliesForCategory(category)
            for family in available_families:
                symbol_ids = family.GetFamilySymbolIds()
                for symbol_id in symbol_ids:
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol and symbol.IsActive:
                        item = ListViewItem(self.GetCategoryName(category))
                        item.SubItems.Add(family.Name)
                        item.SubItems.Add(self.GetSymbolTypeName(symbol))
                        shelf_length = self.GetShelfLength(symbol)
                        item.SubItems.Add(str(shelf_length))
                        item.Tag = symbol
                        item.Checked = False
                        self.lstTagFamilies.Items.Add(item)

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

    def GetTagCategoryId(self, element_category):
        mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_MechanicalEquipment: BuiltInCategory.OST_MechanicalEquipmentTags,
        }

        try:
            if hasattr(element_category, "Id") and element_category.Id.IntegerValue < 0:
                element_cat = BuiltInCategory(element_category.Id.IntegerValue)
                if element_cat in mapping:
                    tag_cat = Category.GetCategory(self.doc, mapping[element_cat])
                    if tag_cat:
                        return tag_cat.Id
        except:
            pass
        return None

    def GenerateSummary(self):
        summary = "СВОДКА ПО ВЫБРАННЫМ ТИПОРАЗМЕРАМ МАРОК И ДЛИНАМ ПОЛОК:\r\n\r\n"
        summary += (
            "Выбрано типоразмеров: "
            + str(len(self.settings.selected_tag_types))
            + "\r\n\r\n"
        )

        summary += "Детали:\r\n"
        for tag_type in self.settings.selected_tag_types:
            shelf_length = self.GetShelfLength(tag_type)
            category_name = self.GetCategoryName(tag_type.Category)
            summary += (
                "- Категория: "
                + category_name
                + ", Семейство: "
                + tag_type.Family.Name
                + ", Типоразмер: "
                + self.GetSymbolTypeName(tag_type)
                + " - Длина полки: "
                + str(shelf_length)
                + " мм\r\n"
            )
        return summary

    def GetShelfLength(self, tag_type):
        param_names = ["Длина полки", "Shelf Length"]
        for param_name in param_names:
            try:
                param = tag_type.LookupParameter(param_name)
                if param and param.HasValue:
                    value = param.AsDouble() * MM_TO_FEET
                    return round(value, 2)
            except:
                pass
        return 0.0

    def OnTabSelecting(self, sender, args):
        if not self.allow_tab_change:
            args.Cancel = True
        else:
            self.allow_tab_change = False


def main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        if doc and uidoc:
            Application.Run(MainForm(doc, uidoc))
        else:
            MessageBox.Show("Нет доступа к документу Revit")
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))


if __name__ == "__main__":
    main()
