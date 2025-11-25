# -*- coding: utf-8 -*-
__title__ = "Объединенный анализ и корректировка марок"
__author__ = "Rage"
__doc__ = "Выбор видов, категорий, параметров; анализ марок; корректировка типоразмеров по длине полки"
__ver__ = "1"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import math

from Autodesk.Revit.DB import *
from System import Array
from System.Drawing import *
from System.Windows.Forms import *

MM_TO_FEET = 304.8


class Settings(object):
    def __init__(self):
        self.selected_views = []
        self.selected_categories = []
        self.available_parameters = {}
        self.selected_parameters = {}


class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = Settings()
        self.all_views_dict = {}
        self.views_dict = {}
        self.category_mapping = {}

        self.InitializeComponent()
        self.Load3DViews()

    def InitializeComponent(self):
        print("Инициализация компонентов формы")
        self.Text = "Объединенный анализ и корректировка марок"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
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
        print("Настройка вкладки 1: Виды")
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
        print("Настройка вкладки 2: Категории и Параметры")
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

    def SetupTab3(self, tab):  # Результаты
        print("Настройка вкладки 3: Результаты")
        self.txtResults = self.CreateControl(
            TextBox,
            Location=Point(10, 40),
            Size=Size(700, 400),
            Multiline=True,
            ReadOnly=True,
            ScrollBars=ScrollBars.Vertical,
        )
        controls = [
            self.CreateControl(
                Label,
                Text="Результаты анализа и корректировки:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.txtResults,
            self.CreateControl(Button, Text="← Назад", Location=Point(340, 450)),
            self.CreateControl(
                Button,
                Text="Выполнить корректировку",
                Location=Point(500, 450),
                Size=Size(200, 30),
            ),
        ]
        self.btnBack2 = controls[2]
        self.btnExecute = controls[3]
        self.btnBack2.Click += self.OnBack2Click
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
        print("Переход к вкладке 2: выбор видов")
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            if self.lstViews.GetItemChecked(i):
                self.settings.selected_views.append(
                    self.all_views_dict[self.lstViews.Items[i]]
                )
        if not self.settings.selected_views:
            print("Выберите хотя бы один вид!")
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

    def CollectParameters(self):
        self.settings.parameters_by_category = {}
        try:
            for category in self.settings.selected_categories:
                collector = FilteredElementCollector(self.doc).OfCategory(
                    BuiltInCategory(category.Id.IntegerValue)
                )
                elements = list(collector)
                params = set()
                for el in elements:
                    if el.Parameters:
                        for param in el.Parameters:
                            if param and param.Definition and param.Definition.Name:
                                if param.Definition.Name not in [
                                    "Имя типа",
                                    "Тип",
                                ]:  # keep "Марка" and others
                                    params.add(param.Definition.Name)
                sorted_params = sorted(list(params))
                self.settings.parameters_by_category[category.Id.IntegerValue] = (
                    sorted_params
                )
                self.settings.selected_parameters[category.Id.IntegerValue] = (
                    sorted_params[0] if sorted_params else None
                )
        except Exception as e:
            print("Ошибка сбора параметров: " + str(e))

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
        if form.ShowDialog() == DialogResult.OK:
            self.settings.selected_parameters[cat_id] = form.SelectedParameter
            selected.SubItems[1].Text = form.SelectedParameter or "Не выбран"
            print(
                "Выбран параметр '{}' для категории '{}' (ID: {})".format(
                    form.SelectedParameter, self.GetCategoryName(category_obj), cat_id
                )
            )

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    # Анализ и корректировка
    # Анализ и корректировка
    def OnExecuteClick(self, sender, args):
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
        changed_tags = 0  # Новый счетчик
        created_symbols = {}  # family_id -> {required_length: new_symbol}
        actual_category_ids = set()

        trans = Transaction(self.doc, "Корректировка типоразмеров марок")
        try:
            trans.Start()
            for view in self.settings.selected_views:
                print("Начинаем обработку вида: {}".format(view.Name))
                collector = FilteredElementCollector(self.doc, view.Id).OfClass(
                    IndependentTag
                )
                for tag in collector:
                    total_tags_all += 1
                    print("Обработка марки ID {}".format(tag.Id))
                    if tag.Category:
                        actual_category_ids.add(tag.Category.Id.IntegerValue)
                        print(
                            "  Категория марки ID: {}".format(
                                tag.Category.Id.IntegerValue
                            )
                        )
                    else:
                        print("  Марка без категории, пропуск")
                        continue
                    if tag.Category and tag.Category.Id.IntegerValue in [
                        c.value__ for c in tag_categories
                    ]:
                        tags_with_category += 1
                        print("  Марка в целевой категории")
                        tagged_elements = tag.GetTaggedLocalElements()
                        if tagged_elements:
                            tags_with_elements += 1
                            element = tagged_elements[0]
                            category = element.Category
                            print(
                                "  Привязанный элемент ID {}, категория: {} (ID: {})".format(
                                    element.Id,
                                    self.GetCategoryName(category),
                                    category.Id.IntegerValue,
                                )
                            )
                            param_name = self.settings.selected_parameters.get(
                                category.Id.IntegerValue
                            )
                            if param_name:
                                tags_with_param += 1
                                print("  Выбранный параметр: '{}'".format(param_name))
                                param = element.LookupParameter(param_name)
                                if param:
                                    print(
                                        "  Параметр найден, HasValue: {}".format(
                                            param.HasValue
                                        )
                                    )
                                    if param and param.HasValue:
                                        tags_with_value += 1
                                        value = str(param.AsString())
                                        char_count = len(value)
                                        required_length = math.ceil(
                                            char_count * 1.6 + 1
                                        )
                                        print(
                                            "  Значение параметра: '{}', длина текста: {}, требуемая длина: {}".format(
                                                value, char_count, required_length
                                            )
                                        )

                                        try:
                                            # Получить базовый символ марки
                                            base_symbol = self.doc.GetElement(
                                                tag.GetTypeId()
                                            )
                                            if not isinstance(
                                                base_symbol, FamilySymbol
                                            ):
                                                print(
                                                    "Марка ID {} не является FamilySymbol, пропуск".format(
                                                        tag.Id
                                                    )
                                                )
                                                continue
                                            family_id = (
                                                base_symbol.Family.Id.IntegerValue
                                            )

                                            # Инициализировать словарь для семейства
                                            if family_id not in created_symbols:
                                                created_symbols[family_id] = {}

                                            # Создать новый символ, если нужен
                                            if (
                                                required_length
                                                not in created_symbols[family_id]
                                            ):
                                                new_name = "Base_{}mm".format(
                                                    required_length
                                                )
                                                # Убедиться в уникальности имени (простая версия, улучшить если нужно)
                                                existing_names = [
                                                    sym.Name
                                                    for sym in FilteredElementCollector(
                                                        self.doc
                                                    )
                                                    .OfClass(FamilySymbol)
                                                    .WhereElementIsElementType()
                                                    .ToElements()
                                                    if sym.Family.Id
                                                    == base_symbol.Family.Id
                                                    and sym.Name != base_symbol.Name
                                                ]
                                                counter = 1
                                                base_new_name = new_name
                                                while any(
                                                    s
                                                    for s in existing_names
                                                    if s == new_name
                                                ):
                                                    new_name = "{}_{}".format(
                                                        base_new_name, counter
                                                    )
                                                    counter += 1

                                                new_symbol = base_symbol.Duplicate(
                                                    new_name
                                                )
                                                # Установить "Длина полки"
                                                param_names = [
                                                    "Длина полки",
                                                    "Shelf Length",
                                                ]
                                                shelf_set = False
                                                for param_name_shelf in param_names:
                                                    shelf_param = (
                                                        new_symbol.LookupParameter(
                                                            param_name_shelf
                                                        )
                                                    )
                                                    if (
                                                        shelf_param
                                                        and not shelf_param.IsReadOnly
                                                    ):
                                                        shelf_param.Set(
                                                            required_length / MM_TO_FEET
                                                        )  # В футах
                                                        shelf_set = True
                                                        break
                                                if not shelf_set:
                                                    print(
                                                        "Предупреждение: Не удалось установить 'Длина полки' для символа {}".format(
                                                            new_name
                                                        )
                                                    )
                                                created_symbols[family_id][
                                                    required_length
                                                ] = new_symbol
                                                print(
                                                    "Создан новый символ: {} с длиной {}мм".format(
                                                        new_name, required_length
                                                    )
                                                )

                                            # Назначить новый символ марке
                                            new_symbol = created_symbols[family_id][
                                                required_length
                                            ]
                                            if tag.GetTypeId() != new_symbol.Id:
                                                tag.ChangeElementType(new_symbol.Id)
                                                changed_tags += 1
                                                print(
                                                    "Изменен тип марки ID {} на {}".format(
                                                        tag.Id, new_symbol.Name
                                                    )
                                                )

                                            results.append(
                                                "Марка ID {}: присвоен тип '{}' с длиной полки {}мм".format(
                                                    tag.Id,
                                                    new_symbol.Name,
                                                    required_length,
                                                )
                                            )
                                        except Exception as e:
                                            print(
                                                "Ошибка при обработке марки ID {}: {}".format(
                                                    tag.Id, str(e)
                                                )
                                            )
                                            continue
                                    else:
                                        print("  Параметр не заполнен, пропуск")
                                else:
                                    print("  Параметр не найден у элемента")
                            else:
                                print("  Параметр не выбран для категории, пропуск")
                        else:
                            print("  Нет привязанных элементов, пропуск")
                    else:
                        print("  Марка не в целевой категории, пропуск")

            trans.Commit()
        except Exception as e:
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
            print("Ошибка в транзакции: {}".format(str(e)))
            results.append("Ошибка: {}".format(str(e)))
        finally:
            trans.Dispose()

        log_text = """Полный лог обработки:
Всего марок: {}
Найденные IDs категорий марок: {}
Ожидаемые IDs: {}
С категорией: {}
С привязанными элементами: {}
С выбранным параметром: {}
С заполненным значением: {}
Изменено: {}""".format(
            total_tags_all,
            ", ".join(str(id) for id in sorted(actual_category_ids)),
            ", ".join(str(c.value__) for c in tag_categories),
            tags_with_category,
            tags_with_elements,
            tags_with_param,
            tags_with_value,
            changed_tags,
        )
        print(log_text)

        self.txtResults.Text = "\n".join(results) if results else "Ничего не найдено."
        print("Общий результат: {} записей".format(len(results)))

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
    print("Запуск скрипта: анализ и корректировка марок")
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        if doc and uidoc:
            Application.Run(MainForm(doc, uidoc))
        else:
            print("Нет доступа к документу Revit")
    except Exception as e:
        print("Ошибка: " + str(e))


if __name__ == "__main__":
    main()
