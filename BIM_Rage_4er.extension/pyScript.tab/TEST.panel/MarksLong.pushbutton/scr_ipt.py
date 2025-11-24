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

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    # Анализ и корректировка
    def OnExecuteClick(self, sender, args):
        MessageBox.Show("Начало корректировки марок")
        tag_categories = [
            BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
        ]

        print(
            "Категории марок: {}".format(
                ", ".join(str(c.value__) for c in tag_categories)
            )
        )

        results = []
        trans = Transaction(self.doc, "Корректировка марок")
        trans.Start()
        try:
            # Counts for logging
            total_tags_all = 0
            tags_with_category = 0
            tags_with_elements = 0
            tags_with_param = 0
            tags_with_value = 0
            tags_length_mismatch = 0
            actual_category_ids = set()
            for view in self.settings.selected_views:
                print("Начинаем обработку вида: {}".format(view.Name))
                collector = FilteredElementCollector(self.doc, view.Id).OfClass(
                    IndependentTag
                )
                for tag in collector:
                    total_tags_all += 1
                    if tag.Category:
                        actual_category_ids.add(tag.Category.Id.IntegerValue)
                    tagged_elements = tag.GetTaggedLocalElements()
                    if tagged_elements:
                        tags_with_category += (
                            1  # renamed for clarity, now means has elements
                        )
                        element = tagged_elements[0]
                        category = element.Category
                        print(
                            "Element category: {} ({})".format(
                                category.Name if category else "None",
                                category.Id.IntegerValue if category else 0,
                            )
                        )
                        if category.Id.IntegerValue in [
                            c.Id.IntegerValue for c in self.settings.selected_categories
                        ]:
                            tags_with_elements += 1  # now means category matches
                            param_name = self.settings.selected_parameters.get(
                                category.Id.IntegerValue
                            )
                            print("Param name for category: {}".format(param_name))
                            if param_name:
                                tags_with_param += 1
                                param = element.LookupParameter(param_name)
                                if param and param.HasValue:
                                    tags_with_value += 1
                                    value = str(param.AsString())
                                    char_count = len(value)
                                    required_length = math.ceil(char_count * 1.6 + 1)
                                    print(
                                        "Value: '{}', required length: {}".format(
                                            value, required_length
                                        )
                                    )

                                    shelf_param = self.doc.GetElement(
                                        tag.Type
                                    ).LookupParameter("Длина полки")
                                    current_length = 0
                                    if shelf_param and shelf_param.HasValue:
                                        current_length = round(
                                            shelf_param.AsDouble() * MM_TO_FEET, 0
                                        )
                                    print(
                                        "Current shelf length: {}".format(
                                            current_length
                                        )
                                    )

                                    if current_length != required_length:
                                        tags_length_mismatch += 1
                                        # Perform changes
                        results = []
                        trans = Transaction(self.doc, "Корректировка марок")
                        trans.Start()
                        try:
                            # First pass: create all type copies
                            length_to_type = {}  # required_length: type_id
                            for view in self.settings.selected_views:
                                print("Начинаем обработку вида: {}".format(view.Name))
                                collector = FilteredElementCollector(
                                    self.doc, view.Id
                                ).OfClass(IndependentTag)
                                for tag in collector:
                                    try:
                                        if (
                                            tag.Category
                                            and tag.Category.Id.IntegerValue
                                            in [c.value__ for c in tag_categories]
                                        ):
                                            tagged_elements = (
                                                tag.GetTaggedLocalElements()
                                            )
                                            if tagged_elements:
                                                element = tagged_elements[0]
                                                category = element.Category
                                                param_name = self.settings.selected_parameters.get(
                                                    category
                                                )
                                                if param_name:
                                                    param = element.LookupParameter(
                                                        param_name
                                                    )
                                                    if param and param.HasValue:
                                                        value = str(param.AsString())
                                                        char_count = len(value)
                                                        required_length = math.ceil(
                                                            char_count * 1.6 + 1
                                                        )

                                                        tag_type_id = tag.GetTypeId()
                                                        if not tag_type_id:
                                                            continue
                                                        family_symbol = (
                                                            self.doc.GetElement(
                                                                tag_type_id
                                                            )
                                                        )
                                                        if not family_symbol:
                                                            continue

                                                        if (
                                                            required_length
                                                            not in length_to_type
                                                        ):
                                                            base_name = (
                                                                family_symbol.Name.split(
                                                                    "_"
                                                                )[0]
                                                                if "_"
                                                                in family_symbol.Name
                                                                else family_symbol.Name
                                                            )
                                                            type_name = "{}_{}".format(
                                                                base_name,
                                                                int(required_length),
                                                            )

                                                            existing_symbols = family_symbol.Family.GetFamilySymbolIds()
                                                            new_symbol = None
                                                            for (
                                                                sym_id
                                                            ) in existing_symbols:
                                                                sym = (
                                                                    self.doc.GetElement(
                                                                        sym_id
                                                                    )
                                                                )
                                                                if (
                                                                    sym.Name
                                                                    == type_name
                                                                ):
                                                                    new_symbol = sym
                                                                    break

                                                            if not new_symbol:
                                                                new_symbol = family_symbol.Duplicate(
                                                                    type_name
                                                                )
                                                                shelf_param_new = new_symbol.LookupParameter(
                                                                    "Длина полки"
                                                                )
                                                                if shelf_param_new:
                                                                    shelf_param_new.Set(
                                                                        required_length
                                                                        / MM_TO_FEET
                                                                    )

                                                            length_to_type[
                                                                required_length
                                                            ] = new_symbol.Id
                                    except:
                                        pass

                            # Second pass: change tag types
                            for view in self.settings.selected_views:
                                collector = FilteredElementCollector(
                                    self.doc, view.Id
                                ).OfClass(IndependentTag)
                                for tag in collector:
                                    try:
                                        if (
                                            tag.Category
                                            and tag.Category.Id.IntegerValue
                                            in [c.value__ for c in tag_categories]
                                        ):
                                            tagged_elements = (
                                                tag.GetTaggedLocalElements()
                                            )
                                            if tagged_elements:
                                                element = tagged_elements[0]
                                                category = element.Category
                                                param_name = self.settings.selected_parameters.get(
                                                    category
                                                )
                                                if param_name:
                                                    param = element.LookupParameter(
                                                        param_name
                                                    )
                                                    if param and param.HasValue:
                                                        value = str(param.AsString())
                                                        char_count = len(value)
                                                        required_length = math.ceil(
                                                            char_count * 1.6 + 1
                                                        )

                                                        tag_type_id = tag.GetTypeId()
                                                        if not tag_type_id:
                                                            continue
                                                        family_symbol = (
                                                            self.doc.GetElement(
                                                                tag_type_id
                                                            )
                                                        )
                                                        if not family_symbol:
                                                            continue
                                                        shelf_param = family_symbol.LookupParameter(
                                                            "Длина полки"
                                                        )
                                                        current_length = 0
                                                        if (
                                                            shelf_param
                                                            and shelf_param.HasValue
                                                        ):
                                                            current_length = round(
                                                                shelf_param.AsDouble()
                                                                * MM_TO_FEET,
                                                                0,
                                                            )

                                                        if (
                                                            current_length
                                                            != required_length
                                                        ):
                                                            type_id = (
                                                                length_to_type.get(
                                                                    required_length
                                                                )
                                                            )
                                                            if type_id:
                                                                self.doc.ChangeElementType(
                                                                    tag,
                                                                    self.doc.GetElement(
                                                                        type_id
                                                                    ),
                                                                )
                                                                new_type_name = (
                                                                    self.doc.GetElement(
                                                                        type_id
                                                                    ).Name
                                                                )
                                                                results.append(
                                                                    "Марка ID {} изменена на тип {}".format(
                                                                        tag.Id,
                                                                        new_type_name,
                                                                    )
                                                                )
                                    except:
                                        pass

                            trans.Commit()
                        except Exception as e:
                            trans.RollBack()
                            print("Ошибка: {}".format(e))
                        finally:
                            trans.Dispose()
            # Show final log
            log_text = """Полный лог обработки:
Всего марок: {}
Найденные IDs категорий марок: {}
Ожидаемые IDs: {}
С категорией: {}
С привязанными элементами: {}
С выбранным параметром: {}
С заполненным значением: {}
С несовпадающей длиной: {}
Изменено: {}""".format(
                total_tags_all,
                ", ".join(str(id) for id in sorted(actual_category_ids)),
                ", ".join(str(c.value__) for c in tag_categories),
                tags_with_category,
                tags_with_elements,
                tags_with_param,
                tags_with_value,
                tags_length_mismatch,
                len(results),
            )
            print(log_text)

            trans.Commit()
            self.txtResults.Text = (
                "\n".join(results) if results else "Ничего не найдено."
            )
            print("Общий результат: {} изменений".format(len(results)))
        except Exception as e:
            trans.RollBack()
            print("Ошибка: " + str(e))
        finally:
            trans.Dispose()

    def GetCategoryName(self, category):
        if not category:
            return "Неизвестная категория"
        try:
            if hasattr(category, "Id") and category.Id.IntegerValue < 0:
                return LabelUtils.GetLabelFor(BuiltInCategory(category.Id.IntegerValue))
        except:
            pass
        return getattr(category, "Name", "Неизвестная категория")

    def OnTabSelecting(self, sender, args):
        args.Cancel = True


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
