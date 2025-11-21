# -*- coding: utf-8 -*-
__title__ = "Параметры элементов"
__author__ = "Rage"
__doc__ = "Сбор параметров элементов по категориям определяем длину значения параметра"
__ver__ = "0"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import math

from Autodesk.Revit.DB import *
from System.Drawing import *
from System.Windows.Forms import *

MM_TO_FEET = 304.8


class Settings(object):
    def __init__(self):
        self.selected_categories = []
        self.parameters_by_category = {}
        self.selected_parameters = {}
        self.lengths_by_category = {}  # {category: {length: set(ids)}}


class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = Settings()
        self.category_mapping = {}

        self.InitializeComponent()
        self.CollectCategories()

    def InitializeComponent(self):
        self.Text = "Параметры элементов по категориям"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill

        tabs = [
            "1. Категории",
            "2. Выбор параметров",
            "3. Длины текста",
        ]
        for i, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            if i == 0:
                self.SetupTab1(tab)
            elif i == 1:
                self.SetupParametersSelectionTab(tab)
            elif i == 2:
                self.SetupLengthsTab(tab)
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

    def SetupParametersTab(self, tab):
        self.txtParameters = TextBox()
        self.txtParameters.Location = Point(10, 40)
        self.txtParameters.Size = Size(700, 400)
        self.txtParameters.Multiline = True
        self.txtParameters.ScrollBars = ScrollBars.Vertical
        self.txtParameters.ReadOnly = True
        self.txtParameters.Font = Font("Courier New", 10)
        self.txtParameters.WordWrap = True

        controls = [
            self.CreateControl(
                Label,
                Text="Параметры элементов по категориям:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.txtParameters,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)
            ),
            self.CreateControl(
                Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.btnBack1, self.btnNext2 = controls[2], controls[3]
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click

        for c in controls:
            tab.Controls.Add(c)

    def SetupParametersSelectionTab(self, tab):
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
        self.btnBack2, self.btnNext2 = controls[2], controls[3]
        self.btnBack2.Click += self.OnBack2Click
        self.btnNext2.Click += self.OnNext2Click

        for c in controls:
            tab.Controls.Add(c)

    def OnNext1Click(self, sender, args):
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            name = self.lstCategories.Items[i]
            if self.lstCategories.GetItemChecked(i) and name in self.category_mapping:
                self.settings.selected_categories.append(self.category_mapping[name])

        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return

        self.CollectParameters()
        self.PopulateParameterSelection()
        self.tabControl.SelectedIndex = 1

    def OnNext2Click(self, sender, args):
        self.CollectLengths()
        self.DisplayLengths()
        self.tabControl.SelectedIndex = 2

    def OnDoneClick(self, sender, args):
        summary = self.GenerateLengthsSummary()
        MessageBox.Show(summary)
        self.Close()

    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedIndex = 0

    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedIndex = 0

    def OnBack2Click(self, sender, args):
        self.tabControl.SelectedIndex = 1

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
                                params.add(param.Definition.Name)
                sorted_params = sorted(list(params))
                self.settings.parameters_by_category[category] = sorted_params
                self.settings.selected_parameters[category] = (
                    sorted_params[0] if sorted_params else None
                )
        except Exception as e:
            MessageBox.Show("Ошибка при сборе параметров: " + str(e))

    def SetupValuesTab(self, tab):
        self.txtValues = TextBox()
        self.txtValues.Location = Point(10, 40)
        self.txtValues.Size = Size(700, 400)
        self.txtValues.Multiline = True
        self.txtValues.ScrollBars = ScrollBars.Vertical
        self.txtValues.ReadOnly = True
        self.txtValues.Font = Font("Courier New", 10)
        self.txtValues.WordWrap = True

        controls = [
            self.CreateControl(
                Label,
                Text="Уникальные значения параметров по категориям:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.txtValues,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)
            ),
            self.CreateControl(
                Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.btnBack3, self.btnNext4 = controls[2], controls[3]
        self.btnBack3.Click += self.OnBack3Click
        self.btnNext4.Click += self.OnNext4Click

        for c in controls:
            tab.Controls.Add(c)

    def PopulateParameterSelection(self):
        self.lstParameterSelection.Items.Clear()
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            param = self.settings.selected_parameters.get(category, "Не выбран")
            item.SubItems.Add(param)
            item.Tag = category
            self.lstParameterSelection.Items.Add(item)

    def OnParameterSelectionDoubleClick(self, sender, args):
        if self.lstParameterSelection.SelectedItems.Count == 0:
            return
        selected = self.lstParameterSelection.SelectedItems[0]
        category = selected.Tag
        available_params = self.settings.parameters_by_category.get(category, [])
        current_param = self.settings.selected_parameters.get(category)
        form = ParameterSelectionForm(
            self.doc, category, available_params, current_param
        )
        if form.ShowDialog() == DialogResult.OK:
            self.settings.selected_parameters[category] = form.SelectedParameter
            selected.SubItems[1].Text = form.SelectedParameter or "Не выбран"

    def CollectValues(self):
        self.settings.values_by_category = {}
        try:
            for category in self.settings.selected_categories:
                param_name = self.settings.selected_parameters.get(category)
                if not param_name:
                    continue
                collector = FilteredElementCollector(self.doc).OfCategory(
                    BuiltInCategory(category.Id.IntegerValue)
                )
                elements = list(collector)
                unique_values = {}  # value: set of element IDs
                for el in elements:
                    param = el.LookupParameter(param_name)
                    if param and param.HasValue:
                        if param.StorageType == StorageType.String:
                            value = param.AsString()
                        elif param.StorageType == StorageType.Double:
                            value = str(param.AsDouble())
                        elif param.StorageType == StorageType.Integer:
                            value = str(param.AsInteger())
                        else:
                            value = param.AsValueString()
                        if value not in unique_values:
                            unique_values[value] = set()
                        unique_values[value].add(el.Id.IntegerValue)
                self.settings.values_by_category[category] = unique_values
        except Exception as e:
            MessageBox.Show("Ошибка при сборе значений: " + str(e))

    def DisplayValues(self):
        text = ""
        for category, val_dict in self.settings.values_by_category.items():
            cat_name = self.GetCategoryName(category)
            param_name = self.settings.selected_parameters.get(category, "")
            text += "{} ({}):\r\n".format(cat_name, param_name)
            for value, ids in sorted(val_dict.items()):
                ids_str = ", ".join(str(id) for id in sorted(ids))
                text += " - {} (IDs: {})\r\n".format(value, ids_str)
            text += "\r\n"
        self.txtValues.Text = text

    def GenerateParameterSummary(self):
        summary = "Выбранные параметры:\n"
        for category in self.settings.selected_categories:
            param = self.settings.selected_parameters.get(category, "Не выбран")
            cat_name = self.GetCategoryName(category)
            summary += "{}: {}\n".format(cat_name, param)
        return summary

    def GenerateValuesSummary(self):
        summary = "Уникальные значения выбранных параметров:\n"
        for category in self.settings.selected_categories:
            val_dict = self.settings.values_by_category.get(category, {})
            cat_name = self.GetCategoryName(category)
            param_name = self.settings.selected_parameters.get(category, "")
            summary += "{} ({}):\n".format(cat_name, param_name)
            for value, ids in sorted(val_dict.items()):
                ids_str = ", ".join(str(id) for id in sorted(ids))
                summary += " - {} (IDs: {})\n".format(value, ids_str)
            summary += "\n"
        return summary

    def SetupLengthsTab(self, tab):
        self.txtLengths = TextBox()
        self.txtLengths.Location = Point(10, 40)
        self.txtLengths.Size = Size(700, 400)
        self.txtLengths.Multiline = True
        self.txtLengths.ScrollBars = ScrollBars.Vertical
        self.txtLengths.ReadOnly = True
        self.txtLengths.Font = Font("Courier New", 10)
        self.txtLengths.WordWrap = True

        controls = [
            self.CreateControl(
                Label,
                Text="Длины текста по категориям:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.txtLengths,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)
            ),
            self.CreateControl(
                Button, Text="Завершить", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.btnBack2, self.btnDone = controls[2], controls[3]
        self.btnBack2.Click += self.OnBack2Click
        self.btnDone.Click += self.OnDoneClick

        for c in controls:
            tab.Controls.Add(c)

    def CollectLengths(self):
        self.settings.lengths_by_category = {}
        try:
            for category in self.settings.selected_categories:
                param_name = self.settings.selected_parameters.get(category)
                if not param_name:
                    continue
                collector = FilteredElementCollector(self.doc).OfCategory(
                    BuiltInCategory(category.Id.IntegerValue)
                )
                elements = list(collector)
                lengths_dict = {}  # length: set of element IDs
                for el in elements:
                    param = el.LookupParameter(param_name)
                    if param and param.HasValue:
                        if param.StorageType == StorageType.String:
                            value = param.AsString()
                        elif param.StorageType == StorageType.Double:
                            value = str(param.AsDouble())
                        elif param.StorageType == StorageType.Integer:
                            value = str(param.AsInteger())
                        else:
                            value = param.AsValueString()
                        char_count = len(value)
                        length = math.ceil(char_count * 1.6 + 1)
                        if length not in lengths_dict:
                            lengths_dict[length] = set()
                        lengths_dict[length].add(el.Id.IntegerValue)
                self.settings.lengths_by_category[category] = lengths_dict
        except Exception as e:
            MessageBox.Show("Ошибка при расчете длин: " + str(e))

    def DisplayLengths(self):
        text = ""
        for category, lengths_dict in self.settings.lengths_by_category.items():
            cat_name = self.GetCategoryName(category)
            param_name = self.settings.selected_parameters.get(category, "")
            text += "{} ({}):\r\n".format(cat_name, param_name)
            for length, ids in sorted(lengths_dict.items()):
                ids_str = ", ".join(str(id) for id in sorted(ids))
                text += " - {:.0f} мм (IDs: {})\r\n".format(length, ids_str)
            text += "\r\n"
        self.txtLengths.Text = text

    def GenerateLengthsSummary(self):
        summary = "Длины текста по категориям:\n"
        for category in self.settings.selected_categories:
            lengths_dict = self.settings.lengths_by_category.get(category, {})
            cat_name = self.GetCategoryName(category)
            param_name = self.settings.selected_parameters.get(category, "")
            summary += "{} ({}):\n".format(cat_name, param_name)
            for length, ids in sorted(lengths_dict.items()):
                ids_str = ", ".join(str(id) for id in sorted(ids))
                summary += " - {:.0f} мм (IDs: {})\n".format(length, ids_str)
            summary += "\n"
        return summary


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
            MessageBox.Show("Выберите параметр!")

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
        else:
            MessageBox.Show("Нет доступа к документу Revit")
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))


if __name__ == "__main__":
    main()



# -*- coding: utf-8 -*-
__title__ = "Анализ марок"
__author__ = "Rage"
__doc__ = "Анализ марок на выбранных 3D-видах: поиск марок категорий DuctTags, вывод параметров и ID маркированных элементов"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from System import Array, Object
from System.Drawing import *
from System.Windows.Forms import *


class TagSettings(object):
    def __init__(self):
        self.selected_views = []


class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.all_views_dict = {}
        self.views_dict = {}
        self.category_mapping = {}

        self.InitializeComponent()
        self.Load3DViews()

    def InitializeComponent(self):
        self.Text = "Анализ марок на 3D-видах"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = ["1. Выбор видов", "2. Результаты"]
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
        self.txtSearchViews = self.CreateControl(
            TextBox, Location=Point(120, 35), Size=Size(140, 20)
        )
        self.btnSelectAll = self.CreateControl(
            Button, Text="Выбрать все", Location=Point(270, 35), Size=Size(100, 25)
        )
        self.btnSelectAll.Click += self.OnSelectAllViews
        self.btnDeselectAll = self.CreateControl(
            Button, Text="Снять выбор", Location=Point(380, 35), Size=Size(100, 25)
        )
        self.btnDeselectAll.Click += self.OnDeselectAllViews
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
        self.txtSearchViews = controls[2]
        self.btnSelectAll = controls[3]
        self.btnDeselectAll = controls[4]
        self.lstViews = controls[5]
        self.btnNext1 = controls[6]
        self.btnNext1.Click += self.OnNext1Click
        for c in controls:
            tab.Controls.Add(c)
        self.txtSearchViews.TextChanged += self.OnSearchViewsTextChanged

    def SetupTab2(self, tab):
        self.lstResults = ListBox()
        self.lstResults.Location = Point(10, 40)
        self.lstResults.Size = Size(700, 400)
        self.lstResults.HorizontalScrollbar = True

        controls = [
            self.CreateControl(
                Label,
                Text="Результаты анализа марок:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.lstResults,
            self.CreateControl(Button, Text="← Назад", Location=Point(340, 450)),
            self.CreateControl(
                Button,
                Text="Выполнить анализ",
                Location=Point(500, 450),
                Size=Size(150, 30),
            ),
        ]
        self.btnBack2, self.btnExecute = controls[2], controls[3]
        self.btnBack2.Click += self.OnBack2Click
        self.btnExecute.Click += self.OnExecuteClick
        for c in controls:
            tab.Controls.Add(c)

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
                if view.CanBePrinted and not view.IsTemplate:
                    view_name = view.Name
                    if view_name.startswith("{"):
                        continue
                    self.lstViews.Items.Add(view_name, False)
                    self.all_views_dict[view_name] = view
            self.UpdateViewsList("")
        except Exception as e:
            MessageBox.Show("Ошибка при загрузке видов: " + str(e))

    def UpdateViewsList(self, filter_text):
        self.lstViews.Items.Clear()
        self.views_dict.clear()
        for name, view in self.all_views_dict.items():
            if not filter_text or filter_text.lower() in name.lower():
                self.lstViews.Items.Add(
                    name, name in [v.Name for v in self.settings.selected_views]
                )
                self.views_dict[name] = view

    def OnNext1Click(self, sender, args):
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            name = self.lstViews.Items[i]
            if self.lstViews.GetItemChecked(i):
                self.settings.selected_views.append(self.all_views_dict[name])
        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return

        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnToggleViews(self, checked):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, checked)

    def OnSelectAllViews(self, sender, args):
        self.OnToggleViews(True)

    def OnDeselectAllViews(self, sender, args):
        self.OnToggleViews(False)

    def OnSearchViewsTextChanged(self, sender, args):
        filter_text = sender.Text
        self.UpdateViewsList(filter_text)

    def OnNext2Click(self, sender, args):
        self.settings.selected_categories = []
        for i in range(self.lstCategories.Items.Count):
            name = self.lstCategories.Items[i]
            if self.lstCategories.GetItemChecked(i) and name in self.category_mapping:
                self.settings.selected_categories.append(self.category_mapping[name])

        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return

        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    def OnExecuteClick(self, sender, args):
        tag_categories = [
            BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipmentTags,
        ]

        tag_list = []
        try:
            for view in self.settings.selected_views:
                collector = FilteredElementCollector(self.doc, view.Id).OfClass(
                    IndependentTag
                )
                for tag in collector:
                    try:
                        if tag.Category and tag.Category.Id.IntegerValue in [
                            cat.value__ for cat in tag_categories
                        ]:
                            param_list = []
                            for param in tag.Parameters:
                                if param.HasValue and not param.IsReadOnly:
                                    try:
                                        value = param.AsValueString() or str(
                                            param.AsString()
                                        )
                                    except:
                                        value = "Неизвестно"
                                    param_list.append(
                                        "{} = {}".format(param.Definition.Name, value)
                                    )

                            tagged_elements = tag.GetTaggedLocalElements()
                            elem_id = (
                                tagged_elements[0].Id.IntegerValue
                                if tagged_elements and len(tagged_elements) > 0
                                else "None"
                            )
                            if param_list:
                                tag_list.append(
                                    "Вид '{}', Марка ID {}: {}, Элемент ID = {}".format(
                                        view.Name,
                                        tag.Id.IntegerValue,
                                        ", ".join(param_list),
                                        elem_id,
                                    )
                                )
                            else:
                                tag_list.append(
                                    "Вид '{}', Марка ID {}: Нет параметров, Элемент ID = {}".format(
                                        view.Name, tag.Id.IntegerValue, elem_id
                                    )
                                )
                    except:
                        continue

            if not tag_list:
                self.lstResults.Items.Add(
                    "На выбранных видах нет марок указанных категорий."
                )
            else:
                self.lstResults.Items.AddRange(Array[Object](tag_list))
                MessageBox.Show(
                    "Анализ завершен. Найдено {} марок.".format(len(tag_list)),
                    "Результат",
                )
        except Exception as e:
            MessageBox.Show("Ошибка при анализе марок: " + str(e))

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 0
        self.tabControl.Selecting += self.OnTabSelecting

    def OnTabSelecting(self, sender, args):
        args.Cancel = True


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



# -*- coding: utf-8 -*-
__title__ = "Длина полки марок"
__author__ = "Rage"
__doc__ = "Получение длины полки из выбранных марок"

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

        self.InitializeComponent()
        self.CollectCategories()

    def InitializeComponent(self):
        self.Text = "Длина полки марок"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = ["1. Категории", "2. Марки", "3. Результат"]
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

        controls = [
            self.CreateControl(
                Label, Text="Результат:", Location=Point(10, 10), Size=Size(300, 20)
            ),
            self.txtSummary,
            self.CreateControl(
                Button, Text="← Назад", Location=Point(600, 450), Size=Size(80, 25)
            ),
        ]
        self.btnBack2 = controls[2]
        self.btnBack2.Click += self.OnBack2Click
        for c in controls:
            tab.Controls.Add(c)

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
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnNext2Click(self, sender, args):
        self.settings.selected_tag_types = []
        for item in self.lstTagFamilies.Items:
            if item.Checked:
                self.settings.selected_tag_types.append(item.Tag)
        if not self.settings.selected_tag_types:
            MessageBox.Show("Выберите хотя бы один типоразмер марки!")
            return
        self.txtSummary.Text = self.GenerateSummary()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack1Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 0
        self.tabControl.Selecting += self.OnTabSelecting

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

                name_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param and name_param.HasValue:
                    symbol_name = name_param.AsString()
                else:
                    symbol_name = tag_type.Name

                new_name = (
                    symbol_name
                    + "_копия_"
                    + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    + "_"
                    + str(duplicated_count + 1)
                )
                new_symbol = tag_type.Duplicate(new_name)
                duplicated_count += 1
            trans.Commit()
            self.lblStatusTab2.Text = "Дублировано {0} марок.".format(duplicated_count)
        except Exception as e:
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
            self.lblStatusTab2.Text = "Ошибка дублирования: {0}".format(e)
        finally:
            trans.Dispose()

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnTabSelecting(self, sender, args):
        args.Cancel = True

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
                        item.SubItems.Add(self.GetElementName(family))
                        item.SubItems.Add(self.GetElementName(symbol))
                        shelf_length = self.GetShelfLength(symbol)
                        item.SubItems.Add(str(shelf_length))
                        item.Tag = symbol
                        item.Checked = False  # Optionally uncheck
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
                    for symbol_id in symbol_ids:
                        symbol = self.doc.GetElement(symbol_id)
                        if symbol and symbol.IsActive:
                            return family, symbol
                    tag_type = self.doc.GetElement(list(symbol_ids)[0])
                    return family, tag_type

        return None, None

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
                        return family_name + " - " + type_name
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

        except:
            pass

        return "Элемент " + str(element.Id.IntegerValue)

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
                + self.GetElementName(tag_type)
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
