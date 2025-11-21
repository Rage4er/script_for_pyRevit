# -*- coding: utf-8 -*-
__title__ = "Параметры элементов"
__author__ = "Rage"
__doc__ = "Сбор параметров элементов по категориям определяем длину значения параметра"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import datetime
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
