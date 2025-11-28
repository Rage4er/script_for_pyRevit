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
