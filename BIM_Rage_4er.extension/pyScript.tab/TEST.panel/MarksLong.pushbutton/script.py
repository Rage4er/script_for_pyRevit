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

        controls = [
            self.CreateControl(
                Label,
                Text="Выберите типоразмеры марок:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.lstTagFamilies,
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

    def OnDuplicateClick(self, sender, args):
        selected_items = []
        for item in self.lstTagFamilies.Items:
            if item.Checked:
                selected_items.append(item)

        if not selected_items:
            self.lblStatusTab2.Text = "Выберите марку для дублирования!"
            return

        selected_item = selected_items[0]
        category = selected_item.Tag
        tag_type = self.settings.category_tag_types.get(category)
        if not tag_type:
            self.lblStatusTab2.Text = "Для этой категории не выбрана марка!"
            return

        if not isinstance(tag_type, FamilySymbol):
            self.lblStatusTab2.Text = "Неверный тип марки: не FamilySymbol"
            return

        trans = Transaction(self.doc, "Дублирование марки")
        try:
            name_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if name_param and name_param.HasValue:
                symbol_name = name_param.AsString()
            else:
                symbol_name = tag_type.Name

            new_name = (
                symbol_name
                + "_копия_"
                + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            )
            trans.Start()
            new_symbol = tag_type.Duplicate(new_name)
            trans.Commit()
            self.lblStatusTab2.Text = "Марка '{0}' дублирована как '{1}'.".format(
                symbol_name, new_name
            )
        except Exception as e:
            if trans.GetStatus() == TransactionStatus.Started:
                trans.RollBack()
            self.lblStatusTab2.Text = "Ошибка дублирования: {0}".format(e)
        finally:
            trans.Dispose()

    def OnDuplicateClick(self, sender, args):
        if self.lstTagFamilies.SelectedItems.Count == 0:
            MessageBox.Show("Выберите марку для дублирования!")
            return
        selected = self.lstTagFamilies.SelectedItems[0]
        category = selected.Tag
        tag_type = self.settings.category_tag_types.get(category)
        if not tag_type:
            MessageBox.Show("Для этой категории не выбрана марка!")
            return

        # Дублировать типоразмер
        trans = Transaction(self.doc, "Дублирование марки")
        trans.Start()
        try:
            new_name = tag_type.Name + "_копия"
            new_symbol = tag_type.Duplicate(new_name)
            trans.Commit()
            MessageBox.Show(
                "Марка '{0}' дублирована как '{1}'.".format(tag_type.Name, new_name)
            )
        except Exception as e:
            trans.RollBack()
            MessageBox.Show("Ошибка дублирования: {0}".format(e))
        finally:
            trans.Dispose()

    def OnTagFamilyDoubleClick(self, sender, args):
        if self.lstTagFamilies.SelectedItems.Count == 0:
            return
        selected = self.lstTagFamilies.SelectedItems[0]
        category = selected.Tag

        available_families = self.GetAvailableTagFamiliesForCategory(category)

        if available_families:
            current_family = self.settings.category_tag_families.get(category)
            current_type = self.settings.category_tag_types.get(category)

            form = TagFamilySelectionForm(
                self.doc, available_families, current_family, current_type
            )
            if (
                form.ShowDialog() == DialogResult.OK
                and form.SelectedFamily
                and form.SelectedType
            ):
                selected.SubItems[1].Text = self.GetElementName(form.SelectedFamily)
                selected.SubItems[2].Text = self.GetElementName(form.SelectedType)
                self.settings.category_tag_families[category] = form.SelectedFamily
                self.settings.category_tag_types[category] = form.SelectedType

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


class TagFamilySelectionForm(Form):
    def __init__(self, doc, available_families, current_family, current_type):
        self.doc = doc
        self.available_families = available_families
        self.selected_family = None
        self.selected_type = None
        self.family_dict = {}
        self.type_dict = {}

        self.InitializeComponent()
        self.PopulateFamilies()

    def InitializeComponent(self):
        self.Text = "Выбор семейства и типоразмера марки"
        self.Size = Size(800, 500)
        self.StartPosition = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        controls = [
            Label(
                Text="Выберите семейство марки:",
                Location=Point(10, 10),
                Size=Size(250, 20),
            ),
            ListBox(Location=Point(10, 40), Size=Size(250, 350)),
            Label(
                Text="Выберите типоразмер:", Location=Point(270, 10), Size=Size(250, 20)
            ),
            ListBox(Location=Point(270, 40), Size=Size(500, 350)),
            Button(Text="OK", Location=Point(400, 400), Size=Size(75, 25)),
            Button(Text="Отмена", Location=Point(485, 400), Size=Size(75, 25)),
        ]

        (
            self.lblFamilies,
            self.lstFamilies,
            self.lblTypes,
            self.lstTypes,
            self.btnOK,
            self.btnCancel,
        ) = controls

        self.lstFamilies.SelectedIndexChanged += self.OnFamilySelected
        self.btnOK.Click += self.OnOKClick
        self.btnCancel.Click += self.OnCancelClick

        for c in controls:
            self.Controls.Add(c)

    def PopulateFamilies(self):
        self.lstFamilies.Items.Clear()
        self.family_dict.clear()

        for family in self.available_families:
            if family:
                family_name = self.GetElementName(family)
                self.lstFamilies.Items.Add(family_name)
                self.family_dict[family_name] = family

        if self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0

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

    def OnFamilySelected(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0:
            selected_name = self.lstFamilies.SelectedItem
            selected_family = self.family_dict.get(selected_name)
            if selected_family:
                self.PopulateTypesList(selected_family)

    def PopulateTypesList(self, family):
        self.lstTypes.Items.Clear()
        self.type_dict.clear()

        try:
            symbol_ids = family.GetFamilySymbolIds()
            if symbol_ids and symbol_ids.Count > 0:
                for symbol_id in list(symbol_ids):
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = self.GetElementName(symbol)
                        status = " (активный)" if symbol.IsActive else " (не активный)"
                        display_name = (
                            symbol_name
                            + status
                            + " [ID:"
                            + str(symbol.Id.IntegerValue)
                            + "]"
                        )

                        self.lstTypes.Items.Add(display_name)
                        self.type_dict[display_name] = symbol

                for i in range(self.lstTypes.Items.Count):
                    display_name = self.lstTypes.Items[i]
                    symbol = self.type_dict.get(display_name)
                    if symbol and symbol.IsActive:
                        self.lstTypes.SelectedIndex = i
                        return

                if self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
            else:
                MessageBox.Show("В выбранном семействе нет типоразмеров")

        except:
            MessageBox.Show("Ошибка при загрузке типоразмеров")

    def OnOKClick(self, sender, args):
        if self.lstFamilies.SelectedIndex >= 0 and self.lstTypes.SelectedIndex >= 0:
            selected_family_name = self.lstFamilies.SelectedItem
            selected_type_name = self.lstTypes.SelectedItem

            self.selected_family = self.family_dict.get(selected_family_name)
            self.selected_type = self.type_dict.get(selected_type_name)

            if self.selected_family and self.selected_type:
                self.DialogResult = DialogResult.OK
                self.Close()
            else:
                MessageBox.Show("Ошибка при получении выбранных объектов!")
        else:
            MessageBox.Show("Выберите семейство и типоразмер марки!")

    def OnCancelClick(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    @property
    def SelectedFamily(self):
        return self.selected_family

    @property
    def SelectedType(self):
        return self.selected_type


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
