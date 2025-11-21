# -*- coding: utf-8 -*-
__title__ = "Анализ марок"
__author__ = "Rage"
__doc__ = "Анализ марок на выбранных 3D-видах: поиск марок категорий DuctTags, вывод параметров и ID маркированных элементов"
__ver__ = "1"

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
