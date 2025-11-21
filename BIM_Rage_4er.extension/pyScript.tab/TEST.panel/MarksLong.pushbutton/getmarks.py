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


class Logger:
    def __init__(self, enabled=False):
        self.messages = []
        self.enabled = enabled

    def add(self, message):
        self.messages.append(message)
        if self.enabled:
            print(message)

    def show(self):
        if not self.messages:
            MessageBox.Show("Нет сообщений логирования.")
            return
        form = Form()
        form.Text = "Логи"
        form.Size = Size(600, 400)
        textbox = TextBox()
        textbox.Multiline = True
        textbox.ReadOnly = True
        textbox.ScrollBars = ScrollBars.Vertical
        textbox.Text = "\n".join(self.messages)
        textbox.Dock = DockStyle.Fill
        form.Controls.Add(textbox)
        form.ShowDialog()


class TagSettings(object):
    def __init__(self):
        self.selected_views = []
        self.selected_categories = []


class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.all_views_dict = {}
        self.views_dict = {}
        self.category_mapping = {}
        self.logger = Logger(enabled=False)

        self.InitializeComponent()
        self.Load3DViews()

    def InitializeComponent(self):
        self.Text = "Анализ марок на 3D-видах"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = ["1. Выбор видов", "2. Категории", "3. Результаты"]
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

    def CreateButton(self, text, location=None, size=None, click_handler=None):
        button = Button()
        button.Text = text
        if location:
            button.Location = location
        if size:
            button.Size = size
        if click_handler:
            button.Click += click_handler
        return button

    def SetupTab1(self, tab):
        self.txtSearchViews = self.CreateControl(
            TextBox, Location=Point(120, 35), Size=Size(140, 20)
        )
        self.btnSelectAll = self.CreateButton(
            text="Выбрать все",
            location=Point(270, 35),
            size=Size(100, 25),
            click_handler=self.OnSelectAllViews,
        )
        self.btnDeselectAll = self.CreateButton(
            text="Снять выбор",
            location=Point(380, 35),
            size=Size(100, 25),
            click_handler=self.OnDeselectAllViews,
        )
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
            self.CreateControl(
                CheckBox,
                Text="Включить логирование",
                Location=Point(10, 440),
                Size=Size(200, 20),
            ),
            self.CreateButton(
                text="Показать логи",
                location=Point(220, 440),
                size=Size(100, 25),
                click_handler=self.OnShowLogsClick,
            ),
            self.CreateButton(
                text="Далее →",
                location=Point(600, 440),
                click_handler=self.OnNext1Click,
            ),
        ]
        (
            self.lblViews,
            self.lblSearch,
            self.txtSearchViews,
            self.btnSelectAll,
            self.btnDeselectAll,
            self.lstViews,
            self.chkLogging,
            self.btnShowLogs,
            self.btnNext1,
        ) = (
            controls[0],
            controls[1],
            controls[2],
            controls[3],
            controls[4],
            controls[5],
            controls[6],
            controls[7],
            controls[8],
        )
        self.chkLogging.CheckedChanged += self.OnLoggingCheckedChanged
        self.txtSearchViews.TextChanged += self.OnSearchViewsTextChanged
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab2(self, tab):
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
            self.CreateButton(
                "← Назад", Point(500, 450), click_handler=self.OnBack1Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 450), click_handler=self.OnNext2Click
            ),
        ]
        self.lblCategories, self.lstCategories, self.btnBack1, self.btnNext2 = controls
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab3(self, tab):
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
            self.CreateButton(
                "← Назад", Point(340, 450), click_handler=self.OnBack2Click
            ),
            self.CreateButton(
                "Выполнить анализ",
                Point(500, 450),
                Size(150, 30),
                click_handler=self.OnExecuteClick,
            ),
        ]
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

    def OnLoggingCheckedChanged(self, sender, args):
        self.logger.enabled = sender.Checked

    def OnShowLogsClick(self, sender, args):
        self.logger.show()

    def OnNext1Click(self, sender, args):
        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            name = self.lstViews.Items[i]
            if self.lstViews.GetItemChecked(i):
                self.settings.selected_views.append(self.all_views_dict[name])
        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return

        self.CollectCategories()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnSelectAllViews(self, sender, args):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, True)

    def OnDeselectAllViews(self, sender, args):
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, False)

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
        except Exception as e:
            self.logger.add("Ошибка получения имени категории: {0}".format(e))
        return getattr(category, "Name", "Неизвестная категория")


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
