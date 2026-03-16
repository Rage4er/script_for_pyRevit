# -*- coding: utf-8 -*-
__title__ = """Марки
 на видах"""
__author__ = "Rage"
__doc__ = "Автоматическое размещение марок на 3D-видах и планах"
__version__ = "1.1"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import codecs
import datetime
import json
import os
import random

from Autodesk.Revit.DB import *
from System.Drawing import *
from System.Windows.Forms import *

# Константы
MM_TO_FEET = 304.8
DEFAULT_OFFSET_X = 60.0
DEFAULT_OFFSET_Y = 30.0
VIEW_TYPES = ["3D виды", "Планы этажей"]  # Доступные типы видов


# Логирование
class Logger:
    """
    Класс для управления логированием.

    Attributes:
        messages (list): Список сообщений лога.
        enabled (bool): Флаг включения логирования.
    """

    def __init__(self, enabled=False):
        """
        Инициализирует логгер.

        Args:
            enabled (bool): Включить логирование.
        """
        self.messages = []
        self.enabled = enabled

    def add(self, message):
        """
        Добавляет сообщение в лог.

        Args:
            message (str): Сообщение для логирования.
        """
        if self.enabled:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            self.messages.append("[{0}] {1}".format(timestamp, message))

    def show(self):
        """
        Показывает лог в окне сообщения или форме.
        """
        if not self.enabled:
            MessageBox.Show("Логирование отключено", "Информация")
            return
        if not self.messages:
            MessageBox.Show("Нет записей в логе", "Информация")
            return
        log_text = "\n".join(self.messages)
        form = Form()
        form.Text = "Логи выполнения"
        form.Size = Size(600, 400)
        textbox = TextBox()
        textbox.Multiline = True
        textbox.Dock = DockStyle.Fill
        textbox.ScrollBars = ScrollBars.Vertical
        textbox.Text = log_text
        form.Controls.Add(textbox)
        form.ShowDialog()


# Настройки
class TagSettings(object):
    """
    Класс для хранения настроек расстановки марок.

    Attributes:
        selected_view_types (list): Выбранные типы видов (3D, Планы).
        selected_views (list): Выбранные виды.
        selected_categories (list): Выбранные категории элементов.
        category_tag_families_3d (dict): Семейства марок по категориям для 3D.
        category_tag_types_3d (dict): Типоразмеры марок по категориям для 3D.
        category_tag_families_plan (dict): Семейства марок по категориям для планов.
        category_tag_types_plan (dict): Типоразмеры марок по категориям для планов.
        offset_x (float): Смещение по X в мм.
        offset_y (float): Смещение по Y в мм.
        orientation (TagOrientation): Ориентация марки.
        use_leader (bool): Использовать выноску.
        enable_logging (bool): Включить логирование.
    """

    def __init__(self):
        """
        Инициализирует настройки с значениями по умолчанию.
        """
        self.selected_view_types = []
        self.selected_views = []
        self.selected_categories = []
        # Раздельные настройки для 3D и планов
        self.category_tag_families_3d = {}
        self.category_tag_types_3d = {}
        self.category_tag_families_plan = {}
        self.category_tag_types_plan = {}
        self.offset_x = DEFAULT_OFFSET_X
        self.offset_y = DEFAULT_OFFSET_Y
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True
        self.enable_logging = False
        self.random_offset = True  # Случайное направление смещения


# Главная форма
class MainForm(Form):
    """
    Основная форма приложения для расстановки марок на 3D-видах.

    Attributes:
        doc (Document): Документ Revit.
        uidoc (UIDocument): UI-документ Revit.
        settings (TagSettings): Настройки приложения.
        views_dict (dict): Словарь видов.
        category_mapping (dict): Маппинг категорий.
        logger (Logger): Экземпляр логгера.
        tag_defaults (dict): Словарь дефолтных марок по категориям.
    """

    def __init__(self, doc, uidoc):
        """
        Инициализирует главную форму.

        Args:
            doc (Document): Активный документ Revit.
            uidoc (UIDocument): UI-документ Revit.
        """
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}
        self.all_views_dict = {}
        self.lstViewsChecked = {}
        self.category_mapping = {}
        self.logger = Logger(self.settings.enable_logging)
        self.tag_defaults = self.LoadTagDefaults()

        self.InitializeComponent()
        self.LoadAllViews()

    def InitializeComponent(self):
        """
        Инициализирует компоненты формы.
        """
        self.Text = "Расстановка марок на видах (3D и планы)"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        self.tabControl.Selecting += self.OnTabSelecting

        tabs = [
            "1. Выбор видов",
            "2. Категории",
            "3. Марки",
            "4. Настройки",
            "5. Выполнение",
        ]
        for i, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            getattr(self, "SetupTab" + str(i + 1))(tab)
            self.tabControl.TabPages.Add(tab)

        self.Controls.Add(self.tabControl)

    def CreateControl(self, control_type, **kwargs):
        """
        Создает контроль с заданными свойствами.

        Args:
            control_type: Тип контроля (например, Label).
            **kwargs: Свойства и их значения.

        Returns:
            Control: Созданный контроль.
        """
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control

    def CreateButton(self, text, location, size=None, click_handler=None):
        """
        Создает кнопку с заданными параметрами.

        Args:
            text (str): Текст кнопки.
            location (Point): Позиция.
            size (Size, optional): Размер. По умолчанию (80, 25).
            click_handler: Обработчик клика.

        Returns:
            Button: Созданная кнопка.
        """
        size = size or Size(80, 25)
        btn = self.CreateControl(Button, Text=text, Location=location, Size=size)
        if click_handler:
            btn.Click += click_handler
        return btn

    def SetupTab1(self, tab):
        """
        Настраивает вкладку 1: Выбор видов.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
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
        """
        Настраивает вкладку 2: Категории (раздельно для 3D и планов).

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        # Заголовки
        lblTitle3D = self.CreateControl(
            Label,
            Text="3D виды - категории:",
            Location=Point(10, 10),
            Size=Size(200, 20),
            Font=Font(self.Font.FontFamily, self.Font.Size, FontStyle.Bold),
            ForeColor=Color.DarkBlue
        )
        tab.Controls.Add(lblTitle3D)

        lblTitlePlan = self.CreateControl(
            Label,
            Text="Планы этажей - категории:",
            Location=Point(10, 240),
            Size=Size(200, 20),
            Font=Font(self.Font.FontFamily, self.Font.Size, FontStyle.Bold),
            ForeColor=Color.DarkGreen
        )
        tab.Controls.Add(lblTitlePlan)

        # Список категорий для 3D
        self.lstCategories3D = CheckedListBox()
        self.lstCategories3D.Location = Point(10, 35)
        self.lstCategories3D.Size = Size(300, 193)
        self.lstCategories3D.CheckOnClick = True
        tab.Controls.Add(self.lstCategories3D)

        # Список категорий для планов
        self.lstCategoriesPlan = CheckedListBox()
        self.lstCategoriesPlan.Location = Point(10, 265)
        self.lstCategoriesPlan.Size = Size(300, 193)
        self.lstCategoriesPlan.CheckOnClick = True
        tab.Controls.Add(self.lstCategoriesPlan)

        # Кнопки выбора
        self.btnSelectAllCats = Button()
        self.btnSelectAllCats.Text = "Выбрать все"
        self.btnSelectAllCats.Location = Point(320, 35)
        self.btnSelectAllCats.Size = Size(120, 25)
        self.btnSelectAllCats.Click += self.OnSelectAllCategoriesClick
        tab.Controls.Add(self.btnSelectAllCats)

        self.btnDeselectAllCats = Button()
        self.btnDeselectAllCats.Text = "Снять все"
        self.btnDeselectAllCats.Location = Point(320, 65)
        self.btnDeselectAllCats.Size = Size(120, 25)
        self.btnDeselectAllCats.Click += self.OnDeselectAllCategoriesClick
        tab.Controls.Add(self.btnDeselectAllCats)

        # Подсказка
        lblHint = self.CreateControl(
            Label,
            Text="Выберите категории отдельно для 3D видов и планов",
            Location=Point(320, 100),
            Size=Size(250, 40),
            ForeColor=Color.DarkSlateGray
        )
        tab.Controls.Add(lblHint)

        # Кнопки навигации
        controls = [
            self.CreateButton(
                "← Назад", Point(500, 455), click_handler=self.OnBack1Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 455), click_handler=self.OnNext2Click
            ),
        ]
        for c in controls:
            tab.Controls.Add(c)

    def OnSelectAllCategoriesClick(self, sender, args):
        """Выбрать все категории в обоих списках"""
        for i in range(self.lstCategories3D.Items.Count):
            self.lstCategories3D.SetItemChecked(i, True)
        for i in range(self.lstCategoriesPlan.Items.Count):
            self.lstCategoriesPlan.SetItemChecked(i, True)

    def OnDeselectAllCategoriesClick(self, sender, args):
        """Снять все категории в обоих списках"""
        for i in range(self.lstCategories3D.Items.Count):
            self.lstCategories3D.SetItemChecked(i, False)
        for i in range(self.lstCategoriesPlan.Items.Count):
            self.lstCategoriesPlan.SetItemChecked(i, False)

    def SetupTab3(self, tab):
        """
        Настраивает вкладку 3: Марки (две отдельные таблицы для 3D и планов).

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        # Заголовки
        lblTitle3D = self.CreateControl(
            Label,
            Text="3D виды - марки:",
            Location=Point(10, 10),
            Size=Size(200, 20),
            Font=Font(self.Font.FontFamily, self.Font.Size, FontStyle.Bold),
            ForeColor=Color.DarkBlue
        )
        tab.Controls.Add(lblTitle3D)

        lblTitlePlan = self.CreateControl(
            Label,
            Text="Планы этажей - марки:",
            Location=Point(10, 215),
            Size=Size(200, 20),
            Font=Font(self.Font.FontFamily, self.Font.Size, FontStyle.Bold),
            ForeColor=Color.DarkGreen
        )
        tab.Controls.Add(lblTitlePlan)

        # Список марок для 3D видов
        self.lstTagFamilies3D = ListView()
        self.lstTagFamilies3D.Location = Point(10, 35)
        self.lstTagFamilies3D.Size = Size(700, 165)
        self.lstTagFamilies3D.View = View.Details
        self.lstTagFamilies3D.FullRowSelect = True
        self.lstTagFamilies3D.GridLines = True
        self.lstTagFamilies3D.Columns.Add("Категория", 180)
        self.lstTagFamilies3D.Columns.Add("Семейство марки", 200)
        self.lstTagFamilies3D.Columns.Add("Типоразмер марки", 300)
        self.lstTagFamilies3D.DoubleClick += self.OnTagFamilyDoubleClick3D
        tab.Controls.Add(self.lstTagFamilies3D)

        # Список марок для планов
        self.lstTagFamiliesPlan = ListView()
        self.lstTagFamiliesPlan.Location = Point(10, 240)
        self.lstTagFamiliesPlan.Size = Size(700, 165)
        self.lstTagFamiliesPlan.View = View.Details
        self.lstTagFamiliesPlan.FullRowSelect = True
        self.lstTagFamiliesPlan.GridLines = True
        self.lstTagFamiliesPlan.Columns.Add("Категория", 180)
        self.lstTagFamiliesPlan.Columns.Add("Семейство марки", 200)
        self.lstTagFamiliesPlan.Columns.Add("Типоразмер марки", 300)
        self.lstTagFamiliesPlan.DoubleClick += self.OnTagFamilyDoubleClickPlan
        tab.Controls.Add(self.lstTagFamiliesPlan)

        # Подсказка
        lblHint = self.CreateControl(
            Label,
            Text="Настройте марки отдельно для 3D видов и планов этажей",
            Location=Point(10, 415),
            Size=Size(500, 20),
            ForeColor=Color.DarkSlateGray
        )
        tab.Controls.Add(lblHint)

        # Кнопки навигации
        controls = [
            self.CreateButton(
                "← Назад", Point(500, 455), click_handler=self.OnBack2Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 455), click_handler=self.OnNext3Click
            ),
        ]
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab4(self, tab):
        """
        Настраивает вкладку 4: Настройки.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        self.cmbOrientation = ComboBox()
        self.cmbOrientation.Location = Point(170, 50)
        self.cmbOrientation.Size = Size(150, 20)
        self.cmbOrientation.Items.Add("Горизонтальная")
        self.cmbOrientation.Items.Add("Вертикальная")
        self.cmbOrientation.SelectedIndex = 0

        self.chkUseLeader = CheckBox()
        self.chkUseLeader.Text = "Использовать выноску"
        self.chkUseLeader.Location = Point(10, 80)
        self.chkUseLeader.Size = Size(200, 20)
        self.chkUseLeader.Checked = True

        self.chkRandomOffset = CheckBox()
        self.chkRandomOffset.Text = "Случайное направление смещения"
        self.chkRandomOffset.Location = Point(10, 105)
        self.chkRandomOffset.Size = Size(250, 20)
        self.chkRandomOffset.Checked = True

        controls = [
            self.CreateControl(
                Label,
                Text="Настройки размещения:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.CreateControl(
                Label, Text="Ориентация:", Location=Point(10, 50), Size=Size(150, 20)
            ),
            self.cmbOrientation,
            self.chkUseLeader,
            self.chkRandomOffset,
            self.CreateButton(
                "← Назад", Point(500, 455), click_handler=self.OnBack3Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 455), click_handler=self.OnNext4Click
            ),
        ]
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab5(self, tab):
        """
        Настраивает вкладку 5: Выполнение.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        self.txtSummary = TextBox()
        self.txtSummary.Location = Point(10, 40)
        self.txtSummary.Size = Size(700, 400)
        self.txtSummary.Multiline = True
        self.txtSummary.ScrollBars = ScrollBars.Vertical
        self.txtSummary.ReadOnly = True

        self.progressBar = ProgressBar()
        self.progressBar.Location = Point(10, 450)
        self.progressBar.Size = Size(320, 20)
        self.progressBar.Minimum = 0
        self.progressBar.Maximum = 100

        controls = [
            self.CreateControl(
                Label,
                Text="Готово к выполнению:",
                Location=Point(10, 10),
                Size=Size(300, 20),
            ),
            self.txtSummary,
            self.progressBar,
            self.CreateButton(
                "← Назад", Point(340, 450), click_handler=self.OnBack4Click
            ),
            self.CreateButton(
                "Выполнить",
                Point(500, 450),
                Size(150, 30),
                click_handler=self.OnExecuteClick,
            ),
        ]
        for c in controls:
            tab.Controls.Add(c)

    def LoadAllViews(self):
        """
        Загружает список всех видов (3D и планы) в список.
        """
        try:
            self.lstViews.Items.Clear()
            self.all_views_dict.clear()
            self.views_dict.clear()
            self.lstViewsChecked.clear()

            # Собираем 3D виды
            views_3d = (
                FilteredElementCollector(self.doc)
                .OfClass(View3D)
                .WhereElementIsNotElementType()
                .ToElements()
            )

            # Собираем планы этажей
            views_plan = (
                FilteredElementCollector(self.doc)
                .OfClass(ViewPlan)
                .WhereElementIsNotElementType()
                .ToElements()
            )

            all_views = list(views_3d) + list(views_plan)

            # Сортируем виды по имени
            all_views_sorted = sorted(all_views, key=lambda v: v.Name or "")

            for view in all_views_sorted:
                if not view.IsTemplate and view.CanBePrinted:
                    # Определяем тип вида
                    view_type = "3D" if isinstance(view, View3D) else "План"
                    name = "{0} [{1}] (ID: {2})".format(
                        view.Name,
                        view_type,
                        str(view.Id.IntegerValue)
                    )
                    self.all_views_dict[name] = view
                    self.lstViewsChecked[name] = False

            self.UpdateViewsList("")  # Показать все
            if self.lstViews.Items.Count > 0:
                self.lstViews.SetItemChecked(0, True)
                self.lstViewsChecked[self.lstViews.Items[0]] = True

        except Exception as e:
            MessageBox.Show("Ошибка загрузки видов: " + str(e))

    def UpdateViewsList(self, filter_text):
        """
        Обновляет список видов с фильтром и сортировкой по имени.

        Args:
            filter_text (str): Текст фильтра.
        """
        self.lstViews.Items.Clear()
        self.views_dict.clear()
        filter_lower = filter_text.lower()
        
        # Сортируем виды по имени перед добавлением
        sorted_views = sorted(
            self.all_views_dict.items(),
            key=lambda x: x[0].lower()  # Сортировка по имени вида (ключ словаря)
        )
        
        for name, view in sorted_views:
            if filter_lower in name.lower():
                self.lstViews.Items.Add(name, self.lstViewsChecked.get(name, False))
                self.views_dict[name] = view

    # Навигация
    def OnNext1Click(self, sender, args):
        self.settings.enable_logging = self.chkLogging.Checked
        self.logger.enabled = self.settings.enable_logging

        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            name = self.lstViews.Items[i]
            checked = self.lstViews.GetItemChecked(i)
            if checked and name in self.views_dict:
                self.settings.selected_views.append(self.views_dict[name])
            self.lstViewsChecked[name] = checked

        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return

        self.CollectCategories()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnSelectAllViews(self, sender, args):
        """
        Обработчик кнопки "Выбрать все".
        """
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, True)
            name = self.lstViews.Items[i]
            self.lstViewsChecked[name] = True

    def OnDeselectAllViews(self, sender, args):
        """
        Обработчик кнопки "Снять выбор".
        """
        for i in range(self.lstViews.Items.Count):
            self.lstViews.SetItemChecked(i, False)
            name = self.lstViews.Items[i]
            self.lstViewsChecked[name] = False

    def OnSearchViewsTextChanged(self, sender, args):
        """
        Обработчик изменения текста поиска видов.
        """
        filter_text = sender.Text
        self.UpdateViewsList(filter_text)

    def OnLoggingCheckedChanged(self, sender, args):
        """
        Обработчик изменения состояния CheckBox логирования.
        """
        self.logger.enabled = sender.Checked

    def OnNext2Click(self, sender, args):
        # Собираем категории из обоих списков
        self.settings.selected_categories = []
        
        # Собираем из списка 3D
        for i in range(self.lstCategories3D.Items.Count):
            name = self.lstCategories3D.Items[i]
            if self.lstCategories3D.GetItemChecked(i) and name in self.category_mapping:
                cat = self.category_mapping[name]
                if cat not in self.settings.selected_categories:
                    self.settings.selected_categories.append(cat)
        
        # Собираем из списка планов
        for i in range(self.lstCategoriesPlan.Items.Count):
            name = self.lstCategoriesPlan.Items[i]
            if self.lstCategoriesPlan.GetItemChecked(i) and name in self.category_mapping:
                cat = self.category_mapping[name]
                if cat not in self.settings.selected_categories:
                    self.settings.selected_categories.append(cat)

        if not self.settings.selected_categories:
            MessageBox.Show("Выберите хотя бы одну категорию!")
            return

        # Заполняем обе таблицы
        self.PopulateTagFamilies3D()
        self.PopulateTagFamiliesPlan()

        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    def OnNext3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 3
        self.tabControl.Selecting += self.OnTabSelecting

    def OnNext4Click(self, sender, args):
        self.settings.orientation = (
            TagOrientation.Horizontal
            if self.cmbOrientation.SelectedIndex == 0
            else TagOrientation.Vertical
        )
        self.settings.use_leader = self.chkUseLeader.Checked
        self.settings.random_offset = self.chkRandomOffset.Checked
        self.txtSummary.Text = self.GenerateSummary()
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 4
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack1Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 0
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack2Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 1
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack3Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 2
        self.tabControl.Selecting += self.OnTabSelecting

    def OnBack4Click(self, sender, args):
        self.tabControl.Selecting -= self.OnTabSelecting
        self.tabControl.SelectedIndex = 3
        self.tabControl.Selecting += self.OnTabSelecting

    def OnTabSelecting(self, sender, args):
        """
        Обработчик выбора вкладки.
        """
        args.Cancel = True

    def CollectCategories(self):
        """
        Собирает категории элементов для маркировки и заполняет оба списка (3D и Планы).
        """
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
            except Exception as e:
                self.logger.add("Ошибка получения категории {0}: {1}".format(cat, e))

        self.settings.selected_categories = list(unique_cats)
        
        # Заполняем оба списка одинаковыми категориями
        self.lstCategories3D.Items.Clear()
        self.lstCategoriesPlan.Items.Clear()
        self.category_mapping.clear()

        for cat in sorted(
            self.settings.selected_categories, key=lambda x: self.GetCategoryName(x)
        ):
            name = self.GetCategoryName(cat)
            # Добавляем в оба списка с одинаковым состоянием
            self.lstCategories3D.Items.Add(name, True)
            self.lstCategoriesPlan.Items.Add(name, True)
            self.category_mapping[name] = cat

    def GetDuctsToTag(self, duct_elements):
        """
        Отбирает воздуховоды для маркировки, оставляя только самый длинный в каждой группе.
        Группировка по "Имя системы" -> "Сечение".

        Args:
            duct_elements (list): Список элементов воздуховодов.

        Returns:
            list: Список элементов воздуховодов для маркировки.
        """
        # Словарь для группировки: {system_name: {section: [ducts]}}
        groups = {}

        # Группируем воздуховоды
        for duct in duct_elements:
            try:
                # Получаем параметр "Имя системы"
                system_name_param = duct.LookupParameter("Имя системы")
                if not system_name_param:
                    # Пытаемся найти параметр по BuiltInParameter
                    system_name_param = duct.get_Parameter(
                        BuiltInParameter.RBS_SYSTEM_NAME_PARAM
                    )

                system_name = ""
                if system_name_param and system_name_param.HasValue:
                    if system_name_param.StorageType == StorageType.String:
                        system_name = system_name_param.AsString()
                    else:
                        system_name = system_name_param.AsValueString()

                # Получаем параметр "Сечение"
                section_param = duct.LookupParameter("Сечение")
                if not section_param:
                    # Для воздуховодов сечение можно получить из размеров
                    section = self.GetDuctSection(duct)
                else:
                    section = ""
                    if section_param.HasValue:
                        if section_param.StorageType == StorageType.String:
                            section = section_param.AsString()
                        else:
                            section = section_param.AsValueString()

                # Получаем длину воздуховода
                length_param = duct.LookupParameter("Длина")
                if not length_param:
                    length_param = duct.get_Parameter(
                        BuiltInParameter.CURVE_ELEM_LENGTH
                    )

                length = 0.0
                if length_param and length_param.HasValue:
                    if length_param.StorageType == StorageType.Double:
                        length = length_param.AsDouble()
                    else:
                        length = float(length_param.AsValueString().replace(",", "."))

                # Создаем ключи для группировки
                if system_name not in groups:
                    groups[system_name] = {}

                if section not in groups[system_name]:
                    groups[system_name][section] = []

                # Добавляем воздуховод в группу
                groups[system_name][section].append({"element": duct, "length": length})

            except Exception as e:
                self.logger.add(
                    "Ошибка обработки воздуховода {0}: {1}".format(duct.Id, e)
                )
                continue

        # Отбираем самые длинные воздуховоды в каждой группе
        selected_ducts = []
        for system_name, sections in groups.items():
            for section, ducts in sections.items():
                if ducts:
                    # Находим воздуховод с максимальной длиной
                    longest_duct = max(ducts, key=lambda x: x["length"])
                    selected_ducts.append(longest_duct["element"])
                    self.logger.add(
                        "Выбран воздуховод ID {0} для системы '{1}', сечения '{2}', длина {3:.2f}".format(
                            longest_duct["element"].Id,
                            system_name,
                            section,
                            longest_duct["length"],
                        )
                    )

        return selected_ducts

    def GetDuctSection(self, duct):
        """
        Получает сечение воздуховода.

        Args:
            duct (Element): Элемент воздуховода.

        Returns:
            str: Строка с описанием сечения.
        """
        try:
            # Для круглых воздуховодов
            diameter_param = duct.LookupParameter("Диаметр")
            if not diameter_param:
                diameter_param = duct.get_Parameter(
                    BuiltInParameter.RBS_CURVE_DIAMETER_PARAM
                )

            if diameter_param and diameter_param.HasValue:
                if diameter_param.StorageType == StorageType.Double:
                    diameter = diameter_param.AsDouble() * 304.8  # в мм
                    return "Ø{:.0f}".format(diameter)

            # Для прямоугольных воздуховодов
            width_param = duct.LookupParameter("Ширина")
            if not width_param:
                width_param = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)

            height_param = duct.LookupParameter("Высота")
            if not height_param:
                height_param = duct.get_Parameter(
                    BuiltInParameter.RBS_CURVE_HEIGHT_PARAM
                )

            if (
                width_param
                and height_param
                and width_param.HasValue
                and height_param.HasValue
            ):
                width = 0.0
                height = 0.0

                if width_param.StorageType == StorageType.Double:
                    width = width_param.AsDouble() * 304.8  # в мм
                else:
                    width = float(width_param.AsValueString().replace(",", "."))

                if height_param.StorageType == StorageType.Double:
                    height = height_param.AsDouble() * 304.8  # в мм
                else:
                    height = float(height_param.AsValueString().replace(",", "."))

                return "{:.0f}x{:.0f}".format(width, height)

            return "Не определено"
        except Exception as e:
            self.logger.add(
                "Ошибка получения сечения воздуховода {0}: {1}".format(duct.Id, e)
            )
            return "Ошибка"

    def GetCategoryName(self, category):
        if not category:
            return "Неизвестная категория"
        try:
            if hasattr(category, "Id") and category.Id.IntegerValue < 0:
                return LabelUtils.GetLabelFor(BuiltInCategory(category.Id.IntegerValue))
        except Exception as e:
            self.logger.add("Ошибка получения имени категории: {0}".format(e))
        return getattr(category, "Name", "Неизвестная категория")

    def PopulateTagFamilies3D(self):
        """Заполняет список марок для 3D видов"""
        self.lstTagFamilies3D.Items.Clear()
        
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            item.Tag = category

            cat_name = self.GetCategoryName(category)
            family_name, type_name = self.GetSavedTagForCurrentView(cat_name, "3D")

            if family_name and type_name:
                tag_family, tag_type = self.FindSavedTag(family_name, type_name)
                if tag_family and tag_type:
                    item.SubItems.Add(self.GetElementName(tag_family))
                    item.SubItems.Add(self.GetElementName(tag_type))
                    self.settings.category_tag_families_3d[category] = tag_family
                    self.settings.category_tag_types_3d[category] = tag_type
                    self.lstTagFamilies3D.Items.Add(item)
                    continue

            tag_family, tag_type = self.FindTagForCategory(category)
            if tag_family and tag_type:
                item.SubItems.Add(self.GetElementName(tag_family))
                item.SubItems.Add(self.GetElementName(tag_type))
                self.settings.category_tag_families_3d[category] = tag_family
                self.settings.category_tag_types_3d[category] = tag_type
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")

            self.lstTagFamilies3D.Items.Add(item)

    def PopulateTagFamiliesPlan(self):
        """Заполняет список марок для планов этажей"""
        self.lstTagFamiliesPlan.Items.Clear()
        
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            item.Tag = category

            cat_name = self.GetCategoryName(category)
            family_name, type_name = self.GetSavedTagForCurrentView(cat_name, "План")

            if family_name and type_name:
                tag_family, tag_type = self.FindSavedTag(family_name, type_name)
                if tag_family and tag_type:
                    item.SubItems.Add(self.GetElementName(tag_family))
                    item.SubItems.Add(self.GetElementName(tag_type))
                    self.settings.category_tag_families_plan[category] = tag_family
                    self.settings.category_tag_types_plan[category] = tag_type
                    self.lstTagFamiliesPlan.Items.Add(item)
                    continue

            tag_family, tag_type = self.FindTagForCategory(category)
            if tag_family and tag_type:
                item.SubItems.Add(self.GetElementName(tag_family))
                item.SubItems.Add(self.GetElementName(tag_type))
                self.settings.category_tag_families_plan[category] = tag_family
                self.settings.category_tag_types_plan[category] = tag_type
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")

            self.lstTagFamiliesPlan.Items.Add(item)

    def GetSavedTagForCurrentView(self, cat_name, view_type):
        """Получает сохранённую марку для указанного типа вида"""
        defaults = self.tag_defaults.get(cat_name, {})
        if view_type == "3D":
            family_name = defaults.get("family_3d")
            type_name = defaults.get("type_3d")
        else:
            family_name = defaults.get("family_plan")
            type_name = defaults.get("type_plan")
        return family_name, type_name

    def SaveTagForCurrentView(self, cat_name, view_type, family_name, type_name):
        """Сохраняет марку для указанного типа вида"""
        if cat_name not in self.tag_defaults:
            self.tag_defaults[cat_name] = {}
        if view_type == "3D":
            self.tag_defaults[cat_name]["family_3d"] = family_name
            self.tag_defaults[cat_name]["type_3d"] = type_name
        else:
            self.tag_defaults[cat_name]["family_plan"] = family_name
            self.tag_defaults[cat_name]["type_plan"] = type_name
        self.logger.add("Сохранение настроек для '{0}' ({1}): {2} - {3}".format(
            cat_name, view_type, family_name, type_name))
        self.SaveTagDefaults()

    def OnTagFamilyDoubleClick3D(self, sender, args):
        """Обработчик двойного клика по таблице 3D"""
        if self.lstTagFamilies3D.SelectedItems.Count == 0:
            return
        selected = self.lstTagFamilies3D.SelectedItems[0]
        category = selected.Tag
        self.OpenTagSelectionDialog(selected, category, "3D")

    def OnTagFamilyDoubleClickPlan(self, sender, args):
        """Обработчик двойного клика по таблице Планов"""
        if self.lstTagFamiliesPlan.SelectedItems.Count == 0:
            return
        selected = self.lstTagFamiliesPlan.SelectedItems[0]
        category = selected.Tag
        self.OpenTagSelectionDialog(selected, category, "План")

    def OpenTagSelectionDialog(self, selected_item, category, view_type):
        """Открывает диалог выбора марки"""
        available_families = self.GetAvailableTagFamiliesForCategory(category)

        if available_families:
            if view_type == "3D":
                current_family = self.settings.category_tag_families_3d.get(category)
                current_type = self.settings.category_tag_types_3d.get(category)
            else:
                current_family = self.settings.category_tag_families_plan.get(category)
                current_type = self.settings.category_tag_types_plan.get(category)

            form = TagFamilySelectionForm(
                self.doc, available_families, current_family, current_type
            )
            if (
                form.ShowDialog() == DialogResult.OK
                and form.SelectedFamily
                and form.SelectedType
            ):
                selected_item.SubItems[1].Text = self.GetElementName(form.SelectedFamily)
                selected_item.SubItems[2].Text = self.GetElementName(form.SelectedType)
                
                cat_name = self.GetCategoryName(category)
                family_name = self.GetElementName(form.SelectedFamily)
                type_name = self.GetElementName(form.SelectedType)
                
                if view_type == "3D":
                    self.settings.category_tag_families_3d[category] = form.SelectedFamily
                    self.settings.category_tag_types_3d[category] = form.SelectedType
                else:
                    self.settings.category_tag_families_plan[category] = form.SelectedFamily
                    self.settings.category_tag_types_plan[category] = form.SelectedType
                
                self.SaveTagForCurrentView(cat_name, view_type, family_name, type_name)
        else:
            MessageBox.Show("Нет доступных семейств марок для этой категории")

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
                    # Ищем активный типоразмер
                    for symbol_id in list(symbol_ids):
                        tag_type = self.doc.GetElement(symbol_id)
                        if tag_type and tag_type.IsActive:
                            return family, tag_type
                    # Если нет активных, берем первый доступный
                    tag_type = self.doc.GetElement(list(symbol_ids)[0])
                    return family, tag_type

        return None, None

    def GetTagCategoryId(self, element_category):
        mapping = {
            BuiltInCategory.OST_DuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_FlexDuctCurves: BuiltInCategory.OST_DuctTags,
            BuiltInCategory.OST_DuctTerminal: BuiltInCategory.OST_DuctTerminalTags,
            BuiltInCategory.OST_DuctAccessory: BuiltInCategory.OST_DuctAccessoryTags,
            BuiltInCategory.OST_DuctInsulations: BuiltInCategory.OST_DuctInsulationsTags,
            BuiltInCategory.OST_MechanicalEquipment: BuiltInCategory.OST_MechanicalEquipmentTags,
        }

        try:
            if hasattr(element_category, "Id") and element_category.Id.IntegerValue < 0:
                element_cat = BuiltInCategory(element_category.Id.IntegerValue)
                if element_cat in mapping:
                    tag_cat = Category.GetCategory(self.doc, mapping[element_cat])
                    if tag_cat:
                        return tag_cat.Id
        except Exception as e:
            self.logger.add("Ошибка получения категории марки: {0}".format(e))
        return None

    def GetElementName(self, element):
        if not element:
            return "Без имени"
        try:
            # Для FamilySymbol (типоразмеров)
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

                    # Пробуем получить имя типа через параметры
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

            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            # Общий случай
            if hasattr(element, "Name") and element.Name:
                name = element.Name
                if name and not name.startswith("IronPython"):
                    return name

        except Exception as e:
            self.logger.add(
                "GetElementName: Ошибка при получении имени элемента: {0}".format(e)
            )

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
        summary = "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ:\r\n\r\n"

        # Подсчёт типов видов
        views_3d_count = sum(1 for v in self.settings.selected_views if isinstance(v, View3D))
        views_plan_count = sum(1 for v in self.settings.selected_views if isinstance(v, ViewPlan))

        summary += "Выбрано видов: " + str(len(self.settings.selected_views)) + "\r\n"
        summary += "  - 3D виды: " + str(views_3d_count) + "\r\n"
        summary += "  - Планы: " + str(views_plan_count) + "\r\n"
        summary += (
            "Выбрано категорий: " + str(len(self.settings.selected_categories)) + "\r\n"
        )

        orientation_text = (
            "Горизонтальная"
            if self.settings.orientation == TagOrientation.Horizontal
            else "Вертикальная"
        )
        summary += "Ориентация: " + orientation_text + "\r\n"
        summary += "Выноска: " + ("Да" if self.settings.use_leader else "Нет") + "\r\n"
        summary += "Случайное смещение: " + ("Да" if self.settings.random_offset else "Нет") + "\r\n"
        summary += (
            "Логирование: "
            + ("Включено" if self.settings.enable_logging else "Отключено")
            + "\r\n\r\n"
        )

        summary += "Детали по категориям:\r\n"
        for category in self.settings.selected_categories:
            # Для 3D
            tag_family_3d = self.settings.category_tag_families_3d.get(category)
            tag_type_3d = self.settings.category_tag_types_3d.get(category)
            if tag_family_3d and tag_type_3d:
                status_3d = self.GetElementName(tag_type_3d)
            else:
                status_3d = "НЕТ МАРКИ"

            # Для планов
            tag_family_plan = self.settings.category_tag_families_plan.get(category)
            tag_type_plan = self.settings.category_tag_types_plan.get(category)
            if tag_family_plan and tag_type_plan:
                status_plan = self.GetElementName(tag_type_plan)
            else:
                status_plan = "НЕТ МАРКИ"

            cat_name = self.GetCategoryName(category)
            summary += "- " + cat_name + ":\r\n"
            summary += "    3D: " + status_3d + "\r\n"
            summary += "    План: " + status_plan + "\r\n"

        return summary

    def OnExecuteClick(self, sender, args):
        """
        Обработчик кнопки 'Выполнить'.
        """
        success_count = 0
        errors = []

        self.logger.add("Начало выполнения расстановки марок.")
        self.logger.add("Выбрано видов: {0}".format(len(self.settings.selected_views)))

        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        try:
            # 1. Собираем элементы по видам и категориям
            elements_by_view_and_category = {}
            for view in self.settings.selected_views:
                # Проверяем тип вида - поддерживаем 3D и Планы
                if not isinstance(view, (View3D, ViewPlan)):
                    error_msg = "Вид '{0}' не поддерживается (требуется 3D или План), пропущен".format(view.Name)
                    errors.append(error_msg)
                    self.logger.add(error_msg)
                    continue

                view_type = "3D" if isinstance(view, View3D) else "План"
                self.logger.add("Обработка вида: {0} [{1}]".format(view.Name, view_type))
                elements_by_view_and_category[view.Id] = {}

                for category in self.settings.selected_categories:
                    elements = list(
                        FilteredElementCollector(self.doc, view.Id)
                        .OfCategoryId(category.Id)
                        .WhereElementIsNotElementType()
                        .ToElements()
                    )
                    elements_by_view_and_category[view.Id][category.Id] = elements
                    self.logger.add(
                        "Категория '{0}': найдено элементов {1}".format(
                            self.GetCategoryName(category), len(elements)
                        )
                    )

            # 2. Применяем фильтрацию для воздуховодов (если нужно)
            # Сначала собираем все элементы для прогресса
            all_elements_for_progress = []

            for view in self.settings.selected_views:
                if not isinstance(view, (View3D, ViewPlan)):
                    continue

                for category in self.settings.selected_categories:
                    elements = elements_by_view_and_category[view.Id][category.Id]

                    # Применяем новую логику только для воздуховодов
                    if category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves):
                        # Отбираем воздуховоды для маркировки
                        elements = self.GetDuctsToTag(elements)
                        self.logger.add(
                            "Для категории '{0}' отобрано {1} воздуховодов для маркировки".format(
                                self.GetCategoryName(category), len(elements)
                            )
                        )

                    # Обновляем элементы в кэше
                    elements_by_view_and_category[view.Id][category.Id] = elements
                    # Добавляем в общий список для прогресса
                    all_elements_for_progress.extend(elements)

            # 3. Настраиваем прогресс-бар на РЕАЛЬНОЕ количество операций
            total_operations = len(all_elements_for_progress)
            self.logger.add("Всего операций после фильтрации: {0}".format(total_operations))

            if total_operations == 0:
                self.progressBar.Maximum = 1
            else:
                self.progressBar.Maximum = total_operations

            self.progressBar.Value = 0
            
            # 4. Основной цикл расстановки марок
            processed_count = 0

            for view in self.settings.selected_views:
                # Пропускаем неподдерживаемые типы видов
                if not isinstance(view, (View3D, ViewPlan)):
                    continue

                view_type = "3D" if isinstance(view, View3D) else "План"
                self.logger.add("  [Вид] {0} [{1}]".format(view.Name, view_type))

                for category in self.settings.selected_categories:
                    elements = elements_by_view_and_category[view.Id][category.Id]

                    self.logger.add("    [{0}] Обработка {1} элементов...".format(
                        self.GetCategoryName(category), len(elements)))

                    for element in elements:
                        # Обновляем прогресс перед каждой операцией
                        processed_count += 1
                        self.progressBar.Value = min(
                            processed_count, self.progressBar.Maximum
                        )

                        # Обновляем UI для плавного отображения прогресса
                        Application.DoEvents()

                        if self.HasExistingTag(element, view):
                            self.logger.add(
                                "  ⊘ Элемент {0} уже имеет марку, пропущен".format(
                                    element.Id
                                )
                            )
                            continue

                        if self.CreateTag(element, view, category):
                            success_count += 1
                        else:
                            error_msg = "Не удалось создать марку для элемента {0}".format(
                                element.Id
                            )
                            errors.append(error_msg)
                            self.logger.add("  ✗ {0}".format(error_msg))

            trans.Commit()
            self.logger.add("Транзакция подтверждена успешно")
            
        except Exception as e:
            trans.RollBack()
            error_msg = "Критическая ошибка: {0}".format(e)
            errors.append(error_msg)
            self.logger.add(error_msg)
        finally:
            trans.Dispose()
            self.logger.add("Транзакция завершена")
        
        # 5. Гарантируем, что прогресс-бар показывает 100% в конце
        self.progressBar.Value = self.progressBar.Maximum
        
        # Сохранить выбранные марки
        self.SaveTagDefaults()

        result_msg = "Успешно расставлено марок: {0}".format(success_count)
        if errors:
            result_msg += "\n\nОшибки ({0}):\n".format(len(errors)) + "\n".join(
                errors[:10]
            )
            if len(errors) > 10:
                result_msg += "\n... и еще {0} ошибок".format(len(errors) - 10)

        MessageBox.Show(result_msg, "Результат")
        if self.settings.enable_logging:
            self.logger.show()
        self.Close()

    def CreateTag(self, element, view, category):
        """
        Создает марку для указанного элемента на виде.

        Args:
            element (Element): Элемент для марки.
            view (View): Вид (3D или План).
            category (Category): Категория элемента.

        Returns:
            bool: True, если марка создана успешно.
        """
        try:
            # Получаем тип марки для текущего вида
            if isinstance(view, View3D):
                tag_type = self.settings.category_tag_types_3d.get(category)
                view_type_name = "3D"
            else:
                tag_type = self.settings.category_tag_types_plan.get(category)
                view_type_name = "План"
            
            if not tag_type:
                self.logger.add("  ⚠️ Тип марки не выбран для категории '{}' ({} вид)".format(
                    self.GetCategoryName(category), view_type_name))
                return False

            # Проверяем bounding box
            bbox = element.get_BoundingBox(view)
            if not bbox or not bbox.Min or not bbox.Max:
                self.logger.add("  ⚠️ Элемент ID {} не имеет bounding box на виде {}".format(
                    element.Id.IntegerValue, view.Name))
                return False

            center = (bbox.Min + bbox.Max) / 2
            scale_factor = 100.0 / view.Scale

            self.logger.add("  -> Элемент ID {}: центр=({}, {}), масштаб={}".format(
                element.Id.IntegerValue,
                round(center.X, 2),
                round(center.Y, 2),
                round(view.Scale, 1)))

            # Информация о типе марки
            if tag_type:
                try:
                    tag_info = "ID={}, Class={}".format(
                        tag_type.Id.IntegerValue,
                        tag_type.__class__.__name__)
                    self.logger.add("  -> Тип марки ({}): {}".format(view_type_name, tag_info))
                except:
                    self.logger.add("  -> Тип марки: ID={}".format(tag_type.Id.IntegerValue))
            else:
                self.logger.add("  -> Тип марки: НЕ УКАЗАН")

            offset_x = (self.settings.offset_x * scale_factor) / MM_TO_FEET
            offset_y = (self.settings.offset_y * scale_factor) / MM_TO_FEET

            # Для 3D видов - случайное направление, для планов - только XY
            if isinstance(view, View3D):
                # 3D вид - случайное направление по всем осям
                if self.settings.random_offset:
                    direction_x = random.choice([-1, 1])
                    direction_y = random.choice([-1, 1])
                    direction_z = random.choice([-1, 1])
                else:
                    direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
                    direction_y = 1 if element.Id.IntegerValue % 3 == 0 else -1
                    direction_z = 1 if element.Id.IntegerValue % 5 == 0 else -1

                tag_point = XYZ(
                    center.X + offset_x * direction_x,
                    center.Y + offset_y * direction_y,
                    center.Z + offset_x * 0.5 * direction_z,  # Меньшее смещение по Z
                )
            else:
                # План этажа - смещение только по XY
                if self.settings.random_offset:
                    direction_x = random.choice([-1, 1])
                    direction_y = random.choice([-1, 1])
                else:
                    direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
                    direction_y = 1 if element.Id.IntegerValue % 3 == 0 else -1

                tag_point = XYZ(
                    center.X + offset_x * direction_x,
                    center.Y + offset_y * direction_y,
                    center.Z,
                )

            self.logger.add("  -> Точка для марки: ({}, {}, {})".format(
                round(tag_point.X, 2),
                round(tag_point.Y, 2),
                round(tag_point.Z, 2)))

            # Проверяем тип марки
            if not tag_type:
                self.logger.add("  ⚠️ Тип марки не указан!")
                return False

            # Получаем имя типа марки для логирования (с обработкой ошибок)
            tag_type_name = None
            try:
                # Пытаемся получить имя через параметр
                name_param = tag_type.LookupParameter("Тип")
                if name_param and name_param.HasValue:
                    tag_type_name = name_param.AsString()
                
                if not tag_type_name:
                    name_param = tag_type.LookupParameter("Type Name")
                    if name_param and name_param.HasValue:
                        tag_type_name = name_param.AsString()
                
                if not tag_type_name:
                    # Пробуем прямое свойство
                    tag_type_name = str(tag_type.Name) if tag_type.Name else None
                    
            except:
                pass
            
            # Если не удалось получить имя, используем ID
            if not tag_type_name:
                tag_type_name = "ID_{}".format(tag_type.Id.IntegerValue)

            # Проверяем и активируем типоразмер если нужно
            if not tag_type.IsActive:
                self.logger.add("  ⚠️ Тип марки '{}' не активен, активирую...".format(tag_type_name))
                try:
                    tag_type.Activate()
                    self.logger.add("  ✓ Тип марки '{}' активирован".format(tag_type_name))
                except Exception as e:
                    self.logger.add("  ⚠️ Не удалось активировать тип '{}': {}".format(tag_type_name, e))
                    # Продолжаем, возможно марка создастся с типом по умолчанию

            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                Reference(element),
                self.settings.use_leader,
                TagMode.TM_ADDBY_CATEGORY,
                self.settings.orientation,
                tag_point,
            )

            if tag:
                self.logger.add("  ✓ Марка создана для элемента ID {}".format(element.Id.IntegerValue))

                # Примечание: Тип конца выноски (Leader End) нельзя изменить программно
                # Это настройка семейства марки, которая задаётся в редакторе семейств Revit

                # Пытаемся установить нужный типоразмер
                try:
                    tag.ChangeTypeId(tag_type.Id)
                    self.logger.add("  ✓ Тип марки изменён на '{}'".format(tag_type_name))
                except Exception as e:
                    self.logger.add("  ⚠️ Не удалось изменить тип марки на '{}': {}".format(tag_type_name, e))
                    self.logger.add("     Примечание: марка создана с типом по умолчанию")
                return True
            else:
                self.logger.add("  ⚠️ IndependentTag.Create вернул None")
                return False

        except Exception as e:
            self.logger.add("  ⚠️ Exception при создании марки: {0}".format(e))
            return False

    def HasExistingTag(self, element, view):
        """
        Проверяет, есть ли уже марка для элемента на виде.

        Args:
            element (Element): Элемент.
            view (View): Вид (3D или План).

        Returns:
            bool: True, если марка существует.
        """
        try:
            for tag in (
                FilteredElementCollector(self.doc, view.Id)
                .OfClass(IndependentTag)
                .ToElements()
            ):
                try:
                    for tagged_elem in tag.GetTaggedLocalElements():
                        if tagged_elem.Id == element.Id:
                            return True
                except Exception as e:
                    self.logger.add("Ошибка проверки марки {0}: {1}".format(tag.Id, e))
                    if (
                        hasattr(tag, "TaggedLocalElementId")
                        and tag.TaggedLocalElementId == element.Id
                    ):
                        return True
            return False
        except Exception as e:
            self.logger.add("Ошибка проверки существующей марки: {0}".format(e))
            return False

    def OnShowLogsClick(self, sender, args):
        """
        Обработчик кнопки 'Показать логи'.
        """
        self.logger.show()

    def LoadTagDefaults(self):
        """
        Загружает сохраненные дефолты марок из файла.

        Returns:
            dict: Словарь с марками по категориям.
        """
        config_path = self.GetConfigPath()
        if os.path.exists(config_path):
            try:
                # Используем codecs для совместимости с IronPython 2.7
                with codecs.open(config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.logger.add("Настройки марок загружены из: {}".format(config_path))
                    return data
            except Exception as e:
                self.logger.add("Ошибка загрузки настроек марок: {0}".format(e))
        else:
            self.logger.add("Файл настроек не найден: {}".format(config_path))
        return {}

    def SaveTagDefaults(self):
        """
        Сохраняет выбранные марки в файл (и для 3D, и для планов).
        """
        defaults = {}

        # Сохраняем настройки для 3D видов
        for cat, family in self.settings.category_tag_families_3d.items():
            cat_name = self.GetCategoryName(cat)
            if cat_name not in defaults:
                defaults[cat_name] = {}

            family_name = self.GetElementName(family)
            type_elem = self.settings.category_tag_types_3d.get(cat)
            type_name = self.GetElementName(type_elem) if type_elem else ""

            if family_name and type_name:
                defaults[cat_name]["family_3d"] = family_name
                defaults[cat_name]["type_3d"] = type_name

        # Сохраняем настройки для планов
        for cat, family in self.settings.category_tag_families_plan.items():
            cat_name = self.GetCategoryName(cat)
            if cat_name not in defaults:
                defaults[cat_name] = {}

            family_name = self.GetElementName(family)
            type_elem = self.settings.category_tag_types_plan.get(cat)
            type_name = self.GetElementName(type_elem) if type_elem else ""

            if family_name and type_name:
                defaults[cat_name]["family_plan"] = family_name
                defaults[cat_name]["type_plan"] = type_name

        config_path = self.GetConfigPath()
        try:
            # Используем codecs для совместимости с IronPython 2.7
            with codecs.open(config_path, "w", encoding="utf-8") as f:
                json.dump(defaults, f, ensure_ascii=False, indent=4)
            self.logger.add("Настройки марок сохранены в: {}".format(config_path))
        except Exception as e:
            self.logger.add("Ошибка сохранения настроек марок: {0}".format(e))

    def FindSavedTag(self, family_name, type_name):
        """
        Находит семейство и тип по именам из сохраненных.

        Args:
            family_name (str): Имя семейства.
            type_name (str): Имя типоразмера.

        Returns:
            tuple: (Family, FamilySymbol) или (None, None).
        """
        collector = FilteredElementCollector(self.doc).OfClass(Family)
        for family in collector:
            if self.GetElementName(family) == family_name:
                symbol_ids = family.GetFamilySymbolIds()
                for symbol_id in symbol_ids:
                    symbol = self.doc.GetElement(symbol_id)
                    if self.GetElementName(symbol) == type_name:
                        return family, symbol
        return None, None

    def GetConfigPath(self):
        """
        Возвращает путь к файлу конфигурации.

        Returns:
            str: Путь к tag_defaults.json.
        """
        script_dir = os.path.dirname(__file__)
        return os.path.join(script_dir, "tag_defaults.json")


# Форма выбора семейства марки
class TagFamilySelectionForm(Form):
    """
    Форма для выбора семейства и типоразмера марки из доступных вариантов.

    Attributes:
        doc (Document): Документ Revit.
        available_families (list): Список доступных семейств.
        selected_family: Выбранное семейство.
        selected_type: Выбранный типоразмер.
    """

    def __init__(self, doc, available_families, current_family, current_type):
        self.doc = doc
        self.available_families = available_families
        self.selected_family = None
        self.selected_type = None
        self.family_dict = {}
        self.type_dict = {}
        self.logger = Logger(enabled=True)  # Логирование для формы выбора

        self.InitializeComponent()
        self.PopulateFamilies()

    def InitializeComponent(self):
        """
        Инициализирует компоненты формы.
        """
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
        self.lstFamilies.DoubleClick += self.OnFamilyDoubleClick
        self.lstTypes.DoubleClick += self.OnTypeDoubleClick
        self.btnOK.Click += self.OnOKClick
        self.btnCancel.Click += self.OnCancelClick

        for c in controls:
            self.Controls.Add(c)

    def PopulateFamilies(self):
        """
        Заполняет список семейств доступными вариантами.
        """
        self.lstFamilies.Items.Clear()
        self.family_dict.clear()

        for family in self.available_families:
            if family:
                family_name = self.GetElementNameImproved(family)
                self.lstFamilies.Items.Add(family_name)
                self.family_dict[family_name] = family

        # Выбор текущего или первого семейства
        if self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0

    def GetElementName(self, element):
        """
        Получает имя элемента Revit.

        Args:
            element (Element): Элемент Revit.

        Returns:
            str: Имя элемента.
        """
        if not element:
            return "Без имени"
        try:
            # Для FamilySymbol (типоразмеров)
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

                    # Пробуем получить имя типа через параметры
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

            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            # Общий случай
            if hasattr(element, "Name") and element.Name:
                name = element.Name
                if name and not name.startswith("IronPython"):
                    return name

        except Exception as e:
            self.logger.add(
                "GetElementName: Ошибка при получении имени элемента: {0}".format(e)
            )

        return "Элемент " + str(element.Id.IntegerValue)

    def GetElementNameImproved(self, element):
        """
        Получает улучшенное имя для элемента Revit.

        Args:
            element (Element): Элемент Revit.

        Returns:
            str: Имя элемента.
        """
        if not element:
            return "Без имени"
        try:
            # Для FamilySymbol (типоразмеров)
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

                    # Пробуем получить имя типа через параметры
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

            # Для Family (семейств)
            elif isinstance(element, Family):
                if hasattr(element, "Name") and element.Name:
                    name = element.Name
                    if name and not name.startswith("IronPython"):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            # Общий случай
            if hasattr(element, "Name") and element.Name:
                name = element.Name
                if name and not name.startswith("IronPython"):
                    return name

        except Exception as e:
            self.logger.add("Ошибка получения имени элемента: {0}".format(e))

        return "Элемент " + str(element.Id.IntegerValue)

    def OnFamilySelected(self, sender, args):
        """
        Обработчик изменения выбора семейства.
        """
        if self.lstFamilies.SelectedIndex >= 0:
            selected_name = self.lstFamilies.SelectedItem
            selected_family = self.family_dict.get(selected_name)
            if selected_family:
                self.PopulateTypesList(selected_family)

    def PopulateTypesList(self, family):
        """
        Заполняет список типоразмеров для выбранного семейства.

        Args:
            family (Family): Семейство марок.
        """
        self.lstTypes.Items.Clear()
        self.type_dict.clear()

        try:
            symbol_ids = family.GetFamilySymbolIds()
            if symbol_ids and symbol_ids.Count > 0:
                for symbol_id in list(symbol_ids):
                    symbol = self.doc.GetElement(symbol_id)
                    if symbol:
                        symbol_name = self.GetElementNameImproved(symbol)
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

                # Выбор первого активного типоразмера
                for i in range(self.lstTypes.Items.Count):
                    display_name = self.lstTypes.Items[i]
                    symbol = self.type_dict.get(display_name)
                    if symbol and symbol.IsActive:
                        self.lstTypes.SelectedIndex = i
                        return

                # Если нет активных, выбираем первый
                if self.lstTypes.Items.Count > 0:
                    self.lstTypes.SelectedIndex = 0
            else:
                MessageBox.Show("В выбранном семействе нет типоразмеров")

        except Exception as e:
            self.logger.add("Ошибка загрузки типоразмеров: {0}".format(e))
            MessageBox.Show("Ошибка при загрузке типоразмеров")

    def OnFamilyDoubleClick(self, sender, args):
        """
        Обработчик двойного клика по списку семейств.
        """
        if self.lstFamilies.SelectedIndex >= 0:
            selected_name = self.lstFamilies.SelectedItem
            selected_family = self.family_dict.get(selected_name)
            if selected_family:
                self.PopulateTypesList(selected_family)

    def OnTypeDoubleClick(self, sender, args):
        """
        Обработчик двойного клика по списку типоразмеров.
        """
        if self.lstTypes.SelectedIndex >= 0:
            self.OnOKClick(sender, args)

    def OnOKClick(self, sender, args):
        """
        Обработчик кнопки OK.
        """
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
        """
        Обработчик кнопки Отмена.
        """
        self.DialogResult = DialogResult.Cancel
        self.Close()

    @property
    def SelectedFamily(self):
        return self.selected_family

    @property
    def SelectedType(self):
        return self.selected_type


# Запуск
def main():
    """
    Главная функция для запуска приложения расстановки марок.
    """
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