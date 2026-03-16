# -*- coding: utf-8 -*-
__title__ = """Марки
 на 3D"""
__author__ = "Rage"
__doc__ = "Автоматическое размещение марок на 3D-видах"
__version__ = "1.0"

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import datetime
import json
import os

from Autodesk.Revit.DB import *
from System.Drawing import *
from System.Windows.Forms import *

# Константы
MM_TO_FEET = 304.8
DEFAULT_OFFSET_X = 60.0
DEFAULT_OFFSET_Y = 30.0


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
        selected_views (list): Выбранные 3D-виды.
        selected_categories (list): Выбранные категории элементов.
        category_tag_families (dict): Семейства марок по категориям.
        category_tag_types (dict): Типоразмеры марок по категориям.
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
        self.selected_views = []
        self.selected_categories = []
        self.category_tag_families = {}
        self.category_tag_types = {}
        self.offset_x = DEFAULT_OFFSET_X
        self.offset_y = DEFAULT_OFFSET_Y
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True
        self.enable_logging = False


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
        self.Load3DViews()

    def InitializeComponent(self):
        """
        Инициализирует компоненты формы.
        """
        self.Text = "Расстановка марок на 3D видах"
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
        Настраивает вкладку 2: Категории.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
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
        """
        Настраивает вкладку 3: Марки.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        self.lstTagFamilies = ListView()
        self.lstTagFamilies.Location = Point(10, 40)
        self.lstTagFamilies.Size = Size(700, 400)
        self.lstTagFamilies.View = View.Details
        self.lstTagFamilies.FullRowSelect = True
        self.lstTagFamilies.GridLines = True
        self.lstTagFamilies.Columns.Add("Категория", 180)
        self.lstTagFamilies.Columns.Add("Семейство марки", 200)
        self.lstTagFamilies.Columns.Add("Типоразмер марки", 300)
        self.lstTagFamilies.DoubleClick += self.OnTagFamilyDoubleClick

        controls = [
            self.CreateControl(
                Label,
                Text="Выберите марки для категорий:",
                Location=Point(10, 10),
                Size=Size(400, 20),
            ),
            self.lstTagFamilies,
            self.CreateButton(
                "← Назад", Point(500, 450), click_handler=self.OnBack2Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 450), click_handler=self.OnNext3Click
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
        self.cmbOrientation.Size = Size(100, 20)
        self.cmbOrientation.Items.Add("Горизонтальная")
        self.cmbOrientation.Items.Add("Вертикальная")
        self.cmbOrientation.SelectedIndex = 0

        self.chkUseLeader = CheckBox()
        self.chkUseLeader.Text = "Использовать выноску"
        self.chkUseLeader.Location = Point(10, 80)
        self.chkUseLeader.Size = Size(200, 20)
        self.chkUseLeader.Checked = True

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
            self.CreateButton(
                "← Назад", Point(500, 450), click_handler=self.OnBack3Click
            ),
            self.CreateButton(
                "Далее →", Point(600, 450), click_handler=self.OnNext4Click
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

    def Load3DViews(self):
        """
        Загружает список 3D-видов в список.
        """
        try:
            views = (
                FilteredElementCollector(self.doc)
                .OfClass(View3D)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            self.lstViews.Items.Clear()
            self.all_views_dict.clear()
            self.views_dict.clear()
            self.lstViewsChecked.clear()
            for view in views:
                if not view.IsTemplate and view.CanBePrinted:
                    name = view.Name + " (ID: " + str(view.Id.IntegerValue) + ")"
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
        Обновляет список видов с фильтром.

        Args:
            filter_text (str): Текст фильтра.
        """
        self.lstViews.Items.Clear()
        self.views_dict.clear()
        filter_lower = filter_text.lower()
        for name, view in self.all_views_dict.items():
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
        Собирает категории элементов для маркировки.
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
        self.lstCategories.Items.Clear()
        self.category_mapping.clear()

        for cat in sorted(
            self.settings.selected_categories, key=lambda x: self.GetCategoryName(x)
        ):
            name = self.GetCategoryName(cat)
            self.lstCategories.Items.Add(name, True)
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

    def PopulateTagFamilies(self):
        self.lstTagFamilies.Items.Clear()
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            item.Tag = category

            # Попытаться найти сохраненную марку
            cat_name = self.GetCategoryName(category)
            default = self.tag_defaults.get(cat_name, {})
            family_name = default.get("family")
            type_name = default.get("type")

            if family_name and type_name:
                # Найти семейство и тип по имени
                tag_family, tag_type = self.FindSavedTag(family_name, type_name)
                if tag_family and tag_type:
                    item.SubItems.Add(self.GetElementName(tag_family))
                    item.SubItems.Add(self.GetElementName(tag_type))
                    self.settings.category_tag_families[category] = tag_family
                    self.settings.category_tag_types[category] = tag_type
                    self.lstTagFamilies.Items.Add(item)
                    continue

            # Если не найдено, найти дефолтное
            tag_family, tag_type = self.FindTagForCategory(category)

            if tag_family and tag_type:
                item.SubItems.Add(self.GetElementName(tag_family))
                item.SubItems.Add(self.GetElementName(tag_type))
                # Сохранить в settings
                self.settings.category_tag_families[category] = tag_family
                self.settings.category_tag_types[category] = tag_type
            else:
                item.SubItems.Add("Нет подходящих марок")
                item.SubItems.Add("")

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
        else:
            MessageBox.Show("Нет доступных семейств марок для этой категории")

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
        summary += "Выбрано видов: " + str(len(self.settings.selected_views)) + "\r\n"
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
        summary += (
            "Логирование: "
            + ("Включено" if self.settings.enable_logging else "Отключено")
            + "\r\n\r\n"
        )

        summary += "Детали по категориям:\r\n"
        for category in self.settings.selected_categories:
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            if tag_family and tag_type:
                status = (
                    self.GetElementName(tag_family)
                    + " ("
                    + self.GetElementName(tag_type)
                    + ")"
                )
            else:
                status = "НЕТ МАРКИ"
            summary += "- " + self.GetCategoryName(category) + ": " + status + "\r\n"
        return summary

    def OnExecuteClick(self, sender, args):
        """
        Обработчик кнопки 'Выполнить'.
        """
        success_count = 0
        errors = []
        
        self.logger.add("Начало выполнения расстановки марок.")
        
        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        try:
            # 1. Собираем элементы по видам и категориям
            elements_by_view_and_category = {}
            for view in self.settings.selected_views:
                if not isinstance(view, View3D):
                    error_msg = "Вид '{0}' не 3D, пропущен".format(view.Name)
                    errors.append(error_msg)
                    self.logger.add(error_msg)
                    continue
                    
                self.logger.add("Обработка вида: {0}".format(view.Name))
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
                if not isinstance(view, View3D):
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
                if not isinstance(view, View3D):
                    continue
                    
                for category in self.settings.selected_categories:
                    elements = elements_by_view_and_category[view.Id][category.Id]
                    tag_type = self.settings.category_tag_types.get(category)

                    if not tag_type:
                        error_msg = "Нет марки для категории '{0}'".format(
                            self.GetCategoryName(category)
                        )
                        errors.append(error_msg)
                        self.logger.add(error_msg)
                        continue

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
                                "Элемент {0} уже имеет марку на виде {1}, пропущен".format(
                                    element.Id, view.Name
                                )
                            )
                            continue
                            
                        if self.CreateTag(element, view, tag_type):
                            success_count += 1
                            self.logger.add(
                                "Марка создана для элемента {0} на виде {1}".format(
                                    element.Id, view.Name
                                )
                            )
                        else:
                            error_msg = "Не удалось создать марку для элемента {0} на виде {1}".format(
                                element.Id, view.Name
                            )
                            errors.append(error_msg)
                            self.logger.add(error_msg)

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

    def CreateTag(self, element, view, tag_type):
        """
        Создает марку для указанного элемента на виде.

        Args:
            element (Element): Элемент для марки.
            view (View3D): 3D-вид.
            tag_type (FamilySymbol): Тип марки.

        Returns:
            bool: True, если марка создана успешно.
        """
        try:
            bbox = element.get_BoundingBox(view)
            if not bbox or not bbox.Min or not bbox.Max:
                return False

            center = (bbox.Min + bbox.Max) / 2
            scale_factor = 100.0 / view.Scale

            offset_x = (self.settings.offset_x * scale_factor) / MM_TO_FEET
            offset_y = (self.settings.offset_y * scale_factor) / MM_TO_FEET

            direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
            direction_y = 1 if element.Id.IntegerValue % 3 == 0 else -1

            tag_point = XYZ(
                center.X + offset_x * direction_x,
                center.Y + offset_y * direction_y,
                center.Z,
            )

            tag = IndependentTag.Create(
                self.doc,
                view.Id,
                Reference(element),
                self.settings.use_leader,
                TagMode.TM_ADDBY_CATEGORY,
                self.settings.orientation,
                tag_point,
            )

            if tag and tag_type:
                try:
                    tag.ChangeTypeId(tag_type.Id)
                except RevitException as e:
                    self.logger.add("Не удалось изменить тип марки: {0}".format(e))
                    return False
                return True
            return False
        except Exception as e:
            self.logger.add("Ошибка создания марки: {0}".format(e))
            return False

    def HasExistingTag(self, element, view):
        """
        Проверяет, есть ли уже марка для элемента на виде.

        Args:
            element (Element): Элемент.
            view (View3D): 3D-вид.

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
                with open(config_path, "r") as f:
                    return json.load(f)
            except Exception as e:
                self.logger.add("Ошибка загрузки настроек марок: {0}".format(e))
        return {}

    def SaveTagDefaults(self):
        """
        Сохраняет выбранные марки в файл.
        """
        defaults = {}
        for cat, family in self.settings.category_tag_families.items():
            cat_name = self.GetCategoryName(cat)
            family_name = self.GetElementName(family)
            type_elem = self.settings.category_tag_types.get(cat)
            type_name = self.GetElementName(type_elem) if type_elem else ""
            if family_name and type_name:
                defaults[cat_name] = {"family": family_name, "type": type_name}

        config_path = self.GetConfigPath()
        try:
            with open(config_path, "w") as f:
                json.dump(defaults, f, ensure_ascii=True, indent=4)
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