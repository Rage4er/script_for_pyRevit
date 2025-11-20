# -*- coding: utf-8 -*-
__title__ = 'Марки на 3D'
__author__ = 'Rage'
__doc__ = 'Автоматическое размещение маркировочных меток на 3D-видах'

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
import datetime

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
        self.category_mapping = {}
        self.logger = Logger(self.settings.enable_logging)

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

        tabs = ["1. Выбор видов", "2. Категории", "3. Марки", "4. Настройки", "5. Выполнение"]
        for i, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            getattr(self, "SetupTab" + str(i+1))(tab)
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
        controls = [
            self.CreateControl(Label, Text="Выберите 3D виды:", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(CheckedListBox, Location=Point(10, 40), Size=Size(600, 400), CheckOnClick=True),
            self.CreateControl(CheckBox, Text="Включить логирование", Location=Point(10, 450), Size=Size(200, 20)),
            self.CreateButton("Показать логи", Point(220, 450), Size(100, 25), self.OnShowLogsClick),
            self.CreateButton("Далее →", Point(600, 450), click_handler=self.OnNext1Click)
        ]
        self.lblViews, self.lstViews, self.chkLogging, self.btnShowLogs, self.btnNext1 = controls
        self.chkLogging.CheckedChanged += self.OnLoggingCheckedChanged
        for c in controls:
            tab.Controls.Add(c)

    def SetupTab2(self, tab):
        """
        Настраивает вкладку 2: Категории.

        Args:
            tab (TabPage): Вкладка для настройки.
        """
        controls = [
            self.CreateControl(Label, Text="Выберите категории элементов:", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(CheckedListBox, Location=Point(10, 40), Size=Size(600, 400), CheckOnClick=True),
            self.CreateButton("← Назад", Point(500, 450), click_handler=self.OnBack1Click),
            self.CreateButton("Далее →", Point(600, 450), click_handler=self.OnNext2Click)
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
            self.CreateControl(Label, Text="Выберите марки для категорий:", Location=Point(10, 10), Size=Size(400, 20)),
            self.lstTagFamilies,
            self.CreateButton("← Назад", Point(500, 450), click_handler=self.OnBack2Click),
            self.CreateButton("Далее →", Point(600, 450), click_handler=self.OnNext3Click)
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
            self.CreateControl(Label, Text="Настройки размещения:", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Label, Text="Ориентация:", Location=Point(10, 50), Size=Size(150, 20)),
            self.cmbOrientation,
            self.chkUseLeader,
            self.CreateButton("← Назад", Point(500, 450), click_handler=self.OnBack3Click),
            self.CreateButton("Далее →", Point(600, 450), click_handler=self.OnNext4Click)
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
            self.CreateControl(Label, Text="Готово к выполнению:", Location=Point(10, 10), Size=Size(300, 20)),
            self.txtSummary,
            self.progressBar,
            self.CreateButton("← Назад", Point(340, 450), click_handler=self.OnBack4Click),
            self.CreateButton("Выполнить", Point(500, 450), Size(150, 30), click_handler=self.OnExecuteClick)
        ]
        for c in controls:
            tab.Controls.Add(c)

    def Load3DViews(self):
        """
        Загружает список 3D-видов в список.
        """
        try:
            views = FilteredElementCollector(self.doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()
            self.lstViews.Items.Clear()
            for view in views:
                if not view.IsTemplate and view.CanBePrinted:
                    name = view.Name + " (ID: " + str(view.Id.IntegerValue) + ")"
                    self.lstViews.Items.Add(name, False)
                    self.views_dict[name] = view
            if self.lstViews.Items.Count > 0:
                self.lstViews.SetItemChecked(0, True)
        except Exception as e:
            MessageBox.Show("Ошибка загрузки видов: " + str(e))

    # Навигация
    def OnNext1Click(self, sender, args):
        self.settings.enable_logging = self.chkLogging.Checked
        self.logger.enabled = self.settings.enable_logging

        self.settings.selected_views = []
        for i in range(self.lstViews.Items.Count):
            name = self.lstViews.Items[i]
            if self.lstViews.GetItemChecked(i) and name in self.views_dict:
                self.settings.selected_views.append(self.views_dict[name])

        if not self.settings.selected_views:
            MessageBox.Show("Выберите хотя бы один вид!")
            return

        self.CollectCategories()
        self.tabControl.SelectedIndex = 1

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
        self.tabControl.SelectedIndex = 2

    def OnNext3Click(self, sender, args):
        self.tabControl.SelectedIndex = 3

    def OnNext4Click(self, sender, args):
        self.settings.orientation = TagOrientation.Horizontal if self.cmbOrientation.SelectedIndex == 0 else TagOrientation.Vertical
        self.settings.use_leader = self.chkUseLeader.Checked
        self.txtSummary.Text = self.GenerateSummary()
        self.tabControl.SelectedIndex = 4

    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedIndex = 0
    def OnBack2Click(self, sender, args):
        self.tabControl.SelectedIndex = 1
    def OnBack3Click(self, sender, args):
        self.tabControl.SelectedIndex = 2
    def OnBack4Click(self, sender, args):
        self.tabControl.SelectedIndex = 3

    def CollectCategories(self):
        categories = [
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_MechanicalEquipment
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

        for cat in sorted(self.settings.selected_categories, key=lambda x: self.GetCategoryName(x)):
            name = self.GetCategoryName(cat)
            self.lstCategories.Items.Add(name, True)
            self.category_mapping[name] = cat

    def GetCategoryName(self, category):
        if not category:
            return "Неизвестная категория"
        try:
            if hasattr(category, 'Id') and category.Id.IntegerValue < 0:
                return LabelUtils.GetLabelFor(BuiltInCategory(category.Id.IntegerValue))
        except Exception as e:
            self.logger.add("Ошибка получения имени категории: {0}".format(e))
        return getattr(category, 'Name', 'Неизвестная категория')

    def PopulateTagFamilies(self):
        self.lstTagFamilies.Items.Clear()
        for category in self.settings.selected_categories:
            item = ListViewItem(self.GetCategoryName(category))
            item.Tag = category
            tag_family, tag_type = self.FindTagForCategory(category)

            if tag_family and tag_type:
                item.SubItems.Add(self.GetElementName(tag_family))
                item.SubItems.Add(self.GetElementName(tag_type))
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
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()

        for family in tag_families:
            if not family or not hasattr(family, 'FamilyCategory'):
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
            BuiltInCategory.OST_MechanicalEquipment: BuiltInCategory.OST_MechanicalEquipmentTags
        }

        try:
            if hasattr(element_category, 'Id') and element_category.Id.IntegerValue < 0:
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
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name

                if hasattr(element, 'Family') and element.Family:
                    family_name = element.Family.Name if hasattr(element.Family, 'Name') and element.Family.Name else ""

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
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            # Общий случай
            if hasattr(element, 'Name') and element.Name:
                name = element.Name
                if name and not name.startswith('IronPython'):
                    return name

        except Exception as e:
            self.logger.add("GetElementName: Ошибка при получении имени элемента: {0}".format(e))

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

            form = TagFamilySelectionForm(self.doc, available_families, current_family, current_type)
            if form.ShowDialog() == DialogResult.OK and form.SelectedFamily and form.SelectedType:
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
        tag_families = collector.OfClass(Family).WhereElementIsNotElementType().ToElements()

        available_families = []
        for family in tag_families:
            if family and hasattr(family, 'FamilyCategory') and family.FamilyCategory and family.FamilyCategory.Id == tag_category_id:
                available_families.append(family)

        return available_families

    def GenerateSummary(self):
        summary = "СВОДКА ПЕРЕД ВЫПОЛНЕНИЕМ:\r\n\r\n"
        summary += "Выбрано видов: " + str(len(self.settings.selected_views)) + "\r\n"
        summary += "Выбрано категорий: " + str(len(self.settings.selected_categories)) + "\r\n"

        orientation_text = "Горизонтальная" if self.settings.orientation == TagOrientation.Horizontal else "Вертикальная"
        summary += "Ориентация: " + orientation_text + "\r\n"
        summary += "Выноска: " + ("Да" if self.settings.use_leader else "Нет") + "\r\n"
        summary += "Логирование: " + ("Включено" if self.settings.enable_logging else "Отключено") + "\r\n\r\n"

        summary += "Детали по категориям:\r\n"
        for category in self.settings.selected_categories:
            tag_family = self.settings.category_tag_families.get(category)
            tag_type = self.settings.category_tag_types.get(category)
            if tag_family and tag_type:
                status = self.GetElementName(tag_family) + " (" + self.GetElementName(tag_type) + ")"
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
        total_operations = sum(
            len(FilteredElementCollector(self.doc, view.Id).OfCategoryId(category.Id).WhereElementIsNotElementType().ToElements())
            for view in self.settings.selected_views for category in self.settings.selected_categories
        )
        self.progressBar.Maximum = total_operations or 1
        self.progressBar.Value = 0

        self.logger.add("Начало выполнения расстановки марок. Всего операций: {0}".format(total_operations))

        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        try:
            # Кэширование элементов
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
                        FilteredElementCollector(self.doc, view.Id).OfCategoryId(category.Id).WhereElementIsNotElementType().ToElements()
                    )
                    elements_by_view_and_category[view.Id][category.Id] = elements
                    self.logger.add("Категория '{0}': найдено элементов {1}".format(self.GetCategoryName(category), len(elements)))

            for view in self.settings.selected_views:
                if not isinstance(view, View3D):
                    continue
                for category in self.settings.selected_categories:
                    elements = elements_by_view_and_category[view.Id][category.Id]
                    tag_type = self.settings.category_tag_types.get(category)

                    if not tag_type:
                        error_msg = "Нет марки для категории '{0}'".format(self.GetCategoryName(category))
                        errors.append(error_msg)
                        self.logger.add(error_msg)
                        continue

                    for element in elements:
                        self.progressBar.Value = min(self.progressBar.Value + 1, self.progressBar.Maximum)
                        if self.HasExistingTag(element, view):
                            self.logger.add("Элемент {0} уже имеет марку на виде {1}, пропущен".format(element.Id, view.Name))
                            continue
                        if self.CreateTag(element, view, tag_type):
                            success_count += 1
                            self.logger.add("Марка создана для элемента {0} на виде {1}".format(element.Id, view.Name))
                        else:
                            error_msg = "Не удалось создать марку для элемента {0} на виде {1}".format(element.Id, view.Name)
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

        result_msg = "Успешно расставлено марок: {0}".format(success_count)
        if errors:
            result_msg += "\n\nОшибки ({0}):\n".format(len(errors)) + "\n".join(errors[:10])
            if len(errors) > 10:
                result_msg += "\n... и еще {0} ошибок".format(len(errors)-10)

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

            tag_point = XYZ(center.X + offset_x * direction_x, center.Y + offset_y * direction_y, center.Z)

            tag = IndependentTag.Create(self.doc, view.Id, Reference(element), self.settings.use_leader,
                                      TagMode.TM_ADDBY_CATEGORY, self.settings.orientation, tag_point)

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
            for tag in FilteredElementCollector(self.doc, view.Id).OfClass(IndependentTag).ToElements():
                try:
                    for tagged_elem in tag.GetTaggedLocalElements():
                        if tagged_elem.Id == element.Id:
                            return True
                except Exception as e:
                    self.logger.add("Ошибка проверки марки {0}: {1}".format(tag.Id, e))
                    if hasattr(tag, 'TaggedLocalElementId') and tag.TaggedLocalElementId == element.Id:
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
            Label(Text="Выберите семейство марки:", Location=Point(10, 10), Size=Size(250, 20)),
            ListBox(Location=Point(10, 40), Size=Size(250, 350)),
            Label(Text="Выберите типоразмер:", Location=Point(270, 10), Size=Size(250, 20)),
            ListBox(Location=Point(270, 40), Size=Size(500, 350)),
            Button(Text="OK", Location=Point(400, 400), Size=Size(75, 25)),
            Button(Text="Отмена", Location=Point(485, 400), Size=Size(75, 25))
        ]

        self.lblFamilies, self.lstFamilies, self.lblTypes, self.lstTypes, self.btnOK, self.btnCancel = controls

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
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name

                if hasattr(element, 'Family') and element.Family:
                    family_name = element.Family.Name if hasattr(element.Family, 'Name') and element.Family.Name else ""

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
                if hasattr(element, 'Name') and element.Name:
                    name = element.Name
                    if name and not name.startswith('IronPython'):
                        return name
                return "Семейство " + str(element.Id.IntegerValue)

            # Общий случай
            if hasattr(element, 'Name') and element.Name:
                name = element.Name
                if name and not name.startswith('IronPython'):
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
                        display_name = symbol_name + status + " [ID:" + str(symbol.Id.IntegerValue) + "]"

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
