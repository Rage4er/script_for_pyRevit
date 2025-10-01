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

# Настройки
class TagSettings(object):
    def __init__(self):
        self.selected_views = []
        self.selected_categories = []
        self.category_tag_families = {}
        self.category_tag_types = {}
        self.offset_x = 60
        self.offset_y = 30
        self.orientation = TagOrientation.Horizontal
        self.use_leader = True
        self.enable_logging = False

# Базовые функции
log_messages = []
enable_logging = False

def add_log(message):
    if not enable_logging: 
        return
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_message = "[{0}] {1}".format(timestamp, message)
    log_messages.append(log_message)

def show_logs():
    if not enable_logging or not log_messages:
        MessageBox.Show("Логирование отключено или нет записей", "Информация")
        return
    log_text = "\n".join(log_messages)
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

# Главная форма
class MainForm(Form):
    def __init__(self, doc, uidoc):
        global enable_logging
        self.doc = doc
        self.uidoc = uidoc
        self.settings = TagSettings()
        self.views_dict = {}
        self.category_mapping = {}
        
        self.InitializeComponent()
        self.Load3DViews()
    
    def InitializeComponent(self):
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
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control
    
    def SetupTab1(self, tab):
        controls = [
            self.CreateControl(Label, Text="Выберите 3D виды:", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(CheckedListBox, Location=Point(10, 40), Size=Size(600, 400), CheckOnClick=True),
            self.CreateControl(CheckBox, Text="Включить логирование", Location=Point(10, 450), Size=Size(200, 20)),
            self.CreateControl(Button, Text="Показать логи", Location=Point(220, 450), Size=Size(100, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.lblViews, self.lstViews, self.chkLogging, self.btnShowLogs, self.btnNext1 = controls
        self.btnShowLogs.Click += self.OnShowLogsClick
        self.btnNext1.Click += self.OnNext1Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab2(self, tab):
        controls = [
            self.CreateControl(Label, Text="Выберите категории элементов:", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(CheckedListBox, Location=Point(10, 40), Size=Size(600, 400), CheckOnClick=True),
            self.CreateControl(Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.lblCategories, self.lstCategories, self.btnBack1, self.btnNext2 = controls
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab3(self, tab):
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
            self.CreateControl(Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack2, self.btnNext3 = controls[2], controls[3]
        self.btnBack2.Click += self.OnBack2Click
        self.btnNext3.Click += self.OnNext3Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab4(self, tab):
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
            self.CreateControl(Button, Text="← Назад", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack3, self.btnNext4 = controls[4], controls[5]
        self.btnBack3.Click += self.OnBack3Click
        self.btnNext4.Click += self.OnNext4Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab5(self, tab):
        self.txtSummary = TextBox()
        self.txtSummary.Location = Point(10, 40)
        self.txtSummary.Size = Size(700, 400)
        self.txtSummary.Multiline = True
        self.txtSummary.ScrollBars = ScrollBars.Vertical
        self.txtSummary.ReadOnly = True
        
        controls = [
            self.CreateControl(Label, Text="Готово к выполнению:", Location=Point(10, 10), Size=Size(300, 20)),
            self.txtSummary,
            self.CreateControl(Button, Text="← Назад", Location=Point(340, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Выполнить", Location=Point(500, 450), Size=Size(150, 30))
        ]
        self.btnBack4, self.btnExecute = controls[2], controls[3]
        self.btnBack4.Click += self.OnBack4Click
        self.btnExecute.Click += self.OnExecuteClick
        for c in controls: 
            tab.Controls.Add(c)
    
    def Load3DViews(self):
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
        global enable_logging
        enable_logging = self.settings.enable_logging = self.chkLogging.Checked
        
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
            except: 
                pass
        
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
        except: 
            pass
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
        except: 
            pass
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
            add_log("GetElementName: Ошибка при получении имени элемента: " + str(e))
        
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
        success_count = 0
        errors = []
        
        trans = Transaction(self.doc, "Расстановка марок")
        trans.Start()
        try:
            for view in self.settings.selected_views:
                if not isinstance(view, View3D):
                    errors.append("Вид '" + view.Name + "' не 3D")
                    continue
                
                for category in self.settings.selected_categories:
                    elements = FilteredElementCollector(self.doc, view.Id).OfCategoryId(category.Id).WhereElementIsNotElementType().ToElements()
                    tag_type = self.settings.category_tag_types.get(category)
                    
                    if not tag_type:
                        errors.append("Нет марки для " + self.GetCategoryName(category))
                        continue
                    
                    for element in elements:
                        if element and not self.HasExistingTag(element, view) and self.CreateTag(element, view, tag_type):
                            success_count += 1
            
            trans.Commit()
        except Exception as e:
            trans.RollBack()
            errors.append("Ошибка: " + str(e))
        finally:
            trans.Dispose()
        
        result_msg = "Успешно расставлено марок: " + str(success_count)
        if errors:
            result_msg += "\n\nОшибки (" + str(len(errors)) + "):\n" + "\n".join(errors[:10])
            if len(errors) > 10: 
                result_msg += "\n... и еще " + str(len(errors)-10) + " ошибок"
        
        MessageBox.Show(result_msg, "Результат")
        if self.settings.enable_logging: 
            show_logs()
        self.Close()
    
    def CreateTag(self, element, view, tag_type):
        try:
            bbox = element.get_BoundingBox(view)
            if not bbox: 
                return False
            
            center = (bbox.Min + bbox.Max) / 2
            scale_factor = 100.0 / view.Scale
            
            offset_x = (self.settings.offset_x * scale_factor) / 304.8
            offset_y = (self.settings.offset_y * scale_factor) / 304.8
            
            direction_x = 1 if element.Id.IntegerValue % 2 == 0 else -1
            direction_y = 1 if element.Id.IntegerValue % 3 == 0 else -1
            
            tag_point = XYZ(center.X + offset_x * direction_x, center.Y + offset_y * direction_y, center.Z)
            
            tag = IndependentTag.Create(self.doc, view.Id, Reference(element), self.settings.use_leader, 
                                      TagMode.TM_ADDBY_CATEGORY, self.settings.orientation, tag_point)
            
            if tag and tag_type:
                try: 
                    tag.ChangeTypeId(tag_type.Id)
                except: 
                    pass
                return True
            return False
        except: 
            return False
    
    def HasExistingTag(self, element, view):
        try:
            for tag in FilteredElementCollector(self.doc, view.Id).OfClass(IndependentTag).ToElements():
                try:
                    for tagged_elem in tag.GetTaggedLocalElements():
                        if tagged_elem.Id == element.Id:
                            return True
                except:
                    if hasattr(tag, 'TaggedLocalElementId') and tag.TaggedLocalElementId == element.Id:
                        return True
            return False
        except: 
            return False
    
    def OnShowLogsClick(self, sender, args): 
        show_logs()

# Форма выбора семейства марки
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
            Label(Text="Выберите семейство марки:", Location=Point(10, 10), Size=Size(250, 20)),
            ListBox(Location=Point(10, 40), Size=Size(250, 350)),
            Label(Text="Выберите типоразмер:", Location=Point(270, 10), Size=Size(250, 20)),
            ListBox(Location=Point(270, 40), Size=Size(500, 350)),
            Button(Text="OK", Location=Point(400, 400), Size=Size(75, 25)),
            Button(Text="Отмена", Location=Point(485, 400), Size=Size(75, 25))
        ]
        
        self.lblFamilies, self.lstFamilies, self.lblTypes, self.lstTypes, self.btnOK, self.btnCancel = controls
        
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
                family_name = self.GetElementNameImproved(family)
                self.lstFamilies.Items.Add(family_name)
                self.family_dict[family_name] = family
        
        # Выбор текущего или первого семейства
        if self.lstFamilies.Items.Count > 0:
            self.lstFamilies.SelectedIndex = 0
    
    def GetElementNameImproved(self, element):
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

# Запуск
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