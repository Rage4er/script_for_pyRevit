# -*- coding: utf-8 -*-

__title__ = """Переименование
видов"""
__author__ = 'Rage'
__doc__ = '''Переименование выбранных видов, 
добавление суффикса, префикса, 
номера'''
__ver__ = "1.0"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
import sys
import datetime

# Глобальные переменные
enable_logging = False  # Логирование отключено по умолчанию
log_messages = []

# Логирование сообщений
def add_log(message):
    global enable_logging
    if not enable_logging:
        return
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_message = "[{0}] {1}".format(timestamp, message)
    log_messages.append(log_message)
    print(log_message)

# Функция для переименования видов
def rename_views(doc, selected_views, prefix, suffix, start_number, number_format, replace_current_name):
    transaction = Transaction(doc, "Переименование видов")
    transaction.Start()
    
    try:
        for i, view in enumerate(selected_views):
            current_name = view.Name
            
            # Форматирование номера (начинаем с start_number)
            if number_format == "1":
                number_str = str(start_number + i)
            elif number_format == "01":
                number_str = str(start_number + i).zfill(2)
            elif number_format == "001":
                number_str = str(start_number + i).zfill(3)
            else:
                number_str = str(start_number + i)
            
            # Формирование нового имени
            if replace_current_name:
                new_name = "{0}{1}{2}".format(prefix, number_str, suffix)
            else:
                new_name = "{0}{1}{2}{3}".format(prefix, current_name, number_str, suffix)
            
            add_log("Попытка переименовать вид: '{0}' -> '{1}'".format(current_name, new_name))
            view.Name = new_name
            add_log("Успешно переименован вид: '{0}' -> '{1}'".format(current_name, new_name))
        
        transaction.Commit()
        add_log("Транзакция завершена успешно")
        return True, "Успешно переименовано {0} видов".format(len(selected_views))
    
    except Exception as e:
        transaction.RollBack()
        add_log("Ошибка в транзакции: {0}".format(str(e)))
        return False, "Ошибка при переименовании: {0}".format(str(e))

# Функция для получения выбранных видов
def get_selected_views(uidoc):
    selected_views = []
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()
    
    add_log("Получено элементов из выделения: {0}".format(len(selected_ids)))
    
    for elem_id in selected_ids:
        elem = uidoc.Document.GetElement(elem_id)
        
        if elem is None:
            add_log("Элемент с ID {0} не найден".format(elem_id))
            continue
        
        # ПРОСТАЯ ПРОВЕРКА - если элемент имеет свойство Name и ViewType, то это вид
        if hasattr(elem, 'Name') and hasattr(elem, 'ViewType'):
            selected_views.append(elem)
            add_log("Добавлен вид: {0} (ID: {1})".format(elem.Name, elem.Id))
        else:
            add_log("Пропущен не-вид: {0}".format(elem.GetType().Name))
    
    return selected_views

# Основная форма
class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.selected_views = []
        self.InitializeComponent()
        
        # Получаем предварительно выбранные виды автоматически
        self.selected_views = get_selected_views(self.uidoc)
        self.UpdateViewsList()
    
    def InitializeComponent(self):
        self.Text = "Пакетное переименование видов"
        self.Size = Size(550, 700)  # Уменьшена ширина до 2/3
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True  # Окно всегда поверх других
        
        # Основные элементы управления
        self.SetupControls()
    
    def SetupControls(self):
        # Заголовок
        label_title = Label()
        label_title.Text = "Пакетное переименование видов"
        label_title.Location = Point(10, 10)
        label_title.Size = Size(400, 25)
        label_title.Font = Font("Arial", 12, FontStyle.Bold)
        
        # Группа выбора видов
        group_selection = GroupBox()
        group_selection.Text = "Выбор видов"
        group_selection.Location = Point(10, 40)
        group_selection.Size = Size(520, 150)
        
        # Кнопки выбора видов (более высокие)
        self.btnRefreshSelection = Button()
        self.btnRefreshSelection.Text = "Обновить выделение"
        self.btnRefreshSelection.Location = Point(10, 25)
        self.btnRefreshSelection.Size = Size(140, 30)
        self.btnRefreshSelection.Click += self.OnRefreshSelectionClick
        
        self.btnAddToSelection = Button()
        self.btnAddToSelection.Text = "Добавить к выделению"
        self.btnAddToSelection.Location = Point(160, 25)
        self.btnAddToSelection.Size = Size(140, 30)
        self.btnAddToSelection.Click += self.OnAddToSelectionClick
        
        # Список видов
        self.listViews = ListBox()
        self.listViews.Location = Point(10, 65)
        self.listViews.Size = Size(500, 80)
        
        # Счетчик выбранных видов
        self.lblSelectedCount = Label()
        self.lblSelectedCount.Text = "Выбрано видов: 0"
        self.lblSelectedCount.Location = Point(10, 150)
        self.lblSelectedCount.Size = Size(150, 20)
        
        # Добавляем элементы в группу
        group_selection.Controls.Add(self.btnRefreshSelection)
        group_selection.Controls.Add(self.btnAddToSelection)
        group_selection.Controls.Add(self.listViews)
        group_selection.Controls.Add(self.lblSelectedCount)
        
        # Группа настроек именования
        group_settings = GroupBox()
        group_settings.Text = "Настройки именования"
        group_settings.Location = Point(10, 200)
        group_settings.Size = Size(520, 150)
        
        # Настройки именования
        label_prefix = Label()
        label_prefix.Text = "Префикс:"
        label_prefix.Location = Point(10, 25)
        label_prefix.Size = Size(100, 20)
        
        self.txtPrefix = TextBox()
        self.txtPrefix.Location = Point(120, 25)
        self.txtPrefix.Size = Size(150, 20)
        
        label_suffix = Label()
        label_suffix.Text = "Суффикс:"
        label_suffix.Location = Point(10, 50)
        label_suffix.Size = Size(100, 20)
        
        self.txtSuffix = TextBox()
        self.txtSuffix.Location = Point(120, 50)
        self.txtSuffix.Size = Size(150, 20)
        
        label_start = Label()
        label_start.Text = "Начальный №:"
        label_start.Location = Point(10, 75)
        label_start.Size = Size(100, 20)
        
        self.numStartNumber = NumericUpDown()
        self.numStartNumber.Location = Point(120, 75)
        self.numStartNumber.Size = Size(80, 20)
        self.numStartNumber.Minimum = 0  # Теперь можно начинать с 0
        self.numStartNumber.Maximum = 1000
        self.numStartNumber.Value = 0   # По умолчанию начинаем с 0
        
        label_format = Label()
        label_format.Text = "Формат номера:"
        label_format.Location = Point(10, 100)
        label_format.Size = Size(100, 20)
        
        self.cmbNumberFormat = ComboBox()
        self.cmbNumberFormat.Location = Point(120, 100)
        self.cmbNumberFormat.Size = Size(80, 20)
        self.cmbNumberFormat.Items.Add("1")
        self.cmbNumberFormat.Items.Add("01")
        self.cmbNumberFormat.Items.Add("001")
        self.cmbNumberFormat.SelectedIndex = 0
        
        self.chkReplaceName = CheckBox()
        self.chkReplaceName.Text = "Заменить текущее имя"
        self.chkReplaceName.Location = Point(10, 125)
        self.chkReplaceName.Size = Size(200, 20)
        self.chkReplaceName.Checked = False
        
        # Кнопка обновления предпросмотра
        self.btnUpdatePreview = Button()
        self.btnUpdatePreview.Text = "Обновить предпросмотр"
        self.btnUpdatePreview.Location = Point(300, 25)
        self.btnUpdatePreview.Size = Size(150, 30)
        self.btnUpdatePreview.Click += self.OnUpdatePreviewClick
        
        # Добавляем элементы в группу настроек
        group_settings.Controls.Add(label_prefix)
        group_settings.Controls.Add(self.txtPrefix)
        group_settings.Controls.Add(label_suffix)
        group_settings.Controls.Add(self.txtSuffix)
        group_settings.Controls.Add(label_start)
        group_settings.Controls.Add(self.numStartNumber)
        group_settings.Controls.Add(label_format)
        group_settings.Controls.Add(self.cmbNumberFormat)
        group_settings.Controls.Add(self.chkReplaceName)
        group_settings.Controls.Add(self.btnUpdatePreview)
        
        # Группа предпросмотра
        group_preview = GroupBox()
        group_preview.Text = "Предпросмотр изменений"
        group_preview.Location = Point(10, 360)
        group_preview.Size = Size(520, 150)
        
        self.listPreview = ListBox()
        self.listPreview.Location = Point(10, 20)
        self.listPreview.Size = Size(500, 120)
        
        group_preview.Controls.Add(self.listPreview)
        
        # Группа выполнения
        group_execute = GroupBox()
        group_execute.Text = "Выполнение"
        group_execute.Location = Point(10, 520)
        group_execute.Size = Size(520, 130)
        
        # Кнопка выполнения (большая и зеленая)
        self.btnExecute = Button()
        self.btnExecute.Text = "ВЫПОЛНИТЬ ПЕРЕИМЕНОВАНИЕ"
        self.btnExecute.Location = Point(10, 25)
        self.btnExecute.Size = Size(200, 40)
        self.btnExecute.BackColor = Color.LightGreen
        self.btnExecute.Font = Font("Arial", 10, FontStyle.Bold)
        self.btnExecute.Click += self.OnExecuteClick
        
        # Поле результатов
        self.txtResults = TextBox()
        self.txtResults.Location = Point(220, 25)
        self.txtResults.Size = Size(290, 70)
        self.txtResults.Multiline = True
        self.txtResults.ReadOnly = True
        self.txtResults.ScrollBars = ScrollBars.Vertical
        
        # Кнопка закрытия
        self.btnClose = Button()
        self.btnClose.Text = "Закрыть"
        self.btnClose.Location = Point(420, 100)
        self.btnClose.Size = Size(90, 25)
        self.btnClose.Click += self.OnCloseClick
        
        group_execute.Controls.Add(self.btnExecute)
        group_execute.Controls.Add(self.txtResults)
        group_execute.Controls.Add(self.btnClose)
        
        # Добавляем все группы на форму
        self.Controls.Add(label_title)
        self.Controls.Add(group_selection)
        self.Controls.Add(group_settings)
        self.Controls.Add(group_preview)
        self.Controls.Add(group_execute)
    
    def UpdateViewsList(self):
        self.listViews.Items.Clear()
        for view in self.selected_views:
            self.listViews.Items.Add("{0} (ID: {1})".format(view.Name, view.Id))
        
        self.lblSelectedCount.Text = "Выбрано видов: {0}".format(len(self.selected_views))
    
    def OnRefreshSelectionClick(self, sender, args):
        try:
            self.selected_views = get_selected_views(self.uidoc)
            self.UpdateViewsList()
            
            if len(self.selected_views) == 0:
                MessageBox.Show("Не выбрано ни одного вида или выбраны не виды! Выберите виды в Revit и нажмите кнопку снова.", "Внимание")
        
        except Exception as e:
            error_msg = "Ошибка при обновлении выделения: {0}".format(str(e))
            MessageBox.Show(error_msg, "Ошибка")
    
    def OnAddToSelectionClick(self, sender, args):
        try:
            new_views = get_selected_views(self.uidoc)
            
            existing_ids = [v.Id for v in self.selected_views]
            for view in new_views:
                if view.Id not in existing_ids:
                    self.selected_views.append(view)
            
            self.UpdateViewsList()
            
        except Exception as e:
            error_msg = "Ошибка при добавлении к выделению: {0}".format(str(e))
            MessageBox.Show(error_msg, "Ошибка")
    
    def OnUpdatePreviewClick(self, sender, args):
        if not self.selected_views:
            MessageBox.Show("Сначала выберите виды", "Внимание")
            return
        
        prefix = self.txtPrefix.Text
        suffix = self.txtSuffix.Text
        start_number = int(self.numStartNumber.Value)
        number_format = self.cmbNumberFormat.SelectedItem.ToString()
        replace_current = self.chkReplaceName.Checked
        
        self.listPreview.Items.Clear()
        
        for i, view in enumerate(self.selected_views):
            current_name = view.Name
            
            # Нумерация начинается с start_number (может быть 0)
            if number_format == "1":
                number_str = str(start_number + i)
            elif number_format == "01":
                number_str = str(start_number + i).zfill(2)
            elif number_format == "001":
                number_str = str(start_number + i).zfill(3)
            else:
                number_str = str(start_number + i)
            
            if replace_current:
                new_name = "{0}{1}{2}".format(prefix, number_str, suffix)
            else:
                new_name = "{0}{1}{2}{3}".format(prefix, current_name, number_str, suffix)
            
            self.listPreview.Items.Add("{0} → {1}".format(current_name, new_name))
    
    def OnExecuteClick(self, sender, args):
        if not self.selected_views:
            MessageBox.Show("Не выбрано ни одного вида", "Ошибка")
            return
        
        prefix = self.txtPrefix.Text
        suffix = self.txtSuffix.Text
        start_number = int(self.numStartNumber.Value)
        number_format = self.cmbNumberFormat.SelectedItem.ToString()
        replace_current = self.chkReplaceName.Checked
        
        self.txtResults.Text = "Начало переименования...\n"
        
        success, message = rename_views(self.doc, self.selected_views, prefix, suffix, start_number, number_format, replace_current)
        
        self.txtResults.Text += message + "\n"
        
        if success:
            self.txtResults.Text += "Переименование завершено успешно!"
            # Обновляем список после переименования
            self.UpdateViewsList()
        else:
            self.txtResults.Text += "Произошла ошибка при переименовании."
    
    def OnCloseClick(self, sender, args):
        self.Close()

def main():
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        
        if doc and uidoc:
            form = MainForm(doc, uidoc)
            Application.Run(form)
        else:
            MessageBox.Show("Нет доступа к документу Revit")
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))

if __name__ == "__main__":
    main()