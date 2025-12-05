# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit.DB import Document
from System.Windows.Forms import *
from System.Drawing import *

# Глобальные переменные
enable_logging = False
log_messages = []

# Логирование сообщений
def add_log(message):
    global enable_logging
    if not enable_logging: 
        return
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_message = "[{0}] {1}".format(timestamp, message)
    log_messages.append(log_message)

# Показать логи
def show_logs():
    global log_messages
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

# Основная форма
class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "Рабочий интерфейс"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        
        tabs = ["1. Начальная страница", "2. Второй этап", "3. Третий этап", "4. Четвертый этап", "5. Завершение"]
        for index, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            if index == 0:
                self.SetupTab1(tab)
            elif index == 1:
                self.SetupTab2(tab)
            elif index == 2:
                self.SetupTab3(tab)
            elif index == 3:
                self.SetupTab4(tab)
            elif index == 4:
                self.SetupTab5(tab)
            self.tabControl.TabPages.Add(tab)
        
        self.Controls.Add(self.tabControl)
    
    def CreateControl(self, control_type, **kwargs):
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control
    
    def SetupTab1(self, tab):
        controls = [
            self.CreateControl(Label, Text="Первая вкладка", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnNext1 = controls[-1]
        self.btnNext1.Click += self.OnNext1Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab2(self, tab):
        controls = [
            self.CreateControl(Label, Text="Вторая вкладка", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Button, Text="Назад ←", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack1, self.btnNext2 = controls[1:]
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab3(self, tab):
        controls = [
            self.CreateControl(Label, Text="Третья вкладка", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Button, Text="Назад ←", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack2, self.btnNext3 = controls[1:]
        self.btnBack2.Click += self.OnBack2Click
        self.btnNext3.Click += self.OnNext3Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab4(self, tab):
        controls = [
            self.CreateControl(Label, Text="Четвёртая вкладка", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Button, Text="Назад ←", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Далее →", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack3, self.btnNext4 = controls[1:]
        self.btnBack3.Click += self.OnBack3Click
        self.btnNext4.Click += self.OnNext4Click
        for c in controls: 
            tab.Controls.Add(c)
    
    def SetupTab5(self, tab):
        controls = [
            self.CreateControl(Label, Text="Заключительная вкладка", Location=Point(10, 10), Size=Size(300, 20)),
            self.CreateControl(Button, Text="Назад ←", Location=Point(500, 450), Size=Size(80, 25)),
            self.CreateControl(Button, Text="Завершить", Location=Point(600, 450), Size=Size(80, 25))
        ]
        self.btnBack4, self.btnFinish = controls[1:]
        self.btnBack4.Click += self.OnBack4Click
        self.btnFinish.Click += self.OnFinishClick
        for c in controls: 
            tab.Controls.Add(c)
    
    # Переход между вкладками
    def OnNext1Click(self, sender, args): 
        self.tabControl.SelectedIndex = 1
    def OnNext2Click(self, sender, args): 
        self.tabControl.SelectedIndex = 2
    def OnNext3Click(self, sender, args): 
        self.tabControl.SelectedIndex = 3
    def OnNext4Click(self, sender, args): 
        self.tabControl.SelectedIndex = 4
    
    def OnBack1Click(self, sender, args): 
        self.tabControl.SelectedIndex = 0
    def OnBack2Click(self, sender, args): 
        self.tabControl.SelectedIndex = 1
    def OnBack3Click(self, sender, args): 
        self.tabControl.SelectedIndex = 2
    def OnBack4Click(self, sender, args): 
        self.tabControl.SelectedIndex = 3
    
    def OnFinishClick(self, sender, args): 
        MessageBox.Show("Работа завершена!")
        self.Close()

# Точка входа
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