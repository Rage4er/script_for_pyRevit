# -*- coding: utf-8 -*-
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import *
from System.Drawing import *

class MovingButtonForm(Form):
    def __init__(self):
        self.Text = "Moving Button Game"
        self.Size = Size(500, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.LightBlue
        
        self.click_count = 0
        self.button_y = 350  # Начальная позиция Y
        
        # Заголовок
        self.title_label = Label()
        self.title_label.Text = "Нажмите 'Да' 3 раза чтобы закрыть\nКнопка будет двигаться вверх!"
        self.title_label.Location = Point(50, 50)
        self.title_label.Size = Size(400, 40)
        self.title_label.Font = Font("Arial", 12, FontStyle.Bold)
        self.title_label.TextAlign = ContentAlignment.MiddleCenter
        self.Controls.Add(self.title_label)
        
        # Счетчик
        self.counter_label = Label()
        self.counter_label.Text = "Нажатий: 0/3"
        self.counter_label.Location = Point(200, 120)
        self.counter_label.Size = Size(100, 20)
        self.counter_label.Font = Font("Arial", 10)
        self.counter_label.TextAlign = ContentAlignment.MiddleCenter
        self.Controls.Add(self.counter_label)
        
        # Кнопка ДА
        self.yes_button = Button()
        self.yes_button.Text = "Да"
        self.yes_button.Location = Point(150, self.button_y)
        self.yes_button.Size = Size(100, 50)
        self.yes_button.Font = Font("Arial", 12)
        self.yes_button.BackColor = Color.LightGreen
        self.yes_button.Click += self.on_yes_click
        self.Controls.Add(self.yes_button)
        
        # Кнопка НЕТ
        self.no_button = Button()
        self.no_button.Text = "Нет"
        self.no_button.Location = Point(270, self.button_y)
        self.no_button.Size = Size(100, 50)
        self.no_button.Font = Font("Arial", 12)
        self.no_button.BackColor = Color.LightCoral
        self.no_button.Click += self.on_no_click
        self.Controls.Add(self.no_button)

    def on_yes_click(self, sender, e):
        self.click_count += 1
        self.button_y -= 100  # Перемещаем на 100 пикселей вверх
        
        # Обновляем позиции кнопок
        self.yes_button.Location = Point(150, self.button_y)
        self.no_button.Location = Point(270, self.button_y)
        
        # Обновляем счетчик
        self.counter_label.Text = "Нажатий: {}/3".format(self.click_count)
        
        # Принудительно обновляем форму
        self.Refresh()
        
        if self.click_count >= 3:
            MessageBox.Show("Поздравляем! Вы нажали 'Да' 3 раза!", "Победа!")
            self.Close()

    def on_no_click(self, sender, e):
        MessageBox.Show("Игра завершена", "Выход")
        self.Close()

# Функции для pyRevit
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True

def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        form = MovingButtonForm()
        form.ShowDialog()
        return True
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))
        return False

# Для тестирования
if __name__ == "__main__":
    Application.Run(MovingButtonForm())