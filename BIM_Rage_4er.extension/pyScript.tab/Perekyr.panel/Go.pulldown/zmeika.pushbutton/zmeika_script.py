# -*- coding: utf-8 -*-
import clr
import random
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import *
from System.Drawing import *
from System.Timers import Timer

class SnakeForm(Form):
    def __init__(self):
        self.Text = "Snake in PyRevit"
        self.Size = Size(400, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.Black
        self.DoubleBuffered = True
        
        # Игровое поле
        self.grid_size = 20
        self.cell_size = 20
        self.grid_width = 15
        self.grid_height = 18
        
        # Змейка
        self.snake = []
        self.direction = "RIGHT"
        self.next_direction = "RIGHT"
        self.score = 0
        self.game_over = False
        self.game_started = False
        
        # Еда
        self.food_pos = None
        
        # Таймер
        self.timer = Timer(200)  # Начальная скорость
        self.timer.Elapsed += self.game_tick
        
        # Управление
        self.KeyPreview = True
        self.KeyDown += self.on_key_down
        
        # Панель управления
        self.control_panel = Panel()
        self.control_panel.Location = Point(0, 0)
        self.control_panel.Size = Size(400, 30)
        self.control_panel.BackColor = Color.DarkGreen
        
        # Статус
        self.status_label = Label()
        self.status_label.Text = "Змейка | Счет: 0 | Нажмите ENTER для старта"
        self.status_label.ForeColor = Color.White
        self.status_label.Location = Point(10, 5)
        self.status_label.Size = Size(380, 20)
        self.control_panel.Controls.Add(self.status_label)
        
        # УБРАЛ КНОПКУ "НОВАЯ ИГРА"
        
        self.Controls.Add(self.control_panel)
        
        self.Shown += self.on_form_shown
        self.new_game()
    
    def on_form_shown(self, sender, e):
        self.Focus()
    
    def new_game(self):
        if self.timer.Enabled:
            self.timer.Stop()
            
        # Инициализация змейки
        self.snake = [
            (self.grid_width // 2, self.grid_height // 2),
            (self.grid_width // 2 - 1, self.grid_height // 2),
            (self.grid_width // 2 - 2, self.grid_height // 2)
        ]
        
        self.direction = "RIGHT"
        self.next_direction = "RIGHT"
        self.score = 0
        self.game_over = False
        self.game_started = False
        self.timer.Interval = 200
        
        self.spawn_food()
        self.update_status()
        self.Invalidate()
        self.Focus()
    
    def spawn_food(self):
        while True:
            x = random.randint(0, self.grid_width - 1)
            y = random.randint(0, self.grid_height - 1)
            if (x, y) not in self.snake:
                self.food_pos = (x, y)
                break
    
    def game_tick(self, sender, e):
        if not self.game_over and self.game_started:
            self.move_snake()
            self.Invoke(System.Action(self.Invalidate))
    
    def move_snake(self):
        if not self.game_started:
            return
            
        # Обновляем направление
        self.direction = self.next_direction
        
        # Получаем голову змейки
        head_x, head_y = self.snake[0]
        
        # Вычисляем новую позицию головы
        if self.direction == "UP":
            new_head = (head_x, head_y - 1)
        elif self.direction == "DOWN":
            new_head = (head_x, head_y + 1)
        elif self.direction == "LEFT":
            new_head = (head_x - 1, head_y)
        elif self.direction == "RIGHT":
            new_head = (head_x + 1, head_y)
        
        # Проверка столкновений со стенами
        if (new_head[0] < 0 or new_head[0] >= self.grid_width or 
            new_head[1] < 0 or new_head[1] >= self.grid_height):
            self.game_over = True
            self.timer.Stop()
            return
        
        # Проверка столкновения с собой
        if new_head in self.snake:
            self.game_over = True
            self.timer.Stop()
            return
        
        # Добавляем новую голову
        self.snake.insert(0, new_head)
        
        # Проверка съедания еды
        if new_head == self.food_pos:
            self.score += 10
            self.spawn_food()
            
            # Увеличиваем скорость каждые 50 очков
            if self.score % 50 == 0:
                new_speed = max(50, 200 - (self.score // 50) * 20)
                self.timer.Interval = new_speed
        else:
            # Удаляем хвост, если не съели еду
            self.snake.pop()
        
        self.update_status()
    
    def on_key_down(self, sender, e):
        if e.KeyCode == Keys.Enter and not self.game_started and not self.game_over:
            self.game_started = True
            self.timer.Start()
            self.update_status()
            return
            
        if self.game_over:
            if e.KeyCode == Keys.Enter or e.KeyCode == Keys.N:
                self.new_game()
            return
        
        if not self.game_started:
            return
            
        # Управление змейкой (без возможности разворота на 180 градусов)
        if e.KeyCode == Keys.Up and self.direction != "DOWN":
            self.next_direction = "UP"
        elif e.KeyCode == Keys.Down and self.direction != "UP":
            self.next_direction = "DOWN"
        elif e.KeyCode == Keys.Left and self.direction != "RIGHT":
            self.next_direction = "LEFT"
        elif e.KeyCode == Keys.Right and self.direction != "LEFT":
            self.next_direction = "RIGHT"
        elif e.KeyCode == Keys.N:
            self.new_game()
    
    def update_status(self):
        if not self.game_started and not self.game_over:
            status = "Змейка | Счет: {} | Нажмите ENTER для старта".format(self.score)
        elif self.game_over:
            speed = int(200 / self.timer.Interval) if self.timer.Interval > 0 else 1
            status = "Игра окончена! Счет: {} | Скорость: {}x | ENTER - новая игра".format(self.score, speed)
        else:
            speed = int(200 / self.timer.Interval)
            status = "Змейка | Счет: {} | Скорость: {}x | Управление: ← → ↑ ↓".format(self.score, speed)
        
        self.status_label.Text = status
    
    def OnPaint(self, e):
        g = e.Graphics
        g.Clear(Color.Black)
        
        # Смещение для центрирования игрового поля
        offset_x = (self.ClientSize.Width - self.grid_width * self.cell_size) // 2
        offset_y = 40
        
        # Рисуем сетку
        for y in range(self.grid_height):
            for x in range(self.grid_width):
                rect = Rectangle(offset_x + x * self.cell_size, 
                               offset_y + y * self.cell_size,
                               self.cell_size - 1, 
                               self.cell_size - 1)
                g.DrawRectangle(Pens.DarkGreen, rect)
        
        # Рисуем змейку
        for i, (x, y) in enumerate(self.snake):
            rect = Rectangle(offset_x + x * self.cell_size + 1,
                           offset_y + y * self.cell_size + 1,
                           self.cell_size - 3,
                           self.cell_size - 3)
            
            # Голова - другой цвет
            if i == 0:
                g.FillRectangle(SolidBrush(Color.Lime), rect)
            else:
                g.FillRectangle(SolidBrush(Color.Green), rect)
        
        # Рисуем еду
        if self.food_pos:
            x, y = self.food_pos
            rect = Rectangle(offset_x + x * self.cell_size + 1,
                           offset_y + y * self.cell_size + 1,
                           self.cell_size - 3,
                           self.cell_size - 3)
            g.FillRectangle(SolidBrush(Color.Red), rect)
        
        # Сообщения
        if not self.game_started and not self.game_over:
            font = Font("Arial", 14, FontStyle.Bold)
            g.DrawString("ЗМЕЙКА", font, Brushes.White, 150, 150)
            g.DrawString("Управление: ← → ↑ ↓", Font("Arial", 12), Brushes.Yellow, 120, 180)
            g.DrawString("ENTER - начать игру", Font("Arial", 12), Brushes.LightGreen, 120, 200)
            g.DrawString("N - новая игра", Font("Arial", 12), Brushes.LightBlue, 120, 220)
        
        elif self.game_over:
            font = Font("Arial", 16, FontStyle.Bold)
            g.DrawString("GAME OVER", font, Brushes.Red, 140, 150)
            g.DrawString("Счет: " + str(self.score), font, Brushes.White, 160, 180)
            g.DrawString("ENTER - новая игра", Font("Arial", 12), Brushes.Yellow, 120, 220)

# Функции для pyRevit
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True

def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        form = SnakeForm()
        form.ShowDialog()
        return True
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))
        return False

if __name__ == "__main__":
    Application.Run(SnakeForm())