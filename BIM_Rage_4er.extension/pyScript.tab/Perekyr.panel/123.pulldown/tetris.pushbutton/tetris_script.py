# -*- coding: utf-8 -*-
import clr
import random
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import *
from System.Drawing import *
from System.Timers import Timer

class TetrisForm(Form):
    def __init__(self):
        self.Text = "Tetris in PyRevit"
        self.Size = Size(400, 550)  # Увеличили ширину для следующей фигуры
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.Black
        self.DoubleBuffered = True
        
        # Игровое поле 10x20
        self.grid_width = 10
        self.grid_height = 20
        self.cell_size = 20
        self.grid = [[0 for _ in range(self.grid_width)] for _ in range(self.grid_height)]
        
        # Фигуры тетриса
        self.shapes = [
            [[1, 1, 1, 1]],  # I
            [[1, 1], [1, 1]],  # O
            [[1, 1, 1], [0, 1, 0]],  # T
            [[1, 1, 1], [1, 0, 0]],  # L
            [[1, 1, 1], [0, 0, 1]],  # J
            [[1, 1, 0], [0, 1, 1]],  # Z
            [[0, 1, 1], [1, 1, 0]]   # S
        ]
        
        self.colors = [
            Color.Cyan,     # I
            Color.Yellow,   # O  
            Color.Purple,   # T
            Color.Orange,   # L
            Color.Blue,     # J
            Color.Red,      # Z
            Color.Green     # S
        ]
        
        self.current_piece = None
        self.current_color = None
        self.next_piece = None
        self.next_color = None
        self.current_x = 0
        self.current_y = 0
        self.score = 0
        self.level = 1
        self.game_over = False
        
        # Таймер для падения
        self.timer = Timer(1000)
        self.timer.Elapsed += self.game_tick
        
        # Управление
        self.KeyPreview = True
        self.KeyDown += self.on_key_down
        
        # Панель управления
        self.control_panel = Panel()
        self.control_panel.Location = Point(0, 0)
        self.control_panel.Size = Size(400, 30)
        self.control_panel.BackColor = Color.DarkGray
        
        # Статус
        self.status_label = Label()
        self.status_label.Text = "Счет: 0 | Ур: 1 | Скорость: 1x"
        self.status_label.ForeColor = Color.White
        self.status_label.Location = Point(10, 5)
        self.status_label.Size = Size(200, 20)
        self.control_panel.Controls.Add(self.status_label)
        
        # Кнопка Новая игра - ВОЗВРАЩАЕМ!
        self.new_game_button = Button()
        self.new_game_button.Text = "Новая игра"
        self.new_game_button.Location = Point(300, 3)
        self.new_game_button.Size = Size(80, 24)
        self.new_game_button.BackColor = Color.LightGreen
        self.new_game_button.Font = Font("Arial", 8)
        self.new_game_button.Click += self.on_new_game_click
        self.new_game_button.TabStop = False
        self.control_panel.Controls.Add(self.new_game_button)
        
        self.Controls.Add(self.control_panel)
        
        self.Shown += self.on_form_shown
        self.new_game()
    
    def on_form_shown(self, sender, e):
        self.Focus()
    
    def on_new_game_click(self, sender, e):
        self.new_game()
        self.Focus()
    
    def new_game(self):
        if self.timer.Enabled:
            self.timer.Stop()
            
        self.grid = [[0 for _ in range(self.grid_width)] for _ in range(self.grid_height)]
        self.score = 0
        self.level = 1
        self.game_over = False
        self.timer.Interval = 1000
        
        # Генерируем первую и следующую фигуры
        self.next_piece = None
        self.next_color = None
        self.spawn_piece()
        
        self.update_status()
        self.timer.Start()
        self.Invalidate()
        self.Focus()
    
    def spawn_piece(self):
        # Если есть следующая фигура - используем ее
        if self.next_piece is not None:
            self.current_piece = self.next_piece
            self.current_color = self.next_color
        else:
            # Иначе генерируем новую
            shape_index = random.randint(0, len(self.shapes) - 1)
            self.current_piece = self.shapes[shape_index]
            self.current_color = self.colors[shape_index]
        
        # Генерируем следующую фигуру
        next_shape_index = random.randint(0, len(self.shapes) - 1)
        self.next_piece = self.shapes[next_shape_index]
        self.next_color = self.colors[next_shape_index]
        
        self.current_x = self.grid_width // 2 - len(self.current_piece[0]) // 2
        self.current_y = 0
        
        if not self.is_valid_position():
            self.game_over = True
            self.timer.Stop()
    
    def is_valid_position(self, piece=None, x=None, y=None):
        if piece is None:
            piece = self.current_piece
        if x is None:
            x = self.current_x
        if y is None:
            y = self.current_y
            
        for row in range(len(piece)):
            for col in range(len(piece[0])):
                if piece[row][col]:
                    new_x = x + col
                    new_y = y + row
                    
                    if (new_x < 0 or new_x >= self.grid_width or 
                        new_y >= self.grid_height or 
                        (new_y >= 0 and self.grid[new_y][new_x])):
                        return False
        return True
    
    def merge_piece(self):
        for row in range(len(self.current_piece)):
            for col in range(len(self.current_piece[0])):
                if self.current_piece[row][col]:
                    actual_y = self.current_y + row
                    if actual_y >= 0:
                        self.grid[actual_y][self.current_x + col] = self.current_color
    
    def clear_lines(self):
        lines_cleared = 0
        for row in range(self.grid_height):
            if all(self.grid[row]):
                for r in range(row, 0, -1):
                    self.grid[r] = self.grid[r-1][:]
                self.grid[0] = [0] * self.grid_width
                lines_cleared += 1
        
        if lines_cleared > 0:
            base_points = lines_cleared * 100
            level_bonus = self.level * 50
            self.score += base_points + level_bonus
            
            new_level = (self.score // 500) + 1
            if new_level > self.level:
                self.level = new_level
                self.update_speed()
            
            self.update_status()
    
    def update_speed(self):
        speed_settings = {
            1: 1000,   # Уровень 1: 1x
            2: 800,    # Уровень 2: 1.25x  
            3: 650,    # Уровень 3: 1.5x
            4: 500,    # Уровень 4: 2x
            5: 400,    # Уровень 5: 2.5x
            6: 300,    # Уровень 6: 3.3x
            7: 250,    # Уровень 7: 4x
            8: 200,    # Уровень 8: 5x
            9: 150,    # Уровень 9: 6.7x
            10: 100    # Уровень 10: 10x
        }
        
        new_interval = speed_settings.get(self.level, 100)
        self.timer.Interval = new_interval
    
    def rotate_piece(self):
        rotated = list(zip(*reversed(self.current_piece)))
        rotated = [list(row) for row in rotated]
        
        if self.is_valid_position(rotated, self.current_x, self.current_y):
            self.current_piece = rotated
            self.Invalidate()
    
    def game_tick(self, sender, e):
        if not self.game_over:
            if self.is_valid_position(y=self.current_y + 1):
                self.current_y += 1
            else:
                self.merge_piece()
                self.clear_lines()
                self.spawn_piece()
            
            self.Invoke(System.Action(self.Invalidate))
    
    def on_key_down(self, sender, e):
        if e.KeyCode == Keys.Space and self.new_game_button.Focused:
            return
            
        if self.game_over:
            if e.KeyCode == Keys.N or e.KeyCode == Keys.Enter:
                self.new_game()
            return
            
        if e.KeyCode == Keys.Left and self.is_valid_position(x=self.current_x - 1):
            self.current_x -= 1
        elif e.KeyCode == Keys.Right and self.is_valid_position(x=self.current_x + 1):
            self.current_x += 1
        elif e.KeyCode == Keys.Down and self.is_valid_position(y=self.current_y + 1):
            self.current_y += 1
        elif e.KeyCode == Keys.Up:
            self.rotate_piece()
        elif e.KeyCode == Keys.Space:
            while self.is_valid_position(y=self.current_y + 1):
                self.current_y += 1
            self.merge_piece()
            self.clear_lines()
            self.spawn_piece()
        elif e.KeyCode == Keys.N:
            self.new_game()
            
        self.Invalidate()
    
    def update_status(self):
        speed_multiplier = int(1000 / self.timer.Interval)
        self.status_label.Text = "Счет: {} | Ур: {} | Скорость: {}x".format(
            self.score, self.level, speed_multiplier)
    
    def OnPaint(self, e):
        if hasattr(self, 'grid'):
            g = e.Graphics
            g.Clear(Color.Black)
            
            # Рисуем основное поле
            for y in range(self.grid_height):
                for x in range(self.grid_width):
                    rect = Rectangle(x * self.cell_size + 2, 
                                   y * self.cell_size + 35,
                                   self.cell_size - 2, 
                                   self.cell_size - 2)
                    
                    if self.grid[y][x]:
                        g.FillRectangle(SolidBrush(self.grid[y][x]), rect)
                    else:
                        g.DrawRectangle(Pens.Gray, rect)
            
            # Рисуем текущую фигуру
            if self.current_piece and not self.game_over:
                for row in range(len(self.current_piece)):
                    for col in range(len(self.current_piece[0])):
                        if self.current_piece[row][col]:
                            rect = Rectangle((self.current_x + col) * self.cell_size + 2,
                                           (self.current_y + row) * self.cell_size + 35,
                                           self.cell_size - 2,
                                           self.cell_size - 2)
                            g.FillRectangle(SolidBrush(self.current_color), rect)
            
            # Рисуем следующую фигуру (справа от основного поля)
            if self.next_piece and not self.game_over:
                # Заголовок "Следующая:"
                g.DrawString("Следующая:", Font("Arial", 10, FontStyle.Bold), Brushes.White, 220, 40)
                
                # Центрируем следующую фигуру в области предпросмотра
                preview_x = 240  # Позиция X для предпросмотра
                preview_y = 70   # Позиция Y для предпросмотра
                
                for row in range(len(self.next_piece)):
                    for col in range(len(self.next_piece[0])):
                        if self.next_piece[row][col]:
                            rect = Rectangle(preview_x + col * self.cell_size,
                                           preview_y + row * self.cell_size,
                                           self.cell_size - 2,
                                           self.cell_size - 2)
                            g.FillRectangle(SolidBrush(self.next_color), rect)
                            g.DrawRectangle(Pens.White, rect)
            
            # Game over сообщение
            if self.game_over:
                font = Font("Arial", 16, FontStyle.Bold)
                g.DrawString("GAME OVER", font, Brushes.Red, 120, 200)
                g.DrawString("Счет: " + str(self.score), font, Brushes.White, 140, 230)
                g.DrawString("Уровень: " + str(self.level), font, Brushes.Yellow, 130, 260)
                g.DrawString("N или кнопка - новая игра", Font("Arial", 9), Brushes.White, 100, 290)

# Функции для pyRevit
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True

def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        form = TetrisForm()
        form.ShowDialog()
        return True
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))
        return False

if __name__ == "__main__":
    Application.Run(TetrisForm())