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
        self.Size = Size(300, 500)
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
        self.current_x = 0
        self.current_y = 0
        self.score = 0
        self.game_over = False
        
        # Таймер для падения
        self.timer = Timer(1000)  # 1 секунда
        self.timer.Elapsed += self.game_tick
        self.timer.Start()
        
        # Управление
        self.KeyPreview = True
        self.KeyDown += self.on_key_down
        
        # Статус
        self.status_label = Label()
        self.status_label.Text = "Счет: 0"
        self.status_label.ForeColor = Color.White
        self.status_label.Location = Point(10, 10)
        self.status_label.Size = Size(200, 20)
        self.Controls.Add(self.status_label)
        
        self.new_game()
    
    def new_game(self):
        self.grid = [[0 for _ in range(self.grid_width)] for _ in range(self.grid_height)]
        self.score = 0
        self.game_over = False
        self.spawn_piece()
        self.update_status()
        self.Invalidate()
    
    def spawn_piece(self):
        shape_index = random.randint(0, len(self.shapes) - 1)
        self.current_piece = self.shapes[shape_index]
        self.current_color = self.colors[shape_index]
        self.current_x = self.grid_width // 2 - len(self.current_piece[0]) // 2
        self.current_y = 0
        
        # Проверка на game over
        if not self.is_valid_position():
            self.game_over = True
            self.timer.Stop()
            MessageBox.Show("Игра окончена! Счет: " + str(self.score))
    
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
                    if actual_y >= 0:  # Только если фигура внутри поля
                        self.grid[actual_y][self.current_x + col] = self.current_color
    
    def clear_lines(self):
        lines_cleared = 0
        for row in range(self.grid_height):
            if all(self.grid[row]):
                # Удаляем строку
                for r in range(row, 0, -1):
                    self.grid[r] = self.grid[r-1][:]
                self.grid[0] = [0] * self.grid_width
                lines_cleared += 1
        
        if lines_cleared > 0:
            self.score += lines_cleared * 100
            self.update_status()
    
    def rotate_piece(self):
        # Поворот матрицы
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
        if self.game_over:
            return
            
        if e.KeyCode == Keys.Left and self.is_valid_position(x=self.current_x - 1):
            self.current_x -= 1
        elif e.KeyCode == Keys.Right and self.is_valid_position(x=self.current_x + 1):
            self.current_x += 1
        elif e.KeyCode == Keys.Down and self.is_valid_position(y=self.current_y + 1):
            self.current_y += 1
        elif e.KeyCode == Keys.Up:
            self.rotate_piece()
        elif e.KeyCode == Keys.Space:  # Быстрое падение
            while self.is_valid_position(y=self.current_y + 1):
                self.current_y += 1
            self.merge_piece()
            self.clear_lines()
            self.spawn_piece()
        elif e.KeyCode == Keys.N:
            self.new_game()
            
        self.Invalidate()
    
    def update_status(self):
        self.status_label.Text = "Счет: {} | Управление: ← → ↓ ↑ Space N".format(self.score)
    
    def OnPaint(self, e):
        if hasattr(self, 'grid'):
            g = e.Graphics
            g.Clear(Color.Black)
            
            # Рисуем сетку
            for y in range(self.grid_height):
                for x in range(self.grid_width):
                    rect = Rectangle(x * self.cell_size + 2, 
                                   y * self.cell_size + 30, 
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
                                           (self.current_y + row) * self.cell_size + 30,
                                           self.cell_size - 2,
                                           self.cell_size - 2)
                            g.FillRectangle(SolidBrush(self.current_color), rect)
            
            # Game over сообщение
            if self.game_over:
                font = Font("Arial", 16, FontStyle.Bold)
                g.DrawString("GAME OVER", font, Brushes.Red, 80, 200)
                g.DrawString("Счет: " + str(self.score), font, Brushes.White, 90, 230)
                g.DrawString("N - Новая игра", Font("Arial", 10), Brushes.Yellow, 80, 260)

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

# Добавить в класс TetrisForm:

def increase_difficulty(self):
    # Увеличиваем скорость каждые 1000 очков
    new_interval = max(100, 1000 - (self.score // 1000) * 100)
    self.timer.Interval = new_interval

# Вызывать в clear_lines после увеличения счета

# Для тестирования
if __name__ == "__main__":
    Application.Run(TetrisForm())