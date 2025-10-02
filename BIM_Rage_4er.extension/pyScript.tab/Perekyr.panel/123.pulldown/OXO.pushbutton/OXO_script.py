# -*- coding: utf-8 -*-
import clr
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import *
from System.Drawing import *
from System.Timers import Timer

class TicTacToeForm(Form):
    def __init__(self):
        self.Text = "Крестики-Нолики in PyRevit"
        self.Size = Size(400, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.LightBlue
        self.DoubleBuffered = True
        
        # Игровое поле 3x3
        self.board_size = 3
        self.cell_size = 80
        self.board = [['' for _ in range(3)] for _ in range(3)]
        
        # Игроки
        self.current_player = 'X'
        self.game_over = False
        self.winner = None
        self.moves_count = 0
        
        # Панель управления
        self.control_panel = Panel()
        self.control_panel.Location = Point(0, 0)
        self.control_panel.Size = Size(400, 40)
        self.control_panel.BackColor = Color.DarkBlue
        
        # Статус
        self.status_label = Label()
        self.status_label.Text = "Крестики-Нолики | Ход: X"
        self.status_label.ForeColor = Color.White
        self.status_label.Location = Point(10, 10)
        self.status_label.Size = Size(380, 20)
        self.status_label.Font = Font("Arial", 10, FontStyle.Bold)
        self.control_panel.Controls.Add(self.status_label)
        
        self.Controls.Add(self.control_panel)
        
        # Кнопка Новая игра
        self.new_game_button = Button()
        self.new_game_button.Text = "Новая игра"
        self.new_game_button.Location = Point(150, 400)
        self.new_game_button.Size = Size(100, 30)
        self.new_game_button.BackColor = Color.LightGreen
        self.new_game_button.Font = Font("Arial", 10)
        self.new_game_button.Click += self.on_new_game_click
        self.new_game_button.TabStop = False
        self.Controls.Add(self.new_game_button)
        
        # Обработчик кликов по полю
        self.MouseClick += self.on_board_click
        
        self.new_game()
    
    def on_new_game_click(self, sender, e):
        self.new_game()
    
    def new_game(self):
        self.board = [['' for _ in range(3)] for _ in range(3)]
        self.current_player = 'X'
        self.game_over = False
        self.winner = None
        self.moves_count = 0
        self.update_status()
        self.Invalidate()
    
    def on_board_click(self, sender, e):
        if self.game_over:
            return
            
        # Определяем ячейку по координатам клика
        offset_x = (self.ClientSize.Width - self.board_size * self.cell_size) // 2
        offset_y = 60
        
        # Проверяем попадание в игровое поле
        if (e.X < offset_x or e.X >= offset_x + self.board_size * self.cell_size or
            e.Y < offset_y or e.Y >= offset_y + self.board_size * self.cell_size):
            return
        
        # Вычисляем индексы ячейки
        col = (e.X - offset_x) // self.cell_size
        row = (e.Y - offset_y) // self.cell_size
        
        # Если ячейка пустая - делаем ход
        if self.board[row][col] == '':
            self.board[row][col] = self.current_player
            self.moves_count += 1
            
            # Проверяем победу
            if self.check_win(row, col):
                self.game_over = True
                self.winner = self.current_player
            # Проверяем ничью
            elif self.moves_count == 9:
                self.game_over = True
                self.winner = 'Draw'
            else:
                # Смена игрока
                self.current_player = 'O' if self.current_player == 'X' else 'X'
            
            self.update_status()
            self.Invalidate()
    
    def check_win(self, row, col):
        player = self.board[row][col]
        
        # Проверка строки
        if all(self.board[row][c] == player for c in range(3)):
            return True
        
        # Проверка столбца
        if all(self.board[r][col] == player for r in range(3)):
            return True
        
        # Проверка диагоналей
        if row == col and all(self.board[i][i] == player for i in range(3)):
            return True
        
        if row + col == 2 and all(self.board[i][2-i] == player for i in range(3)):
            return True
        
        return False
    
    def update_status(self):
        if self.game_over:
            if self.winner == 'Draw':
                status = "Ничья! | Нажмите 'Новая игра'"
            else:
                status = "Победил: {}! | Нажмите 'Новая игра'".format(self.winner)
        else:
            status = "Крестики-Нолики | Ход: {}".format(self.current_player)
        
        self.status_label.Text = status
    
    def OnPaint(self, e):
        g = e.Graphics
        g.Clear(Color.LightBlue)
        
        # Смещение для центрирования игрового поля
        offset_x = (self.ClientSize.Width - self.board_size * self.cell_size) // 2
        offset_y = 60
        
        # Рисуем сетку
        pen = Pen(Color.DarkBlue, 3)
        for i in range(4):
            # Вертикальные линии
            g.DrawLine(pen, 
                      offset_x + i * self.cell_size, offset_y,
                      offset_x + i * self.cell_size, offset_y + 3 * self.cell_size)
            # Горизонтальные линии
            g.DrawLine(pen,
                      offset_x, offset_y + i * self.cell_size,
                      offset_x + 3 * self.cell_size, offset_y + i * self.cell_size)
        
        # Рисуем крестики и нолики
        for row in range(3):
            for col in range(3):
                cell_x = offset_x + col * self.cell_size
                cell_y = offset_y + row * self.cell_size
                
                if self.board[row][col] == 'X':
                    self.draw_x(g, cell_x, cell_y)
                elif self.board[row][col] == 'O':
                    self.draw_o(g, cell_x, cell_y)
        
        # Подпись игры
        title_font = Font("Arial", 16, FontStyle.Bold)
        g.DrawString("КРЕСТИКИ-НОЛИКИ", title_font, Brushes.DarkBlue, 100, 350)
        
        # Инструкция
        instr_font = Font("Arial", 10)
        g.DrawString("Кликайте по ячейкам чтобы сделать ход", instr_font, Brushes.DarkBlue, 80, 380)
    
    def draw_x(self, g, x, y):
        pen = Pen(Color.Red, 4)
        margin = 15
        
        # Рисуем крестик
        g.DrawLine(pen, 
                  x + margin, y + margin,
                  x + self.cell_size - margin, y + self.cell_size - margin)
        g.DrawLine(pen,
                  x + self.cell_size - margin, y + margin,
                  x + margin, y + self.cell_size - margin)
    
    def draw_o(self, g, x, y):
        pen = Pen(Color.Green, 4)
        margin = 15
        diameter = self.cell_size - 2 * margin
        
        # Рисуем нолик
        g.DrawEllipse(pen,
                     x + margin, y + margin,
                     diameter, diameter)

# Функции для pyRevit
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True

def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        form = TicTacToeForm()
        form.ShowDialog()
        return True
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))
        return False

if __name__ == "__main__":
    Application.Run(TicTacToeForm())