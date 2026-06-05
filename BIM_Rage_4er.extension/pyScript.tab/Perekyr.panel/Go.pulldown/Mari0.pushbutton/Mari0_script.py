# -*- coding: utf-8 -*-
"""
Mario Style Platformer Game для pyRevit
Управление: стрелки влево/вправо, пробел - прыжок
Собирай монеты и избегай врагов!
"""
import clr
import random
import math

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Form, Timer, Label, Panel
from System.Windows.Forms import FormStartPosition, Padding, DockStyle, Keys
from System.Drawing import Color, SolidBrush, Rectangle, Size, Point, Font, FontStyle, Graphics, ContentAlignment
from System.Drawing.Drawing2D import LinearGradientBrush, LinearGradientMode, SmoothingMode
from System.Drawing import Pen as SystemPen
from System import Array

class PlatformerGame(Form):
    def __init__(self):
        self.Text = "Mario Style Platformer"
        self.Width = 1000
        self.Height = 600
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(135, 206, 235)
        self.DoubleBuffered = True
        self.KeyPreview = True
        
        self.gravity = 0.5
        self.jump_power = -10
        self.player_speed = 5
        self.score = 0
        self.game_over = False
        self.game_won = False
        
        self.player = Player(100, 400, 30, 30)
        
        self.platforms = []
        self.coins = []
        self.enemies = []
        
        self.camera_x = 0
        
        self.generate_level()
        
        self.left_pressed = False
        self.right_pressed = False
        self.jump_pressed = False
        
        self.timer = Timer()
        self.timer.Interval = 16
        self.timer.Tick += self.update_game
        
        self.setup_ui()
        
        self.KeyDown += self.on_key_down
        self.KeyUp += self.on_key_up
        self.Paint += self.on_paint
        self.Resize += self.on_resize
        
        self.timer.Start()
    
    def generate_level(self):
        # Земля
        self.platforms.append(Platform(0, 550, 2000, 30, Color.SaddleBrown))
        
        # Платформы
        platforms_data = [
            (200, 500, 100, 20), (350, 450, 100, 20), (550, 400, 100, 20),
            (750, 350, 100, 20), (900, 300, 100, 20), (1100, 400, 100, 20),
            (1300, 350, 100, 20), (1500, 300, 100, 20), (1650, 250, 100, 20),
            (1800, 200, 100, 20), (1950, 300, 100, 20), (2100, 400, 100, 20),
            (2250, 350, 100, 20), (2400, 300, 150, 20)
        ]
        
        for x, y, w, h in platforms_data:
            self.platforms.append(Platform(x, y, w, h, Color.Peru))
        
        # Монеты
        coins_positions = [
            (250, 470), (400, 420), (600, 370), (800, 320), (950, 270),
            (1150, 370), (1350, 320), (1550, 270), (1700, 220), (1850, 170),
            (2000, 270), (2150, 370), (2300, 320), (2450, 270), (2500, 250)
        ]
        
        for x, y in coins_positions:
            self.coins.append(Coin(x, y, 12, 12))  # Добавлены width и height
        
        # Враги
        enemies_positions = [(300, 520), (600, 370), (1000, 270), (1400, 320), (2000, 270)]
        for x, y in enemies_positions:
            self.enemies.append(Enemy(x, y - 20, 25, 25, 2))
        
        # Финиш
        self.flag = Flag(2550, 470, 30, 70)
    
    def setup_ui(self):
        control_panel = Panel()
        control_panel.BackColor = Color.FromArgb(50, 50, 70)
        control_panel.Height = 40
        control_panel.Dock = DockStyle.Top
        control_panel.Padding = Padding(5)
        self.Controls.Add(control_panel)
        
        self.lbl_score = Label()
        self.lbl_score.Text = "Score: 0"
        self.lbl_score.Size = Size(150, 30)
        self.lbl_score.Location = Point(10, 5)
        self.lbl_score.ForeColor = Color.Gold
        self.lbl_score.Font = Font("Segoe UI", 12, FontStyle.Bold)
        control_panel.Controls.Add(self.lbl_score)
        
        info = Label()
        info.Text = "Controls: LEFT / RIGHT / SPACE"
        info.Size = Size(250, 30)
        info.Location = Point(200, 5)
        info.ForeColor = Color.White
        info.Font = Font("Segoe UI", 10)
        control_panel.Controls.Add(info)
        
        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Size = Size(300, 30)
        self.lbl_status.Location = Point(650, 5)
        self.lbl_status.ForeColor = Color.Red
        self.lbl_status.Font = Font("Segoe UI", 12, FontStyle.Bold)
        control_panel.Controls.Add(self.lbl_status)
    
    def update_game(self, sender, e):
        if self.game_over or self.game_won:
            return
        
        # Горизонтальное движение
        if self.left_pressed and self.player.x > 0:
            self.player.vx = -self.player_speed
        elif self.right_pressed:
            self.player.vx = self.player_speed
        else:
            self.player.vx *= 0.9
        
        # Прыжок
        if self.jump_pressed and self.player.on_ground:
            self.player.vy = self.jump_power
            self.player.on_ground = False
        
        # Обновление физики игрока
        self.player.update(self.gravity)
        
        # Проверка столкновений с платформами
        self.player.on_ground = False
        for platform in self.platforms:
            if self.check_collision(self.player, platform):
                self.handle_platform_collision(self.player, platform)
        
        # Проверка столкновений с монетами
        for coin in self.coins[:]:
            if self.check_collision(self.player, coin):
                self.coins.remove(coin)
                self.score += 10
                self.lbl_score.Text = "Score: " + str(self.score)
        
        # Проверка столкновений с врагами
        for enemy in self.enemies[:]:
            if self.check_collision(self.player, enemy):
                if self.player.vy > 0 and self.player.y + self.player.height - enemy.y < 20:
                    # Прыжок на врага
                    self.enemies.remove(enemy)
                    self.player.vy = self.jump_power / 2
                    self.score += 50
                    self.lbl_score.Text = "Score: " + str(self.score)
                else:
                    # Столкновение с врагом сбоку
                    self.game_over = True
                    self.timer.Stop()
                    self.lbl_status.Text = "GAME OVER! Press R to restart"
                    self.lbl_status.ForeColor = Color.Red
        
        # Обновление врагов
        for enemy in self.enemies:
            enemy.update()
            for platform in self.platforms:
                if self.check_collision(enemy, platform):
                    if enemy.vy == 0:
                        enemy.y = platform.y - enemy.height
                        if enemy.x + enemy.width > platform.x and enemy.x < platform.x + platform.width:
                            enemy.vy = 0
                            enemy.on_ground = True
        
        # Обновление камеры
        self.camera_x = self.player.x - 400
        if self.camera_x < 0:
            self.camera_x = 0
        if self.camera_x > 2000:
            self.camera_x = 2000
        
        # Проверка финиша
        if self.check_collision(self.player, self.flag):
            self.game_won = True
            self.timer.Stop()
            self.lbl_status.Text = "YOU WIN! Score: " + str(self.score)
            self.lbl_status.ForeColor = Color.Green
        
        # Проверка падения в пропасть
        if self.player.y > 600:
            self.game_over = True
            self.timer.Stop()
            self.lbl_status.Text = "GAME OVER! Press R to restart"
            self.lbl_status.ForeColor = Color.Red
        
        self.Refresh()
    
    def check_collision(self, obj1, obj2):
        """Проверка коллизии двух прямоугольников"""
        return (obj1.x < obj2.x + obj2.width and
                obj1.x + obj1.width > obj2.x and
                obj1.y < obj2.y + obj2.height and
                obj1.y + obj1.height > obj2.y)
    
    def handle_platform_collision(self, player, platform):
        """Обработка столкновения с платформой"""
        overlap_left = (player.x + player.width) - platform.x
        overlap_right = (platform.x + platform.width) - player.x
        overlap_top = (player.y + player.height) - platform.y
        overlap_bottom = (platform.y + platform.height) - player.y
        
        min_overlap = min(overlap_left, overlap_right, overlap_top, overlap_bottom)
        
        if min_overlap == overlap_top and player.vy >= 0:
            # Столкновение сверху
            player.y = platform.y - player.height
            player.vy = 0
            player.on_ground = True
        elif min_overlap == overlap_bottom and player.vy < 0:
            # Столкновение снизу
            player.y = platform.y + platform.height
            player.vy = 0
        elif min_overlap == overlap_left:
            # Столкновение слева
            player.x = platform.x - player.width
        elif min_overlap == overlap_right:
            # Столкновение справа
            player.x = platform.x + platform.width
    
    def on_paint(self, sender, e):
        g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        
        # Небо с градиентом
        rect = Rectangle(0, 0, self.Width, self.Height)
        gradient = LinearGradientBrush(rect, Color.FromArgb(135, 206, 235), 
                                       Color.FromArgb(100, 150, 200), 
                                       LinearGradientMode.Vertical)
        g.FillRectangle(gradient, rect)
        
        # Облака
        self.draw_cloud(g, 100 - self.camera_x, 80)
        self.draw_cloud(g, 400 - self.camera_x, 120)
        self.draw_cloud(g, 700 - self.camera_x, 60)
        self.draw_cloud(g, 1000 - self.camera_x, 100)
        self.draw_cloud(g, 1500 - self.camera_x, 70)
        self.draw_cloud(g, 2000 - self.camera_x, 130)
        
        # Платформы
        for platform in self.platforms:
            self.draw_platform(g, platform)
        
        # Монеты
        for coin in self.coins:
            self.draw_coin(g, coin)
        
        # Враги
        for enemy in self.enemies:
            self.draw_enemy(g, enemy)
        
        # Финишный флаг
        self.draw_flag(g, self.flag)
        
        # Игрок
        self.draw_player(g, self.player)
    
    def draw_cloud(self, g, x, y):
        """Рисует облако"""
        cloud_brush = SolidBrush(Color.FromArgb(255, 255, 255))
        g.FillEllipse(cloud_brush, x, y, 40, 30)
        g.FillEllipse(cloud_brush, x + 20, y - 10, 50, 35)
        g.FillEllipse(cloud_brush, x + 50, y, 40, 30)
    
    def draw_platform(self, g, platform):
        """Рисует платформу"""
        brush = SolidBrush(platform.color)
        x = platform.x - int(self.camera_x)
        y = platform.y
        g.FillRectangle(brush, x, y, platform.width, platform.height)
        
        # Текстура
        pen = SystemPen(Color.FromArgb(100, 80, 50), 1)
        for i in range(0, platform.width, 20):
            g.DrawLine(pen, x + i, y, x + i, y + 5)
    
    def draw_coin(self, g, coin):
        """Рисует монету"""
        x = coin.x - int(self.camera_x)
        y = coin.y
        coin_brush = SolidBrush(Color.Gold)
        g.FillEllipse(coin_brush, x, y, coin.width, coin.height)
        
        # Блик
        highlight_brush = SolidBrush(Color.FromArgb(150, 255, 255, 200))
        g.FillEllipse(highlight_brush, x + 3, y + 3, 4, 4)
    
    def draw_enemy(self, g, enemy):
        """Рисует врага"""
        x = enemy.x - int(self.camera_x)
        y = enemy.y
        brush = SolidBrush(Color.Purple)
        g.FillRectangle(brush, x, y, enemy.width, enemy.height)
        
        # Глаза
        eye_brush = SolidBrush(Color.White)
        pupil_brush = SolidBrush(Color.Black)
        g.FillEllipse(eye_brush, x + 5, y + 5, 6, 6)
        g.FillEllipse(eye_brush, x + 14, y + 5, 6, 6)
        g.FillEllipse(pupil_brush, x + 7, y + 7, 3, 3)
        g.FillEllipse(pupil_brush, x + 16, y + 7, 3, 3)
    
    def draw_flag(self, g, flag):
        """Рисует финишный флаг"""
        x = flag.x - int(self.camera_x)
        y = flag.y
        
        # Древко
        pen = SystemPen(Color.Brown, 3)
        g.DrawLine(pen, x + 15, y, x + 15, y + flag.height)
        
        # Флаг
        flag_brush = SolidBrush(Color.Red)
        points = Array[Point]([
            Point(x + 15, y),
            Point(x + 35, y + 15),
            Point(x + 15, y + 30)
        ])
        g.FillPolygon(flag_brush, points)
    
    def draw_player(self, g, player):
        """Рисует игрока (Марио)"""
        x = player.x - int(self.camera_x)
        y = player.y
        
        # Тело
        body_brush = SolidBrush(Color.FromArgb(255, 0, 0))
        g.FillRectangle(body_brush, x, y, player.width, player.height)
        
        # Голова
        head_brush = SolidBrush(Color.FromArgb(255, 200, 150))
        g.FillEllipse(head_brush, x + 5, y - 15, 20, 20)
        
        # Шляпа
        hat_brush = SolidBrush(Color.FromArgb(200, 0, 0))
        g.FillRectangle(hat_brush, x + 3, y - 20, 24, 8)
        
        # Усы
        moustache_brush = SolidBrush(Color.Brown)
        g.FillEllipse(moustache_brush, x + 8, y - 5, 6, 4)
        g.FillEllipse(moustache_brush, x + 16, y - 5, 6, 4)
        
        # Глаза
        eye_brush = SolidBrush(Color.White)
        pupil_brush = SolidBrush(Color.Black)
        g.FillEllipse(eye_brush, x + 10, y - 10, 4, 4)
        g.FillEllipse(eye_brush, x + 18, y - 10, 4, 4)
        g.FillEllipse(pupil_brush, x + 11, y - 9, 2, 2)
        g.FillEllipse(pupil_brush, x + 19, y - 9, 2, 2)
        
        # Комбинезон
        overall_brush = SolidBrush(Color.FromArgb(0, 0, 255))
        g.FillRectangle(overall_brush, x + 8, y + 5, 6, 15)
        g.FillRectangle(overall_brush, x + 16, y + 5, 6, 15)
    
    def on_key_down(self, sender, e):
        if e.KeyCode == Keys.Left:
            self.left_pressed = True
        elif e.KeyCode == Keys.Right:
            self.right_pressed = True
        elif e.KeyCode == Keys.Space:
            self.jump_pressed = True
        elif e.KeyCode == Keys.R:
            self.restart_game()
    
    def on_key_up(self, sender, e):
        if e.KeyCode == Keys.Left:
            self.left_pressed = False
        elif e.KeyCode == Keys.Right:
            self.right_pressed = False
        elif e.KeyCode == Keys.Space:
            self.jump_pressed = False
    
    def restart_game(self):
        """Перезапуск игры"""
        self.score = 0
        self.game_over = False
        self.game_won = False
        self.lbl_score.Text = "Score: 0"
        self.lbl_status.Text = ""
        
        self.player = Player(100, 400, 30, 30)
        self.platforms = []
        self.coins = []
        self.enemies = []
        
        self.generate_level()
        
        self.left_pressed = False
        self.right_pressed = False
        self.jump_pressed = False
        
        self.timer.Start()
        self.Refresh()
    
    def on_resize(self, sender, e):
        self.Refresh()
    
    def on_form_closed(self, sender, e):
        self.timer.Stop()

class Player:
    def __init__(self, x, y, width, height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.vx = 0
        self.vy = 0
        self.on_ground = False
    
    def update(self, gravity):
        self.x += self.vx
        self.vy += gravity
        self.y += self.vy

class Platform:
    def __init__(self, x, y, width, height, color):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.color = color

class Coin:
    def __init__(self, x, y, width, height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height

class Enemy:
    def __init__(self, x, y, width, height, speed):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.speed = speed
        self.direction = 1
        self.vy = 0
        self.on_ground = False
    
    def update(self):
        self.x += self.speed * self.direction
    
    def change_direction(self):
        self.direction *= -1

class Flag:
    def __init__(self, x, y, width, height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height

# Запуск игры
form = PlatformerGame()
form.Closed += form.on_form_closed
form.ShowDialog()