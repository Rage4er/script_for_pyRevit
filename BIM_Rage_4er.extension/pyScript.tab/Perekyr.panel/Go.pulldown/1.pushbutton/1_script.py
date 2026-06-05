# -*- coding: utf-8 -*-
"""
Расширенная анимация в Windows Forms для pyRevit
Стильные закругленные кнопки с эффектами
"""
import clr
import random
import math

# Подключаем WinForms и Drawing
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Form, Timer, Label, Button, TrackBar, Panel
from System.Windows.Forms import FormStartPosition, Padding, DockStyle, Cursors, TickStyle
from System.Drawing import Color, SolidBrush, Rectangle, Size, Point, Font, FontStyle, Graphics, ContentAlignment
from System.Drawing.Drawing2D import LinearGradientBrush, LinearGradientMode, SmoothingMode, GraphicsPath
from System.Drawing import Pen as SystemPen

class RoundedButton(Button):
    """Кнопка с закругленными углами и градиентом"""
    def __init__(self):
        self._corner_radius = 10
        self._hover = False
        self._gradient_start = Color.FromArgb(70, 70, 90)
        self._gradient_end = Color.FromArgb(50, 50, 70)
        self.FlatStyle = 0
        self.FlatAppearance.BorderSize = 0
        self.BackColor = Color.Transparent
        self.ForeColor = Color.White
        self.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.Cursor = Cursors.Hand
        self.TextAlign = ContentAlignment.MiddleCenter
        self.MouseEnter += self.on_mouse_enter
        self.MouseLeave += self.on_mouse_leave
    
    def on_mouse_enter(self, sender, e):
        self._hover = True
        self._gradient_start = Color.FromArgb(100, 100, 120)
        self._gradient_end = Color.FromArgb(80, 80, 100)
        self.Invalidate()
    
    def on_mouse_leave(self, sender, e):
        self._hover = False
        self._gradient_start = Color.FromArgb(70, 70, 90)
        self._gradient_end = Color.FromArgb(50, 50, 70)
        self.Invalidate()
    
    def set_colors(self, start_color, end_color):
        self._gradient_start = start_color
        self._gradient_end = end_color
    
    def OnPaint(self, e):
        g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        
        rect = Rectangle(0, 0, self.Width, self.Height)
        
        path = GraphicsPath()
        radius = self._corner_radius
        path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90)
        path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90)
        path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90)
        path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90)
        path.CloseFigure()
        
        gradient = LinearGradientBrush(rect, self._gradient_start, self._gradient_end, LinearGradientMode.Vertical)
        g.FillPath(gradient, path)
        
        pen = SystemPen(Color.FromArgb(120, 120, 140), 1)
        g.DrawPath(pen, path)
        
        if self._hover:
            glow_pen = SystemPen(Color.FromArgb(80, 200, 200, 255), 2)
            g.DrawPath(glow_pen, path)
        
        text_size = g.MeasureString(self.Text, self.Font)
        x = (self.Width - text_size.Width) / 2
        y = (self.Height - text_size.Height) / 2
        text_brush = SolidBrush(self.ForeColor)
        g.DrawString(self.Text, self.Font, text_brush, x, y)

class AdvancedAnimationForm(Form):
    def __init__(self):
        self.Text = "Advanced Balls Animation"
        self.Width = 900
        self.Height = 700
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(20, 20, 40)
        self.DoubleBuffered = True
        
        self.gravity = 0.3
        self.balls = []
        self.particles = []
        self.score = 0
        self.collision_count = 0
        
        self.margin = 80
        self.bounds_min_x = self.margin
        self.bounds_min_y = self.margin
        
        self.timer = Timer()
        self.timer.Interval = 16
        self.timer.Tick += self.on_timer_tick
        
        self.is_animating = False
        self.show_trails = True
        self.background_angle = 0
        
        self.setup_controls()
        self.create_balls()
        
        self.MouseClick += self.on_form_click
        self.Paint += self.on_paint
        self.Resize += self.on_resize
        
        self.update_bounds()
    
    def setup_controls(self):
        control_panel = Panel()
        control_panel.BackColor = Color.FromArgb(30, 30, 50)
        control_panel.Height = 80
        control_panel.Dock = DockStyle.Bottom
        control_panel.Padding = Padding(10)
        self.Controls.Add(control_panel)
        
        # Кнопка Старт/Стоп
        self.btn_toggle = RoundedButton()
        self.btn_toggle.Text = "START"
        self.btn_toggle.Size = Size(100, 45)
        self.btn_toggle.Location = Point(10, 17)
        self.btn_toggle.set_colors(Color.FromArgb(0, 150, 0), Color.FromArgb(0, 100, 0))
        self.btn_toggle.Click += self.on_toggle_click
        control_panel.Controls.Add(self.btn_toggle)
        
        # Кнопка Добавить шарик
        btn_add = RoundedButton()
        btn_add.Text = "ADD BALL"
        btn_add.Size = Size(110, 45)
        btn_add.Location = Point(125, 17)
        btn_add.set_colors(Color.FromArgb(0, 100, 150), Color.FromArgb(0, 70, 110))
        btn_add.Click += self.on_add_ball
        control_panel.Controls.Add(btn_add)
        
        # Кнопка Сброс
        btn_reset = RoundedButton()
        btn_reset.Text = "RESET"
        btn_reset.Size = Size(100, 45)
        btn_reset.Location = Point(250, 17)
        btn_reset.set_colors(Color.FromArgb(150, 100, 0), Color.FromArgb(110, 70, 0))
        btn_reset.Click += self.on_reset
        control_panel.Controls.Add(btn_reset)
        
        # Кнопка Следы
        self.chk_trails = RoundedButton()
        self.chk_trails.Text = "TRAILS: ON"
        self.chk_trails.Size = Size(120, 45)
        self.chk_trails.Location = Point(365, 17)
        self.chk_trails.set_colors(Color.FromArgb(100, 70, 120), Color.FromArgb(70, 50, 90))
        self.chk_trails.Click += self.on_toggle_trails
        control_panel.Controls.Add(self.chk_trails)
        
        # Счетчик столкновений
        self.lbl_score = Label()
        self.lbl_score.Text = "COLLISIONS: 0"
        self.lbl_score.Size = Size(180, 35)
        self.lbl_score.Location = Point(500, 22)
        self.lbl_score.ForeColor = Color.Gold
        self.lbl_score.Font = Font("Segoe UI", 11, FontStyle.Bold)
        self.lbl_score.TextAlign = ContentAlignment.MiddleCenter
        control_panel.Controls.Add(self.lbl_score)
        
        # Метка Гравитация
        lbl_gravity = Label()
        lbl_gravity.Text = "GRAVITY:"
        lbl_gravity.Size = Size(90, 25)
        lbl_gravity.Location = Point(690, 15)
        lbl_gravity.ForeColor = Color.White
        lbl_gravity.Font = Font("Segoe UI", 9, FontStyle.Bold)
        lbl_gravity.TextAlign = ContentAlignment.MiddleLeft
        control_panel.Controls.Add(lbl_gravity)
        
        # Ползунок гравитации
        self.track_gravity = TrackBar()
        self.track_gravity.Minimum = 0
        self.track_gravity.Maximum = 20
        self.track_gravity.Value = 3
        self.track_gravity.Size = Size(120, 45)
        self.track_gravity.Location = Point(690, 40)
        self.track_gravity.TickFrequency = 2
        self.track_gravity.TickStyle = TickStyle.BottomRight
        self.track_gravity.BackColor = Color.FromArgb(40, 40, 60)
        self.track_gravity.ValueChanged += self.on_gravity_changed
        control_panel.Controls.Add(self.track_gravity)
        
        # Значение гравитации
        self.lbl_gravity_val = Label()
        self.lbl_gravity_val.Text = "0.3"
        self.lbl_gravity_val.Size = Size(40, 25)
        self.lbl_gravity_val.Location = Point(820, 45)
        self.lbl_gravity_val.ForeColor = Color.Gold
        self.lbl_gravity_val.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lbl_gravity_val.TextAlign = ContentAlignment.MiddleLeft
        control_panel.Controls.Add(self.lbl_gravity_val)
        
        # Информационная строка
        lbl_info = Label()
        lbl_info.Text = "CLICK ON BALL -> START ANIMATION | CLICK ON EMPTY SPACE -> ADD BALL"
        lbl_info.Size = Size(700, 25)
        lbl_info.Location = Point(100, 62)
        lbl_info.ForeColor = Color.FromArgb(150, 150, 200)
        lbl_info.Font = Font("Segoe UI", 8, FontStyle.Bold)
        lbl_info.TextAlign = ContentAlignment.MiddleCenter
        control_panel.Controls.Add(lbl_info)
    
    def create_balls(self):
        colors = [Color.OrangeRed, Color.DeepSkyBlue, Color.LimeGreen]
        positions = [(250, 200), (450, 150), (350, 300)]
        velocities = [(4, -3), (-3, -4), (3, -2)]
        radii = [20, 18, 22]
        masses = [1.0, 0.8, 1.2]
        
        for i in range(3):
            ball = Ball(
                positions[i][0], positions[i][1],
                velocities[i][0], velocities[i][1],
                colors[i], radii[i], masses[i]
            )
            self.balls.append(ball)
    
    def add_particles(self, x, y, color, count=10):
        for _ in range(count):
            angle = random.uniform(0, 2 * math.pi)
            speed = random.uniform(2, 8)
            vx = math.cos(angle) * speed
            vy = math.sin(angle) * speed
            particle = Particle(x, y, vx, vy, color)
            self.particles.append(particle)
    
    def update_bounds(self):
        self.bounds_max_x = self.ClientSize.Width - self.margin
        self.bounds_max_y = self.ClientSize.Height - self.margin - 80
    
    def check_ball_collisions(self):
        for i in range(len(self.balls)):
            for j in range(i + 1, len(self.balls)):
                ball1 = self.balls[i]
                ball2 = self.balls[j]
                
                dx = ball2.x - ball1.x
                dy = ball2.y - ball1.y
                distance = math.sqrt(dx * dx + dy * dy)
                min_distance = ball1.radius + ball2.radius
                
                if distance < min_distance:
                    overlap = min_distance - distance
                    angle = math.atan2(dy, dx)
                    
                    move_x = math.cos(angle) * overlap / 2
                    move_y = math.sin(angle) * overlap / 2
                    ball1.x -= move_x
                    ball1.y -= move_y
                    ball2.x += move_x
                    ball2.y += move_y
                    
                    nx = dx / distance
                    ny = dy / distance
                    relative_vel_x = ball2.vx - ball1.vx
                    relative_vel_y = ball2.vy - ball1.vy
                    speed = relative_vel_x * nx + relative_vel_y * ny
                    
                    if speed < 0:
                        restitution = 0.8
                        impulse = (1 + restitution) * speed / (1/ball1.mass + 1/ball2.mass)
                        ball1.vx += impulse * nx / ball1.mass
                        ball1.vy += impulse * ny / ball1.mass
                        ball2.vx -= impulse * nx / ball2.mass
                        ball2.vy -= impulse * ny / ball2.mass
                        
                        self.collision_count += 1
                        self.score += 10
                        self.lbl_score.Text = "COLLISIONS: " + str(self.collision_count)
                        
                        temp_color = ball1.color
                        ball1.color = ball2.color
                        ball2.color = temp_color
                        
                        self.add_particles((ball1.x + ball2.x)/2, (ball1.y + ball2.y)/2, 
                                         Color.Yellow, 15)
    
    def on_timer_tick(self, sender, e):
        if not self.is_animating:
            return
        
        self.update_bounds()
        
        for ball in self.balls:
            collision, normal = ball.update(
                self.gravity,
                self.bounds_min_x, self.bounds_min_y,
                self.bounds_max_x, self.bounds_max_y
            )
            if collision:
                self.add_particles(ball.x, ball.y, ball.color, 8)
                self.collision_count += 1
                self.lbl_score.Text = "COLLISIONS: " + str(self.collision_count)
                
                colors = [Color.Orange, Color.HotPink, Color.Cyan, Color.Magenta, Color.Gold]
                ball.color = random.choice(colors)
        
        self.check_ball_collisions()
        self.particles = [p for p in self.particles if p.update()]
        self.background_angle += 0.01
        self.Refresh()
    
    def on_paint(self, sender, e):
        g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        
        rect = Rectangle(0, 0, self.Width, self.Height)
        gradient = LinearGradientBrush(rect, 
                                       Color.FromArgb(20, 20, 40),
                                       Color.FromArgb(40, 20, 60),
                                       LinearGradientMode.Vertical)
        g.FillRectangle(gradient, rect)
        
        pen = SystemPen(Color.FromArgb(30, 100, 100, 150), 1)
        for i in range(0, self.Width, 40):
            g.DrawLine(pen, i, 0, i, self.Height)
        for i in range(0, self.Height, 40):
            g.DrawLine(pen, 0, i, self.Width, i)
        
        game_rect = Rectangle(
            self.margin - 3, self.margin - 3,
            self.bounds_max_x - self.bounds_min_x + 6,
            self.bounds_max_y - self.bounds_min_y + 6
        )
        border_gradient = LinearGradientBrush(game_rect,
                                             Color.FromArgb(80, 100, 200, 255),
                                             Color.FromArgb(80, 200, 100, 255),
                                             LinearGradientMode.ForwardDiagonal)
        pen = SystemPen(border_gradient, 3)
        g.DrawRectangle(pen, game_rect)
        
        inner_rect = Rectangle(
            self.margin - 1, self.margin - 1,
            self.bounds_max_x - self.bounds_min_x + 2,
            self.bounds_max_y - self.bounds_min_y + 2
        )
        inner_brush = SolidBrush(Color.FromArgb(50, 0, 0, 0))
        g.FillRectangle(inner_brush, inner_rect)
        
        for ball in self.balls:
            ball.draw(g, self.show_trails)
        
        for particle in self.particles:
            particle.draw(g)
        
        score_font = Font("Segoe UI", 16, FontStyle.Bold)
        score_brush = SolidBrush(Color.FromArgb(255, 255, 215, 0))
        score_text = "SCORE: " + str(self.score)
        g.DrawString(score_text, score_font, score_brush, 
                    self.margin, self.bounds_max_y + 15)
    
    def on_form_click(self, sender, e):
        clicked_on_ball = False
        for ball in self.balls:
            dx = e.X - ball.x
            dy = e.Y - ball.y
            distance = math.sqrt(dx * dx + dy * dy)
            if distance <= ball.radius:
                clicked_on_ball = True
                if not self.is_animating:
                    self.start_animation()
                break
        
        if not clicked_on_ball and self.is_animating:
            self.add_ball_at_position(e.X, e.Y)
    
    def add_ball_at_position(self, x, y):
        if len(self.balls) < 12:
            colors = [Color.Red, Color.Blue, Color.Green, Color.Purple, 
                     Color.Orange, Color.Pink, Color.Cyan, Color.Yellow]
            color = random.choice(colors)
            radius = random.randint(12, 25)
            mass = radius / 15.0
            vx = random.uniform(-5, 5)
            vy = random.uniform(-5, 5)
            
            ball = Ball(x, y, vx, vy, color, radius, mass)
            self.balls.append(ball)
            self.add_particles(x, y, Color.White, 20)
    
    def on_add_ball(self, sender, e):
        if len(self.balls) < 12:
            x = random.randint(self.margin + 30, self.bounds_max_x - 30)
            y = random.randint(self.margin + 30, self.bounds_max_y - 30)
            self.add_ball_at_position(x, y)
    
    def on_reset(self, sender, e):
        self.is_animating = False
        self.timer.Stop()
        self.btn_toggle.Text = "START"
        self.btn_toggle.set_colors(Color.FromArgb(0, 150, 0), Color.FromArgb(0, 100, 0))
        
        self.balls = []
        self.particles = []
        self.score = 0
        self.collision_count = 0
        self.lbl_score.Text = "COLLISIONS: 0"
        self.create_balls()
        self.Refresh()
    
    def start_animation(self):
        self.is_animating = True
        self.timer.Start()
        self.btn_toggle.Text = "STOP"
        self.btn_toggle.set_colors(Color.FromArgb(200, 0, 0), Color.FromArgb(150, 0, 0))
    
    def stop_animation(self):
        self.is_animating = False
        self.timer.Stop()
        self.btn_toggle.Text = "START"
        self.btn_toggle.set_colors(Color.FromArgb(0, 150, 0), Color.FromArgb(0, 100, 0))
    
    def on_toggle_click(self, sender, e):
        if self.is_animating:
            self.stop_animation()
        else:
            self.start_animation()
    
    def on_toggle_trails(self, sender, e):
        self.show_trails = not self.show_trails
        if self.show_trails:
            self.chk_trails.Text = "TRAILS: ON"
            self.chk_trails.set_colors(Color.FromArgb(100, 70, 120), Color.FromArgb(70, 50, 90))
        else:
            self.chk_trails.Text = "TRAILS: OFF"
            self.chk_trails.set_colors(Color.FromArgb(80, 80, 80), Color.FromArgb(60, 60, 60))
    
    def on_gravity_changed(self, sender, e):
        self.gravity = self.track_gravity.Value / 10.0
        self.lbl_gravity_val.Text = "{:.1f}".format(self.gravity)
        self.Refresh()
    
    def on_resize(self, sender, e):
        self.update_bounds()
        self.Refresh()
    
    def on_form_closed(self, sender, e):
        self.timer.Stop()

class Particle:
    def __init__(self, x, y, vx, vy, color):
        self.x = x
        self.y = y
        self.vx = vx
        self.vy = vy
        self.color = color
        self.life = 1.0
        self.size = random.randint(2, 6)
    
    def update(self):
        self.x += self.vx
        self.y += self.vy
        self.vy += 0.2
        self.life -= 0.03
        return self.life > 0
    
    def draw(self, g):
        alpha = int(255 * self.life)
        color = Color.FromArgb(alpha, self.color)
        brush = SolidBrush(color)
        g.FillEllipse(brush, self.x - self.size/2, self.y - self.size/2, self.size, self.size)

class Ball:
    def __init__(self, x, y, vx, vy, color, radius, mass):
        self.x = x
        self.y = y
        self.vx = vx
        self.vy = vy
        self.color = color
        self.radius = radius
        self.mass = mass
        self.trail = []
        self.max_trail = 15
    
    def update(self, gravity, bounds_min_x, bounds_min_y, bounds_max_x, bounds_max_y):
        self.vy += gravity
        
        new_x = self.x + self.vx
        new_y = self.y + self.vy
        
        collision = False
        
        if new_x <= bounds_min_x + self.radius:
            new_x = bounds_min_x + self.radius
            self.vx = -self.vx * 0.95
            collision = True
        elif new_x >= bounds_max_x - self.radius:
            new_x = bounds_max_x - self.radius
            self.vx = -self.vx * 0.95
            collision = True
        
        if new_y <= bounds_min_y + self.radius:
            new_y = bounds_min_y + self.radius
            self.vy = -self.vy * 0.95
            collision = True
        elif new_y >= bounds_max_y - self.radius:
            new_y = bounds_max_y - self.radius
            self.vy = -self.vy * 0.95
            collision = True
        
        self.x = new_x
        self.y = new_y
        
        self.trail.append((self.x, self.y))
        if len(self.trail) > self.max_trail:
            self.trail.pop(0)
        
        return collision, None
    
    def draw(self, g, show_trail=True):
        if show_trail and len(self.trail) > 1:
            for i in range(1, len(self.trail)):
                alpha = int(100 * i / len(self.trail))
                color = Color.FromArgb(alpha, self.color)
                pen = SystemPen(color, 2)
                g.DrawLine(pen, 
                          int(self.trail[i-1][0]), int(self.trail[i-1][1]),
                          int(self.trail[i][0]), int(self.trail[i][1]))
        
        for glow_size in range(4, 0, -1):
            alpha = 30 - glow_size * 5
            glow_color = Color.FromArgb(alpha, self.color)
            glow_brush = SolidBrush(glow_color)
            g.FillEllipse(glow_brush,
                         int(self.x - self.radius - glow_size),
                         int(self.y - self.radius - glow_size),
                         (self.radius + glow_size) * 2,
                         (self.radius + glow_size) * 2)
        
        brush = SolidBrush(self.color)
        g.FillEllipse(brush,
                     int(self.x - self.radius),
                     int(self.y - self.radius),
                     self.radius * 2,
                     self.radius * 2)
        
        highlight_brush = SolidBrush(Color.FromArgb(80, Color.White))
        g.FillEllipse(highlight_brush,
                     int(self.x - self.radius * 0.5),
                     int(self.y - self.radius * 0.5),
                     self.radius,
                     self.radius * 0.6)

# Запускаем форму
form = AdvancedAnimationForm()
form.Closed += form.on_form_closed
form.ShowDialog()