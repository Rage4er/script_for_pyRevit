# -*- coding: utf-8 -*-
import random

import clr
import System

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Drawing import *
from System.Windows.Forms import *


class RevitQuizForm(Form):
    def __init__(self):
        self.Text = "Викторина по Revit"
        self.Size = Size(500, 600)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.LightBlue
        self.DoubleBuffered = True

        # Вопросы и ответы
        self.questions = [
            {
                "question": "Какой горячей клавишей создается новый лист в Revit?",
                "options": ["Ctrl+N", "Ctrl+L", "Ctrl+Shift+N", "Alt+L"],
                "correct": 2,
                "explanation": "Ctrl+Shift+N - стандартная горячая клавиша для создания нового листа",
            },
            {
                "question": "Что такое Worksharing в Revit?",
                "options": [
                    "Совместная работа над проектом",
                    "Автосохранение файлов",
                    "Работа с семействами",
                    "Экспорт в другие форматы",
                ],
                "correct": 0,
                "explanation": "Worksharing позволяет нескольким пользователям работать над одним проектом одновременно",
            },
            {
                "question": "Какой параметр отвечает за уровень этажа?",
                "options": ["Level", "Floor", "Story", "Elevation"],
                "correct": 0,
                "explanation": "Level - это параметр, определяющий уровень этажа в Revit",
            },
            {
                "question": "Что означает аббревиатура BIM?",
                "options": [
                    "Building Information Modeling",
                    "Building Intelligent Model",
                    "Basic Information Management",
                    "Building Integrated Method",
                ],
                "correct": 0,
                "explanation": "BIM - Building Information Modeling (Информационное моделирование зданий)",
            },
            {
                "question": "Как создать местный вид в Revit?",
                "options": [
                    "View → Create → Callout",
                    "Architecture → Callout View",
                    "View → Callout",
                    "Modify → Create Callout",
                ],
                "correct": 2,
                "explanation": "Callout создается через вкладку View → Callout",
            },
            {
                "question": "Что такое семейство (Family) в Revit?",
                "options": [
                    "Группа связанных элементов",
                    "Тип проекта",
                    "Категория материалов",
                    "Набор параметров",
                ],
                "correct": 0,
                "explanation": "Семейство - это группа элементов с общими параметрами и поведением",
            },
            {
                "question": "Какой инструмент используется для выравнивания элементов?",
                "options": ["Align", "Match", "Snap", "Adjust"],
                "correct": 0,
                "explanation": "Инструмент Align (Выравнивание) на вкладке Modify",
            },
            {
                "question": "Что делает команда 'Pin'?",
                "options": [
                    "Фиксирует элемент на месте",
                    "Добавляет комментарий",
                    "Создает связь с файлом",
                    "Отмечает элемент как важный",
                ],
                "correct": 0,
                "explanation": "Pin фиксирует элемент, предотвращая его случайное перемещение",
            },
            {
                "question": "Как создать этаж в Revit?",
                "options": [
                    "Architecture → Floor",
                    "Structure → Floor",
                    "Both A and B",
                    "Modify → Create Floor",
                ],
                "correct": 2,
                "explanation": "Этаж можно создать как через Architecture, так и через Structure вкладки",
            },
            {
                "question": "Что такое 'Phasing' в Revit?",
                "options": [
                    "Разделение проекта на этапы",
                    "Настройка фаз материалов",
                    "Автоматическое сохранение",
                    "Синхронизация с облаком",
                ],
                "correct": 0,
                "explanation": "Phasing позволяет управлять этапами строительства (существующее, новое, снос)",
            },
        ]

        self.current_question = 0
        self.score = 0
        self.total_questions = len(self.questions)
        self.selected_answer = None
        self.quiz_completed = False

        # Панель управления
        self.control_panel = Panel()
        self.control_panel.Location = Point(0, 0)
        self.control_panel.Size = Size(500, 60)
        self.control_panel.BackColor = Color.DarkBlue

        # Статус
        self.status_label = Label()
        self.status_label.Text = "Викторина по Revit | Вопрос 1 из {}".format(
            self.total_questions
        )
        self.status_label.ForeColor = Color.White
        self.status_label.Location = Point(10, 10)
        self.status_label.Size = Size(480, 20)
        self.status_label.Font = Font("Arial", 10, FontStyle.Bold)
        self.control_panel.Controls.Add(self.status_label)

        # Счет
        self.score_label = Label()
        self.score_label.Text = "Счет: 0"
        self.score_label.ForeColor = Color.Yellow
        self.score_label.Location = Point(10, 35)
        self.score_label.Size = Size(200, 20)
        self.score_label.Font = Font("Arial", 10, FontStyle.Bold)
        self.control_panel.Controls.Add(self.score_label)

        self.Controls.Add(self.control_panel)

        # Вопрос
        self.question_label = Label()
        self.question_label.Location = Point(20, 80)
        self.question_label.Size = Size(460, 60)
        self.question_label.Font = Font("Arial", 11, FontStyle.Bold)
        self.question_label.Text = ""
        self.Controls.Add(self.question_label)

        # Варианты ответов
        self.option_buttons = []
        for i in range(4):
            btn = Button()
            btn.Location = Point(50, 160 + i * 50)
            btn.Size = Size(400, 40)
            btn.Font = Font("Arial", 10)
            btn.BackColor = Color.White
            btn.Click += lambda s, e, idx=i: self.select_answer(idx)
            btn.TabStop = False
            self.option_buttons.append(btn)
            self.Controls.Add(btn)

        # Кнопка далее
        self.next_button = Button()
        self.next_button.Text = "Далее →"
        self.next_button.Location = Point(350, 400)
        self.next_button.Size = Size(100, 30)
        self.next_button.BackColor = Color.LightGreen
        self.next_button.Font = Font("Arial", 10)
        self.next_button.Click += self.next_question
        self.next_button.Enabled = False
        self.next_button.TabStop = False
        self.Controls.Add(self.next_button)

        # Объяснение
        self.explanation_label = Label()
        self.explanation_label.Location = Point(20, 370)
        self.explanation_label.Size = Size(460, 60)
        self.explanation_label.Font = Font("Arial", 9)
        self.explanation_label.ForeColor = Color.DarkBlue
        self.explanation_label.Visible = False
        self.Controls.Add(self.explanation_label)

        # Результат
        self.result_label = Label()
        self.result_label.Location = Point(50, 450)
        self.result_label.Size = Size(400, 100)
        self.result_label.Font = Font("Arial", 14, FontStyle.Bold)
        self.result_label.TextAlign = ContentAlignment.MiddleCenter
        self.result_label.Visible = False
        self.Controls.Add(self.result_label)

        # Кнопка новая игра
        self.new_game_button = Button()
        self.new_game_button.Text = "Новая викторина"
        self.new_game_button.Location = Point(175, 500)
        self.new_game_button.Size = Size(150, 35)
        self.new_game_button.BackColor = Color.Orange
        self.new_game_button.Font = Font("Arial", 10)
        self.new_game_button.Click += self.new_quiz
        self.new_game_button.Visible = False
        self.new_game_button.TabStop = False
        self.Controls.Add(self.new_game_button)

        self.load_question()

    def load_question(self):
        if self.current_question < self.total_questions:
            question_data = self.questions[self.current_question]

            self.question_label.Text = question_data["question"]

            for i, option in enumerate(question_data["options"]):
                self.option_buttons[i].Text = "{}. {}".format(chr(65 + i), option)
                self.option_buttons[i].BackColor = Color.White
                self.option_buttons[i].Enabled = True

            self.explanation_label.Visible = False
            self.next_button.Enabled = False
            self.selected_answer = None

            self.status_label.Text = "Викторина по Revit | Вопрос {} из {}".format(
                self.current_question + 1, self.total_questions
            )
            self.score_label.Text = "Счет: {}".format(self.score)

    def select_answer(self, index):
        if self.selected_answer is not None:
            return

        self.selected_answer = index
        question_data = self.questions[self.current_question]

        # Подсвечиваем ответы
        for i in range(4):
            if i == question_data["correct"]:
                self.option_buttons[i].BackColor = Color.LightGreen  # Правильный
            elif i == index:
                self.option_buttons[i].BackColor = Color.LightPink  # Неправильный выбор
            else:
                self.option_buttons[i].BackColor = Color.LightGray  # Остальные

        # Отключаем кнопки
        for btn in self.option_buttons:
            btn.Enabled = False

        # Проверяем ответ
        if index == question_data["correct"]:
            self.score += 1
            self.score_label.Text = "Счет: {}".format(self.score)

        # Показываем объяснение
        self.explanation_label.Text = "Объяснение: {}".format(
            question_data["explanation"]
        )
        self.explanation_label.Visible = True

        self.next_button.Enabled = True

    def next_question(self, sender, e):
        self.current_question += 1

        if self.current_question < self.total_questions:
            self.load_question()
        else:
            self.show_results()

    def show_results(self):
        self.quiz_completed = True

        # Скрываем элементы вопроса
        self.question_label.Visible = False
        for btn in self.option_buttons:
            btn.Visible = False
        self.next_button.Visible = False
        self.explanation_label.Visible = False

        # Показываем результаты
        percentage = (float(self.score) / float(self.total_questions)) * 100

        if percentage >= 90:
            message = "Отлично! 🎉\nВы эксперт Revit!\n\nСчет: {}/{} ({:.0f}%)".format(
                self.score, self.total_questions, percentage
            )
            color = Color.Green
        elif percentage >= 70:
            message = "Хорошо! 👍\nВы уверенный пользователь!\n\nСчет: {}/{} ({:.0f}%)".format(
                self.score, self.total_questions, percentage
            )
            color = Color.Blue
        elif percentage >= 50:
            message = (
                "Неплохо! 😊\nЕсть что повторить!\n\nСчет: {}/{} ({:.0f}%)".format(
                    self.score, self.total_questions, percentage
                )
            )
            color = Color.Orange
        else:
            message = "Пора учиться! 📚\nОсвойте основы Revit!\n\nСчет: {}/{} ({:.0f}%)".format(
                self.score, self.total_questions, percentage
            )
            color = Color.Red

        self.result_label.Text = message
        self.result_label.ForeColor = color
        self.result_label.Visible = True

        self.new_game_button.Visible = True

        self.status_label.Text = "Викторина завершена!"

    def new_quiz(self):
        # Сбрасываем игру
        self.current_question = 0
        self.score = 0
        self.quiz_completed = False

        # Показываем элементы вопроса
        self.question_label.Visible = True
        for btn in self.option_buttons:
            btn.Visible = True
        self.next_button.Visible = True
        self.result_label.Visible = False
        self.new_game_button.Visible = False

        # Перемешиваем вопросы
        random.shuffle(self.questions)

        self.load_question()


# Функции для pyRevit
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    return True


def __invoke__(script_cmp, ui_button_cmp, __rvt__):
    try:
        form = RevitQuizForm()
        form.ShowDialog()
        return True
    except Exception as e:
        MessageBox.Show("Ошибка: " + str(e))
        return False


if __name__ == "__main__":
    Application.Run(RevitQuizForm())
