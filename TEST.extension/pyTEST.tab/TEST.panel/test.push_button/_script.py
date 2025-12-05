# -*- coding: utf-8 -*-

from pyrevit import forms

# Определение нашей функции для показа диалога
def show_dialog():
    # Показываем диалоговое окно с вопросом
    result = forms.alert(
        "Выберите Да или Нет?",
        title="Подтверждение",
        yes=True,
        no=True
    )

    # Проверяем результат
    if result:
        print("Вы выбрали Да.")
    else:
        print("Вы выбрали Нет.")

# Выполняем нашу функцию
show_dialog()