"""
Система автоматизации составления отчетов
Главный файл запуска приложения
"""

import tkinter as tk
from gui import ReportApp
from database import init_database
import os

def main():
    """Главная функция запуска приложения"""

    # Создаем необходимые папки, если их нет
    os.makedirs("формы", exist_ok=True)
    os.makedirs("отчеты", exist_ok=True)
    os.makedirs("шаблоны", exist_ok=True)

    # Инициализируем базу данных
    init_database()

    # Создаем главное окно приложения
    root = tk.Tk()
    app = ReportApp(root)

    # Запускаем главный цикл
    root.mainloop()

if __name__ == "__main__":
    main()