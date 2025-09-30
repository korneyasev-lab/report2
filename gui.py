"""
GUI модуль для системы автоматизации отчетов
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
from logic import ReportLogic
import subprocess
import os


class ReportApp:
    """Главный класс приложения с GUI"""

    def __init__(self, root):
        self.root = root
        self.root.title("Система автоматизации отчетов")
        self.root.geometry("1200x800")

        self.logic = ReportLogic()
        self.current_block_widgets = {}

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        os.makedirs("шаблоны", exist_ok=True)

        self.show_main_menu()

    def clear_frame(self):
        """Очистка главного контейнера"""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def show_main_menu(self):
        """Показать главное меню"""
        self.clear_frame()

        tk.Label(self.main_frame, text="Система автоматизации отчетов", font=("Arial", 24, "bold")).pack(pady=50)

        btn_frame = tk.Frame(self.main_frame)
        btn_frame.pack(pady=20)

        tk.Button(btn_frame, text="Создать новый отчет", font=("Arial", 16), width=30, height=2, command=self.show_report_creation).pack(pady=10)
        tk.Button(btn_frame, text="Открыть сохраненный отчет", font=("Arial", 16), width=30, height=2, command=self.show_saved_reports).pack(pady=10)
        tk.Button(btn_frame, text="Просмотр всех отчетов", font=("Arial", 16), width=30, height=2, command=self.show_all_reports_list).pack(pady=10)
        tk.Button(btn_frame, text="Выход", font=("Arial", 16), width=30, height=2, command=self.root.quit).pack(pady=10)

    def show_report_creation(self):
        """Экран выбора формы и параметров отчета"""
        self.clear_frame()

        tk.Label(self.main_frame, text="Новый отчет", font=("Arial", 20, "bold")).pack(pady=30)

        form_frame = tk.Frame(self.main_frame)
        form_frame.pack(pady=20)

        tk.Label(form_frame, text="Выберите форму:", font=("Arial", 16)).grid(row=0, column=0, sticky="w", pady=10)
        self.form_var = tk.StringVar()
        forms = self.logic.load_forms_list()

        if not forms:
            messagebox.showerror("Ошибка", "Не найдены файлы форм в папке 'формы/'")
            self.show_main_menu()
            return

        ttk.Combobox(form_frame, textvariable=self.form_var, values=forms, font=("Arial", 16), width=30, state="readonly").grid(row=0, column=1, pady=10, padx=10)

        tk.Label(form_frame, text="Месяц:", font=("Arial", 16)).grid(row=1, column=0, sticky="w", pady=10)
        self.month_var = tk.StringVar()
        months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
        ttk.Combobox(form_frame, textvariable=self.month_var, values=months, font=("Arial", 16), width=30, state="readonly").grid(row=1, column=1, pady=10, padx=10)

        tk.Label(form_frame, text="Год:", font=("Arial", 16)).grid(row=2, column=0, sticky="w", pady=10)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        tk.Entry(form_frame, textvariable=self.year_var, font=("Arial", 16), width=32).grid(row=2, column=1, pady=10, padx=10)

        tk.Label(form_frame, text="Дата создания:", font=("Arial", 16)).grid(row=3, column=0, sticky="w", pady=10)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        tk.Entry(form_frame, textvariable=self.date_var, font=("Arial", 16), width=32).grid(row=3, column=1, pady=10, padx=10)

        btn_frame = tk.Frame(self.main_frame)
        btn_frame.pack(pady=30)
        tk.Button(btn_frame, text="Начать заполнение", font=("Arial", 16), width=20, command=self.start_filling).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Назад", font=("Arial", 16), width=20, command=self.show_main_menu).pack(side=tk.LEFT, padx=10)

    def start_filling(self):
        """Начать заполнение отчета"""
        form_name = self.form_var.get()
        month = self.month_var.get()
        year = self.year_var.get()
        report_date = self.date_var.get()

        if not all([form_name, month, year, report_date]):
            messagebox.showerror("Ошибка", "Заполните все поля!")
            return

        excel_file = f"формы/{form_name}.xlsx"
        if not self.logic.load_questions_from_excel(excel_file):
            messagebox.showerror("Ошибка", f"Не удалось загрузить вопросы из файла {excel_file}")
            return

        self.logic.init_report(form_name, month, year, report_date)
        self.show_questions_screen()

    def show_questions_screen(self):
        """Экран заполнения вопросов блоками"""
        self.clear_frame()

        start, end = self.logic.get_current_block_questions()
        total = len(self.logic.questions_list)

        header = f"Отчет: {self.logic.current_report_data['form_name']} {self.logic.current_report_data['month']} {self.logic.current_report_data['year']}"
        tk.Label(self.main_frame, text=header, font=("Arial", 16, "bold")).pack(pady=5)
        tk.Label(self.main_frame, text=f"Вопросы {start + 1}-{end} из {total}", font=("Arial", 14)).pack(pady=2)

        canvas_frame = tk.Frame(self.main_frame)
        canvas_frame.pack(pady=5, fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(canvas_frame)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.current_block_widgets = {}
        for i in range(start, end):
            self.create_question_widget(scrollable_frame, i)

        btn_frame = tk.Frame(self.main_frame)
        btn_frame.pack(pady=10)

        btn_prev = tk.Button(btn_frame, text="← Назад", font=("Arial", 18), width=15, command=self.on_prev_block)
        btn_prev.pack(side=tk.LEFT, padx=5)
        if start == 0:
            btn_prev.config(state=tk.DISABLED)

        btn_next = tk.Button(btn_frame, text="Далее →" if end < total else "Завершить", font=("Arial", 18), width=15, command=self.on_next_block)
        btn_next.pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Сохранить", font=("Arial", 18), width=15, command=self.on_save_report).pack(side=tk.LEFT, padx=5)

    def create_question_widget(self, parent, question_index):
        """Создать виджет для одного вопроса"""
        question = self.logic.questions_list[question_index]
        answer_data = self.logic.answers_list[question_index]

        q_frame = tk.LabelFrame(parent, text=f"Вопрос {question_index + 1}", font=("Arial", 16, "bold"), padx=10, pady=5)
        q_frame.pack(pady=5, padx=20, fill=tk.X)

        top_frame = tk.Frame(q_frame)
        top_frame.pack(fill=tk.X, pady=2)

        answer_frame = tk.Frame(top_frame)
        answer_frame.pack(side=tk.LEFT, padx=(0, 15))

        current_answer = answer_data['answer_yes_no']

        btn_yes = tk.Button(
            answer_frame, text="ДА", font=("Arial", 18, "bold"), width=5,
            bg="green" if current_answer == "Да" else "lightgray",
            fg="white" if current_answer == "Да" else "darkgreen",
            activebackground="darkgreen",
            activeforeground="white",
            relief=tk.SUNKEN if current_answer == "Да" else tk.RAISED,
            highlightthickness=2,
            highlightbackground="green" if current_answer == "Да" else "gray",
            command=lambda idx=question_index: self.set_answer(idx, "Да")
        )
        btn_yes.pack(side=tk.LEFT, padx=3)

        btn_no = tk.Button(
            answer_frame, text="НЕТ", font=("Arial", 18, "bold"), width=5,
            bg="red" if current_answer == "Нет" else "lightgray",
            fg="white" if current_answer == "Нет" else "darkred",
            activebackground="darkred",
            activeforeground="white",
            relief=tk.SUNKEN if current_answer == "Нет" else tk.RAISED,
            highlightthickness=2,
            highlightbackground="red" if current_answer == "Нет" else "gray",
            command=lambda idx=question_index: self.set_answer(idx, "Нет")
        )
        btn_no.pack(side=tk.LEFT, padx=3)

        tk.Label(top_frame, text=question['question'], font=("Arial", 18), wraplength=800, justify=tk.LEFT, anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True)

        buttons_frame = tk.Frame(top_frame)
        buttons_frame.pack(side=tk.RIGHT, padx=5)

        tk.Button(buttons_frame, text="?", font=("Arial", 16, "bold"), width=3,
                  command=lambda: self.show_help(question)).pack(side=tk.LEFT, padx=2)

        if question.get('documents'):
            tk.Button(buttons_frame, text="📄", font=("Arial", 16, "bold"), width=3,
                      command=lambda: self.show_documents(question)).pack(side=tk.LEFT, padx=2)

        comment_frame = tk.Frame(q_frame)
        comment_frame.pack(fill=tk.X, pady=2)
        tk.Label(comment_frame, text="Комментарий:", font=("Arial", 18)).pack(side=tk.LEFT, padx=(0, 5))

        comment_text = tk.Text(comment_frame, height=2, font=("Arial", 18), wrap=tk.WORD)
        comment_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if answer_data['comment']:
            comment_text.insert(1.0, answer_data['comment'])

        def adjust_height(event=None):
            lines = comment_text.get("1.0", tk.END).count('\n')
            comment_text.config(height=max(2, min(6, lines + 1)))

        comment_text.bind('<KeyRelease>', adjust_height)
        adjust_height()

        self.current_block_widgets[question_index] = {'btn_yes': btn_yes, 'btn_no': btn_no, 'comment': comment_text}

    def set_answer(self, question_index, answer):
        """Установить ответ с инверсией цвета кнопок"""
        widgets = self.current_block_widgets[question_index]

        if answer == "Да":
            widgets['btn_yes'].config(bg="green", fg="white", relief=tk.SUNKEN, highlightbackground="green", highlightthickness=2)
            widgets['btn_no'].config(bg="lightgray", fg="darkred", relief=tk.RAISED, highlightbackground="gray", highlightthickness=2)
        else:
            widgets['btn_yes'].config(bg="lightgray", fg="darkgreen", relief=tk.RAISED, highlightbackground="gray", highlightthickness=2)
            widgets['btn_no'].config(bg="red", fg="white", relief=tk.SUNKEN, highlightbackground="red", highlightthickness=2)

        widgets['btn_yes'].update()
        widgets['btn_no'].update()

        comment = widgets['comment'].get(1.0, tk.END).strip()
        self.logic.save_answer(question_index, answer, comment)

    def show_help(self, question):
        """Показать справку"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Справка")
        help_window.geometry("900x600")

        tk.Label(help_window, text="ГОСТ ИСО 9001:", font=("Arial", 18, "bold")).pack(pady=10)
        gost_text = scrolledtext.ScrolledText(help_window, height=10, width=100, font=("Arial", 16), wrap=tk.WORD)
        gost_text.pack(pady=5, padx=10)
        gost_text.insert(1.0, question['gost'] or "Информация отсутствует")
        gost_text.config(state=tk.DISABLED)

        tk.Label(help_window, text="Руководство по качеству:", font=("Arial", 18, "bold")).pack(pady=10)
        quality_text = scrolledtext.ScrolledText(help_window, height=10, width=100, font=("Arial", 16), wrap=tk.WORD)
        quality_text.pack(pady=5, padx=10)
        quality_text.insert(1.0, question['quality'] or "Информация отсутствует")
        quality_text.config(state=tk.DISABLED)

        tk.Button(help_window, text="Закрыть", font=("Arial", 16), command=help_window.destroy).pack(pady=15)

    def show_documents(self, question):
        """Показать связанные документы и открыть папку шаблонов"""
        docs = question.get('documents', '')

        if not docs:
            messagebox.showinfo("Документы", "Для этого вопроса нет связанных документов")
            return

        doc_window = tk.Toplevel(self.root)
        doc_window.title("Связанные документы")
        doc_window.geometry("600x400")

        tk.Label(doc_window, text="Документы для этого вопроса:", font=("Arial", 16, "bold")).pack(pady=10)

        text_widget = scrolledtext.ScrolledText(doc_window, height=15, width=70, font=("Arial", 14), wrap=tk.WORD)
        text_widget.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        text_widget.insert(1.0, docs)
        text_widget.config(state=tk.DISABLED)

        btn_frame = tk.Frame(doc_window)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Открыть папку шаблонов", font=("Arial", 14),
                  command=self.open_templates_folder).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Закрыть", font=("Arial", 14),
                  command=doc_window.destroy).pack(side=tk.LEFT, padx=5)

    def open_templates_folder(self):
        """Открыть папку шаблонов в проводнике"""
        templates_path = os.path.abspath("шаблоны")

        if not os.path.exists(templates_path):
            os.makedirs(templates_path)
            messagebox.showinfo("Информация", "Папка 'шаблоны' создана.\nПоместите туда файлы шаблонов документов.")

        try:
            if os.name == 'nt':
                os.startfile(templates_path)
            elif os.name == 'posix':
                subprocess.Popen(['xdg-open', templates_path])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")

    def save_current_block(self):
        """Сохранить ответы текущего блока"""
        for idx, widgets in self.current_block_widgets.items():
            comment = widgets['comment'].get(1.0, tk.END).strip()
            current_answer = self.logic.answers_list[idx]['answer_yes_no']
            if current_answer:
                self.logic.save_answer(idx, current_answer, comment)

    def on_prev_block(self):
        """Обработка кнопки Назад"""
        self.save_current_block()
        self.logic.prev_block()
        self.show_questions_screen()

    def on_next_block(self):
        """Обработка кнопки Далее/Завершить"""
        for idx in self.current_block_widgets.keys():
            if not self.logic.answers_list[idx]['answer_yes_no']:
                messagebox.showwarning("Внимание", f"Пожалуйста, ответьте на вопрос {idx + 1}")
                return

        self.save_current_block()

        if not self.logic.next_block():
            self.save_report()
        else:
            self.show_questions_screen()

    def on_save_report(self):
        """Обработка кнопки Сохранить"""
        self.save_current_block()

        all_ok, question_num = self.logic.check_all_answered()
        if not all_ok:
            messagebox.showwarning("Внимание", f"Вопрос {question_num} не заполнен")
            return

        if messagebox.askyesno("Сохранение", "Сохранить отчет в базу данных и экспортировать в Excel?"):
            self.save_report()

    def save_report(self):
        """Сохранить отчет"""
        success, result = self.logic.save_report()
        if success:
            messagebox.showinfo("Успех", f"Отчет сохранен!\n\nФайл: {result}\nПапка: отчеты/")
            self.show_main_menu()
        else:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{result}")

    def show_saved_reports(self):
        """Список сохраненных отчетов"""
        self.clear_frame()
        tk.Label(self.main_frame, text="Сохраненные отчеты", font=("Arial", 18, "bold")).pack(pady=20)

        reports = self.logic.get_all_reports_from_db()
        if not reports:
            tk.Label(self.main_frame, text="Нет сохраненных отчетов", font=("Arial", 12)).pack(pady=20)
        else:
            tree = self.create_reports_tree(reports, ["ID", "Форма", "Месяц", "Год", "Дата создания"])

            btn_frame = tk.Frame(self.main_frame)
            btn_frame.pack(pady=20)
            tk.Button(btn_frame, text="Открыть отчет", font=("Arial", 12), width=20, command=lambda: self.open_report(tree)).pack(side=tk.LEFT, padx=10)
            tk.Button(btn_frame, text="Экспортировать в Excel", font=("Arial", 12), width=20, command=lambda: self.export_report(tree)).pack(side=tk.LEFT, padx=10)

        tk.Button(self.main_frame, text="Назад", font=("Arial", 12), width=20, command=self.show_main_menu).pack(pady=10)

    def show_all_reports_list(self):
        """Список всех отчетов с удалением"""
        self.clear_frame()
        tk.Label(self.main_frame, text="Все отчеты", font=("Arial", 18, "bold")).pack(pady=20)

        reports = self.logic.get_all_reports_from_db()
        if not reports:
            tk.Label(self.main_frame, text="Нет сохраненных отчетов", font=("Arial", 12)).pack(pady=20)
        else:
            tree = self.create_reports_tree(reports, ["ID", "Форма", "Месяц", "Год", "Дата создания", "Файл"], [50, 200, 100, 80, 150, 300])

            btn_frame = tk.Frame(self.main_frame)
            btn_frame.pack(pady=20)
            tk.Button(btn_frame, text="Удалить отчет", font=("Arial", 12), width=20, command=lambda: self.delete_report(tree)).pack(side=tk.LEFT, padx=10)

        tk.Button(self.main_frame, text="Назад", font=("Arial", 12), width=20, command=self.show_main_menu).pack(pady=10)

    def create_reports_tree(self, reports, columns, widths=None):
        """Создать таблицу отчетов"""
        tree_frame = tk.Frame(self.main_frame)
        tree_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=scrollbar.set)
        scrollbar.config(command=tree.yview)

        for i, col in enumerate(columns):
            tree.heading(col, text=col)
            tree.column(col, width=widths[i] if widths else 150)

        for report in reports:
            values = [report['id'], report['form_name'], report['month'], report['year'], report['created_at']]
            if 'file_path' in report:
                values.append(report['file_path'])
            tree.insert("", tk.END, values=values)

        tree.pack(fill=tk.BOTH, expand=True)
        return tree

    def open_report(self, tree):
        """Открыть выбранный отчет"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите отчет из списка")
            return

        report_id = tree.item(selected[0])['values'][0]
        report_data = self.logic.get_report_from_db(report_id)

        if report_data:
            self.view_report(report_data)
        else:
            messagebox.showerror("Ошибка", "Не удалось загрузить отчет")

    def view_report(self, report_data):
        """Просмотр отчета"""
        self.clear_frame()

        header = f"{report_data['form_name']} - {report_data['month']} {report_data['year']}"
        tk.Label(self.main_frame, text=header, font=("Arial", 16, "bold")).pack(pady=20)

        info = f"Дата отчета: {report_data['report_date']}\nСоздан: {report_data['created_at']}"
        tk.Label(self.main_frame, text=info, font=("Arial", 11)).pack(pady=5)

        text_frame = tk.Frame(self.main_frame)
        text_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        text_widget = scrolledtext.ScrolledText(text_frame, font=("Arial", 11), wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True)

        for i, answer in enumerate(report_data['answers']):
            text_widget.insert(tk.END, f"\n{i+1}. {answer['question_text']}\n", "bold")
            text_widget.insert(tk.END, f"Ответ: {answer['answer_yes_no']}\n")
            if answer['comment']:
                text_widget.insert(tk.END, f"Комментарий: {answer['comment']}\n")
            text_widget.insert(tk.END, "\n" + "-"*80 + "\n")

        text_widget.tag_config("bold", font=("Arial", 11, "bold"))
        text_widget.config(state=tk.DISABLED)

        tk.Button(self.main_frame, text="Назад", font=("Arial", 12), width=20, command=self.show_saved_reports).pack(pady=20)

    def export_report(self, tree):
        """Экспортировать отчет в Excel"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите отчет из списка")
            return

        report_id = tree.item(selected[0])['values'][0]
        report_data = self.logic.get_report_from_db(report_id)

        if report_data:
            success, result = self.logic.export_report_to_word(report_data)
            if success:
                messagebox.showinfo("Успех", f"Отчет экспортирован:\n{result}")
            else:
                messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{result}")
        else:
            messagebox.showerror("Ошибка", "Не удалось загрузить отчет")

    def delete_report(self, tree):
        """Удалить отчет"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите отчет из списка")
            return

        item = tree.item(selected[0])
        report_id = item['values'][0]
        report_name = f"{item['values'][1]} {item['values'][2]} {item['values'][3]}"

        if messagebox.askyesno("Подтверждение", f"Вы уверены, что хотите удалить отчет:\n{report_name}?"):
            success, message = self.logic.delete_report_from_db(report_id)
            if success:
                messagebox.showinfo("Успех", message)
                self.show_all_reports_list()
            else:
                messagebox.showerror("Ошибка", f"Ошибка при удалении:\n{message}")