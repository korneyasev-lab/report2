"""
Модуль бизнес-логики для системы автоматизации отчетов
"""

import os
import openpyxl
from database import save_report_to_db, get_all_reports, get_report_by_id, delete_report
from export_excel import create_excel_report


class ReportLogic:
    """Класс с бизнес-логикой приложения"""

    def __init__(self):
        self.current_report_data = {}
        self.questions_list = []
        self.answers_list = []
        self.current_question_index = 0
        self.questions_per_page = 5

    def load_forms_list(self):
        """Загрузка списка форм из папки 'формы/'"""
        forms_dir = "формы"
        if not os.path.exists(forms_dir):
            return []
        files = os.listdir(forms_dir)
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
        return [os.path.splitext(f)[0] for f in excel_files]

    def load_questions_from_excel(self, file_path):
        """Загрузка вопросов из Excel файла"""
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            self.questions_list = []

            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:
                    self.questions_list.append({
                        'question': row[0] if row[0] else "",
                        'gost': row[1] if len(row) > 1 and row[1] else "",
                        'quality': row[2] if len(row) > 2 and row[2] else "",
                        'documents': row[3] if len(row) > 3 and row[3] else ""
                    })

            wb.close()
            return len(self.questions_list) > 0

        except Exception as e:
            print(f"Ошибка при загрузке Excel: {e}")
            return False

    def init_report(self, form_name, month, year, report_date):
        """Инициализация нового отчета"""
        self.current_report_data = {
            'form_name': form_name,
            'month': month,
            'year': year,
            'report_date': report_date
        }

        self.answers_list = [{
            'question_text': q['question'],
            'answer_yes_no': '',
            'comment': '',
            'gost_text': q['gost'],
            'quality_text': q['quality'],
            'documents_text': q['documents']
        } for q in self.questions_list]

        self.current_question_index = 0
        return True

    def get_current_block_questions(self):
        """Получить вопросы текущего блока"""
        start = self.current_question_index
        end = min(start + self.questions_per_page, len(self.questions_list))
        return start, end

    def save_answer(self, question_index, answer_yes_no, comment):
        """Сохранить ответ на вопрос"""
        if 0 <= question_index < len(self.answers_list):
            self.answers_list[question_index]['answer_yes_no'] = answer_yes_no
            self.answers_list[question_index]['comment'] = comment
            return True
        return False

    def check_all_answered(self):
        """Проверить что все вопросы отвечены"""
        for i, answer in enumerate(self.answers_list):
            if not answer['answer_yes_no']:
                return False, i + 1
        return True, None

    def next_block(self):
        """Переход к следующему блоку"""
        next_start = self.current_question_index + self.questions_per_page
        if next_start >= len(self.questions_list):
            return False
        self.current_question_index = next_start
        return True

    def prev_block(self):
        """Возврат к предыдущему блоку"""
        prev_start = self.current_question_index - self.questions_per_page
        if prev_start < 0:
            prev_start = 0
        self.current_question_index = prev_start
        return True

    def save_report(self):
        """Сохранение отчета в БД и экспорт в Excel"""
        try:
            report_name = f"Отчет: {self.current_report_data['form_name']} {self.current_report_data['month']} {self.current_report_data['year']}"

            file_path = create_excel_report(
                report_name=report_name,
                form_name=self.current_report_data['form_name'],
                month=self.current_report_data['month'],
                year=self.current_report_data['year'],
                answers=self.answers_list
            )

            report_id = save_report_to_db(self.current_report_data, self.answers_list, file_path)

            return True, os.path.basename(file_path)

        except Exception as e:
            return False, str(e)

    def get_all_reports_from_db(self):
        """Получить все отчеты из БД"""
        return get_all_reports()

    def get_report_from_db(self, report_id):
        """Получить конкретный отчет из БД"""
        return get_report_by_id(report_id)

    def delete_report_from_db(self, report_id):
        """Удалить отчет из БД"""
        try:
            delete_report(report_id)
            return True, "Отчет удален"
        except Exception as e:
            return False, str(e)

    def export_report_to_word(self, report_data):
        """Экспортировать отчет в Excel заново"""
        try:
            report_name = f"Отчет: {report_data['form_name']} {report_data['month']} {report_data['year']}"

            file_path = create_excel_report(
                report_name=report_name,
                form_name=report_data['form_name'],
                month=report_data['month'],
                year=report_data['year'],
                answers=report_data['answers']
            )

            return True, os.path.basename(file_path)
        except Exception as e:
            return False, str(e)