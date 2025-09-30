"""
Модуль для работы с базой данных SQLite
Управление отчетами и ответами
"""

import sqlite3
from datetime import datetime


def get_connection():
    """Получить соединение с БД"""
    conn = sqlite3.connect('reports.db')
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_database():
    """Инициализация базы данных - создание таблиц"""
    conn = get_connection()
    cursor = conn.cursor()

    # Таблица отчетов
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            form_name TEXT NOT NULL,
            month TEXT NOT NULL,
            year INTEGER NOT NULL,
            report_date TEXT NOT NULL,
            created_at TEXT NOT NULL,
            file_path TEXT NOT NULL
        )
    ''')

    # Таблица ответов
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS answers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id INTEGER NOT NULL,
            question_text TEXT NOT NULL,
            answer_yes_no TEXT NOT NULL,
            comment TEXT,
            gost_text TEXT,
            quality_text TEXT,
            documents_text TEXT,
            FOREIGN KEY (report_id) REFERENCES reports (id) ON DELETE CASCADE
        )
    ''')

    # Индексы
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_answers_report_id 
        ON answers(report_id)
    ''')

    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_reports_created_at 
        ON reports(created_at)
    ''')

    conn.commit()
    conn.close()
    print("База данных инициализирована")


def save_report_to_db(report_data, answers_list, file_path):
    """Сохранить отчет и ответы в базу данных"""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute('''
            INSERT INTO reports (form_name, month, year, report_date, created_at, file_path)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            report_data['form_name'],
            report_data['month'],
            report_data['year'],
            report_data['report_date'],
            datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
            file_path
        ))

        report_id = cursor.lastrowid

        for answer in answers_list:
            cursor.execute('''
                INSERT INTO answers (report_id, question_text, answer_yes_no, comment, gost_text, quality_text, documents_text)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                report_id,
                answer['question_text'],
                answer['answer_yes_no'],
                answer['comment'],
                answer['gost_text'],
                answer['quality_text'],
                answer.get('documents_text', '')
            ))

        conn.commit()
        print(f"Отчет сохранен в БД с ID: {report_id}")
        return report_id

    except Exception as e:
        conn.rollback()
        print(f"Ошибка при сохранении в БД: {e}")
        raise
    finally:
        conn.close()


def get_all_reports():
    """Получить список всех отчетов"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT id, form_name, month, year, report_date, created_at, file_path
        FROM reports
        ORDER BY created_at DESC
    ''')

    rows = cursor.fetchall()
    conn.close()

    reports = []
    for row in rows:
        reports.append({
            'id': row['id'],
            'form_name': row['form_name'],
            'month': row['month'],
            'year': row['year'],
            'report_date': row['report_date'],
            'created_at': row['created_at'],
            'file_path': row['file_path']
        })

    return reports


def get_report_by_id(report_id):
    """Получить отчет по ID со всеми ответами"""
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT id, form_name, month, year, report_date, created_at, file_path
        FROM reports
        WHERE id = ?
    ''', (report_id,))

    report_row = cursor.fetchone()

    if not report_row:
        conn.close()
        return None

    cursor.execute('''
        SELECT question_text, answer_yes_no, comment, gost_text, quality_text, documents_text
        FROM answers
        WHERE report_id = ?
        ORDER BY id
    ''', (report_id,))

    answer_rows = cursor.fetchall()
    conn.close()

    report_data = {
        'id': report_row['id'],
        'form_name': report_row['form_name'],
        'month': report_row['month'],
        'year': report_row['year'],
        'report_date': report_row['report_date'],
        'created_at': report_row['created_at'],
        'file_path': report_row['file_path'],
        'answers': []
    }

    for answer_row in answer_rows:
        report_data['answers'].append({
            'question_text': answer_row['question_text'],
            'answer_yes_no': answer_row['answer_yes_no'],
            'comment': answer_row['comment'],
            'gost_text': answer_row['gost_text'],
            'quality_text': answer_row['quality_text'],
            'documents_text': answer_row.get('documents_text', '')
        })

    return report_data


def delete_report(report_id):
    """Удалить отчет из базы данных"""
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute('DELETE FROM reports WHERE id = ?', (report_id,))
        conn.commit()
        print(f"Отчет {report_id} удален из БД")
    except Exception as e:
        conn.rollback()
        print(f"Ошибка при удалении: {e}")
        raise
    finally:
        conn.close()