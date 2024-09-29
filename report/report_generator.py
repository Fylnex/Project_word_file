from .processors import (
    ExamReportProcessor,
    CouncilReportProcessor,
    StudentReportProcessor,
    EmployeeReportProcessor,
)
from .utils import convert_docx_to_pdf
import os
from datetime import datetime

class ReportGenerator:
    def __init__(self):
        # Вы можете добавить любые общие настройки или переменные здесь
        pass

    def exam_report(self, template_path=None, replacements=None, students_data=None, output_path=None):
        # Установка пути к шаблону по умолчанию, если не указан
        if template_path is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, 'templates', 'exam', '24_form_20.docx')


        # Создание названия файла, если output_path не указан
        if output_path is None:
            name_file_word = f"exam_{datetime.today().date()}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        # Текущая дата
        date_today = str(datetime.today().date()).split('-')

        # Замены по умолчанию, если replacements не указаны
        if replacements is None:
            replacements = {
                '${s1}': '240025',
                '${s2}': '2',
                '${s3}': '2023',
                '${s4}': '2024',
                '${s5}': '31.05.01 ЛЕЧЕБНОЕ ДЕЛО',
                '${s6}': 'Лечебное дело',
                '${s7}': '3',
                '${s8}': 'Л-301',
                '${s9}': 'Факультет Лечебный факультет',
                '${s10}': 'Биоэтика',
                '${s11}': '72 часов / 2 з.ед.',
                '${s12}': 'Бухарин Василий Фёдорович',
                '${d1}': f"{date_today[2]}.{date_today[1]}.{date_today[0]} ",
                '${p50}': 'ФИО зав кафедры',
                '${p100}': 'ФИО декана факультета',
                '${p1000}': 'Декан факультета',
                '${p201}': 'Абдрахимов Радик Робертович',
                '${p202}': 'Абдрахимов Радик Робертович',
                '${p203}': 'Абдрахимов Радик Робертович',
                '${p204}': 'Абдрахимов Радик Робертович',
                '${p205}': 'Абдрахимов Радик Робертович'
            }

        # Данные студентов по умолчанию, если не указаны
        if students_data is None:
            students_data = [
                {
                    "Фамилия и инициалы": "Абдрахимов Радик Робертович",
                    "Группа": "Л-301",
                    "№ зач. книжки": "Л21013"
                },
                # Добавьте другие записи студентов при необходимости
            ]

        # Создание и запуск нужного процессора
        processor = ExamReportProcessor(template_path, output_path, replacements, students_data)
        if processor.process():
            convert_docx_to_pdf(output_path)
            print(f"Отчет об экзамене сохранен в {output_path}")


    def council_report(self, template_path=None, replacements=None, employees_data=None, output_path=None):
        if template_path is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, 'templates', 'council', 'form_app_01.docx')



        if output_path is None:
            name_file_word = f"council_{datetime.today().date()}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        date_today = str(datetime.today().date()).split('-')

        if employees_data is None:
            employees_data = [
                {
                    'ФАМИЛИЯ, И., О.': 'Туричин Глеб Андреевич',
                    'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': 'д.т.н., доцент',
                    'ДОЛЖНОСТЬ': 'ректор',
                },
                # Добавьте другие записи сотрудников при необходимости
            ]

        if replacements is None:
            replacements = {
                '${d1}': f"{date_today[2]}.{date_today[1]}.{date_today[0]}",
                '${s1}': '08',
                '${s2}': '24',
                '${s3}': f'{len(employees_data)}',
                '${s4}': f"{int(len(employees_data)*2/3)}",
            }

        processor = CouncilReportProcessor(template_path, output_path, replacements, employees_data)
        if processor.process():
            convert_docx_to_pdf(output_path)
            print(f"Отчет о совете сохранен в {output_path}")

    def student_report(self, template_path=None, replacements=None, students_data=None, output_path=None):
        if template_path is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, 'templates', 'student', 'form_app_01.docx')


        if output_path is None:
            name_file_word = f"student_{datetime.today().date()}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        date_today = str(datetime.today().date()).split('-')

        if students_data is None:
            students_data = [
                {
                    'ФИО': 'Иванов Иван Иванович',
                    'Группа': 'Л-301',
                    'Зачетная книжка': 'Л21013',
                },
                # Добавьте другие записи студентов при необходимости
            ]

        if replacements is None:
            replacements = {
                '${d1}': f"{date_today[2]}.{date_today[1]}.{date_today[0]}",
                '${s1}': '08',
                '${s2}': '24',
                '${s3}': f'{len(students_data)}',
                '${s4}': f"{int(len(students_data)*2/3)}",
            }

        processor = StudentReportProcessor(template_path, output_path, replacements, students_data)
        if processor.process():
            convert_docx_to_pdf(output_path)
            print(f"Отчет о студенте сохранен в {output_path}")

    def employee_report(self, template_path=None, replacements=None, employees_data=None, output_path=None):
        if template_path is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, 'templates', 'employee', 'form_app_01.docx')



        if output_path is None:
            name_file_word = f"employee_{datetime.today().date()}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        date_today = str(datetime.today().date()).split('-')

        if employees_data is None:
            employees_data = [
                {
                    'ФАМИЛИЯ, И., О.': 'Петров Петр Петрович',
                    'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': 'к.т.н., доцент',
                    'ДОЛЖНОСТЬ': 'старший преподаватель',
                },
                # Добавьте другие записи сотрудников при необходимости
            ]

        if replacements is None:
            replacements = {
                '${d1}': f"{date_today[2]}.{date_today[1]}.{date_today[0]}",
                '${s1}': '08',
                '${s2}': '24',
                '${s3}': f'{len(employees_data)}',
                '${s4}': f"{int(len(employees_data)*2/3)}",
            }

        processor = EmployeeReportProcessor(template_path, output_path, replacements, employees_data)
        if processor.process():
            convert_docx_to_pdf(output_path)
            print(f"Отчет о сотруднике сохранен в {output_path}")
