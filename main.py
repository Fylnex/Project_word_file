from processors import ExamReportProcessor, CouncilReportProcessor ,StudentReportProcessor,EmployeeReportProcessor
from utils import convert_docx_to_pdf
import os
import random
from datetime import datetime


def report_exam():
    # путь к шаблону
    template_path = "templates/exam/24_form_22.docx"
    # template_path="templates/exam/24_form_21.docx"
    # template_path="templates/exam/24_form_20.docx"

    template_filename = template_path.split('/') # получение названия директории, для удобного сохранения файла

    # name_file_word = f"{template_filename[1]}_{datetime.today().date()}_{random.randrange(10 ** 8, 10 ** 10)}.docx" # создание названия файла
    name_file_word = f"{template_filename[1]}_{datetime.today().date()}.docx" # создание названия файла, для вывода в будущем

    # путь для сохранения документа
    download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    output_path = os.path.join(download_path, f"{name_file_word}")

    # Даты
    date_today = str(datetime.today().date()).split('-')

    # Пример значений для замены
    replacements = {
        '${s1}': '240025',
        '${s2}': '2',
        '${s3}': '2023',
        '${s4}': '2024',
        '${s5}': '31.05.01 ЛЕЧЕБНОЕ ДЕЛО',
        '${s6}': '',
        '${s7}': '3',
        '${s8}': 'Л-301',
        '${s9}': 'Факультет Лечебный факультет',
        '${s10}': 'Биоэтика',
        '${s11}': '72 часов / 2 з.ед.',
        '${s12}': 'Бухарин Василий Фёдорович',
        '${d1}': date_today[2],
        '${d2}': date_today[1],
        '${d3}': f'{date_today[0]} ',
        '${p50}': 'ФИО зав кафедры',
        '${p100}': 'ФИО декана факультета',
        '${p1000}': 'Декан факультета',
        '${p201}': 'ФИО преподавателя 1',
        '${p202}': 'ФИО преподавателя 2',
        '${p203}': 'ФИО преподавателя 3',
        '${p204}': 'ФИО преподавателя 4',
        '${p205}': 'ФИО преподавателя 5'
    }

    # Пример данных студентов
    students_data = [
        {
            'Фамилия и инициалы': 'Абдрахимов Радик Робертович',
            'Группа': 'Л-301',
            '№ зач. книжки': 'Л21013',
            'Экзаменационная оценка цифрой': '',
            'Экзаменационная оценка прописью': '',
            'Подпись экзаменатора': ''
        },
        {
            'Фамилия и инициалы': 'Абдрахимов Радик Робертович',
            'Группа': 'Л-301',
            '№ зач. книжки': 'Л21013',
            'Экзаменационная оценка цифрой': '',
            'Экзаменационная оценка прописью': '',
            'Подпись экзаменатора': ''
        },
        {
            'Фамилия и инициалы': 'Радик Абдрахимов Робертович',
            'Группа': 'Л-301',
            '№ зач. книжки': 'Л21013',
            'Экзаменационная оценка цифрой': '',
            'Экзаменационная оценка прописью': '',
            'Подпись экзаменатора': ''
        },
        {
            'Фамилия и инициалы': 'Радик Абдрахимов Робертович',
            'Группа': 'Л-301',
            '№ зач. книжки': 'Л21013',
            'Экзаменационная оценка цифрой': '',
            'Экзаменационная оценка прописью': '',
            'Подпись экзаменатора': ''
        }
    ]

    print(name_file_word)


    # Выбираем и запускаем нужный процессор
    processor = ExamReportProcessor(template_path, output_path, replacements, students_data)
    if processor.process():
        convert_docx_to_pdf(output_path)

def report_council():

    # путь к шаблону
    template_path = "templates/council/form_app_05.docx"


    template_filename = template_path.split('/')  # получение названия директории, для удобного сохранения файла

    # name_file_word = f"{template_filename[1]}_{datetime.today().date()}_{random.randrange(10 ** 8, 10 ** 10)}.docx" # создание названия файла
    name_file_word = f"{template_filename[1]}_{datetime.today().date()}.docx"  # создание названия файла, для вывода в будущем

    # путь для сохранения документа
    download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    output_path = os.path.join(download_path, f"{name_file_word}")


    # Даты
    date_today =str(datetime.today().date()).split('-')

    # Замены в шаблоне
    replacements = {
        '${d1}': date_today[2],
        '${d2}': date_today[1],
        '${d3}': date_today[0],

    }

    # Пример данных участников совета
    employees_data = [
        {
            'ФИО': 'Абдрахимов Радик Робертович',
            'Ученая степень': 'доцент',
            'Должность': 'старший преподаватель',
            'Организация': 'ИТМО',
            'Подпись': '',
        },
        {
            'ФИО': 'Абдрахимов Радик Робертович',
            'Ученая степень': 'доцент',
            'Должность': 'старший преподаватель',
            'Организация': 'ИТМО',
            'Подпись': '',
        },
        {
            'ФИО': 'Абдрахимов Радик Робертович',
            'Ученая степень': 'доцент',
            'Должность': 'старший преподаватель',
            'Организация': 'ИТМО',
            'Подпись': '',
        },
    ]

    processor = CouncilReportProcessor(template_path, output_path, replacements, employees_data)
    if processor.process():
        convert_docx_to_pdf(output_path)


if __name__ == "__main__":
    report_council()
