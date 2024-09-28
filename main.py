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

    # Дата
    date_today = str(datetime.today().date()).split('-')

    # Пример значений для замены
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
        '${d1}': date_today[2],
        '${d2}': date_today[1],
        '${d3}': date_today[0],
        '${p50}': 'ФИО зав кафедры',
        '${p100}': 'ФИО декана факультета',
        '${p1000}': 'Декан факультета',
        '${p201}': 'Абдрахимов Радик Робертович',
        '${p202}': 'Абдрахимов Радик Робертович',
        '${p203}': 'Абдрахимов Радик Робертович',
        '${p204}': 'Абдрахимов Радик Робертович',
        '${p205}': 'Абдрахимов Радик Робертович'
    }

    # Студенты
    students_data = [
        {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },
  {
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        },{
            "Фамилия и инициалы": "Абдрахимов Радик Робертович",
            "Группа": "Л-301",
            "№ зач. книжки": "Л21013"

        }
    ]

    #
    # print(name_file_word)


    # Выбираем и запускаем нужный процессор
    processor = ExamReportProcessor(template_path, output_path, replacements, students_data)
    if processor.process():
        convert_docx_to_pdf(output_path)



def report_council():

    # путь к шаблону
    template_path = "templates/council/form_app_01.docx"

    # получение названия директории
    template_filename = template_path.split('/')


    # создание названия файла, для вывода в будущем
    # name_file_word = f"{template_filename[1]}_{datetime.today().date()}_{random.randrange(10 ** 8, 10 ** 10)}.docx" # создание названия файла
    name_file_word = f"{template_filename[1]}_{datetime.today().date()}.docx"

    # путь для сохранения документа
    download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    output_path = os.path.join(download_path, f"{name_file_word}")


    # Дата
    date_today =str(datetime.today().date()).split('-')

    # Пример данных участников совета

    employees_data = [
        {
            'ФАМИЛИЯ, И., О.': 'Туричин Глеб Андреевич',
            'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': 'д.т.н., доцент',
            'ДОЛЖНОСТЬ': 'ректор',

        },
        {
            'ФАМИЛИЯ, И., О.': 'Сайченко  Ольга Анатольевна ',
            'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': 'к.э.н., доцент',
            'ДОЛЖНОСТЬ': 'проректор по ОД',

        },
        {
            'ФАМИЛИЯ, И., О.': 'Кузнецов Денис Иванович',
            'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': 'д.т.н., доцент',
            'ДОЛЖНОСТЬ': 'проректор  по научной работе',

        },
        {
            'ФАМИЛИЯ, И., О.': 'Акопян Альберт Беникович',
            'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': '-, кап. 1-го ранга запаса',
            'ДОЛЖНОСТЬ': 'проректор по воспитательной работе',

        },

        {
            'ФАМИЛИЯ, И., О.': 'Прокопенко Андрей Петрович',
            'УЧЕНАЯ СТЕПЕНЬ, УЧЕНОЕ ЗВАНИЕ': '-',
            'ДОЛЖНОСТЬ': 'проректор по безопасности образовательного процесса',

        }




    ]

    # Замены в шаблоне
    replacements = {
        '${d1}': f"{date_today[2]}.{date_today[1]}.{date_today[0]}", #date_today,
        '${s1}': '08',
        '${s2}': '24',
        '${s3}': f'{len(employees_data)}',
        '${s4}': f"{int(len(employees_data)*2/3)}",

    }



    processor = CouncilReportProcessor(template_path, output_path, replacements, employees_data)
    if processor.process():
        convert_docx_to_pdf(output_path)







if __name__ == "__main__":
    a= input("1 - отчет о совете\n2 - отчет о экзамене\nВыберете состояние:")
    if a == '1':
        report_council()
    elif a == '2':
        report_exam()
    # report_council()
