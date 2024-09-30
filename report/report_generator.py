from .processors import AcademicReportProcessor
import os
from datetime import datetime

class ReportGenerator:
    def __init__(self):
        # Вы можете добавить любые общие настройки или переменные здесь
        pass


    def report_academic(self, template_path=None, replacements=None, employees_data=None, output_path=None):
        if template_path is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, 'templates', 'tmp_1.docx')


        n=1

        if output_path is None:
            name_file_word = f"практика_{n}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        date_today = str(datetime.today().date()).split('-')



        if replacements is None:
            replacements = {
                '${d1}': f"{date_today[0]}",
                '${s1}': 'Защита информации от утечки по техническим каналам',
                '${s2}': 'практической',
                '${s3}': n,
                '${s4}': "2251",
                '${s10}': "Чертков Д. Г.",
                '${s20}': "Гусева Е.С.",
                '${s30}': "Скрыленко В.Е.",
                '${s100}': "Кардакова М.В.",

            }

        processor = AcademicReportProcessor(template_path, output_path, replacements, employees_data)
        if processor.process():
            print(f"Отчет сохранен в {output_path}")
