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
            # template_path = os.path.join(base_dir, 'templates', 'tmp_1.docx')
            template_path = os.path.join(base_dir, 'templates', 'tmp_3_diploma.docx')


        n=1

        if output_path is None:
            name_file_word = f"tmp_{datetime.today().date()}.docx"
            download_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            output_path = os.path.join(download_path, f"{name_file_word}")

        date_today = str(datetime.today().date()).split('-')



        if replacements is None:
            replacements = {
                '${d1}': f"{date_today[0]}",
                '${s1}': 'название учебно предмета',
                '${s2}': 'тип работы (практическая/лабораторная)',
                '${s3}': "номер работы",
                '${s4}': "номер группы",
                '${s5}': "",
                '${s10}': "ФИО1",
                '${s20}': "ФИО2",
                '${s30}': "ФИО3",
                '${s101}': "должность преподавателя",
                '${s100}': "ФИО преподавателя"
            }

        processor = AcademicReportProcessor(template_path, output_path, replacements, employees_data)
        if processor.process():
            print(f"Отчет сохранен в {output_path}")
