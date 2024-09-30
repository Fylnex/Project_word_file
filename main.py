from report import ReportGenerator
from datetime import datetime

if __name__ == "__main__":
    generator = ReportGenerator()

    generator.report_academic(replacements = {
                '${d1}': str(datetime.today().date()).split('-')[0],
                '${s1}': input("название учебного предмета:"),
                '${s2}': input("тип работы:"),
                '${s3}': input("номер работы:"),
                '${s4}': input("номер группы:"),
                '${s10}': input("ФИО:"),
                # '${s20}': input("ФИО:"),
                # '${s30}': input("ФИО:"),
                '${s101}': input("должность преподавателя:"),
                '${s100}': input("ФИО преподавателя:"),

            })