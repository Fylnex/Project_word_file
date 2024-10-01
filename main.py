
from report import ReportGenerator

if __name__ == "__main__":
    generator = ReportGenerator()
    choice = input(
        "1 - Отчет о совете\n"
        "2 - Отчет об экзамене\n"
        "3 - Отчет о студенте\n"
        "4 - Отчет о сотруднике\n"
        "Выберите вариант ( ͡° ͜ʖ ͡°): "
    )
    if choice == '1':
        generator.council_report()
    elif choice == '2':
        generator.exam_report()
    elif choice == '3':
        generator.student_report()
    elif choice == '4':
        generator.employee_report()
    else:
        print("¯\_(ツ)_/¯")
