from .document_processor import DocumentProcessor  # Импорт базового класса
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx import Document

class ExamReportProcessor(DocumentProcessor):
    """
    Класс для обработки отчета по экзамену.
    """
    def __init__(self, template_path, output_path, replacements, students_data):
        super().__init__(template_path, output_path, replacements)
        self.students_data = students_data

    def process(self):
        """Процесс обработки отчета по экзамену."""
        try:
            self.load_template()
            self.replace_flags()
            self.insert_table()  # Вставка таблицы на место флага ${table}
            self.save_document()
            return True
        except Exception as e:
            print(f"Ошибка при обработке документа: {e}")
            return False

    def insert_table(self):
        """Вставка таблицы с данными студентов на место флага ${table}."""
        for i, paragraph in enumerate(self.doc.paragraphs):
            if "${table}" in paragraph.text:
                # Заменяем флаг ${table} на пустую строку
                paragraph.text = paragraph.text.replace("${table}", "")

                # Вставляем таблицу на место флага
                table = self.create_table_with_borders()

                # Вставляем таблицу после абзаца с флагом
                paragraph._element.addnext(table._element)
                break

    def create_table_with_borders(self):
        """Создание таблицы с границами и правильным форматированием столбцов."""
        # Создаем таблицу с двумя строками для заголовков и количеством столбцов 7
        table = self.doc.add_table(rows=2, cols=7)

        # Применяем стиль к таблице для отображения всех границ
        # table.style = 'Table Grid'

        # Объединяем ячейки заголовков с нижними ячейками
        table.cell(0, 0).merge(table.cell(1, 0))  # №
        table.cell(0, 1).merge(table.cell(1, 1))  # Фамилия и инициалы
        table.cell(0, 2).merge(table.cell(1, 2))  # Группа
        table.cell(0, 3).merge(table.cell(1, 3))  # № зач. книжки
        table.cell(0, 4).merge(table.cell(0, 5))  # Экзаменационная оценка
        table.cell(0, 6).merge(table.cell(1, 6))  # Подпись экзаменатора

        # Устанавливаем текст для объединенных ячеек
        table.cell(0, 0).text = '№'
        table.cell(0, 1).text = 'Фамилия и инициалы'
        table.cell(0, 2).text = 'Группа'
        table.cell(0, 3).text = '№ зач. книжки'
        table.cell(0, 4).text = 'Экзаменационная оценка'
        table.cell(1, 4).text = 'цифрой'  # Подзаголовок
        table.cell(1, 5).text = 'прописью'  # Подзаголовок
        table.cell(0, 6).text = 'Подпись экзаменатора'

        # Добавляем строки для каждого студента
        for student in self.students_data:
            row = table.add_row().cells
            row[0].text = student.get('№', '')
            row[1].text = student.get('Фамилия и инициалы', '')
            row[2].text = student.get('Группа', '')
            row[3].text = student.get('№ зач. книжки', '')
            row[4].text = student.get('Экзаменационная оценка цифрой', '')
            row[5].text = student.get('Экзаменационная оценка прописью', '')
            row[6].text = student.get('Подпись экзаменатора', '')

        # Устанавливаем границы для каждой ячейки таблицы
        self.set_table_borders(table)

        return table

    def set_table_borders(self, table):
        """Добавление всех границ для каждой ячейки таблицы."""
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._element.get_or_add_tcPr()
                tcBorders = OxmlElement('w:tcBorders')

                # Создаем границы для каждой стороны ячейки
                for side in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{side}')
                    border.set(qn('w:val'), 'single')  # Сплошная линия
                    border.set(qn('w:sz'), '4')  # Толщина границы
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), '000000')  # Черный цвет границы
                    tcBorders.append(border)

                tcPr.append(tcBorders)