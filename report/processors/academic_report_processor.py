from .document_processor import DocumentProcessor



class AcademicReportProcessor(DocumentProcessor):
    """
    Класс для обработки отчета по сотрудникам.
    """
    def __init__(self, template_path, output_path, replacements, employees_data):
        super().__init__(template_path, output_path, replacements)
        self.employees_data = employees_data

    def process(self):
        """Процесс обработки отчета по научному совету."""
        try:
            self.load_template()
            self.replace_flags()

            self.save_document()
            return True
        except Exception as e:
            print(f"Ошибка при обработке документа: {e}")
            return False

    def replace_flags(self):
        """Замена флагов в документе, включая таблицы."""
        # Проход по всем параграфам документа
        for paragraph in self.doc.paragraphs:
            for flag, value in self.replacements.items():
                if flag in paragraph.text:
                    paragraph.text = paragraph.text.replace(flag, str(value))

        # Проход по всем таблицам в документе
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Проход по всем параграфам в каждой ячейке
                    for paragraph in cell.paragraphs:
                        for flag, value in self.replacements.items():
                            if flag in paragraph.text:
                                paragraph.text = paragraph.text.replace(flag, str(value))


