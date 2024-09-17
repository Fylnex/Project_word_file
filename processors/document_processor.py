from docx import Document

class DocumentProcessor:
    """
    Базовый класс для обработки документов.
    """
    def __init__(self, template_path, output_path, replacements):
        self.template_path = template_path
        self.output_path = output_path
        self.replacements = replacements
        self.doc = None

    def load_template(self):
        """Загрузка шаблона документа."""
        self.doc = Document(self.template_path)
        print(f"Шаблон загружен: {self.template_path}")

    def replace_flags(self):
        """Замена флагов в документе."""
        for paragraph in self.doc.paragraphs:
            for flag, value in self.replacements.items():
                if flag in paragraph.text:
                    paragraph.text = paragraph.text.replace(flag, str(value))
        print("Флаги успешно заменены в документе.")

    def save_document(self):
        """Сохранение документа."""
        self.doc.save(self.output_path)
        print(f"Документ сохранен по пути: {self.output_path}")
