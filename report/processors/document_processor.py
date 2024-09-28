from abc import ABC, abstractmethod
from docx import Document

class DocumentProcessor(ABC):
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

    @abstractmethod
    def process(self):
        """Абстрактный метод для обработки документа."""
        pass

    def save_document(self):
        """Сохранение документа."""
        self.doc.save(self.output_path)
        print(f"Документ сохранен по пути: {self.output_path}")
