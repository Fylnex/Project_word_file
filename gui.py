import eel
from processors import AcademicProcessor
import base64
import os

eel.init('web')

@eel.expose
def generate_document(title, authors, abstract, keywords,
                      paragraph, heading, subheading,
                      imageCaption, codeTitle, codeContent,
                      numberedList, unorderedList, references):

    processor = AcademicProcessor()

    # Добавление заголовка статьи
    processor.add_title(title, authors, abstract, keywords)

    # Добавление заголовков и параграфов
    if heading:
        processor.add_heading(heading)
    if paragraph:
        processor.add_paragraph(paragraph)
    if subheading:
        processor.add_subheading(subheading)

    # Добавление изображения
    # Для передачи изображения из браузера в Python необходимо использовать Base64 или загрузить файл на сервер.
    # В данном случае, для простоты, пропустим добавление изображения.

    # Добавление блока кода
    if codeContent:
        processor.add_code_block(codeTitle, codeContent)

    # Добавление списков
    if numberedList and numberedList != ['']:
        processor.add_numbered_list(numberedList)
    if unorderedList and unorderedList != ['']:
        processor.add_unordered_list(unorderedList)

    # Добавление списка литературы
    if references and references != ['']:
        processor.add_bibliography(references)

    # Сохранение документа
    output_path = 'article.docx'
    processor.save_to_word(output_path)

    return f"Документ '{output_path}' успешно создан."

# Запуск приложения Eel
eel.start('index.html', size=(900, 1000))
