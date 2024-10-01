import os
from docx2pdf import convert

def convert_docx_to_pdf(docx_path):
    """
    Функция для конвертации .docx в .pdf с использованием docx2pdf.
    """
    try:
        # Создаем директорию для сохранения PDF, если её нет
        pdf_dir = os.path.dirname(docx_path)
        pdf_path = os.path.join(pdf_dir, f'{os.path.splitext(os.path.basename(docx_path))[0]}.pdf')

        # Конвертируем файл
        convert(docx_path, pdf_path)
        print(f"Документ успешно сконвертирован в PDF по пути: {pdf_path}")

    except Exception as e:
        print(f"Ошибка при конвертации в PDF: {e}")
