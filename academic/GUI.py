from processors import AcademicProcessor



processor = AcademicProcessor()

# Добавление заголовка статьи
processor.add_title(
    title='Пример статьи',
    authors='Иванов И.И., Петров П.П.',
    abstract='В данной статье рассматриваются методы ...',
    keywords=['методы', 'пример', 'статья']
)

# Добавление содержания
processor.add_heading('Введение')
processor.add_paragraph('Текст введения...')

processor.add_subheading('Цели и задачи')
processor.add_paragraph('Текст о целях и задачах...')

# Добавление изображения
processor.add_image('path/to/image.jpg', 'Описание изображения')

# Добавление нумерованного списка
processor.add_numbered_list(['Пункт 1', 'Пункт 2', 'Пункт 3'])

# Добавление блока кода
processor.add_code_block('Пример кода', 'print("Hello, World!")')

# Добавление списка литературы
processor.add_bibliography(['Источник 1', 'Источник 2', 'Источник 3'])

# Сохранение документа
processor.save_to_word('article.docx')
