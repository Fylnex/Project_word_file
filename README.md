# Проект: Модуль автоматизации генерации отчетов в формате DOCX и PDF

Этот проект предназначен для автоматизации процесса создания отчетов в формате DOCX и их последующей конвертации в PDF. Он построен на основе модульной архитектуры с использованием принципов ООП, что позволяет легко добавлять новые сценарии и типы отчетов.

## Структура проекта

Проект состоит из следующих модулей:

    /report
    ├── __init__.py # Инициализация пакета 
    │
    └── generator_report.py # Основной скрипт для запуска 
                             │
                             ├── processors│ 
                             │             ├── init.py                       # Инициализация пакета 
                             │             ├── document_processor.py         # Базовый класс для обработки документов
                             │             └── academic_report_processor.py   # Класс для обработки шаблона
                             │               
                             ├── utils│   
                             │        ├── init.py       # Инициализация пакета 
                             │        └── file_utils.py # Утилитарные функции для работы с файлами 
                             │
                             └── templates│ 
                                          │ 
                                          ├── tmp_1.docx # Шаблон
                                          └── tmp_2.docx # Шаблон 

                               

### Описание модулей

#### 1. `report/generator_report.py`

Основной скрипт для запуска проекта. В этом файле происходит выбор нужного сценария обработки документа (например, отчет по экзамену) и его выполнение. 

Можно добавить новый сценарий, просто импортировав соответствующий класс процессора и создав его экземпляр.



#### 2. `processors/document_processor.py`

Базовый класс `DocumentProcessor` для всех типов процессоров документов. Этот класс определяет общие методы и интерфейс, которые должны быть реализованы в каждом подклассе:
- `load_template()`: Загружает шаблон документа.
- `process()`: Абстрактный метод для обработки документа, который должен быть реализован в каждом подклассе.
- `save_document()`: Сохраняет обработанный документ.



#### 3. `processors/academic_report_processor.py`

Класс `AcademicReportProcessor` — процессор для создания отчета по экзамену. Наследует от `DocumentProcessor` и реализует:
- `process()`: Метод для выполнения всех шагов обработки документа.
- `replace_flags()`: Замена флагов в шаблоне на реальные данные.
- `insert_table()`: Вставка таблицы с данными о студентах на место флага `${table}`.
- `create_table_with_borders()`: Создание таблицы с границами и правильным форматированием.


#### 7. `utils/file_utils.py`

Утилитарные функции для работы с файлами, такие как:
- `convert_docx_to_pdf(docx_path)`: Конвертация документа DOCX в PDF с использованием библиотеки `docx2pdf`.


### Флаги для состояния 'научный совет'
```
replacements = {
                '${d1}': f"{date_today[0]}",
                '${s1}': 'Защита информации от утечки по техническим каналам',
                '${s2}': 'практической',
                '${s3}': n,
                '${s4}': "2251",
                '${s10}': "Чертков Д. Г.",
                '${s20}': "Гусева Е.С.",
                '${s30}': "Скрыленко В.Е.",
                '${s100}': "Кардакова М.В.",

            }
```