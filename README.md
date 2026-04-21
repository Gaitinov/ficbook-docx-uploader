# Ficbook Text Processor

## English

A tool for convenient preparation and uploading of chapters to Ficbook.net from Word documents.

The application splits large `.docx` files (such as fanfiction or manuscripts) into separate chapters based on headings, while preserving text formatting and styles. It also fixes incorrect paragraph indentation to ensure clean formatting for publication.

Includes a graphical interface for selecting files and directories, along with a progress bar for processing status.

---

## Русский

Инструмент для подготовки и удобной загрузки глав на Ficbook.net из Word-документов.

Программа разбивает большие `.docx` файлы (фанфики, тексты и др.) на отдельные главы по заголовкам, сохраняя форматирование и стили текста. Также исправляет некорректные отступы абзацев для аккуратного отображения при публикации.

Есть графический интерфейс для выбора файлов и папок, а также индикатор прогресса обработки.

---

## Возможности

* Извлечение глав из одного большого `.docx` файла по заголовкам
* Сохранение форматирования и стилей
* Удаление пустых абзацев
* Прогресс-бар для долгих операций (CustomTkinter)
* GUI для выбора файлов и папок

## Использование

* Убедитесь, что заголовки глав оформлены стилем заголовка ("Заголовок 1" или аналогичный)
* Выберите исходный `.docx` файл
* Укажите папку для сохранения или выходной файл
* Следите за прогрессом обработки

## Зависимости

* [python-docx](https://pypi.org/project/python-docx/)
* [customtkinter](https://github.com/TomSchimansky/CustomTkinter)
