# Утилита сбора статистики

## Проект на начальной стадии разработки

### Описание
Этот проект представляет собой утилиту для анализа статистики документов DOCX и XLSX.

### Функциональность
**Для документов XLSX:**
- Размер рабочей области (указанный и фактический)
- Общее количество ячеек на листе
- Количество ячеек с данными, числовыми данными, текстовыми данными и пустыми ячейками
- Количество ячеек с формулами и частота встречаемости различных формул
- Количество правил условного форматирования и их типы
- Количество ячеек с внешними ссылками и заливкой
- Примененные фильтры и количество сводных таблиц, умных таблиц и диаграмм
- Количество изображений и их размеры

**Для документов DOCX:**
- Количество изображений в документе
- Количество таблиц в документе и их размерность
- Метрики документа: количество слов, предложений, абзацев и символов
- Метрики заголовков: уровень заголовка и частота встречаемости
- Подсчет символов форматирования: жирный, курсивный и подчеркнутый форматы

### Использование
Просто запустите скрипт main.py, указав путь к анализируемому документу в качестве аргумента командной строки.

```bash
python main.py <путь_к_документу>