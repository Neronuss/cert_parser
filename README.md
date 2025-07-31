# Certificate Parser

Программа для анализа SSL-сертификатов

## Особенности
- Многопоточная обработка файлов
- Экспорт в Excel
- Поддержка Linux (требуется Zenity)

## Установка

```bash
# Установка зависимостей
sudo apt install zenity
pip install -r requirements.txt

# Запуск
python3 src/cert_parser.py
```

## Использование
1. Выберите папку с сертификатами (.cer)
2. Укажите файл для сохранения (.xlsx)
3. Дождитесь завершения обработки
