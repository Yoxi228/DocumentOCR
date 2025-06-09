# Document OCR

Приложение для распознавания текста из документов DOCX и PDF, включая текст из изображений внутри документов.

> **Примечание**: В текущей версии приложение работает только на Windows. Поддержка Linux и macOS планируется в будущих версиях.

## Требования

- Windows 10/11
- Python 3.8+
- Tesseract OCR
- Виртуальное окружение Python (рекомендуется)

## Установка

1. Клонируйте репозиторий:
```bash
git clone [URL репозитория]
cd [папка проекта]
```

2. Создайте и активируйте виртуальное окружение:
```bash
python -m venv venv
venv\Scripts\activate
```

3. Установите зависимости:
```bash
pip install -r requirements.txt
```

4. Установите Tesseract OCR:
- Скачайте установщик с [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- Установите в папку по умолчанию (C:\Program Files\Tesseract-OCR)
- При установке выберите дополнительные языки (Russian, English)

## Использование

1. Запустите приложение:
```bash
python gui.py
```

2. В интерфейсе:
- Выберите DOCX или PDF файл
- Выберите язык OCR (Русский/English/Русский+English)
- Нажмите "Распознать"
- Просмотрите результаты
- Сохраните результат в TXT файл

## Поддерживаемые форматы

- DOCX (Microsoft Word)
- PDF

## Поддерживаемые языки OCR

- Русский
- English
- Русский + English

## Сборка EXE

Для создания исполняемого файла:
```bash
pip install pyinstaller
pyinstaller app.spec
```

Готовый EXE файл будет находиться в папке `dist`. 