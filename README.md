# DocumentOCR

Приложение для извлечения изображений из документов Word (.docx) и их сохранения.

## Требования

- Python 3.8 или выше
- Windows 10/11
- Tesseract OCR

## Установка и запуск

1. Клонируйте репозиторий:
```bash
git clone https://github.com/Yoxi228/DocumentOCR.git
cd DocumentOCR
```

2. Создайте виртуальное окружение:
```bash
python -m venv venv
```

3. Активируйте виртуальное окружение:
```bash
# В PowerShell:
.\venv\Scripts\Activate.ps1

# В Command Prompt (cmd.exe):
.\venv\Scripts\activate.bat
```

4. Установите зависимости:
```bash
pip install -r requirements.txt
```

5. Установите Tesseract OCR:
   - Скачайте установщик с [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
   - Установите в папку по умолчанию (C:\Program Files\Tesseract-OCR)
   - При установке выберите дополнительные языки (Russian, English)

## Запуск приложения

### Вариант 1: Графический интерфейс (рекомендуется)
```bash
python gui.py
```

В открывшемся окне:
- Выберите .docx файл через кнопку "Выбрать файл"
- Выберите язык OCR (Русский/English/Русский+English)
- Нажмите "Обработать"

### Вариант 2: Консольный режим
```bash
python docx_image_parser.py
```

В консольном режиме:
- Откроется диалоговое окно для выбора файла
- После выбора файла начнется автоматическая обработка
- Результаты будут сохранены в папку `extracted_images`

## Результаты

- Все извлеченные изображения сохраняются в папку `extracted_images`
- Поддерживаются форматы изображений: PNG, JPEG, GIF
- Приложение автоматически создает папку для сохранения изображений, если она не существует
- Результаты обработки сохраняются в файл `extracted_text.txt` в папке `extracted_images` 