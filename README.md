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

6. Запустите приложение:
```bash
python gui.py
```

## Использование

1. Запустите приложение через `gui.py`
2. В открывшемся окне:
   - Выберите .docx файл через кнопку "Выбрать файл"
   - Выберите язык OCR (Русский/English/Русский+English)
   - Нажмите "Обработать"
3. Изображения будут автоматически извлечены и сохранены в папку `extracted_images`

## Примечания

- Все извлеченные изображения сохраняются в папку `extracted_images`
- Поддерживаются форматы изображений: PNG, JPEG, GIF
- Приложение автоматически создает папку для сохранения изображений, если она не существует 