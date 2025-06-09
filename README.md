# DocumentOCR

Приложение для извлечения изображений из документов Word (.docx) и их сохранения.

## Требования

- Python 3.8 или выше
- Windows 10/11

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

5. Запустите приложение:
```bash
python docx_image_parser.py
```

## Использование

1. Запустите приложение
2. Выберите .docx файл через диалоговое окно
3. Изображения будут автоматически извлечены и сохранены в папку `extracted_images`

## Примечания

- Все извлеченные изображения сохраняются в папку `extracted_images`
- Поддерживаются форматы изображений: PNG, JPEG, GIF
- Приложение автоматически создает папку для сохранения изображений, если она не существует 