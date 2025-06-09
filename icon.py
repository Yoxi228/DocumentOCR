from PIL import Image, ImageDraw

# Создаем изображение 256x256 пикселей
img = Image.new('RGBA', (256, 256), color=(0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Рисуем круг
draw.ellipse([20, 20, 236, 236], fill='#0078d4')

# Рисуем букву "O" внутри круга
draw.ellipse([60, 60, 196, 196], fill='#1e1e1e')

# Сохраняем как ICO файл
img.save('app_icon.ico', format='ICO') 