FROM python:3.11-slim

# Рабочая папка внутри контейнера
WORKDIR /app

# Копируем зависимости
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# Копируем весь проект внутрь контейнера
COPY . .

# Экспортируем порт Django
EXPOSE 8000

# Копируем скрипт запуска
COPY start.sh .
RUN chmod +x start.sh

# Запускаем start.sh, который поднимает Django + ready.py
CMD ["./start.sh"]


