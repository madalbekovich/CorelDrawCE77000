#!/bin/bash

# Запускаем Django в фоне
python core/manage.py runserver 0.0.0.0:8000 &

# Запускаем скрипт ready.py в фоне
python ready.py &

# Держим контейнер живым
tail -f /dev/null
