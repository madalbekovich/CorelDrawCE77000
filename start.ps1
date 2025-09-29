Start-Process "python" -ArgumentList "ready.py"
Start-Process "python" -ArgumentList "core\\manage.py runserver 0.0.0.0:8000"

while ($true) { Start-Sleep -Seconds 3600 }
