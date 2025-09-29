# from django.db import models
# from django.core.files.base import ContentFile
# from PIL import Image, EpsImagePlugin
# import io, os
# import subprocess
#
# # Путь к Ghostscript
# EpsImagePlugin.gs_windows_binary = r"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe"
#
# class EventsHandler(models.Model):
#     title = models.CharField("Название", max_length=100)
#     created_at = models.DateTimeField("Дата создание")
#     plotter = models.CharField("Плоттер", max_length=50)
#     preview = models.ImageField("Превью", upload_to="previews/", blank=True, null=True)
#     status = models.CharField("Статус", max_length=20, default="UNKNOWN")
#
#     def __str__(self):
#         return f"{self.title} ({self.plotter})"
#
#     def save(self, *args, **kwargs):
#         """
#         Конвертирует .eps или .cdr в PNG и сохраняет в preview.
#         """
#         super().save(*args, **kwargs)
#
#         if self.preview and self.preview.name.lower().endswith((".eps", ".cdr")):
#             input_path = self.preview.path
#             base, ext = os.path.splitext(self.preview.name)
#             png_name = f"{base}.png"
#
#             if ext.lower() == ".eps":
#                 with Image.open(input_path) as im:
#                     im.load(scale=2)
#                     buffer = io.BytesIO()
#                     im.save(buffer, format="PNG")
#                     buffer.seek(0)
#             elif ext.lower() == ".cdr":
#                 # Используем ImageMagick для конвертации .cdr
#                 output_path = os.path.join(os.path.dirname(input_path), os.path.basename(png_name))
#                 try:
#                     subprocess.run([
#                         "magick", input_path, "-flatten", output_path
#                     ], check=True, capture_output=True)
#                     with open(output_path, "rb") as f:
#                         buffer = io.BytesIO(f.read())
#                     os.remove(output_path)  # Удаляем временный PNG
#                 except subprocess.CalledProcessError as e:
#                     print(f"Ошибка конвертации .cdr: {e}")
#                     return  # Пропускаем сохранение, если конвертация не удалась
#
#             self.preview.save(
#                 os.path.basename(png_name),
#                 ContentFile(buffer.read()),
#                 save=False
#             )
#
#             try:
#                 os.remove(input_path)
#             except OSError:
#                 pass
#
#             super().save(update_fields=["preview"])

from django.db import models
from django.core.files.base import ContentFile
from PIL import Image, EpsImagePlugin
import io, os
import subprocess

# Путь к Ghostscript
EpsImagePlugin.gs_windows_binary = r"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe"

class EventsHandler(models.Model):
    title = models.CharField("Название", max_length=100)
    created_at = models.DateTimeField("Дата создание")
    plotter = models.CharField("Плоттер", max_length=50)
    preview = models.ImageField("Превью", upload_to="previews/", blank=True, null=True)
    status = models.CharField("Статус", max_length=20, default="UNKNOWN")

    def __str__(self):
        return f"{self.title} ({self.plotter})"

    def save(self, *args, **kwargs):
        """
        Конвертирует .eps или .cdr в PNG и сохраняет в preview.
        """
        super().save(*args, **kwargs)

        if self.preview and self.preview.name.lower().endswith((".eps", ".cdr")):
            input_path = self.preview.path
            base, ext = os.path.splitext(self.preview.name)
            png_name = f"{base}.png"

            if ext.lower() == ".eps":
                with Image.open(input_path) as im:
                    im.load(scale=2)
                    buffer = io.BytesIO()
                    im.save(buffer, format="PNG")
                    buffer.seek(0)
            elif ext.lower() == ".cdr":
                output_path = os.path.join(os.path.dirname(input_path), os.path.basename(png_name))
                try:
                    subprocess.run([
                        "magick", input_path, "-flatten", output_path
                    ], check=True, capture_output=True)
                    with open(output_path, "rb") as f:
                        buffer = io.BytesIO(f.read())
                    os.remove(output_path)  # Удаляем временный PNG
                except subprocess.CalledProcessError as e:
                    print(f"Ошибка конвертации .cdr: {e}")
                    return  # Пропускаем сохранение, если конвертация не удалась

            self.preview.save(
                os.path.basename(png_name),
                ContentFile(buffer.read()),
                save=False
            )

            try:
                os.remove(input_path)
            except OSError:
                pass

            super().save(update_fields=["preview"])