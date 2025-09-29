import time
import getpass
import os
import shutil
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import datetime, timezone, timedelta
import difflib
import hashlib
import threading
from PIL import Image, ImageDraw, ImageFont, EpsImagePlugin
import io
import subprocess
import tempfile
import win32com.client
import win32gui
import win32con
from PIL import ImageGrab
import win32ui
import win32api
from ctypes import windll
import psutil

# --- Настройки ---
PATHS_TO_WATCH = [
    r"C:\Program Files (x86)\Graphtec\Cutting Master 4\Jobs and Settings\Jobs"
]
ERROR_LOG = r"C:\Program Files (x86)\Graphtec\Cutting Master 4\Temp\saLog.log"
DESIGNS_DIR = r"C:\Users\Lenovo\Desktop\bronoskins"
DJANGO_API = "https://coreldrawce77000.onrender.com/"
MEDIA_PREVIEWS = r"C:\Users\Lenovo\PyCharmMiscProject\media\previews"

PRINTER = "Graphtec CE7000"
USER = getpass.getuser()

# --- Новые настройки для скриншотов Cutting Master 4 ---
ENABLE_CUTTING_MASTER_SCREENSHOT = True  # Включить скриншоты Cutting Master 4
CUTTING_MASTER_WINDOW_TITLES = [  # Возможные заголовки окон Cutting Master
    "cutting master",
    "graphtec cutting master",
    "cutting master 4"
]
SCREENSHOT_DELAY = 0.5  # Задержка перед скриншотом (секунды)
CUTTING_MASTER_CROP = {  # Обрезка интерфейса (в пикселях)
    'top': 20,  # Убрать меню и панели инструментов
    'bottom': 20,  # Убрать строку состояния
    'left': 20,  # Убрать левые панели
    'right': 20  # Убрать правые панели
}

# --- Существующие настройки ---
RETRY_DELAYS = [1, 3, 5]
RECENT_FILE_THRESHOLD = 300
ENABLE_DESIGNS_MONITORING = True

# --- Настройки для конвертации ---
GHOSTSCRIPT_PATH = r"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe"
RASTER_EXTENSIONS = [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff"]
VECTOR_EXTENSIONS = [".eps", ".ai", ".pdf"]
NATIVE_EXTENSIONS = [".cdr"]
CONVERTIBLE_EXTENSIONS = RASTER_EXTENSIONS + VECTOR_EXTENSIONS + NATIVE_EXTENSIONS

# Настройка Pillow для EPS
if os.path.exists(GHOSTSCRIPT_PATH):
    EpsImagePlugin.gs_windows_binary = GHOSTSCRIPT_PATH
    print(f"📦 Ghostscript найден: {GHOSTSCRIPT_PATH}")
else:
    print(f"⚠️ Ghostscript не найден по пути: {GHOSTSCRIPT_PATH}")

# Глобальные переменные для мониторинга
designs_file_cache = {}
pending_jobs = {}

# Настройки CorelDRAW
ENABLE_COREL_AUTOMATION = True
COREL_ZOOM_FIT = True
COREL_RETRY_COUNT = 3
SCREENSHOT_QUALITY = 100


def find_cutting_master_window():
    """
    Находит главное окно Cutting Master 4
    """
    print(f"[CM_WINDOW] 🔍 Ищу окно Cutting Master 4...")

    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if window_text and any(title.lower() in window_text.lower() for title in CUTTING_MASTER_WINDOW_TITLES):
                windows.append((hwnd, window_text))
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)

    if windows:
        # Сортируем по приоритету (главное окно обычно содержит "4" или больше текста)
        main_windows = [w for w in windows if len(w[1]) > 10]  # Окна с длинными заголовками
        if main_windows:
            selected = main_windows[0]
            print(f"[CM_WINDOW] ✅ Найдено главное окно: '{selected[1]}'")
            return selected
        else:
            selected = windows[0]
            print(f"[CM_WINDOW] ✅ Найдено окно: '{selected[1]}'")
            return selected

    print(f"[CM_WINDOW] ❌ Окно Cutting Master 4 не найдено")
    print(f"[CM_WINDOW] 💡 Возможные причины:")
    print(f"[CM_WINDOW]   • Программа не запущена")
    print(f"[CM_WINDOW]   • Окно свернуто")
    print(f"[CM_WINDOW]   • Другое название окна")

    # Показываем все видимые окна для отладки
    all_windows = []
    win32gui.EnumWindows(enum_windows_callback, all_windows)
    visible_windows = [(hwnd, text) for hwnd, text in all_windows if text and len(text.strip()) > 0]

    print(f"[CM_WINDOW] 🪟 Найдено видимых окон: {len(visible_windows)}")
    for hwnd, text in visible_windows[:10]:  # Показываем первые 10
        if 'cutting' in text.lower() or 'master' in text.lower() or 'graphtec' in text.lower():
            print(f"[CM_WINDOW]   📋 ПОДОЗРИТЕЛЬНОЕ: '{text}'")
        elif len(text) > 3:  # Показываем только содержательные заголовки
            print(f"[CM_WINDOW]   🪟 '{text}'")

    return None


def screenshot_cutting_master_window(hwnd, save_path, job_title=""):
    """
    Делает скриншот всего экрана с максимальным качеством
    """
    print(f"[CM_SCREENSHOT] 📸 Делаю скриншот всего экрана...")
    if job_title:
        print(f"[CM_SCREENSHOT] 📋 Для job: '{job_title}'")

    try:
        img = ImageGrab.grab()

        img.save(save_path, 'PNG', quality=100, optimize=False)

        final_width, final_height = img.size
        print(f"[CM_SCREENSHOT] 💾 Скриншот сохранен: {save_path}")
        print(f"[CM_SCREENSHOT] 📐 Финальный размер: {final_width}x{final_height}")
        return True

    except Exception as e:
        print(f"[CM_SCREENSHOT] ❌ Ошибка создания скриншота: {e}")
        return False

def capture_cutting_master_screenshot(job_title):
    """
    Основная функция захвата скриншота Cutting Master 4
    """
    if not ENABLE_CUTTING_MASTER_SCREENSHOT:
        print(f"[CM_CAPTURE] ⚠️ Скриншоты Cutting Master 4 отключены")
        return None

    print(f"\n[CM_CAPTURE] =====================")
    print(f"[CM_CAPTURE] 📸 Захватываю скриншот для job: '{job_title}'")

    # Находим окно Cutting Master
    cm_window = find_cutting_master_window()
    if not cm_window:
        print(f"[CM_CAPTURE] ❌ Не найдено окно Cutting Master 4")
        return None

    hwnd, window_title = cm_window
    print(f"[CM_CAPTURE] 🪟 Используем окно: '{window_title}'")

    # Подготавливаем путь для скриншота
    safe_title = "".join(c for c in job_title if c.isalnum() or c in (' ', '-', '_')).rstrip()
    safe_title = safe_title.replace(" ", "_").replace("/", "_")
    timestamp = int(time.time())
    screenshot_filename = f"cm4_{safe_title}_{timestamp}.png"
    screenshot_path = os.path.join(MEDIA_PREVIEWS, screenshot_filename)

    print(f"[CM_CAPTURE] 💾 Целевой файл: {screenshot_path}")

    # Делаем скриншот
    if screenshot_cutting_master_window(hwnd, screenshot_path, job_title):
        # Проверяем, что файл создан и не пустой
        if os.path.exists(screenshot_path) and os.path.getsize(screenshot_path) > 1000:
            print(f"[CM_CAPTURE] ✅ Скриншот успешно создан")
            return screenshot_path
        else:
            print(f"[CM_CAPTURE] ❌ Скриншот не создан или поврежден")
            return None
    else:
        print(f"[CM_CAPTURE] ❌ Не удалось создать скриншот")
        return None


def normalize_name(s: str) -> str:
    return s.lower().replace(" ", "").replace("_", "")


def convert_to_preview_format(src_path, title):
    """
    Конвертирует различные форматы файлов в PNG для превью
    """
    print(f"\n[CONVERT] =====================")
    print(f"[CONVERT] Конвертирую файл: {src_path}")

    if not os.path.exists(src_path):
        print(f"[CONVERT] ❌ Файл не существует: {src_path}")
        return None

    file_ext = os.path.splitext(src_path)[1].lower()
    print(f"[CONVERT] Расширение файла: {file_ext}")

    # Если файл уже в растровом формате - возвращаем как есть
    if file_ext in RASTER_EXTENSIONS:
        print(f"[CONVERT] ✅ Растровый файл, конвертация не нужна")
        return src_path

    # Подготавливаем целевой файл
    os.makedirs(MEDIA_PREVIEWS, exist_ok=True)
    safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
    safe_title = safe_title.replace(" ", "_").replace("/", "_")
    hash_suffix = hashlib.md5(src_path.encode()).hexdigest()[:6]
    target_filename = f"{safe_title}_{hash_suffix}_converted.png"
    target_path = os.path.join(MEDIA_PREVIEWS, target_filename)

    print(f"[CONVERT] Целевой файл: {target_path}")

    try:
        # Конвертация EPS/AI/PDF через Pillow + Ghostscript
        if file_ext in ['.eps', '.ai', '.pdf']:
            print(f"[CONVERT] 📄 Конвертирую {file_ext.upper()} через Pillow...")

            if not os.path.exists(GHOSTSCRIPT_PATH):
                print(f"[CONVERT] ❌ Ghostscript не найден для конвертации {file_ext.upper()}")
                return None

            with Image.open(src_path) as im:
                # Загружаем с высоким разрешением
                im.load(scale=2)

                # Конвертируем в RGB если нужно
                if im.mode in ('RGBA', 'LA', 'P'):
                    background = Image.new('RGB', im.size, (255, 255, 255))
                    if im.mode == 'P':
                        im = im.convert('RGBA')
                    background.paste(im, mask=im.split()[-1] if im.mode in ('RGBA', 'LA') else None)
                    im = background
                elif im.mode not in ('RGB', 'L'):
                    im = im.convert('RGB')

                # Изменяем размер для превью (максимум 800x600)
                im.thumbnail((800, 600), Image.Resampling.LANCZOS)

                im.save(target_path, 'PNG', quality=95, optimize=True)
                print(f"[CONVERT] ✅ {file_ext.upper()} конвертирован в PNG")
                return target_path

        # Конвертация CDR
        elif file_ext == '.cdr':
            print(f"[CONVERT] 🎨 Попытка конвертации CDR...")
            return create_corel_preview(src_path, title)

        else:
            print(f"[CONVERT] ❌ Неподдерживаемый формат: {file_ext}")
            return None

    except Exception as e:
        print(f"[CONVERT] ❌ Ошибка конвертации: {e}")
        return None


def find_corel_process():
    """
    Находит запущенный процесс CorelDRAW
    """
    corel_processes = []
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            if proc.info['name'] and 'corel' in proc.info['name'].lower():
                corel_processes.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    return corel_processes


def get_corel_window():
    """
    Находит главное окно CorelDRAW
    """

    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if 'coreldraw' in window_text.lower():
                windows.append((hwnd, window_text))
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)

    # Сортируем по приоритету (главное окно обычно содержит версию)
    main_windows = [w for w in windows if 'coreldraw' in w[1].lower() and ('2' in w[1] or 'x' in w[1])]
    if main_windows:
        return main_windows[0]

    return windows[0] if windows else None


def screenshot_corel_window(hwnd, save_path):
    """
    Делает скриншот окна CorelDRAW с улучшенной обработкой ошибок
    """
    print(f"[COREL_SCREENSHOT] 📸 Делаю скриншот окна CorelDRAW...")

    try:
        # Получаем размеры окна
        rect = win32gui.GetWindowRect(hwnd)
        width = rect[2] - rect[0]
        height = rect[3] - rect[1]

        print(f"[COREL_SCREENSHOT] Размер окна: {width}x{height}")

        # Пробуем активировать окно (игнорируем ошибки SetForegroundWindow)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except:
            print(f"[COREL_SCREENSHOT] ⚠️ SetForegroundWindow не удался, продолжаем...")

        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except:
            print(f"[COREL_SCREENSHOT] ⚠️ ShowWindow не удался, продолжаем...")

        time.sleep(0.8)  # Увеличили время ожидания

        # Создаем контекст устройства для окна
        hwnd_dc = None
        mfc_dc = None
        save_dc = None
        save_bitmap = None

        try:
            hwnd_dc = win32gui.GetWindowDC(hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()

            # Создаем bitmap
            save_bitmap = win32ui.CreateBitmap()
            save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(save_bitmap)

            # Копируем содержимое окна - пробуем несколько методов
            result = False

            # Метод 1: PrintWindow с полным содержимым
            try:
                result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 3)  # PW_RENDERFULLCONTENT
                print(f"[COREL_SCREENSHOT] PrintWindow (метод 1): {'успех' if result else 'неудача'}")
            except Exception as e:
                print(f"[COREL_SCREENSHOT] PrintWindow (метод 1) ошибка: {e}")

            # Метод 2: PrintWindow без флагов
            if not result:
                try:
                    result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)
                    print(f"[COREL_SCREENSHOT] PrintWindow (метод 2): {'успех' if result else 'неудача'}")
                except Exception as e:
                    print(f"[COREL_SCREENSHOT] PrintWindow (метод 2) ошибка: {e}")

            # Метод 3: BitBlt (fallback)
            if not result:
                try:
                    result = save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)
                    print(f"[COREL_SCREENSHOT] BitBlt (метод 3): {'успех' if result else 'неудача'}")
                except Exception as e:
                    print(f"[COREL_SCREENSHOT] BitBlt (метод 3) ошибка: {e}")

            if result:
                # Получаем данные bitmap
                bmp_info = save_bitmap.GetInfo()
                bmp_str = save_bitmap.GetBitmapBits(True)

                # Конвертируем в PIL Image
                img = Image.frombuffer(
                    'RGB',
                    (bmp_info['bmWidth'], bmp_info['bmHeight']),
                    bmp_str, 'raw', 'BGRX', 0, 1
                )

                # Обрезаем только рабочую область (убираем меню и панели)
                crop_top = min(100, height // 4)  # Убираем меню и панели инструментов
                crop_bottom = min(50, height // 8)  # Убираем строку состояния
                crop_left = min(50, width // 8)  # Убираем левые панели
                crop_right = min(50, width // 8)  # Убираем правые панели

                cropped_img = img.crop((
                    crop_left,
                    crop_top,
                    max(width - crop_right, width * 3 // 4),
                    max(height - crop_bottom, height * 3 // 4)
                ))

                # Изменяем размер для превью
                cropped_img.thumbnail((800, 600), Image.Resampling.LANCZOS)

                # Сохраняем
                cropped_img.save(save_path, 'PNG', quality=SCREENSHOT_QUALITY, optimize=True)

                print(f"[COREL_SCREENSHOT] ✅ Скриншот сохранен: {save_path}")
                return True
            else:
                print(f"[COREL_SCREENSHOT] ❌ Все методы захвата не сработали")
                return False

        except Exception as e:
            print(f"[COREL_SCREENSHOT] ❌ Ошибка в процессе создания bitmap: {e}")
            return False
        finally:
            # Освобождаем ресурсы
            try:
                if save_bitmap:
                    win32gui.DeleteObject(save_bitmap.GetHandle())
                if save_dc:
                    save_dc.DeleteDC()
                if mfc_dc:
                    mfc_dc.DeleteDC()
                if hwnd_dc:
                    win32gui.ReleaseDC(hwnd, hwnd_dc)
            except:
                pass

    except Exception as e:
        print(f"[COREL_SCREENSHOT] ❌ Общая ошибка создания скриншота: {e}")
        return False


def corel_automation_screenshot(cdr_path, title):
    """
    Автоматизация CorelDRAW для создания скриншота файла
    """
    if not ENABLE_COREL_AUTOMATION:
        print(f"[COREL_AUTO] ⚠️ Автоматизация CorelDRAW отключена")
        return None

    print(f"\n[COREL_AUTO] =====================")
    print(f"[COREL_AUTO] 🎨 Автоматизация CorelDRAW для: {cdr_path}")

    # Проверяем, существует ли файл
    if not os.path.exists(cdr_path):
        print(f"[COREL_AUTO] ❌ Файл не найден: {cdr_path}")
        return None

    corel_app = None
    screenshot_path = None

    try:
        # Подготавливаем путь для скриншота
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_title = safe_title.replace(" ", "_").replace("/", "_")
        timestamp = int(time.time())
        screenshot_filename = f"corel_{safe_title}_{timestamp}.png"
        screenshot_path = os.path.join(MEDIA_PREVIEWS, screenshot_filename)

        print(f"[COREL_AUTO] Целевой файл скриншота: {screenshot_path}")

        # Попробуем подключиться к существующему экземпляру CorelDRAW
        try:
            print(f"[COREL_AUTO] 🔗 Подключаюсь к CorelDRAW...")
            corel_app = win32com.client.GetActiveObject("CorelDRAW.Application")
            print(f"[COREL_AUTO] ✅ Подключен к запущенному CorelDRAW")
        except:
            # Если нет запущенного экземпляра, запускаем новый
            print(f"[COREL_AUTO] 🚀 Запускаю новый экземпляр CorelDRAW...")
            corel_app = win32com.client.Dispatch("CorelDRAW.Application")
            corel_app.Visible = True
            time.sleep(3)  # Ждем запуска
            print(f"[COREL_AUTO] ✅ CorelDRAW запущен")

        # Получаем версию CorelDRAW
        try:
            version = corel_app.VersionMajor
            print(f"[COREL_AUTO] 📋 Версия CorelDRAW: {version}")
        except:
            print(f"[COREL_AUTO] ⚠️ Не удалось получить версию CorelDRAW")

        # Открываем документ
        print(f"[COREL_AUTO] 📂 Открываю файл: {os.path.basename(cdr_path)}")
        doc = corel_app.OpenDocument(cdr_path)

        if not doc:
            print(f"[COREL_AUTO] ❌ Не удалось открыть документ")
            return None

        print(f"[COREL_AUTO] ✅ Документ открыт: {doc.Name}")

        # Активируем документ
        doc.Activate()

        # Подгоняем масштаб
        if COREL_ZOOM_FIT:
            try:
                print(f"[COREL_AUTO] 🔍 Подгоняю масштаб...")
                active_view = corel_app.ActiveView
                if active_view:
                    active_view.FitToPage()  # Подгоняем под страницу
                    time.sleep(0.5)
                    print(f"[COREL_AUTO] ✅ Масштаб подогнан")
            except Exception as e:
                print(f"[COREL_AUTO] ⚠️ Не удалось подогнать масштаб: {e}")

        # Ждем полной загрузки
        time.sleep(1)

        # Находим окно CorelDRAW
        corel_window = get_corel_window()
        if not corel_window:
            print(f"[COREL_AUTO] ❌ Не найдено окно CorelDRAW")
            return None

        hwnd, window_title = corel_window
        print(f"[COREL_AUTO] 🪟 Найдено окно: {window_title}")

        # Делаем скриншот
        for attempt in range(COREL_RETRY_COUNT):
            print(f"[COREL_AUTO] 📸 Попытка скриншота {attempt + 1}/{COREL_RETRY_COUNT}")

            if screenshot_corel_window(hwnd, screenshot_path):
                # Проверяем, что файл создан и не пустой
                if os.path.exists(screenshot_path) and os.path.getsize(screenshot_path) > 1000:
                    print(f"[COREL_AUTO] ✅ Скриншот успешно создан")

                    # Закрываем документ (не сохраняя)
                    try:
                        doc.Close(False)  # False = не сохранять
                        print(f"[COREL_AUTO] 📄 Документ закрыт")
                    except:
                        print(f"[COREL_AUTO] ⚠️ Не удалось закрыть документ")

                    return screenshot_path
                else:
                    print(f"[COREL_AUTO] ❌ Скриншот не создан или поврежден")

            if attempt < COREL_RETRY_COUNT - 1:
                time.sleep(1)  # Ждем перед следующей попыткой

        print(f"[COREL_AUTO] ❌ Не удалось создать скриншот после {COREL_RETRY_COUNT} попыток")
        return None

    except Exception as e:
        print(f"[COREL_AUTO] ❌ Ошибка автоматизации CorelDRAW: {e}")
        return None
    finally:
        # Очистка
        if corel_app and screenshot_path and not os.path.exists(screenshot_path):
            try:
                # Если скриншот не создан, все равно пытаемся закрыть документ
                active_doc = corel_app.ActiveDocument
                if active_doc:
                    active_doc.Close(False)
            except:
                pass


def create_corel_preview(cdr_path, title):
    """
    Создает превью для CDR файла с приоритетом автоматизации CorelDRAW
    """
    print(f"\n[COREL_PREVIEW] =====================")
    print(f"[COREL_PREVIEW] 🎨 Создаю превью для CDR: {cdr_path}")

    # Стратегия 1: Автоматизация CorelDRAW (приоритет)
    if ENABLE_COREL_AUTOMATION:
        screenshot_path = corel_automation_screenshot(cdr_path, title)
        if screenshot_path and os.path.exists(screenshot_path):
            print(f"[COREL_PREVIEW] ✅ Превью создано через CorelDRAW")
            return screenshot_path

    # Стратегия 2: ImageMagick (fallback)
    try:
        result = subprocess.run(['magick', '-version'],
                                capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print(f"[COREL_PREVIEW] 🔧 Пробую ImageMagick как fallback...")

            safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_title = safe_title.replace(" ", "_").replace("/", "_")
            timestamp = int(time.time())
            target_filename = f"magick_{safe_title}_{timestamp}.png"
            target_path = os.path.join(MEDIA_PREVIEWS, target_filename)

            cmd = [
                'magick',
                cdr_path + '[0]',
                '-thumbnail', '800x600>',
                '-background', 'white',
                '-flatten',
                target_path
            ]

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

            if result.returncode == 0 and os.path.exists(target_path):
                print(f"[COREL_PREVIEW] ✅ Превью создано через ImageMagick")
                return target_path
            else:
                print(f"[COREL_PREVIEW] ❌ ImageMagick не смог обработать файл: {result.stderr}")

    except (subprocess.TimeoutExpired, FileNotFoundError):
        print(f"[COREL_PREVIEW] ⚠️ ImageMagick недоступен")

    # Стратегия 3: Информативный placeholder
    print(f"[COREL_PREVIEW] 📝 Создаю placeholder для CDR...")
    return create_cdr_placeholder(cdr_path, title)


def create_cdr_placeholder(cdr_path, title):
    """
    Создает информативный placeholder для CDR файлов
    """
    print(f"\n[CDR_PLACEHOLDER] =====================")
    print(f"[CDR_PLACEHOLDER] Создаю placeholder для CDR: {cdr_path}")

    try:
        # Получаем информацию о файле
        file_stat = os.stat(cdr_path)
        file_size = file_stat.st_size
        file_mtime = datetime.fromtimestamp(file_stat.st_mtime)

        # Создаем изображение 400x300
        width, height = 400, 300
        image = Image.new('RGB', (width, height), '#f8f9fa')
        draw = ImageDraw.Draw(image)

        # Рисуем градиент (серо-голубой)
        for y in range(height):
            r = int(248 - (y / height) * 20)
            g = int(249 - (y / height) * 15)
            b = int(250 - (y / height) * 10)
            color = (r, g, b)
            draw.line([(0, y), (width, y)], fill=color)

        # Рисуем рамку
        draw.rectangle([0, 0, width - 1, height - 1], outline='#6c757d', width=2)

        # Загружаем шрифты
        try:
            font_title = ImageFont.truetype("arial.ttf", 18)
            font_info = ImageFont.truetype("arial.ttf", 12)
            font_small = ImageFont.truetype("arial.ttf", 10)
        except:
            font_title = ImageFont.load_default()
            font_info = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Логотип CDR (просто текст)
        draw.text((20, 20), "CDR", fill='#dc3545', font=font_title)
        draw.text((60, 25), "CorelDRAW", fill='#6c757d', font=font_small)

        # Заголовок файла
        title_text = title[:25] + "..." if len(title) > 25 else title
        title_bbox = draw.textbbox((0, 0), title_text, font=font_title)
        title_x = (width - (title_bbox[2] - title_bbox[0])) // 2
        draw.text((title_x, 80), title_text, fill='#212529', font=font_title)

        # Информация о файле
        size_mb = file_size / (1024 * 1024)
        info_lines = [
            f"Размер: {size_mb:.1f} MB",
            f"Изменен: {file_mtime.strftime('%d.%m.%Y %H:%M')}",
            "Требует CorelDRAW для",
            "полного просмотра"
        ]

        y_pos = 130
        for line in info_lines:
            line_bbox = draw.textbbox((0, 0), line, font=font_info)
            line_x = (width - (line_bbox[2] - line_bbox[0])) // 2
            color = '#6c757d' if "Требует" in line or "полного" in line else '#495057'
            draw.text((line_x, y_pos), line, fill=color, font=font_info)
            y_pos += 20

        # Текущее время
        time_text = datetime.now().strftime("Создано: %d.%m.%Y %H:%M")
        time_bbox = draw.textbbox((0, 0), time_text, font=font_small)
        time_x = (width - (time_bbox[2] - time_bbox[0])) // 2
        draw.text((time_x, 260), time_text, fill='#868e96', font=font_small)

        # Сохраняем
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_title = safe_title.replace(" ", "_").replace("/", "_")
        timestamp = int(time.time())
        placeholder_filename = f"cdr_{safe_title}_{timestamp}.png"
        placeholder_path = os.path.join(MEDIA_PREVIEWS, placeholder_filename)

        image.save(placeholder_path, 'PNG', quality=95)

        print(f"[CDR_PLACEHOLDER] ✅ CDR placeholder создан: {placeholder_path}")
        return placeholder_path

    except Exception as e:
        print(f"[CDR_PLACEHOLDER] ❌ Ошибка создания CDR placeholder: {e}")
        return None


def create_placeholder_preview(title, reason="Новый макет"):
    """
    Создает placeholder превью для макетов без найденного файла
    """
    print(f"\n[PLACEHOLDER] =====================")
    print(f"[PLACEHOLDER] Создаю placeholder для: '{title}'")
    print(f"[PLACEHOLDER] Причина: {reason}")

    try:
        # Создаем изображение 400x300 с градиентом
        width, height = 400, 300
        image = Image.new('RGB', (width, height), '#f0f0f0')
        draw = ImageDraw.Draw(image)

        # Рисуем градиент
        for y in range(height):
            r = int(240 + (y / height) * 15)
            g = int(240 + (y / height) * 15)
            b = int(250 + (y / height) * 5)
            color = (r, g, b)
            draw.line([(0, y), (width, y)], fill=color)

        # Рисуем рамку
        draw.rectangle([0, 0, width - 1, height - 1], outline='#cccccc', width=2)

        # Пытаемся загрузить шрифт, если не получается - используем дефолтный
        try:
            font_title = ImageFont.truetype("arial.ttf", 20)
            font_subtitle = ImageFont.truetype("arial.ttf", 14)
            font_small = ImageFont.truetype("arial.ttf", 10)
        except:
            font_title = ImageFont.load_default()
            font_subtitle = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Добавляем текст
        # Заголовок
        title_text = title[:30] + "..." if len(title) > 30 else title
        title_bbox = draw.textbbox((0, 0), title_text, font=font_title)
        title_x = (width - (title_bbox[2] - title_bbox[0])) // 2
        draw.text((title_x, 80), title_text, fill='#333333', font=font_title)

        # Подзаголовок
        subtitle_bbox = draw.textbbox((0, 0), reason, font=font_subtitle)
        subtitle_x = (width - (subtitle_bbox[2] - subtitle_bbox[0])) // 2
        draw.text((subtitle_x, 120), reason, fill='#666666', font=font_subtitle)

        # Время создания
        time_text = datetime.now().strftime("%d.%m.%Y %H:%M")
        time_bbox = draw.textbbox((0, 0), time_text, font=font_small)
        time_x = (width - (time_bbox[2] - time_bbox[0])) // 2
        draw.text((time_x, 250), time_text, fill='#999999', font=font_small)

        # Сохраняем во временный файл
        os.makedirs(MEDIA_PREVIEWS, exist_ok=True)
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_title = safe_title.replace(" ", "_").replace("/", "_")
        timestamp = int(time.time())
        placeholder_filename = f"placeholder_{safe_title}_{timestamp}.png"
        placeholder_path = os.path.join(MEDIA_PREVIEWS, placeholder_filename)

        image.save(placeholder_path, 'PNG', quality=95)

        print(f"[PLACEHOLDER] ✅ Создан placeholder: {placeholder_path}")
        return placeholder_filename  # Возвращаем только имя файла для rel_path

    except Exception as e:
        print(f"[PLACEHOLDER] ❌ Ошибка создания placeholder: {e}")
        return None


def scan_designs_directory():
    """
    Сканирует папку с макетами и создает кеш файлов с временными метками
    """
    global designs_file_cache
    print(f"\n[CACHE_SCAN] =====================")
    print(f"[CACHE_SCAN] Сканирую папку макетов: {DESIGNS_DIR}")

    new_cache = {}
    file_count = 0

    try:
        for root, dirs, files in os.walk(DESIGNS_DIR):
            for f in files:
                name, ext = os.path.splitext(f)
                if ext.lower() in CONVERTIBLE_EXTENSIONS:
                    full_path = os.path.join(root, f)
                    try:
                        mtime = os.path.getmtime(full_path)
                        rel_path = os.path.relpath(os.path.dirname(full_path), DESIGNS_DIR).replace("\\", "/")
                        if rel_path == ".":
                            rel_path = ""

                        # Ключ для быстрого поиска
                        norm_name = normalize_name(name)
                        cache_key = f"{rel_path}|{norm_name}" if rel_path else norm_name

                        # Определяем приоритет файла (растровые важнее векторных)
                        priority = 1  # Высокий приоритет для растровых
                        if ext.lower() in VECTOR_EXTENSIONS:
                            priority = 2  # Средний для векторных
                        elif ext.lower() in NATIVE_EXTENSIONS:
                            priority = 3  # Низкий для нативных форматов

                        new_cache[cache_key] = {
                            'full_path': full_path,
                            'name': name,
                            'normalized_name': norm_name,
                            'folder': rel_path,
                            'mtime': mtime,
                            'ext': ext,
                            'priority': priority
                        }
                        file_count += 1
                    except OSError:
                        continue

        designs_file_cache = new_cache
        print(f"[CACHE_SCAN] ✅ Обновлен кеш: {file_count} файлов")

        # Статистика по типам файлов
        raster_count = sum(1 for info in new_cache.values() if info['ext'].lower() in RASTER_EXTENSIONS)
        vector_count = sum(1 for info in new_cache.values() if info['ext'].lower() in VECTOR_EXTENSIONS)
        native_count = sum(1 for info in new_cache.values() if info['ext'].lower() in NATIVE_EXTENSIONS)

        print(f"[CACHE_SCAN] 📊 Статистика файлов:")
        print(f"[CACHE_SCAN]   Растровые (PNG/JPG): {raster_count}")
        print(f"[CACHE_SCAN]   Векторные (EPS/AI/PDF): {vector_count}")
        print(f"[CACHE_SCAN]   Нативные (CDR): {native_count}")

        # Показываем недавно измененные файлы
        recent_threshold = time.time() - RECENT_FILE_THRESHOLD
        recent_files = [
            info for info in new_cache.values()
            if info['mtime'] > recent_threshold
        ]

        if recent_files:
            print(f"[CACHE_SCAN] 🕒 Недавно измененные файлы ({len(recent_files)}):")
            for info in sorted(recent_files, key=lambda x: x['mtime'], reverse=True)[:10]:
                mtime_str = datetime.fromtimestamp(info['mtime']).strftime("%H:%M:%S")
                rel_path = os.path.relpath(info['full_path'], DESIGNS_DIR)
                file_type = "🖼️" if info['ext'].lower() in RASTER_EXTENSIONS else "📄"
                print(f"[CACHE_SCAN]   {mtime_str} {file_type} {rel_path}")

    except Exception as e:
        print(f"[CACHE_SCAN] ❌ Ошибка сканирования: {e}")


def find_by_time_proximity(search_filename, search_folder=None, job_creation_time=None):
    """
    Ищет файлы по временной близости к созданию job файла
    """
    if not job_creation_time:
        job_creation_time = time.time()

    print(f"\n[TIME_SEARCH] =====================")
    print(f"[TIME_SEARCH] Поиск по времени для: '{search_filename}' в папке '{search_folder}'")
    print(f"[TIME_SEARCH] Время создания job: {datetime.fromtimestamp(job_creation_time).strftime('%H:%M:%S')}")

    candidates = []
    norm_search = normalize_name(search_filename)

    # Ищем файлы, созданные/измененные в разумном временном окне (до 10 минут до и после)
    time_window = 600  # 10 минут

    for cache_key, info in designs_file_cache.items():
        time_diff = abs(info['mtime'] - job_creation_time)

        if time_diff <= time_window:
            # Проверяем релевантность по имени
            name_ratio = difflib.SequenceMatcher(None, norm_search, info['normalized_name']).ratio()

            # Проверяем папку если указана
            folder_match = True
            if search_folder and info['folder'] != search_folder:
                folder_match = False

            if name_ratio > 0.3 or info['normalized_name'] in norm_search or norm_search in info['normalized_name']:
                score = name_ratio
                if folder_match:
                    score += 0.2
                # Бонус за близость по времени
                time_bonus = max(0, (time_window - time_diff) / time_window) * 0.1
                score += time_bonus

                candidates.append({
                    'path': info['full_path'],
                    'name': info['name'],
                    'score': score,
                    'time_diff': time_diff,
                    'folder_match': folder_match
                })

                print(f"[TIME_SEARCH]   🕒 Кандидат: {info['name']} (score={score:.3f}, время±{time_diff:.0f}с)")

    if candidates:
        candidates.sort(key=lambda x: x['score'], reverse=True)
        best = candidates[0]
        print(f"[TIME_SEARCH] ✅ Лучший по времени: {best['name']} (score={best['score']:.3f})")
        return best['path']

    print(f"[TIME_SEARCH] ❌ Не найдено подходящих файлов по времени")
    return None


def extract_path_from_job_path(job_file_path, watch_dirs):
    """
    Извлекает относительный путь к папке из полного пути к .job файлу
    """
    print(f"\n[PATH_EXTRACT] =====================")
    print(f"[PATH_EXTRACT] Анализирую путь: {job_file_path}")

    # Находим базовую папку Jobs
    base_jobs_dir = None
    for watch_dir in watch_dirs:
        if watch_dir in job_file_path:
            base_jobs_dir = watch_dir
            print(f"[PATH_EXTRACT] Базовая Jobs папка: {base_jobs_dir}")
            break

    if not base_jobs_dir:
        print(f"[PATH_EXTRACT] ❌ Не найдена базовая папка Jobs в пути")
        return None, os.path.basename(job_file_path)

    # Получаем относительный путь от базовой папки Jobs до файла
    try:
        rel_path = os.path.relpath(job_file_path, base_jobs_dir)
        print(f"[PATH_EXTRACT] Относительный путь: {rel_path}")

        # Разделяем на папку и файл
        folder_part = os.path.dirname(rel_path)
        file_part = os.path.basename(rel_path)

        print(f"[PATH_EXTRACT] Папка: '{folder_part}'")
        print(f"[PATH_EXTRACT] Файл: '{file_part}'")

        # Убираем расширение .job из имени файла
        filename_without_ext = os.path.splitext(file_part)[0]
        print(f"[PATH_EXTRACT] Имя без расширения: '{filename_without_ext}'")

        return folder_part if folder_part != "." else None, filename_without_ext

    except Exception as e:
        print(f"[PATH_EXTRACT] ❌ Ошибка при извлечении пути: {e}")
        return None, os.path.splitext(os.path.basename(job_file_path))[0]


def parse_job_title(title):
    """
    Старый метод парсинга из названия (fallback)
    """
    print(f"[TITLE_PARSE] Парсинг названия: '{title}'")
    if "/" in title:
        parts = title.split("/")
        if len(parts) >= 2:
            folder_path = "/".join(parts[:-1])
            filename = parts[-1]
            print(f"[TITLE_PARSE] Из названия извлечено - папка: '{folder_path}', файл: '{filename}'")
            return folder_path, filename
    print(f"[TITLE_PARSE] Разделителей не найдено, только файл: '{title}'")
    return "", title


def find_design_file_advanced(title, job_file_path=None, job_creation_time=None):
    """
    Расширенный поиск дизайн-файла с множественными стратегиями
    """
    print(f"\n[DESIGN_SEARCH] =====================")
    print(f"[DESIGN_SEARCH] Расширенный поиск для JOB: '{title}'")
    if job_file_path:
        print(f"[DESIGN_SEARCH] Полный путь к job: '{job_file_path}'")

    # Попробуем извлечь путь из полного пути к job файлу
    folder_from_path = None
    filename_from_path = None

    if job_file_path:
        folder_from_path, filename_from_path = extract_path_from_job_path(job_file_path, PATHS_TO_WATCH)

    # Fallback к старому методу парсинга названия
    folder_from_title, filename_from_title = parse_job_title(title)

    # Определяем окончательные значения для поиска
    search_folder = folder_from_path if folder_from_path else folder_from_title
    search_filename = filename_from_path if filename_from_path else filename_from_title

    print(f"[DESIGN_SEARCH] ИТОГОВЫЕ параметры поиска:")
    print(f"[DESIGN_SEARCH]   Папка: '{search_folder}'")
    print(f"[DESIGN_SEARCH]   Файл: '{search_filename}'")

    norm_filename = normalize_name(search_filename)
    print(f"[DESIGN_SEARCH]   Нормализованное имя: '{norm_filename}'")

    # СТРАТЕГИЯ 1: Поиск в кеше (быстро)
    print(f"\n[DESIGN_SEARCH] 📋 Стратегия 1: Поиск в кеше")

    # Точное совпадение по пути и имени
    if search_folder:
        exact_key = f"{search_folder}|{norm_filename}"
        if exact_key in designs_file_cache:
            result = designs_file_cache[exact_key]['full_path']
            print(f"[DESIGN_SEARCH] ✅ Найдено в кеше (точное): {result}")
            return result

    # Поиск только по имени (если папка не указана или не найдена)
    name_matches = []
    for key, info in designs_file_cache.items():
        if info['normalized_name'] == norm_filename:
            name_matches.append(info)

    if name_matches:
        # Сортируем по приоритету (растровые файлы первые)
        name_matches.sort(key=lambda x: (x['priority'], -x['mtime']))  # Высокий приоритет и свежий mtime

        # Если есть совпадения, выбираем из правильной папки или лучший по приоритету
        if search_folder:
            folder_matches = [
                info for info in name_matches
                if search_folder in info['folder']
            ]
            if folder_matches:
                best_match = folder_matches[0]
                print(
                    f"[DESIGN_SEARCH] ✅ Найдено в кеше (папка+имя): {best_match['full_path']} (приоритет {best_match['priority']})")
                return best_match['full_path']

        best_match = name_matches[0]
        print(
            f"[DESIGN_SEARCH] ✅ Найдено в кеше (только имя): {best_match['full_path']} (приоритет {best_match['priority']})")
        return best_match['full_path']

    # СТРАТЕГИЯ 2: Поиск по времени создания
    print(f"\n[DESIGN_SEARCH] 🕒 Стратегия 2: Поиск по времени")
    time_result = find_by_time_proximity(search_filename, search_folder, job_creation_time)
    if time_result:
        return time_result

    # СТРАТЕГИЯ 3: Fuzzy поиск в кеше
    print(f"\n[DESIGN_SEARCH] 🔍 Стратегия 3: Fuzzy поиск")
    fuzzy_candidates = []

    for cache_key, info in designs_file_cache.items():
        norm_name = info['normalized_name']

        # Проверка папки
        folder_bonus = 0
        if search_folder and info['folder'] == search_folder:
            folder_bonus = 0.1

        # Различные алгоритмы похожести
        if len(norm_filename) >= 2 and len(norm_name) >= 2:
            ratio = difflib.SequenceMatcher(None, norm_filename, norm_name).ratio()

            # Для коротких запросов
            if len(norm_filename) <= 4:
                starts_with = norm_name.startswith(norm_filename)
                ends_with = norm_name.endswith(norm_filename)
                contains = norm_filename in norm_name

                if starts_with or ends_with or contains:
                    position_bonus = 0.2 if (starts_with or ends_with) else 0.1
                    final_score = min(0.8 + position_bonus + folder_bonus, 1.0)
                    fuzzy_candidates.append((info['full_path'], norm_name, final_score))
            else:
                # Для длинных запросов
                if ratio >= 0.7:
                    final_score = ratio + folder_bonus
                    fuzzy_candidates.append((info['full_path'], norm_name, final_score))

    if fuzzy_candidates:
        fuzzy_candidates.sort(key=lambda x: x[2], reverse=True)
        best_match, best_name, best_score = fuzzy_candidates[0]
        print(f"[DESIGN_SEARCH] ✅ Найден fuzzy: {best_match} (score={best_score:.3f})")
        return best_match

    # СТРАТЕГИЯ 4: Полный пересcan (если кеш устарел)
    print(f"\n[DESIGN_SEARCH] 🔄 Стратегия 4: Пересканирование")
    scan_designs_directory()  # Обновляем кеш

    # Повторяем поиск в обновленном кеше
    if search_folder:
        exact_key = f"{search_folder}|{norm_filename}"
        if exact_key in designs_file_cache:
            result = designs_file_cache[exact_key]['full_path']
            print(f"[DESIGN_SEARCH] ✅ Найдено после пересканирования: {result}")
            return result

    print(f"[DESIGN_SEARCH] ❌ Файл не найден всеми стратегиями")
    return None


def find_design_file_with_retry(title, job_file_path=None, job_creation_time=None):
    """
    Поиск дизайн-файла с повторными попытками
    """
    print(f"\n[RETRY_SEARCH] =====================")
    print(f"[RETRY_SEARCH] Поиск с retry для: '{title}'")

    # Первая попытка
    result = find_design_file_advanced(title, job_file_path, job_creation_time)
    if result:
        print(f"[RETRY_SEARCH] ✅ Найден с первой попытки")
        return result

    # Повторные попытки с задержками
    for i, delay in enumerate(RETRY_DELAYS):
        print(f"[RETRY_SEARCH] ⏳ Попытка {i + 2} через {delay} секунд...")
        time.sleep(delay)

        # Обновляем кеш перед каждой попыткой
        scan_designs_directory()

        result = find_design_file_advanced(title, job_file_path, job_creation_time)
        if result:
            print(f"[RETRY_SEARCH] ✅ Найден с попытки {i + 2}")
            return result

    print(f"[RETRY_SEARCH] ❌ Не найден после всех попыток")
    return None


def copy_to_media(src_path, title):
    """
    Конвертирует файл в подходящий формат и копирует в медиа-директорию
    """
    if not src_path:
        print(f"[COPY_MEDIA] ❌ Исходный файл не указан")
        return None

    print(f"\n[COPY_MEDIA] =====================")
    print(f"[COPY_MEDIA] Обрабатываю файл: {src_path}")

    # Проверяем существование файла
    if not os.path.exists(src_path):
        print(f"[COPY_MEDIA] ❌ Файл не существует: {src_path}")
        return None

    # Пробуем конвертировать файл
    converted_path = convert_to_preview_format(src_path, title)

    if converted_path and os.path.exists(converted_path):
        # Если конвертация успешна, возвращаем относительный путь
        rel_path = os.path.relpath(converted_path, os.path.dirname(MEDIA_PREVIEWS))
        print(f"[COPY_MEDIA] ✅ Файл обработан: {rel_path}")
        return rel_path

    # Если конвертация не удалась, но файл растровый - копируем как есть
    file_ext = os.path.splitext(src_path)[1].lower()
    if file_ext in RASTER_EXTENSIONS:
        print(f"[COPY_MEDIA] 📋 Копирую растровый файл как есть...")

        os.makedirs(MEDIA_PREVIEWS, exist_ok=True)
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_title = safe_title.replace(" ", "_").replace("/", "_")
        hash_suffix = hashlib.md5(src_path.encode()).hexdigest()[:6]
        dst_filename = f"{safe_title}_{hash_suffix}{file_ext}"
        dst_path = os.path.join(MEDIA_PREVIEWS, dst_filename)

        print(f"[COPY_MEDIA] Целевой файл: {dst_path}")

    try:
        shutil.copy2(src_path, dst_path)  # Используем copy2 для сохранения метаданных
        rel_path = f"previews/{dst_filename}"
        print(f"[COPY_MEDIA] ✅ Скопировано: {rel_path}")
        return rel_path
    except Exception as e:
        print(f"[COPY_MEDIA] ❌ Ошибка копирования: {e}")
        return None


    print(f"[COPY_MEDIA] ❌ Не удалось обработать файл")
    return None


def send_event(title, status, job_file_path=None):
    """
    Отправляет событие в Django API с улучшенной обработкой и скриншотом Cutting Master 4
    """
    print(f"\n[SEND_EVENT] =====================")
    print(f"[SEND_EVENT] Обрабатываю событие: '{title}' - {status}")
    if job_file_path:
        print(f"[SEND_EVENT] Путь к job файлу: {job_file_path}")

    ts = datetime.now(timezone.utc).isoformat()

    # Определяем время создания job файла
    job_creation_time = time.time()
    if job_file_path and os.path.exists(job_file_path):
        try:
            job_creation_time = os.path.getmtime(job_file_path)
            print(f"[SEND_EVENT] Время создания job: {datetime.fromtimestamp(job_creation_time).strftime('%H:%M:%S')}")
        except:
            pass

    # НОВОЕ: Создаем скриншот Cutting Master 4 в момент создания/изменения job
    cutting_master_screenshot_path = None
    if ENABLE_CUTTING_MASTER_SCREENSHOT:
        # Делаем скриншот для CREATED и MODIFIED событий
        print(f"[SEND_EVENT] 📸 Создание скриншота Cutting Master 4...")
        cutting_master_screenshot_path = capture_cutting_master_screenshot(title)

    # Расширенный поиск дизайн-файла
    design_file = find_design_file_with_retry(title, job_file_path, job_creation_time)

    preview_rel = None
    preview_full = None

    # Приоритет превью:
    # 1. Скриншот Cutting Master 4 (наивысший приоритет)
    # 2. Найденный дизайн-файл
    # 3. Placeholder для новых макетов

    if cutting_master_screenshot_path and os.path.exists(cutting_master_screenshot_path):
        print(f"[SEND_EVENT] 📸 Используется скриншот Cutting Master 4 как основное превью")
        preview_rel = f"previews/{os.path.basename(cutting_master_screenshot_path)}"
        preview_full = cutting_master_screenshot_path
    elif design_file:
        print(f"[SEND_EVENT] 🎨 Используется найденный дизайн-файл как превью")
        preview_rel = copy_to_media(design_file, title)
        if preview_rel:
            preview_full = os.path.join(MEDIA_PREVIEWS, os.path.basename(preview_rel))
    else:
        # Создаем placeholder для новых макетов
        print(f"[SEND_EVENT] 🆕 Создаю placeholder для нового макета")
        placeholder_filename = create_placeholder_preview(title, "Новый макет - файл не найден")
        if placeholder_filename:
            preview_rel = f"previews/{placeholder_filename}"
            preview_full = os.path.join(MEDIA_PREVIEWS, placeholder_filename)

    # Определяем источник превью для дополнительной информации
    preview_source = "unknown"
    if cutting_master_screenshot_path and os.path.exists(cutting_master_screenshot_path):
        preview_source = "cutting_master_screenshot"
    elif design_file:
        preview_source = "design_file"
    else:
        preview_source = "placeholder"

    payload = {
        "title": title,
        "created_at": ts,
        "plotter": PRINTER,
        "status": status,
        "is_new_design": design_file is None,  # Флаг нового дизайна
        "preview_source": preview_source,  # НОВОЕ: Источник превью
        "has_cutting_master_screenshot": cutting_master_screenshot_path is not None,
        # НОВОЕ: Флаг наличия скриншота CM4
    }

    print(f"[SEND_EVENT] Payload: {payload}")
    print(f"[SEND_EVENT] Preview: {preview_rel}")
    print(f"[SEND_EVENT] Preview source: {preview_source}")

    files = None
    if preview_full and os.path.exists(preview_full):
        try:
            files = {"preview": open(preview_full, "rb")}
            print(f"[SEND_EVENT] 📂 Готов upload превью: {preview_full}")
        except Exception as e:
            print(f"[SEND_EVENT] ❌ Ошибка открытия превью: {e}")

    try:
        print(f"[SEND_EVENT] Отправляю запрос к Django API...")
        r = requests.post(DJANGO_API, data=payload, files=files, timeout=10)
        print(f"[SEND_EVENT] 🔗 Ответ Django: {r.status_code}")
        print(f"[SEND_EVENT] Текст ответа: {r.text[:500]}")
        r.raise_for_status()
        print(f"[SEND_EVENT] ✅ Успешно отправлено")
    except Exception as e:
        print(f"[SEND_EVENT] ❌ Ошибка отправки: {e}")
    finally:
        if files and "preview" in files:
            files["preview"].close()


def delayed_retry_search(title, job_file_path, status, delay):
    """
    Отложенный поиск файла (для случаев когда файл сохраняется после создания job)
    """

    def retry_worker():
        print(f"\n[DELAYED_RETRY] =====================")
        print(f"[DELAYED_RETRY] Отложенный поиск через {delay}с для: '{title}'")
        time.sleep(delay)

        job_creation_time = time.time()
        if job_file_path and os.path.exists(job_file_path):
            try:
                job_creation_time = os.path.getmtime(job_file_path)
            except:
                pass

        # Обновляем кеш
        scan_designs_directory()

        # Ищем файл
        design_file = find_design_file_advanced(title, job_file_path, job_creation_time)

        if design_file:
            print(f"[DELAYED_RETRY] ✅ Найден при отложенном поиске: {design_file}")
            # Отправляем обновленное событие
            send_event(f"{title} [ОБНОВЛЕНО]", f"{status}_UPDATED", job_file_path)
        else:
            print(f"[DELAYED_RETRY] ❌ Не найден и при отложенном поиске")

    # Запускаем в отдельном потоке
    thread = threading.Thread(target=retry_worker)
    thread.daemon = True
    thread.start()


class DesignsHandler(FileSystemEventHandler):
    """
    Обработчик событий для папки с макетами
    """

    def on_created(self, event):
        if not event.is_directory:
            name, ext = os.path.splitext(os.path.basename(event.src_path))
            if ext.lower() in CONVERTIBLE_EXTENSIONS:
                print(f"\n[DESIGNS_EVENT] =====================")
                print(f"[DESIGNS_EVENT] 🆕 Новый дизайн-файл: {event.src_path}")

                # Обновляем кеш
                scan_designs_directory()

                # Проверяем, есть ли отложенные задачи для этого файла
                norm_name = normalize_name(name)
                rel_folder = os.path.relpath(os.path.dirname(event.src_path), DESIGNS_DIR).replace("\\", "/")
                if rel_folder == ".":
                    rel_folder = ""

                print(f"[DESIGNS_EVENT] Ищу отложенные задачи для: '{norm_name}' в папке '{rel_folder}'")

    def on_modified(self, event):
        if not event.is_directory:
            name, ext = os.path.splitext(os.path.basename(event.src_path))
            if ext.lower() in CONVERTIBLE_EXTENSIONS:
                print(f"\n[DESIGNS_EVENT] =====================")
                print(f"[DESIGNS_EVENT] ✏️  Изменен дизайн-файл: {event.src_path}")

                # Обновляем кеш
                scan_designs_directory()


class JobHandler(FileSystemEventHandler):
    """
    Обработчик событий файловой системы для .job файлов
    """

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".job"):
            print(f"\n[FILE_EVENT] =====================")
            print(f"[FILE_EVENT] 🆕 СОЗДАН файл: {event.src_path}")

            # Извлекаем информацию о папке и файле
            folder_name = os.path.basename(os.path.dirname(event.src_path))
            fname = os.path.basename(event.src_path)
            title = os.path.splitext(fname)[0]

            print(f"[FILE_EVENT] Папка: {folder_name}")
            print(f"[FILE_EVENT] Файл: {fname}")
            print(f"[FILE_EVENT] Title: {title}")

            # Отправляем событие (теперь с возможностью скриншота Cutting Master 4)
            send_event(title, "CREATED", event.src_path)

            # Планируем отложенный поиск для случая, если файл появится позже
            delayed_retry_search(title, event.src_path, "CREATED", 30)  # Через 30 секунд

    def on_modified(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".job"):
            print(f"\n[FILE_EVENT] =====================")
            print(f"[FILE_EVENT] ✏️  ИЗМЕНЁН файл: {event.src_path}")

            # Извлекаем информацию о папке и файле
            folder_name = os.path.basename(os.path.dirname(event.src_path))
            fname = os.path.basename(event.src_path)
            title = os.path.splitext(fname)[0]

            print(f"[FILE_EVENT] Папка: {folder_name}")
            print(f"[FILE_EVENT] Файл: {fname}")
            print(f"[FILE_EVENT] Title: {title}")

            send_event(title, "MODIFIED", event.src_path)


if __name__ == "__main__":
    print(f"🚀 Запуск улучшенного Job Watcher с скриншотами Cutting Master 4")
    print(f"📁 Папки для наблюдения: {PATHS_TO_WATCH}")
    print(f"🎨 Папка с макетами: {DESIGNS_DIR}")
    print(f"🖼️  Папка для превью: {MEDIA_PREVIEWS}")
    print(f"🌐 Django API: {DJANGO_API}")
    print(f"👤 Пользователь: {USER}")
    print(f"🖨️  Принтер: {PRINTER}")
    print(f"🔄 Retry задержки: {RETRY_DELAYS}")
    print(f"📊 Мониторинг макетов: {'Включен' if ENABLE_DESIGNS_MONITORING else 'Выключен'}")
    print(f"📸 Скриншоты Cutting Master 4: {'Включены' if ENABLE_CUTTING_MASTER_SCREENSHOT else 'Выключены'}")

    # Показываем настройки скриншотов
    if ENABLE_CUTTING_MASTER_SCREENSHOT:
        print(f"📸 НАСТРОЙКИ СКРИНШОТОВ CUTTING MASTER 4:")
        print(f"   Заголовки окон: {CUTTING_MASTER_WINDOW_TITLES}")
        print(f"   Задержка перед скриншотом: {SCREENSHOT_DELAY}с")
        print(f"   Обрезка интерфейса: {CUTTING_MASTER_CROP}")

    # Создаем папку для превью
    os.makedirs(MEDIA_PREVIEWS, exist_ok=True)

    # Первоначальное сканирование папки макетов
    print(f"\n🔍 Первоначальное сканирование макетов...")
    scan_designs_directory()

    # Проверяем доступность Cutting Master 4 при запуске
    if ENABLE_CUTTING_MASTER_SCREENSHOT:
        print(f"\n📸 Проверка доступности Cutting Master 4...")
        cm_window = find_cutting_master_window()
        if cm_window:
            print(f"✅ Cutting Master 4 найден и готов для скриншотов")
        else:
            print(f"⚠️  Cutting Master 4 не найден. Скриншоты будут пропущены до запуска программы.")

    observers = []

    # Настройка наблюдения за job файлами
    for p in PATHS_TO_WATCH:
        if os.path.exists(p):
            obs = Observer()
            obs.schedule(JobHandler(), p, recursive=True)
            obs.start()
            observers.append(obs)
            print(f"👀 Watching jobs: {p}")

            # Показываем содержимое папки для отладки
            print(f"📋 Содержимое папки Jobs:")
            try:
                for root, dirs, files in os.walk(p):
                    level = root.replace(p, '').count(os.sep)
                    indent = ' ' * 2 * level
                    rel_path = os.path.relpath(root, p) if root != p else ""
                    print(f"{indent}📁 {rel_path}/")
                    subindent = ' ' * 2 * (level + 1)
                    job_files = [f for f in files if f.lower().endswith('.job')]
                    for f in job_files[:5]:  # Показываем первые 5 job файлов
                        print(f"{subindent}📄 {f}")
                    if len(job_files) > 5:
                        print(f"{subindent}... и еще {len(job_files) - 5} job файлов")
                    if len(files) - len(job_files) > 0:
                        print(f"{subindent}... и {len(files) - len(job_files)} других файлов")
            except Exception as e:
                print(f"❌ Ошибка сканирования папки Jobs: {e}")
        else:
            print(f"❌ Папка Jobs не найдена: {p}")

    # Настройка наблюдения за папкой макетов (опционально)
    if ENABLE_DESIGNS_MONITORING and os.path.exists(DESIGNS_DIR):
        designs_obs = Observer()
        designs_obs.schedule(DesignsHandler(), DESIGNS_DIR, recursive=True)
        designs_obs.start()
        observers.append(designs_obs)
        print(f"👀 Watching designs: {DESIGNS_DIR}")

        # Показываем статистику папки макетов
        total_designs = len([info for info in designs_file_cache.values()])
        recent_threshold = time.time() - RECENT_FILE_THRESHOLD
        recent_designs = len([
            info for info in designs_file_cache.values()
            if info['mtime'] > recent_threshold
        ])

        print(f"📊 Статистика макетов:")
        print(f"   Всего файлов: {total_designs}")
        print(f"   Недавних (за {RECENT_FILE_THRESHOLD // 60} мин): {recent_designs}")

        # Показываем структуру папок с макетами
        folders = set(info['folder'] for info in designs_file_cache.values() if info['folder'])
        print(f"   Папок с макетами: {len(folders)}")
        for folder in sorted(list(folders)[:10]):  # Первые 10 папок
            folder_files = [info for info in designs_file_cache.values() if info['folder'] == folder]
            print(f"     📁 {folder}/ ({len(folder_files)} файлов)")
        if len(folders) > 10:
            print(f"     ... и еще {len(folders) - 10} папок")

    print(f"\n✅ Job Watcher запущен! Нажмите Ctrl+C для остановки.")
    print(f"🔧 Возможности:")
    print(f"   • 📸 НОВОЕ: Автоматические скриншоты Cutting Master 4")
    print(f"   • Поиск по структуре папок")
    print(f"   • Retry с задержками {RETRY_DELAYS}")
    print(f"   • Поиск по времени создания")
    print(f"   • Placeholder для новых макетов")
    print(f"   • Кеширование файлов")
    print(f"   • Мониторинг папки макетов")
    print(f"   • Отложенный поиск через 30с")
    print(f"   • Конвертация векторных файлов")
    print(f"   • Приоритет растровым форматам")
    print(f"   • Поддержка EPS/AI/PDF → PNG")
    print(f"   • Специальные placeholder для CDR")

    if ENABLE_CUTTING_MASTER_SCREENSHOT:
        print(f"   • 🎯 Приоритет скриншотов Cutting Master 4 над макетами")
        print(f"   • 🪟 Автоматическое определение окна программы")
        print(f"   • ✂️  Умная обрезка интерфейса")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print(f"\n🛑 Останавливаю наблюдение...")
        for o in observers:
            o.stop()
        for o in observers:
            o.join()
        print(f"✅ Job Watcher остановлен")