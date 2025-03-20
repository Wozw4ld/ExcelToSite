import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Глобальная переменная для хранения объекта драйвера
driver = None

def select_file():
    """Функция для кнопки 'Обзор' – выбираем Excel-файл."""
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filepath)

def login_browser():
    """Открывает браузер с нужным профилем и не закрывается до нажатия 'Выйти'."""
    global driver
    if driver is not None:
        messagebox.showinfo("Информация", "Браузер уже запущен.")
        return

    try:
        options = Options()
        # Указываем путь к браузеру Yandex (или измените на нужный)
        options.binary_location = 'C:\\Users\\asaparbaev\\AppData\\Local\\Yandex\\YandexBrowser\\Application\\browser.exe'
        
        # Задаём профиль – логика профиля не изменяется
        profile_dir = "asd"  # Укажите имя профиля, как было в вашем коде
     
        service = Service(executable_path='yandexdriver.exe')
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://yandex.ru")  # Можно заменить на нужный URL
        status_label.config(text="Браузер запущен!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть браузер: {e}")

def start_automation():
    """Читает Excel и заполняет поля на странице в уже открытом браузере."""
    global driver
    if driver is None:
        messagebox.showerror("Ошибка", "Сначала нажмите 'Войти' для открытия браузера.")
        return

    file_path = file_entry.get()
    url = url_entry.get()
    start_cell = start_cell_entry.get()
    end_cell = end_cell_entry.get()
    
    if not file_path or not url or not start_cell or not end_cell:
        messagebox.showerror("Ошибка", "Необходимо указать файл, URL, начальную и конечную ячейки.")
        return
    
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active  
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")
        return

    try:
        excel_data = []
        cell_range = sheet[start_cell:end_cell]
        for row in cell_range:
            for cell in row:
                excel_data.append(cell.value if cell.value is not None else "")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось считать диапазон {start_cell}:{end_cell} - {e}")
        return

    if not excel_data:
        messagebox.showerror("Ошибка", "В указанном диапазоне не найдены данные.")
        return

    try:
        # Переходим по указанному URL (можно заменить, если требуется другой URL)
        driver.get(url)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть URL: {e}")
        return

    time.sleep(5)  # Ждём загрузки страницы

    try:
        # Поиск всех input с нужным стилем
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[style*='width:100%;text-align:Right;background-color:']")
        if not inputs:
            messagebox.showerror("Ошибка", "На странице не найдено элементов input с нужным стилем.")
            return

        for i, input_element in enumerate(inputs):
            if i < len(excel_data):
                try:
                    input_element.clear()
                    input_element.send_keys(str(excel_data[i]))
                except Exception as e:
                    print(f"Ошибка при заполнении input[{i}]: {e}")
            else:
                break

        messagebox.showinfo("Успех", "Данные успешно переданы на сайт!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при заполнении данных: {e}")

def logout_browser():
    """Закрывает браузер и очищает объект драйвера."""
    global driver
    if driver is not None:
        try:
            driver.quit()
            driver = None
            status_label.config(text="Браузер закрыт.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось закрыть браузер: {e}")
    else:
        messagebox.showinfo("Информация", "Браузер не запущен.")

# Создание окна Tkinter
root = tk.Tk()
root.title("Автоматизация заполнения с Excel в Selenium (Yandex)")

# Рамка для управления файлом и диапазонами Excel
tk.Label(root, text="Выберите Excel файл:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Обзор", command=select_file).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Ссылка на сайт:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Начальная ячейка (например, E7):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
start_cell_entry = tk.Entry(root, width=10)
start_cell_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

tk.Label(root, text="Конечная ячейка (например, AI15):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
end_cell_entry = tk.Entry(root, width=10)
end_cell_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

# Кнопки управления браузером и автоматизацией
button_frame = tk.Frame(root)
button_frame.grid(row=4, column=0, columnspan=3, pady=10)

login_button = tk.Button(button_frame, text="Войти (Открыть браузер)", width=25, command=login_browser)
login_button.grid(row=0, column=0, padx=5, pady=5)

automation_button = tk.Button(button_frame, text="Запустить автоматизацию", width=25, command=start_automation)
automation_button.grid(row=0, column=1, padx=5, pady=5)

logout_button = tk.Button(button_frame, text="Выйти (Закрыть браузер)", width=25, command=logout_browser)
logout_button.grid(row=0, column=2, padx=5, pady=5)

status_label = tk.Label(root, text="Ожидание действий...", fg="blue")
status_label.grid(row=5, column=0, columnspan=3, pady=10)

root.mainloop()
