import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import time
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service  # Импортируем Service
from webdriver_manager.chrome import ChromeDriverManager
def select_file():
    """Функция для кнопки 'Обзор' – выбираем Excel-файл."""
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filepath)

def start_automation():
    """Основная логика: читаем Excel и заполняем поля на сайте."""
    file_path = file_entry.get()
    url = url_entry.get()
    start_cell = start_cell_entry.get()
    end_cell = end_cell_entry.get()
    
    # Проверяем, что все поля заполнены
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
        user_data_dir = "D:\\Users\\Admin\\AppData\\Local\\Chromium\\User Data\\Default"  # Пример для Windows

        options = Options()
        options.add_argument(f"--user-data-dir={user_data_dir}")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(url)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть браузер или URL: {e}")
        return
    
    time.sleep(5)  

    # 3. Поиск всех input с нужным стилем
    inputs = driver.find_elements(By.CSS_SELECTOR, "input[style*='width:100%;text-align:Right;background-color:']")
    if not inputs:
        messagebox.showerror("Ошибка", "На странице не найдено элементов input с нужным стилем.")
        driver.quit()
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

root = tk.Tk()
root.title("Автоматизация заполнения с Excel в Selenium")

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

tk.Button(root, text="Запустить автоматизацию", command=start_automation).grid(row=4, column=1, padx=5, pady=10)

root.mainloop()
