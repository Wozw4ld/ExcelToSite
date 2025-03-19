from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service  # Импортируем Service
from webdriver_manager.chrome import ChromeDriverManager

# Настройки профиля
user_data_dir = "D:\\Users\\asaparbaev\\AppData\\Local\\Chromium\\User Data\\Default"  # Пример для Windows
profile_dir = "asd"  # Или "Guest Profile"

options = Options()
options.add_argument(f"--user-data-dir={user_data_dir}")
options.add_argument(f"--profile-directory={profile_dir}")

# Вариант 1: Использование Service с автоматическим менеджером
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Вариант 2: Указать путь к драйверу вручную (если нужно)
# service = Service(executable_path="C:/path/to/chromedriver.exe")
# driver = webdriver.Chrome(service=service, options=options)

driver.get("https://vk.com/feed")
a = input()