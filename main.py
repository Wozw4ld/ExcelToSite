import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# Пути
yandex_browser_path = 'C:\\Users\\asaparbaev\\AppData\\Local\\Yandex\\YandexBrowser\\Application\\browser.exe'
user_data_dir = 'C:\\Users\\asaparbaev\\AppData\\Local\\Yandex\\YandexBrowser\\User Data'
profile_dir = 'Default'  # Имя профиля
driver_path = 'yandexdriver.exe'  # Путь к драйверу

# Создаем папку профиля, если её нет
def create_profile_if_not_exists(user_data_dir, profile_dir):
    profile_path = os.path.join(user_data_dir, profile_dir)
    if not os.path.exists(profile_path):
        print(f"Профиль {profile_dir} не найден. Создаем новый профиль...")
        os.makedirs(profile_path)
    else:
        print(f"Профиль {profile_dir} уже существует.")

# Создаем профиль, если его нет
create_profile_if_not_exists(user_data_dir, profile_dir)

# Настройки браузера
options = Options()
options.binary_location = yandex_browser_path
options.add_argument(f"--user-data-dir={user_data_dir}")
options.add_argument(f"--profile-directory={profile_dir}")

# Обязательные аргументы
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--remote-allow-origins=*")
options.add_argument("--disable-gpu")
options.add_argument("--start-maximized")

# Инициализация драйвера
try:
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    print("Браузер успешно запущен!")
    driver.get('https://yandex.ru')
    input("Нажмите Enter для завершения...")
except Exception as e:
    print(f"Ошибка: {e}")
finally:
    if 'driver' in locals():
        driver.quit()