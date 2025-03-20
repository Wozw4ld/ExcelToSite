from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# Настройки профиля
options = Options()

# Путь к исполняемому файлу Яндекс.Браузера
options.binary_location = 'C:\\Users\\asaparbaev\\AppData\\Local\\Yandex\\YandexBrowser\\Application\\browser.exe'
 
user_data_dir = "C:\\Users\\Admin\\AppData\\Local\\Chromium\\User Data"
profile_dir = "asd"
options.add_argument(f"--user-data-dir={user_data_dir}")
options.add_argument(f"--profile-directory={profile_dir}")	
service = Service(executable_path='yandexdriver.exe')

driver = webdriver.Chrome(service=service, options=options)
driver.get('https://yandex.ru')
a = input()
driver.quit()