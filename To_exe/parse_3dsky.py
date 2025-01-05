# Парсинг топа продаж 3dsky
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import pandas as pd
import time

# Настройка Selenium
chrome_options = Options()
chrome_options.add_argument("--headless")  # Запуск без интерфейса (опционально)
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_service = Service("chromedriver.exe")  # Замените на путь к вашему ChromeDriver

# Инициализация констант и переменных
result = []                                                         # итоговый список
maxpage = 2                                                         # сколько страниц топа грузим
basic_url = "https://3dsky.org/3dmodels?order=sell_rating&page="    # адрес (без номера страницы)
pause_time = 5                                                      # время задержки на загрузку в сек
page_size = 60                                                      # число моделей на странице
attempts = 3                                                        # число попыток парсинга каждой модели
basic_width = 4                                                     # количество параметров модели из топа
full_width = 14                                                     # полное число параметров модели
url_position = 3                                                    # на каком месте находится параметр url
file_path = 'output.xlsx'                                           # имя файла

# Формируем список моделей
for page in range(maxpage):
    # Инициализация браузера
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    
    url = basic_url + str(page + 1)
    driver.get(url)

    # Небольшая пауза для загрузки контента
    time.sleep(pause_time)

    # Парсинг данных
    try:
        # Получаем список моделей с заданной страницы топа
        models = driver.find_elements(By.CSS_SELECTOR, "div.model-title")

        for n, model in enumerate(models, 1):
            # Извлечение данных
            curdate = str(datetime.now().date())
            title = model.find_element(By.TAG_NAME, "a").get_attribute("title") if model.find_elements(By.TAG_NAME, "a") else "No title"
            link = model.find_element(By.TAG_NAME, "a").get_attribute("href") if model.find_elements(By.TAG_NAME, "a") else "No link"
            #print(f"Название: {title}, Ссылка: {link}")
            
            # Формируем заготовку строки: номер строки, текущая дата, название модели, ссылка              
            line_num = page*page_size + n
            result.append([line_num, curdate, title, link])

    finally:
        # Закрытие браузера
        driver.quit()

# В цикле по списку получаем подробную информацию по каждой модели
for model in result:
    for attempt in range(attempts):
    
        # Инициализация браузера
        driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
        
        # Открытие страницы конкретной модели
        url = model[url_position]
        driver.get(url)
        
        # Небольшая пауза для загрузки контента
        time.sleep(pause_time)
        
        # Парсинг данных
        try:
            category = driver.find_elements(By.CSS_SELECTOR, "body > app-root > app-model > app-base-wide > main > section.container.main-base-container > div.row.white-background > div.col-md-12.model-page-top.ng-tns-c55-0 > div.favourite-and-category.ng-tns-c55-0.ng-star-inserted > div:nth-child(1) > div.category.ng-tns-c55-0 > a > span")
            subcategory = driver.find_elements(By.CSS_SELECTOR, "body > app-root > app-model > app-base-wide > main > section.container.main-base-container > div.row.white-background > div.col-md-12.model-page-top.ng-tns-c55-0 > div.favourite-and-category.ng-tns-c55-0.ng-star-inserted > div:nth-child(1) > div.subcategory.ng-tns-c55-0 > a > span")
            platform = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div.model-info-block.ng-tns-c55-0 > table > tbody > tr:nth-child(1) > td:nth-child(2)")
            renders = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div.model-info-block.ng-tns-c55-0 > table > tbody > tr:nth-child(2) > td:nth-child(2) > div")
            published = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div.model-info-block.ng-tns-c55-0 > div.publication-date.ng-tns-c55-0.ng-star-inserted")
            username = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div:nth-child(1) > div > div > a > div > div > div.model-user-name.ng-tns-c55-0")
            followers = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div:nth-child(1) > div > div > a > div > div > div.model-subscribe-count.ng-tns-c55-0.ng-star-inserted")
            selected = driver.find_elements(By.CSS_SELECTOR, "#info-desktop > div:nth-child(1) > div > div > div:nth-child(3) > div.price-block.ng-tns-c55-0.ng-star-inserted > div.bookmarks-block-wrapper-mobile.ng-tns-c55-0.ng-star-inserted > div > div")
            
            # Добавляем полученные данные в строку: категория, подкатегория, платформа, рендер
            if category:
                model.append(category[0].text)
            if subcategory:
                model.append(subcategory[0].text)
            if platform:
                model.append(platform[0].text)
            
            if renders:
                render_text = renders[0].text
            if ('Corona' in render_text) and ('V-Ray' in render_text):
                model.extend(['Corona','V-Ray','None'])
            elif ('Corona' in render_text):
                model.extend(['Corona','None','None'])
            elif ('V-Ray' in render_text):
                model.extend(['None','V-Ray','None'])
            elif ('Standard' in render_text):
                model.extend(['None','None','Standard'])
            else:
                model.extend(['None','None','None'])
            
            if published:
                pub_date = " ".join(str(published[0].text).split(" ")[1:])
                date_object = datetime.strptime(pub_date, "%d %B %Y")
                formatted_date = date_object.strftime("%Y-%m-%d")
                model.append(formatted_date)
            if username:
                model.append(username[0].text)
            if followers:
                model.append(int(followers[0].text.split(" ")[0]))
            if selected:
                selected_text = driver.execute_script("return arguments[0].innerText;", selected[0])
                model.append(int(selected_text.strip()))

        finally:
            # Закрытие браузера
            driver.quit()
            
        # Проверяем, все ли параметры считались, и если нет, то пробуем ещё несколько раз
        if len(model) == full_width:
            print(f'Success: {model}, attempt: {attempt+1}')
            break
        else:
            if attempt+1 == attempts:
                print(f'Failed: {model}, attempt: {attempt+1}')
            else:
                print(f'Pending: {model}, attempt: {attempt+1}')
                if len(model) > basic_width:
                    del model[basic_width:]

# Чтение данных
existing_data = pd.read_excel(file_path)

# Преобразование текущих данных DataFrame
new_data_df = pd.DataFrame(result[0:], columns=['N','Curdate','Title','Link','Category','Subcategory','Platform','Corona','V-Ray','Standard','Pubdate','Username','Followers','Selected'])

# Конкатенация и запись результата в Excel
updated_data = pd.concat([existing_data, new_data_df], ignore_index=True)
updated_data.to_excel(file_path, index=False)

# Выводим результат для контроля
print(f'Данные успешно сохранены в {file_path}')