import logging
import os
from datetime import datetime
import time
import random
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from typing import List, Set


def setup_logging():
    """Настройка логирования"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_file = os.path.join(log_dir, f"vc_parser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    logger = logging.getLogger("vc_parser")
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


logger = setup_logging()


class VCParserException(Exception):
    """Базовый класс для исключений парсера"""
    pass


def setup_driver() -> webdriver.Chrome:
    """Настройка драйвера Chrome"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Запуск без GUI
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')

    # Установка и настройка драйвера
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def scroll_page(driver: webdriver.Chrome, pbar: tqdm) -> None:
    """Прокрутка страницы до конца с подсчетом статей"""
    last_height = driver.execute_script("return document.body.scrollHeight")
    articles_count = 0
    no_new_content_count = 0

    while True:
        # Прокрутка вниз
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(1, 2))  # Ждем загрузку контента

        # Получаем новую высоту страницы
        new_height = driver.execute_script("return document.body.scrollHeight")

        # Подсчет текущего количества статей
        current_articles = len(driver.find_elements(By.CSS_SELECTOR, 'div.feed__item'))

        if current_articles > articles_count:
            # Обновляем прогресс
            new_articles = current_articles - articles_count
            pbar.update(new_articles)
            articles_count = current_articles
            no_new_content_count = 0
        else:
            no_new_content_count += 1

        # Если высота не изменилась или долго нет новых статей - прекращаем
        if new_height == last_height or no_new_content_count >= 3:
            break

        last_height = new_height


def extract_titles(driver: webdriver.Chrome) -> Set[str]:
    """Извлечение заголовков из загруженной страницы"""
    titles = set()

    # Ждем появления статей
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.feed__item"))
        )
    except TimeoutException:
        logger.error("Timeout waiting for articles to load")
        return titles

    # Получаем HTML и парсим с помощью BeautifulSoup для удобства
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Ищем все заголовки статей
    for article in soup.select('div.feed__item'):
        title_elem = article.select_one('h2.content-title')
        if title_elem:
            title_text = title_elem.get_text(strip=True)
            if title_text:
                titles.add(title_text)

    return titles


def get_articles() -> List[str]:
    """Получение статей с сайта"""
    titles = set()
    url = "https://vc.ru/marketing"

    # Создаем и настраиваем драйвер
    driver = setup_driver()

    try:
        # Загружаем страницу
        driver.get(url)
        logger.info("Страница загружена успешно")

        # Создаем прогресс-бар
        pbar = tqdm(desc="Загрузка статей", unit="статья")

        # Прокручиваем страницу и собираем статьи
        scroll_page(driver, pbar)

        # Извлекаем заголовки
        titles = extract_titles(driver)

        pbar.close()
        logger.info(f"Найдено статей: {len(titles)}")

    except Exception as e:
        logger.error(f"Ошибка при парсинге: {str(e)}")
        raise VCParserException(f"Ошибка при парсинге: {str(e)}")
    finally:
        driver.quit()

    if not titles:
        raise VCParserException("Не удалось получить статьи")

    return list(sorted(titles))


def save_results(titles: List[str]) -> None:
    """Сохранение результатов в Excel файл"""
    try:
        results_dir = "results"
        if not os.path.exists(results_dir):
            os.makedirs(results_dir)

        filename = os.path.join(results_dir, f"articles_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

        # Создаем DataFrame с нумерацией
        df = pd.DataFrame({
            '№': range(1, len(titles) + 1),
            'Заголовок статьи': titles,
            'Дата сбора': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * len(titles)
        })

        # Сохраняем в Excel с форматированием
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Статьи')

            # Получаем рабочий лист для форматирования
            worksheet = writer.sheets['Статьи']

            # Устанавливаем ширину столбцов
            worksheet.column_dimensions['A'].width = 5  # Для номера
            worksheet.column_dimensions['B'].width = 100  # Для заголовка
            worksheet.column_dimensions['C'].width = 20  # Для даты

            # Форматируем заголовки
            for cell in worksheet[1]:
                cell.style = 'Headline 1'

        print(f"\nРезультаты сохранены в файл: {filename}")
        logger.info(f"Результаты сохранены в файл: {filename}")

    except Exception as e:
        error_msg = f"Ошибка при сохранении результатов: {str(e)}"
        print(f"\n{error_msg}")
        logger.error(error_msg)


def main():
    try:
        print("\nНачало работы парсера")
        logger.info("Начало работы парсера")

        titles = get_articles()

        if titles:
            print(f"\nНайдено {len(titles)} статей")
            logger.info(f"Найдено {len(titles)} статей")
            save_results(titles)
        else:
            print("\nНе удалось получить заголовки статей")
            logger.warning("Не удалось получить заголовки статей")

    except Exception as e:
        error_msg = f"Критическая ошибка: {str(e)}"
        print(f"\n{error_msg}")
        logger.error(f"Критическая ошибка в main(): {str(e)}")
    finally:
        print("\nЗавершение работы парсера")
        logger.info("Завершение работы парсера")


if __name__ == "__main__":
    main()
