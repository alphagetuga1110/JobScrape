# scrapers/doda_scraper.py

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup

def scrape(list_url, num_to_scrape, msg_queue):
    # (scrape_doda_data関数の中身をここにペースト)
    chrome_options = Options()
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    chrome_options.add_argument(f'user-agent={user_agent}')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_window_size(1280, 800)
    except Exception as e:
        msg_queue.put(f"ChromeDriverの起動に失敗しました: {e}")
        return []

    scraped_data = []
    try:
        msg_queue.put("ブラウザを起動し、DODAの求人一覧ページにアクセスします...")
        driver.get(list_url)
        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".jobCard-card")))
        msg_queue.put("ページが読み込まれました。")
        
        job_cards = driver.find_elements(By.CSS_SELECTOR, ".jobCard-card")
        job_count = len(job_cards)
        loop_count = min(job_count, num_to_scrape)
        msg_queue.put(f"ページ上の{job_count}件中、指定された{loop_count}件を処理します...")
        msg_queue.put("-" * 50)
        
        main_window_handle = driver.current_window_handle
        for i in range(loop_count):
            current_card = driver.find_elements(By.CSS_SELECTOR, ".jobCard-card")[i]
            company_name = current_card.find_element(By.TAG_NAME, 'h2').text
            msg_queue.put(f"({i+1}/{loop_count}) {company_name} の情報を取得中...")
            try:
                detail_button = WebDriverWait(current_card, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.Button-module_button--green__Zirc1")))
                driver.execute_script("arguments[0].click();", detail_button)
            except Exception:
                msg_queue.put(" -> 詳細ボタンのクリックに失敗。スキップします。")
                continue
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            new_window_handle = [h for h in driver.window_handles if h != main_window_handle][0]
            driver.switch_to.window(new_window_handle)
            address, corporate_url = "取得失敗", "URLなし"
            try:
                try:
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//dt[text()='所在地']")))
                except TimeoutException:
                    job_detail_tab = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '求人詳細')]")))
                    job_detail_tab.click()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//dt[text()='所在地']")))
                address_element = driver.find_element(By.XPATH, "//dt[text()='所在地']/following-sibling::dd/p")
                inner_html = address_element.get_attribute('innerHTML')
                soup = BeautifulSoup(inner_html, 'html.parser')
                address = ' '.join(soup.get_text(separator=' ', strip=True).split())
                try:
                    url_element = driver.find_element(By.XPATH, "//dt[text()='企業URL']/following-sibling::dd/a")
                    corporate_url = url_element.get_attribute('href')
                except NoSuchElementException:
                    corporate_url = "URLなし"
            except Exception as e:
                msg_queue.put(f" -> 詳細ページでの情報取得に失敗しました: {e}")
            scraped_data.append([company_name, address, corporate_url])
            msg_queue.put(f" -> 完了")
            driver.close()
            driver.switch_to.window(main_window_handle)
            time.sleep(1)
    finally:
        if driver:
            msg_queue.put("-" * 50)
            msg_queue.put("スクレイピング処理が完了しました。")
            driver.quit()
    return scraped_data