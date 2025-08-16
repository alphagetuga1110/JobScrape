# scrapers/workport_scraper.py

import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def scrape(list_url, num_to_scrape, msg_queue):
    """
    Workportから指定件数のデータをスクレイピングし、リストとして返す関数
    """
    # ... (scrape_workport_data関数の中身をここにペースト) ...
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
        msg_queue.put("ブラウザを起動し、Workportの求人一覧ページにアクセスします...")
        driver.get(list_url)
        msg_queue.put("ページが読み込まれました。")
        
        try:
            msg_queue.put("「検索」ボタンをクリックします...")
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.SearchBtnEnter.js-searchbtn"))
            )
            search_button.click()
            msg_queue.put("ボタンをクリックしました。検索結果の表示を待機します...")
        except TimeoutException:
            msg_queue.put("「検索」ボタンが見つかりませんでした。処理を続行します。")

        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "searchAll_box")))
        msg_queue.put("検索結果が表示されました。")
        
        job_cards = driver.find_elements(By.CLASS_NAME, "searchAll_box")
        job_count = len(job_cards)
        loop_count = min(job_count, num_to_scrape)
        msg_queue.put(f"ページ上の{job_count}件中、指定された{loop_count}件を処理します...")
        msg_queue.put("-" * 50)
        
        for i in range(loop_count):
            try:
                current_card = driver.find_elements(By.CLASS_NAME, "searchAll_box")[i]
                
                try:
                    temp_company_name = current_card.find_element(By.CLASS_NAME, 'company').text.strip()
                except:
                    temp_company_name = f"{(i+1)}番目の求人"
                msg_queue.put(f"({i+1}/{loop_count}) {temp_company_name} の情報を取得中...")

                detail_button = WebDriverWait(current_card, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.enter_btn a"))
                )
                driver.execute_script("arguments[0].click();", detail_button)

                company_name, address, corporate_url = "取得失敗", "取得失敗", "N/A"
                
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "page3")))
                
                try:
                    name_element = driver.find_element(By.XPATH, "//h3[text()='会社名']/following-sibling::p[@class='txt']")
                    company_name = name_element.text.strip()
                except NoSuchElementException:
                    company_name = temp_company_name

                try:
                    address_element = driver.find_element(By.XPATH, "//h3[text()='本社所在地']/following-sibling::p[@class='txt']")
                    address = address_element.text.strip()
                except NoSuchElementException:
                    address = "所在地なし"

                scraped_data.append([company_name, address, corporate_url])
                msg_queue.put(f" -> 完了")
                
                driver.back()
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "searchAll_box")))
                
            except Exception as e:
                msg_queue.put(f" -> この求人の処理中にエラーが発生したため、スキップします。エラー: {e}")
                driver.get(list_url)
                try:
                    search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.SearchBtnEnter.js-searchbtn")))
                    search_button.click()
                except:
                    pass
                continue
            
    except Exception as e:
        msg_queue.put(f"処理全体で重大なエラーが発生しました: {e}")
    finally:
        if driver:
            msg_queue.put("-" * 50)
            msg_queue.put("スクレイピング処理が完了しました。")
            driver.quit()
            
    return scraped_data