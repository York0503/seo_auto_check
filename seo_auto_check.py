from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
import subprocess
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os
from urllib.parse import urlparse

from selenium.webdriver.chrome.service import Service

def load_keywords():
    """從 Excel 檔案讀取關鍵字列表"""
    keyword_file = "seo_search_keyword.xlsx"

    # 檢查檔案是否存在
    if not os.path.exists(keyword_file):
        print(f"錯誤：找不到關鍵字檔案 {keyword_file}")
        return []

    try:
        # 載入 Excel 檔案
        wb = load_workbook(keyword_file)
        ws = wb.active  # 使用第一個工作表

        keywords = []
        # 從第二列開始讀取（跳過標題列），讀取第二欄（B欄）
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
            keyword = row[0]
            # 略過空白儲存格
            if keyword and str(keyword).strip():
                keywords.append(str(keyword).strip())

        print(f"✓ 成功載入 {len(keywords)} 個關鍵字")
        return keywords

    except Exception as e:
        print(f"讀取關鍵字檔案時發生錯誤: {e}")
        return []

def save_to_excel(results, keyword):
    """將搜尋結果儲存到 Excel 檔案"""
    # 取得當前日期和月份
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d %H:%M:%S")
    month_str = now.strftime("%Y-%m")

    # Excel 檔案路徑
    excel_file = "seo_search_results.xlsx"

    # 檢查檔案是否存在
    if os.path.exists(excel_file):
        # 載入現有檔案
        wb = load_workbook(excel_file)
    else:
        # 建立新檔案
        wb = Workbook()
        # 移除預設的工作表
        wb.remove(wb.active)

    # 檢查當月分頁是否存在
    if month_str in wb.sheetnames:
        ws = wb[month_str]
    else:
        # 建立新的月份分頁
        ws = wb.create_sheet(month_str)
        # 加入標題列
        ws.append(["日期", "關鍵字", "網址", "目標網站", "排名", "標題"])

    # 寫入搜尋結果
    for rank, (title, url) in enumerate(results, 1):
        # 檢查網址是否為目標網域
        try:
            parsed_url = urlparse(url)
            # www.scincotaiwan.tw 或其子網域
            is_target_domain = 1 if parsed_url.netloc == 'www.scincotaiwan.tw' or parsed_url.netloc == 'scincotaiwan.tw' or parsed_url.netloc.endswith('.scincotaiwan.tw') else 0
        except Exception:
            is_target_domain = 0 # 如果 URL 格式有問題

        ws.append([date_str, keyword, url, is_target_domain, rank, title])

    # 儲存檔案
    wb.save(excel_file)
    print(f"\n✓ 搜尋結果已儲存至 {excel_file}，分頁：{month_str}")

def search_keyword(driver, keyword):
    """搜尋單一關鍵字並回傳結果，會自動翻頁直到收集到 10 筆資料"""
    try:
        # 開啟 Google
        driver.get("https://www.google.com")

        # 等待頁面載入
        time.sleep(2)

        # 找到搜尋欄位
        search_box = driver.find_element(By.NAME, "q")

        # 清空搜尋欄位
        search_box.clear()

        # 模擬人類打字速度
        for char in keyword:
            search_box.send_keys(char)

        # 按下 Enter
        search_box.send_keys(Keys.RETURN)

        # 等待查看結果
        time.sleep(5)

        print(f"\n關鍵字「{keyword}」的搜尋結果：")
        print("=" * 50)

        # 收集搜尋結果
        results = []
        page_count = 1
        # 最多翻 3 頁，或直到收集到 10 筆結果
        while len(results) < 10 and page_count <= 3:
            # 在迴圈開始時檢查是否需要翻頁
            if page_count > 1:
                print(f"\n... 前往第 {page_count} 頁 ...")
                time.sleep(3) # 等待新頁面載入

            # 取得當前頁面的搜尋結果
            # 我們尋找包含 h3 標籤的 <a> 標籤
            search_links = driver.find_elements(By.XPATH, "//a[h3]")

            for link_element in search_links:
                try:
                    title = link_element.find_element(By.TAG_NAME, 'h3').text
                    url = link_element.get_attribute('href')

                    # 略過空白標題和無效網址，並檢查是否重複
                    if title.strip() and url and not any(r[1] == url for r in results):
                        results.append((title, url))
                        print(f"{len(results)}. {title}")
                        if len(results) == 10:
                            break
                except Exception:
                    # 忽略無法解析的元素
                    continue
            
            # 如果還沒滿 10 筆，且還沒達到頁數上限，就嘗試翻頁
            if len(results) < 10 and page_count < 3:
                try:
                    # 嘗試多種方式尋找 "下一頁" 按鈕
                    next_button = None
                    try:
                        # 英文版 Google
                        next_button = driver.find_element(By.CSS_SELECTOR, 'a[aria-label="Next page"]')
                    except Exception:
                        try:
                            # 中文版 Google
                            next_button = driver.find_element(By.CSS_SELECTOR, 'a[aria-label="下一頁"]')
                        except Exception:
                             try:
                                # 舊版 ID
                                next_button = driver.find_element(By.ID, "pnnext")
                             except Exception:
                                print("\n... 找不到 '下一頁' 按鈕，可能已達搜尋結果末頁 ...")
                                break
                    
                    if next_button:
                        # 使用 JavaScript 點擊以避免元素被遮擋
                        driver.execute_script("arguments[0].click();", next_button)
                        page_count += 1
                    else:
                        break

                except Exception as e:
                    print(f"\n翻頁時發生錯誤: {e}")
                    break
            else:
                # 已集滿 10 筆或已達頁數上限
                break

        # 儲存結果到 Excel
        save_to_excel(results, keyword)

        return results

    except Exception as e:
        print(f"搜尋關鍵字「{keyword}」時發生錯誤: {e}")
        return []

def main():
    """主程式"""
    print("=" * 50)
    print("SEO 自動檢查工具")
    print("=" * 50)

    # 載入關鍵字
    keywords = load_keywords()
    if not keywords:
        print("沒有找到關鍵字，程式結束。")
        return

    # 設定 Chrome 無痕模式
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    # 設定 Chrome 選項以避免被偵測為機器人
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    # 建立 Service 物件
    service = Service()
    # 開啟 Chrome 瀏覽器
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # 對每個關鍵字執行搜尋
        for i, keyword in enumerate(keywords, 1):
            print(f"\n[{i}/{len(keywords)}] 正在搜尋關鍵字：{keyword}")
            search_keyword(driver, keyword)

            # 在關鍵字之間等待（避免被偵測為機器人）
            if i < len(keywords):
                wait_time = 5
                print(f"\n等待 {wait_time} 秒後搜尋下一個關鍵字...")
                time.sleep(wait_time)

        print("\n" + "=" * 50)
        print(f"✓ 所有關鍵字搜尋完成！共處理 {len(keywords)} 個關鍵字")
        print("=" * 50)

    finally:
        # 模擬真人點擊關閉按鈕 (macOS 專用)
        try:
            # 使用 AppleScript 關閉最前面的 Chrome 視窗
            applescript = '''
            tell application "Google Chrome"
                if (count of windows) > 0 then
                    close (window 1)
                end if
            end tell
            '''
            subprocess.run(['osascript', '-e', applescript], check=True, timeout=5)
            time.sleep(1)  # 等待視窗關閉動畫
        except Exception as e:
            print(f"AppleScript 關閉視窗失敗: {e}")

        # 確保 driver 和 service 清理
        try:
            driver.quit()
        except Exception as e:
            print(f"driver.quit() 失敗: {e}")

        try:
            service.stop()
        except Exception as e:
            print(f"service.stop() 失敗: {e}")

if __name__ == "__main__":
    main()
