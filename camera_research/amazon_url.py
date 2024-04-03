import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

def wait_for_page_load(driver):
    driver.implicitly_wait(10)

def search_amazon(keyword):
    try:
        driver = webdriver.Chrome()
        driver.get("https://www.amazon.co.jp/")
        wait_for_page_load(driver)

        search_box = driver.find_element(By.ID, "twotabsearchtextbox")
        search_box.send_keys(keyword)
        search_box.submit()

        results = []

        while True:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "search")))
            product_elements = driver.find_elements(By.CSS_SELECTOR, "div[data-component-type='s-search-result']")
            for product_element in product_elements:
                try:
                    product_data = {}
                    product_data["商品名"] = product_element.find_element(By.CSS_SELECTOR, "h2 a").text
                    product_data["URL"] = product_element.find_element(By.CSS_SELECTOR, "h2 a").get_attribute("href")
                    results.append(product_data)
                except Exception as e:
                    print(f"商品情報の取得中にエラーが発生しました:", e)
                    continue

            # 次のページへ遷移
            try:
                next_page_link = driver.find_element(By.XPATH, "//a[contains(@class, 's-pagination-next')]")
                next_page_link.click()
                wait_for_page_load(driver)  # ページの読み込みを待つ
            except NoSuchElementException:
                print("最後のページに達しました。")
                break
        return results
    finally:
        driver.quit()

def main():
    excel_path = r"C:\Users\pyuta\OneDrive\ドキュメント\03_プロジェクト\003_カメラリサーチ\検索リスト.xlsx"
    df = pd.read_excel(excel_path, sheet_name="商品リスト", header=None)
    df = df.dropna()  # 空白行を削除
    keywords = df.iloc[:, 0].tolist()

    all_results = {}
    for keyword in keywords:
        print(f"検索キーワード: {keyword}")
        results = search_amazon(keyword)
        all_results[keyword] = results

    # 検索結果をエクセルファイルに出力
    output_excel_path = r"C:\Users\pyuta\OneDrive\ドキュメント\03_プロジェクト\003_カメラリサーチ\amazon_urls_list.xlsx"
    with pd.ExcelWriter(output_excel_path) as writer:
        for keyword, results in all_results.items():
            df = pd.DataFrame(results)
            df.to_excel(writer, sheet_name=keyword, index=False)

if __name__ == "__main__":
    main()
