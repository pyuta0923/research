import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time

def wait_for_page_load(driver):
    # ページの読み込みが完了するまで待つ
    time.sleep(5)  # 適切な待機時間に調整する必要があります

def wait_for_product_list(driver):
    # 商品一覧エリアの要素が表示されるまで待機
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "product-list")))

def search_kitamura(keyword):
    try:
        # WebDriverを起動
        driver = webdriver.Chrome()

        # キタムラオンラインショップにアクセス
        driver.get("https://shop.kitamura.jp/")

        # ページの読み込みが完了するまで待つ
        wait_for_page_load(driver)

        # 検索ボックスにキーワードを入力して検索
        search_box = driver.find_element(By.ID, "search-keyword")
        search_box.send_keys(keyword)

        # 検索ボタンをクリック
        search_button = driver.find_element(By.CSS_SELECTOR, "button.v-icon.far.fa-search")
        search_button.click()

        # 商品一覧ページの表示が完了するまで待機
        wait_for_product_list(driver)

        # 商品情報を取得
        products = []
        product_elements = driver.find_elements(By.CLASS_NAME, "product-info")
        for index, element in enumerate(product_elements, start=1):
            product_data = {}
            try:
                product_data["メーカ名"] = element.find_element(By.CLASS_NAME, "product-maker-name").text
                product_data["商品名"] = element.find_element(By.CLASS_NAME, "product-name").text
                product_data["価格"] = element.find_element(By.CSS_SELECTOR, ".product-price-area .product-price").text
                product_data["中古"] = element.find_element(By.CSS_SELECTOR, ".product-price-area .product-used-price").text
                product_data["URL"] = driver.current_url
            except NoSuchElementException as e:
                print(f"商品名: {keyword} の製品情報({index}番目)の取得中にエラーが発生しました:", e)
                continue

            products.append(product_data)

        return products
    finally:
        # WebDriverを終了
        driver.quit()

def main():
    # Excelファイルから検索キーワードを読み込む
    file_path = r"C:\Users\pyuta\OneDrive\ドキュメント\03_プロジェクト\003_カメラリサーチ\検索リスト.xlsx"
    df = pd.read_excel(file_path, sheet_name="商品リスト", header=None)
    
    # データフレームの最初の列から空白セルに達するまでの行を取得
    keyword_list = df.iloc[:, 0].tolist()
    keyword_list = [keyword for keyword in keyword_list if pd.notnull(keyword)]

    for keyword in keyword_list:
        products = search_kitamura(keyword)
        if products:
            print(f"{keyword} の検索結果:")
            for product in products:
                print(product)
        else:
            print(f"{keyword} の検索結果: 商品が見つかりませんでした。")

if __name__ == "__main__":
    main()
