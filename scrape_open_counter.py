"""
調達ポータル オープンカウンター案件 自動取得スクリプト

使い方:
  pip install selenium pandas
  # ChromeDriver は Chrome と同じバージョンが自動インストールされます
  python scrape_open_counter.py

出力: open_counter_cases.csv
"""

import time
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# ChromeDriverを自動管理するため selenium-manager を使用（Selenium 4.6以上で自動）
def create_driver():
    options = Options()
    # options.add_argument("--headless")  # ← ヘッドレスにしたい場合はコメント解除
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1280,900")
    return webdriver.Chrome(options=options)


def fetch_open_counter_cases():
    driver = create_driver()
    wait = WebDriverWait(driver, 15)
    cases = []

    try:
        print("ポータルにアクセス中...")
        driver.get("https://www.p-portal.go.jp/pps-web-biz/UAA01/OAA0101")
        time.sleep(3)

        # 調達方式「オープンカウンタ（少額）」を選択
        try:
            調達方式 = wait.until(EC.presence_of_element_located((By.NAME, "調達方式")))
            select = Select(調達方式)
            # 選択肢の中からオープンカウンタを選ぶ
            for option in select.options:
                if "オープンカウンタ" in option.text or "open" in option.text.lower():
                    option.click()
                    break
        except Exception:
            # セレクトボックスが見つからない場合、チェックボックス等を探す
            try:
                checkboxes = driver.find_elements(By.XPATH, "//*[contains(text(),'オープンカウンタ')]")
                for cb in checkboxes:
                    cb.click()
                    break
            except Exception as e:
                print(f"調達方式の選択をスキップ: {e}")

        # 検索ボタンをクリック
        try:
            search_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(text(),'検索')] | //input[@value='検索']")
            ))
            search_btn.click()
            time.sleep(3)
        except Exception as e:
            print(f"検索ボタンが見つかりません: {e}")

        # ページネーションを含む全案件を取得
        page = 1
        while True:
            print(f"  ページ {page} を取得中...")
            time.sleep(2)

            # テーブル行を取得（サイトの構造に合わせて調整）
            rows = driver.find_elements(By.CSS_SELECTOR, "table tr, .case-row, .list-row, [class*='row']")
            found = 0
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 3:
                    case = [c.text.strip() for c in cells]
                    if any(case):  # 空行を除外
                        cases.append(case)
                        found += 1

            print(f"    → {found} 件取得")

            # 次のページへ
            try:
                next_btn = driver.find_element(By.XPATH,
                    "//a[contains(text(),'次へ')] | //button[contains(text(),'次へ')] | //a[contains(@class,'next')]"
                )
                if not next_btn.is_enabled():
                    break
                next_btn.click()
                page += 1
            except Exception:
                break  # 次ページなし

        # ページのHTMLをそのままダンプして確認用に保存
        with open("portal_page.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("portal_page.html にページ内容を保存しました（デバッグ用）")

    finally:
        driver.quit()

    return cases


def save_csv(cases, filename="open_counter_cases.csv"):
    if not cases:
        print("案件データが取得できませんでした。portal_page.html を確認してください。")
        return

    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["案件名", "発注機関", "公告日", "締切日", "予定金額", "業種・内容", "備考"])
        for case in cases:
            # 列数が少ない場合は空白で補完
            padded = case + [""] * (7 - len(case))
            writer.writerow(padded[:7])

    print(f"\n✅ {len(cases)} 件を {filename} に保存しました")


if __name__ == "__main__":
    cases = fetch_open_counter_cases()
    save_csv(cases)
    print("\n--- 取得結果プレビュー（先頭5件）---")
    for c in cases[:5]:
        print(c)
