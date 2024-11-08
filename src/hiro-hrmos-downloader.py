import random
import time
import csv
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ChromeDriverのパスを指定
driver_path = "C:/Users/USER/Desktop/chromedriver.exe"
service = Service(driver_path)

# WebDriverの設定
driver = webdriver.Chrome(service=service)

# ハーモス勤怠の管理ページにアクセス
driver.get("https://f.ieyasu.co/hirovet-ah/login")

# メモ書き用のCSVファイルを開く（追記モードで開く）
with open('C:/Users/USER/Desktop/Hiro名簿変換前/download_log.csv', mode='a', newline='', encoding='utf-8-sig') as log_file:
    log_writer = csv.writer(log_file)
    
    # ログイン処理
    try:
        driver.find_element(By.ID, "user_login_id").send_keys("result@hirovet.com")  # IDを入力
        driver.find_element(By.ID, "user_password").send_keys("Nodapopo-")  # パスワードを入力
        driver.find_element(By.CLASS_NAME, "btnSubmit").click()

        time.sleep(2)  # ページの読み込みを待機
        print("ログインが成功しました。")



        try:
            monthly_link = driver.find_element(By.ID, "montyl_working_summary_link")
            monthly_link.click()
            print("「月別データ」リンクをクリックしました。")
            time.sleep(2)  # リンククリック後のページ読み込みを待機
        except Exception as e:
            print("「月別データ」リンクをクリックできませんでした:", e)

        # カレンダーアイコンのボタンをクリック
        try:
            calendar_button = driver.find_element(By.XPATH, '//button/img[@src="/asset/design1/img/templates/form_icon_calendar_16px.png"]/parent::button')
            calendar_button.click()
            print("カレンダーアイコンのボタンをクリックしました。")
            time.sleep(2)  # カレンダーが表示されるのを待機
        except Exception as e:
            print("カレンダーアイコンのボタンをクリックできませんでした:", e)

        # 「10月」のリンクをクリック
        try:
            october_button = driver.find_element(By.CLASS_NAME, "htBlock-calendarMonth_dialog_month9")
            october_button.click()
            print("「10月」のボタンをクリックしました。")
            time.sleep(2)  # 「10月」選択後の待機
        except Exception as e:
            print("「10月」のボタンをクリックできませんでした:", e)

        # 「表示」ボタンをクリックして「10月」を確定
        try:
            display_button = driver.find_element(By.ID, "display_button")
            display_button.click()
            print("「表示」ボタンをクリックしました。")
            time.sleep(2)  # 「表示」ボタンをクリック後のページ読み込みを待機

        except Exception as e:
            print("「表示」ボタンをクリックできませんでした:", e)

        # 従業員ごとに操作を繰り返す
        for employee_id in ["1401689245", "1401689258"]:
            try:
                # 従業員のボタンをクリック
                target_button = driver.find_element(By.ID, f"button_{employee_id}")
                target_button.click()
                print(f"従業員 {employee_id} のボタンをクリックしました。")
                time.sleep(2)  # ページ読み込み待機
                
                # 従業員名を取得
                employee_info = driver.find_element(By.XPATH, '//div[@class="htBlock-headerPanel_content_row"]/p')
                employee_name = employee_info.find_elements(By.TAG_NAME, 'span')[-1].text.replace(" ", "_")  # 最後のspan要素から取得
                print(f"従業員名を取得しました: {employee_name}")

                # 現在の日時を取得（ダウンロード時のタイムスタンプとして使用）
                download_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # 従業員ID、名前、タイムスタンプをCSVファイルに記録
                log_writer.writerow([employee_id, employee_name, download_timestamp])

                # セレクトボックスで「EXCEL」を選択
                select_element = Select(driver.find_element(By.ID, "export_timecard_select"))
                select_element.select_by_value("EXCEL")
                print(f"従業員 {employee_id} の「EXCEL」を選択しました。")
                time.sleep(2)  # 選択後の待機

                # 「出力」ボタンがクリック可能になるまで待機してクリック
                wait = WebDriverWait(driver, 10)  # 最大10秒待機
                output_button = wait.until(EC.element_to_be_clickable((By.ID, "excel_out_button")))
                output_button.click()
                print(f"従業員 {employee_id} の「出力」ボタンをクリックしました。")
                time.sleep(2)  # 「出力」ボタンをクリック後の待機

                # 「ダウンロード」ボタンが有効になるのを待ってクリック
                download_button = wait.until(EC.text_to_be_present_in_element((By.ID, "excel_out_button"), "ダウンロード"))
                driver.find_element(By.ID, "excel_out_button").click()
                print(f"従業員 {employee_id} の「ダウンロード」ボタンをクリックしました。")
                time.sleep(2)  # ダウンロード後の待機

                # 前のページに戻る（Chromeの戻るボタンを模倣）
                driver.back()
                print("前のページに戻りました。")
                time.sleep(5)  # ページ読み込み待機
                
            except Exception as e:
                print(f"従業員 {employee_id} のデータ取得中にエラーが発生しました:", e)

        # セッションの確認として現在のページURLを表示
        current_url = driver.current_url
        print("現在のURL:", current_url)

        # ページタイトルの確認
        page_title = driver.title
        print("現在のページタイトル:", page_title)

    except Exception as e:
        print("エラーが発生しました:", e)

    finally:
        # WebDriverを終了
        driver.quit()
