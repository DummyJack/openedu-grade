from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os
import re
from dotenv import load_dotenv


# 載入環境變數
load_dotenv(".env")

# 帳密
email = os.getenv("EMAIL")
password = os.getenv("PASSWORD")

# 檢查是否成功載入帳密
if not email or not password:
    print("錯誤：無法從 .env 檔案中讀取帳號或密碼")
    exit(1)

print(f"✅ 成功載入帳號: {email}")

# 啟動 Chrome
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--disable-web-security")
options.add_argument("--disable-features=VizDisplayCompositor")

def create_driver():
    """建立 ChromeDriver 實例"""
    try:
        # 嘗試清除舊的 ChromeDriver 快取
        import shutil

        cache_path = os.path.expanduser("~/.wdm")
        if os.path.exists(cache_path):
            shutil.rmtree(cache_path, ignore_errors=True)
            print("✅ 已清除 ChromeDriver 快取")
        service = Service(ChromeDriverManager().install())
        driver_instance = webdriver.Chrome(service=service, options=options)
        print("✅ ChromeDriver 啟動成功")
        return driver_instance
    except Exception as e:
        print(f"❌ ChromeDriver 啟動失敗: {e}")
        print("嘗試重新啟動 ChromeDriver...")
        driver_instance = webdriver.Chrome(options=options)
        print("✅ 使用系統 ChromeDriver 啟動成功")
        return driver_instance


def login_dashboard(driver_instance):
    """登入儀表板頁面"""
    driver_instance.get(os.getenv("URL2"))

    email_input = WebDriverWait(driver_instance, 10).until(
        EC.presence_of_element_located((By.ID, "login-email"))
    )
    password_input = WebDriverWait(driver_instance, 10).until(
        EC.presence_of_element_located((By.ID, "login-password"))
    )

    email_input.send_keys(email)
    password_input.send_keys(password)

    login_button = WebDriverWait(driver_instance, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.login-button"))
    )
    login_button.click()


def recreate_driver_with_login(old_driver=None):
    """重建 driver 並重新登入（用於容錯復原）"""
    if old_driver is not None:
        try:
            old_driver.quit()
        except Exception:
            pass

    new_driver = create_driver()
    login_dashboard(new_driver)
    return new_driver


driver = recreate_driver_with_login()

# 讀取Excel中的帳戶名稱
try:
    df = pd.read_excel("openedu_all_videos.xlsx")
    accounts = df["帳戶名稱"].unique().tolist()
    print(f"✅ 成功讀取 {len(accounts)} 個帳戶名稱")
except Exception as e:
    print(f"❌ 讀取Excel檔案失敗: {e}")
    driver.quit()
    exit(1)

# 定義成績欄位
grade_columns = [
    "帳戶名稱",
    "隨測_第一章",
    "隨測_第二章",
    "隨測_第三章",
    "隨測_第四章",
    "隨測_第五章",
    "隨測_第六章",
    "評量_第一章",
    "評量_第二章",
    "評量_第三章",
    "評量_第四章",
    "評量_第五章",
    "評量_第六章",
    "OJ_第一章",
    "OJ_第二章",
    "OJ_第三章",
    "OJ_第四章",
    "OJ_第五章", 
    "OJ_第七章"
]

# 儲存所有成績資料
all_grades_data = []
error_data = []


def save_openedu_grades_score(results_df, grade_columns, choice):
    """選項 2／3：與現有 openedu_grades_score.xlsx 合併，只更新／新增本次結果中的帳戶，其餘列保留。"""
    excel_path = "./scores/openedu_grades_score.xlsx"
    if not os.path.exists("./scores"):
        os.makedirs("./scores")

    if results_df.empty:
        return

    if not os.path.exists(excel_path):
        results_df.to_excel(excel_path, index=False)
        print(
            f"\n✅ 成功建立新檔案並儲存 {len(results_df)} 筆成績資料到 openedu_grades_score.xlsx"
        )
        return

    try:
        existing_df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"⚠️ 讀取現有檔案失敗，將直接覆蓋: {e}")
        results_df.to_excel(excel_path, index=False)
        print(
            f"\n✅ 成功儲存 {len(results_df)} 筆成績資料到 openedu_grades_score.xlsx"
        )
        return

    if "帳戶名稱" not in existing_df.columns:
        results_df.to_excel(excel_path, index=False)
        print(
            f"\n✅ 現有檔缺少「帳戶名稱」欄位，已覆蓋寫入 {len(results_df)} 筆"
        )
        return

    new_set = set(
        results_df["帳戶名稱"].dropna().astype(str).str.strip().tolist()
    )
    existing_key = existing_df["帳戶名稱"].fillna("").astype(str).str.strip()
    existing_accounts = existing_key.tolist()
    new_list = list(new_set)
    updated_accounts = [a for a in new_list if a in existing_accounts]
    new_accounts = [a for a in new_list if a not in existing_accounts]

    all_cols = list(dict.fromkeys(list(existing_df.columns) + list(grade_columns)))
    existing_aligned = existing_df.reindex(columns=all_cols)
    results_aligned = results_df.reindex(columns=all_cols)

    combined_df = pd.concat(
        [existing_aligned[~existing_key.isin(new_set)], results_aligned],
        ignore_index=True,
    )
    combined_df.to_excel(excel_path, index=False)

    print(
        f"\n✅ 成功處理 {len(results_df)} 個帳戶的成績資料（已合併至檔案，其餘帳戶列保留）"
    )
    if choice == "2":
        if updated_accounts:
            print(f"   🔄 更新的帳戶: {updated_accounts}")
        if new_accounts:
            print(f"   ➕ 新增的帳戶: {new_accounts}")


# 對帳戶進行處理
test_accounts = None
choice = None  # 全域變數，用於追蹤使用者選擇

while test_accounts is None:
    print("\n請選擇處理模式：")
    print("1. 處理全部帳戶")
    print("2. 指定特定帳戶")
    print("3. 重新處理錯誤帳戶（errors/grade_errors.xlsx）")

    choice = input("請輸入選擇 (1、2 或 3): ").strip()

    if choice == "1":
        test_accounts = accounts
        print(f"✅ 已選擇處理全部 {len(accounts)} 個帳戶")
    elif choice == "2":
        while test_accounts is None:
            print("\n請輸入要處理的帳戶名稱（多個帳戶請用 , 分隔）：")
            input_accounts = input("帳戶名稱: ").strip()
            
            if not input_accounts:
                print("❌ 未輸入帳戶名稱，請至少輸入一個帳戶名稱")
                continue
            
            # 分割輸入的帳戶名稱並去除空白
            specified_accounts = [acc.strip() for acc in input_accounts.split(",") if acc.strip()]
            
            if not specified_accounts:
                print("❌ 未輸入有效的帳戶名稱，請至少輸入一個帳戶名稱")
                continue
            
            # 檢查輸入的帳戶是否存在於帳戶列表中
            valid_accounts = []
            invalid_accounts = []
            
            for acc in specified_accounts:
                if acc in accounts:
                    valid_accounts.append(acc)
                else:
                    invalid_accounts.append(acc)
            
            if invalid_accounts:
                print(f"⚠️ 以下帳戶不存在於帳戶列表中: {invalid_accounts}")
            
            if valid_accounts:
                test_accounts = valid_accounts
                print(f"✅ 將處理以下 {len(valid_accounts)} 個帳戶: {valid_accounts}")
            else:
                print("❌ 沒有有效的帳戶，請重新輸入")
    elif choice == "3":
        error_file_path = "./errors/grade_errors.xlsx"
        if not os.path.exists(error_file_path):
            print(f"❌ 找不到錯誤帳戶檔案: {error_file_path}")
            print("   請先執行一次完整處理，或改用選項 1 / 2")
            continue

        try:
            error_df = pd.read_excel(error_file_path)
            if "帳戶名稱" not in error_df.columns:
                print("❌ 錯誤帳戶檔案缺少「帳戶名稱」欄位，請檢查檔案格式")
                continue

            retry_accounts = (
                error_df["帳戶名稱"]
                .dropna()
                .astype(str)
                .str.strip()
            )
            retry_accounts = [acc for acc in retry_accounts.tolist() if acc]

            if not retry_accounts:
                print("❌ 錯誤帳戶檔案中沒有可重算的帳戶")
                continue

            # 只處理仍存在於總帳戶列表中的帳戶，避免無效輸入
            account_set = set(str(acc).strip() for acc in accounts)
            valid_retry_accounts = [acc for acc in retry_accounts if acc in account_set]
            invalid_retry_accounts = [acc for acc in retry_accounts if acc not in account_set]

            if invalid_retry_accounts:
                print(f"⚠️ 以下錯誤帳戶不存在於來源清單，將略過: {invalid_retry_accounts}")

            if not valid_retry_accounts:
                print("❌ 錯誤帳戶清單皆不在來源帳戶中，無法處理")
                continue

            # 去重複並維持原順序
            valid_retry_accounts = list(dict.fromkeys(valid_retry_accounts))
            test_accounts = valid_retry_accounts
            print(f"✅ 已載入錯誤帳戶，共 {len(test_accounts)} 人，將重新計算成績")
        except Exception as e:
            print(f"❌ 讀取錯誤帳戶檔案失敗: {e}")
            continue
    else:
        print("❌ 無效的選擇，請輸入 1、2 或 3")

for idx, account in enumerate(test_accounts, 1):
    print(f"\n處理帳戶 {idx}/{len(test_accounts)}: {account}")

    # 錯誤重試機制：先標準策略，失敗後激進策略
    max_account_retries = 3
    account_success = False

    for attempt in range(max_account_retries):
        if attempt > 0:
            print(f"  ⚠️ 重試第 {attempt} 次...")

        try:
            # 判斷是否使用激進策略
            use_aggressive_strategy = attempt > 0

            if use_aggressive_strategy:
                if attempt == 1:
                    driver.get(os.getenv("URL2"))
                else:
                    driver.refresh()

            select_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "action-preview-select"))
            )
            select = Select(select_element)

            # 選擇「特定學生」選項
            select.select_by_value("specific student")
            print("✅ 已選擇'特定學生'選項")

            # 重新找到使用者名稱輸入框
            username_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "action-preview-username"))
            )

            # 清空輸入框並輸入帳戶名稱
            username_input.clear()
            username_input.send_keys(account)

            # 按下Enter鍵提交表單
            username_input.send_keys(Keys.RETURN)

            time.sleep(1)

            # 驗證使用者名稱是否正確
            h2_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "h2.progress-certificates-title")
                )
            )

            h2_text = re.search(r"'([^']+)'", h2_element.text).group(1)

            if h2_text == account:
                print(
                    f"      ✅ 使用者名稱驗證成功：輸入 {account}、頁面顯示 {h2_text}"
                )
            else:
                print(
                    f"      ❌ 使用者名稱驗證失敗：輸入 {account}，但頁面顯示 {h2_text}"
                )
                raise Exception(
                    f"使用者名稱驗證失敗：輸入 {account}，但頁面顯示 {h2_text}"
                )

            h4_elements = WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "h4"))
            )

            # 初始化帳戶成績和章節分數統計
            account_grades = {"帳戶名稱": account}
            for col in grade_columns[1:]:  # 跳過帳戶名稱欄位
                account_grades[col] = 0  # 改為數字0而非空字串

            # 用於統計各章節的其他分數（非自我測驗、非程式練習）
            chapter_other_scores = {1: [], 2: [], 3: [], 4: [], 5: [], 6: []}

            # 用索引方式重新抓取h4元素，避免stale element錯誤
            h4_count = len(h4_elements)

            for i in range(h4_count):
                try:
                    # 每次重新尋找h4元素，避免stale element錯誤
                    current_h4_elements = driver.find_elements(By.TAG_NAME, "h4")
                    if i >= len(current_h4_elements):
                        continue

                    h4 = current_h4_elements[i]

                    # 重新取得h4的文字內容和span元素
                    h4_text = h4.text.strip()
                    span_elements = h4.find_elements(By.TAG_NAME, "span")

                    if span_elements:
                        for span in span_elements:
                            try:
                                span_text = span.text.strip()
                                if span_text and "(" in span_text and "/" in span_text:
                                    # 使用正則表達式提取括號中斜線前的數字（支援科學記號）
                                    match = re.search(r"\(([0-9.eE+-]+)/", span_text)
                                    if match:
                                        try:
                                            # 將科學記號轉換為數字，然後轉為整數
                                            score = int(float(match.group(1)))
                                        except (ValueError, OverflowError):
                                            # 如果轉換失敗，嘗試直接轉為整數
                                            try:
                                                score = int(match.group(1))
                                            except ValueError:
                                                print(
                                                    f"      ⚠️ 無法轉換分數: {match.group(1)}"
                                                )
                                                continue

                                        # 判斷是哪種類型的項目
                                        if "自我測驗" in h4_text:
                                            # 自我測驗 -> 評量欄位
                                            if "1.6" in h4_text:
                                                account_grades["評量_第一章"] = score
                                                print(f"      ✅ 評量_第一章: {score}")
                                            elif "2.7" in h4_text:
                                                account_grades["評量_第二章"] = score
                                                print(f"      ✅ 評量_第二章: {score}")
                                            elif "3.4" in h4_text:
                                                account_grades["評量_第三章"] = score
                                                print(f"      ✅ 評量_第三章: {score}")
                                            elif "4.6" in h4_text:
                                                account_grades["評量_第四章"] = score
                                                print(f"      ✅ 評量_第四章: {score}")
                                            elif "5.7" in h4_text:
                                                account_grades["評量_第五章"] = score
                                                print(f"      ✅ 評量_第五章: {score}")
                                            elif "6.5" in h4_text:
                                                account_grades["評量_第六章"] = (
                                                    score * 50
                                                )
                                                print(
                                                    f"      ✅ 評量_第六章: {score*50}"
                                                )

                                        elif "程式練習" in h4_text:
                                            # 程式練習 -> OJ欄位
                                            if "1.5" in h4_text:
                                                account_grades["OJ_第一章"] = score
                                                print(f"      ✅ OJ_第一章: {score}")
                                            elif "2.8" in h4_text:
                                                account_grades["OJ_第二章"] = score
                                                print(f"      ✅ OJ_第二章: {score}")
                                            elif "3.5" in h4_text:
                                                account_grades["OJ_第三章"] = score
                                                print(f"      ✅ OJ_第三章: {score}")
                                            elif "4.7" in h4_text:
                                                account_grades["OJ_第四章"] = score
                                                print(f"      ✅ OJ_第四章: {score}")
                                            elif "5.8" in h4_text:
                                                account_grades["OJ_第五章"] = score
                                                print(f"      ✅ OJ_第五章: {score}")
                                            elif "7.4" in h4_text:
                                                account_grades["OJ_第七章"] = score
                                                print(f"      ✅ OJ_第七章: {score}")

                                        else:
                                            # 計入對應章節的隨測總分
                                            chapter = None
                                            if h4_text.startswith("1."):
                                                chapter = 1
                                            elif h4_text.startswith("2."):
                                                chapter = 2
                                            elif h4_text.startswith("3."):
                                                chapter = 3
                                            elif h4_text.startswith("4."):
                                                chapter = 4
                                            elif h4_text.startswith("5."):
                                                chapter = 5
                                            elif h4_text.startswith("6."):
                                                chapter = 6

                                            if (
                                                chapter
                                                and chapter in chapter_other_scores
                                            ):
                                                chapter_other_scores[chapter].append(
                                                    score
                                                )
                                            else:
                                                print(
                                                    f"      ⚠️ 無法分類的項目: {h4_text}"
                                                )
                            except Exception as span_e:
                                print(f"    ⚠️ 處理span元素時發生錯誤: {span_e}")
                                continue

                except Exception as e:
                    continue

            # 計算各章節隨測總分
            for chapter, scores in chapter_other_scores.items():
                total_score = sum(scores)
                column_name = (
                    f"隨測_第{['', '一', '二', '三', '四', '五', '六'][chapter]}章"
                )
                account_grades[column_name] = total_score
                print(f"      {column_name}: {scores} = {total_score}")

            if not os.path.exists("./scores"):
                os.makedirs("./scores")

            all_grades_data.append(account_grades)
            print(f"  ✅ 帳戶 {account} 處理完成")

            # 儲存成功的成績資料：選項 1 整檔覆寫；選項 2／3 與現有檔合併
            if all_grades_data:
                results_df = pd.DataFrame(all_grades_data, columns=grade_columns)
                if choice == "1":
                    results_df.to_excel(
                        "./scores/openedu_grades_score.xlsx", index=False
                    )
                    print(
                        f"\n✅ 成功儲存 {len(all_grades_data)} 筆成績資料到 openedu_grades_score.xlsx"
                    )
                else:
                    save_openedu_grades_score(results_df, grade_columns, choice)
                    
            account_success = True
            break  # 成功後跳出重試迴圈

        except Exception as e:
            if not os.path.exists("./errors"):
                os.makedirs("./errors")
            # 提取 Alert Text 內容
            error_message = str(e)
            if "Message:" in error_message:
                # 提取 Message: 後面的內容，直到換行或結束
                match = re.search(r"Message:\s*(.+?)(?:\n|$)", error_message, re.DOTALL)
                if match:
                    error_message = match.group(1).strip()

            print(
                f"  ❌ 嘗試 {attempt + 1}/{max_account_retries} 處理帳戶 {account} 時發生錯誤: {error_message}"
            )
            if attempt == max_account_retries - 1:  # 最後一次重試也失敗
                # 將錯誤資料儲存到專門的錯誤列表
                error_data.append(
                    {
                        "帳戶名稱": account,
                        "錯誤訊息": str(error_message),
                        "處理順序": idx,
                    }
                )
                error_df = pd.DataFrame(error_data)
                error_df.to_excel("./errors/grade_errors.xlsx", index=False)
                print(f"✅ 成功儲存 {len(error_data)} 筆錯誤資料到 grade_errors.xlsx")
            else:
                print(f"    準備重試帳戶 {account}")

# 顯示總結
print(f"\n總結:")
print(f"帳戶總數: {len(accounts)}")
print(f"成功處理帳戶數: {len(all_grades_data)}")
print(f"失敗帳戶數: {len(error_data)}")

driver.quit()
