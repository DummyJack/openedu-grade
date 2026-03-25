from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import pandas as pd
import re
import os
from dotenv import load_dotenv

# 載入環境變數
load_dotenv(".env")

# 帳密
email = os.getenv("EMAIL")
password = os.getenv("PASSWORD")

# 檢查是否成功載入帳密
if not email or not password:
    print("錯誤：無法從 .env 檔案中讀取帳號或密碼")
    print("請確保 .env 檔案存在且包含 EMAIL 和 PASSWORD")
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

try:
    # 嘗試清除舊的 ChromeDriver 快取
    import shutil

    cache_path = os.path.expanduser("~/.wdm")
    if os.path.exists(cache_path):
        shutil.rmtree(cache_path, ignore_errors=True)
        print("✅ 已清除 ChromeDriver 快取")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("✅ ChromeDriver 啟動成功")
except Exception as e:
    print(f"❌ ChromeDriver 啟動失敗: {e}")
    print("嘗試重新啟動 ChromeDriver...")
    driver = webdriver.Chrome(options=options)
    print("✅ 使用系統 ChromeDriver 啟動成功")

driver.get(os.getenv("URL1"))

# Login 按鈕
login_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "Login"))
)
login_btn.click()

# Email/密碼
email_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "login-email"))
)
password_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "login-password"))
)
email_input.send_keys(email)
password_input.send_keys(password)

driver.find_element(By.CSS_SELECTOR, "button.login-button").click()

# 主頁資料表
rows = WebDriverWait(driver, 20).until(
    EC.presence_of_all_elements_located(
        (By.CSS_SELECTOR, "#sample_editable_1 tbody tr")
    )
)

# 爬取帳戶清單

main_window = driver.current_window_handle
rows = driver.find_elements(By.CSS_SELECTOR, "#sample_editable_1 tbody tr")

all_data = []
video_errors = []  # 格式: [{"帳戶名稱": account, "錯誤原因": reason}]

# account_count = 0

total_accounts = len(rows)
for idx, row in enumerate(rows, start=1):
    account_cells = row.find_elements(By.TAG_NAME, "td")
    if not account_cells:
        continue
    account = account_cells[0].text.strip()
    onclick = row.get_attribute("onclick")
    if not onclick:
        continue

    m = re.search(r'window\.open\((["\'])(.+?)\1\)', onclick)
    if not m:
        print("找不到連結！", onclick)
        video_errors.append({"帳戶名稱": account, "錯誤原因": f"找不到連結: {onclick}"})
        continue
    url = "https://dashboard.openedu.tw" + m.group(2)
    print("抓取:", account, url)

    # 開新分頁
    before_windows = driver.window_handles
    driver.execute_script(f"window.open('{url}');")

    # 等新分頁出現（最多等10秒）
    WebDriverWait(driver, 10).until(
        lambda d: len(d.window_handles) > len(before_windows)
    )

    # 找新分頁 handle
    new_window = [w for w in driver.window_handles if w not in before_windows][0]
    driver.switch_to.window(new_window)

    # 找到影片表格
    try:
        # 等分頁內容載入（如表格）
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "sample_editable_1"))
        )
        vrows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
        for vrow in vrows:
            vtd = vrow.find_elements(By.TAG_NAME, "td")
            video_name = vtd[0].text.strip()
            watch_count = vtd[1].text.strip()
            watch_time = vtd[2].text.strip()
            video_length = vtd[3].text.strip()
            finish_rate = vtd[4].text.strip()
            all_data.append(
                {
                    "帳戶名稱": account,
                    "影片名稱": video_name,
                    "觀看數": watch_count,
                    "觀看時間(分)": watch_time,
                    "影片長度(分)": video_length,
                    "完看率": finish_rate,
                }
            )
        # 匯出資料
        df = pd.DataFrame(all_data)
        df.to_excel("openedu_all_videos.xlsx", index=False)
    except Exception as e:
        # 新增：即使沒資料也寫一筆空資料
        all_data.append(
            {
                "帳戶名稱": account,
                "影片名稱": "",
                "觀看數": "",
                "觀看時間(分)": "",
                "影片長度(分)": "",
                "完看率": "",
            }
        )
        # 匯出資料
        df = pd.DataFrame(all_data)
        df.to_excel("openedu_all_videos.xlsx", index=False)

    # 關掉分頁，切回主分頁
    driver.close()
    driver.switch_to.window(main_window)
    WebDriverWait(driver, 5).until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "#sample_editable_1 tbody tr")
        )
    )

    print(
        f"已處理帳號數: {idx}/{total_accounts}，剩餘: {total_accounts-idx}，目前帳號: {account}"
    )

    # account_count += 1
    # if account_count >= 2:
    #     break

# 最終資料處理：檢查並修正重複的 "Video: for range"
print("檢查重複的影片記錄...")


def fix_video(df):
    """檢查並修正重複的 'Video: for range' 記錄"""
    modified = False

    # 按帳戶分組處理
    for account in df["帳戶名稱"].unique():
        # 找到該帳戶的所有 "Video: for range" 記錄
        account_mask = df["帳戶名稱"] == account
        video_for_range_mask = df["影片名稱"] == "Video: for range"
        duplicate_indices = df[account_mask & video_for_range_mask].index.tolist()

        if len(duplicate_indices) > 1:
            print(
                f"發現帳戶 '{account}' 有 {len(duplicate_indices)} 個 'Video: for range' 記錄"
            )
            print(f"將第二個記錄改為 'Video: for range2'")

            # 將第二個記錄改名為 "Video: for range2"
            df.loc[duplicate_indices[1], "影片名稱"] = "Video: for range2"
            modified = True

    if modified:
        print("✅ 已修正重複的影片記錄")
    else:
        print("✅ 沒有發現重複的 'Video: for range' 記錄")

    return df


# 修正重複記錄
df_final = fix_video(df)

# 最終匯出
df_final.to_excel("openedu_all_videos.xlsx", index=False)
print("✅ 最終資料已匯出至 openedu_all_videos.xlsx")

# 匯出錯誤帳號
if video_errors:
    if not os.path.exists("./errors"):
        os.makedirs("./errors")
    print("有錯誤帳號，請檢查 video_errors.xlsx")
    pd.DataFrame(video_errors).to_excel("./errors/video_errors.xlsx", index=False)

driver.quit()
