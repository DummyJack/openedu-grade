import pandas as pd
import os
import re

score_xlsx = pd.ExcelFile("./video_length_python.xlsx")

def get_chapter_from_video_name(video_name):
    """從影片名稱中提取章節編號"""
    if pd.isna(video_name):
        return None
    video_str = str(video_name)
    if video_str.startswith("#"):
        # 處理 "# 第一章" 格式
        match = re.search(r"第([一二三四五六七\d]+)章", video_str)
        if match:
            chapter_text = match.group(1)
            # 轉換中文數字為阿拉伯數字
            chinese_to_num = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7}
            return chinese_to_num.get(chapter_text, int(chapter_text) if chapter_text.isdigit() else None)
    return None

def get_available_chapters(video_df):
    """獲取可用的章節列表"""
    chapters = set()
    current_chapter = None
    
    for _, row in video_df.iterrows():
        video_name = row["Video"]
        chapter = get_chapter_from_video_name(video_name)
        if chapter is not None:
            current_chapter = chapter
            chapters.add(chapter)
        elif current_chapter is not None:
            # 如果當前影片不是章節標題，但有當前章節，則屬於該章節
            continue
    
    return sorted(list(chapters))

def filter_videos_by_chapters(video_df, start_chapter=None, end_chapter=None):
    """根據章節範圍過濾影片"""
    if start_chapter is None and end_chapter is None:
        # 排除章節標題行（#開頭）、合計行和純數字的總和項目
        filtered_df = video_df[~video_df["Video"].astype(str).str.startswith("#")]
        filtered_df = filtered_df[filtered_df["Video"] != "合計"]
        filtered_df = filtered_df[~filtered_df["Video"].astype(str).str.strip().str.isdigit()]
        return filtered_df
    
    result_videos = []
    current_chapter = None
    include_current = False
    
    for _, row in video_df.iterrows():
        video_name = row["Video"]
        chapter = get_chapter_from_video_name(video_name)
        
        if chapter is not None:
            # 這是章節標題
            current_chapter = chapter
            if start_chapter is None:
                include_current = chapter <= end_chapter
            elif end_chapter is None:
                include_current = chapter >= start_chapter
            else:
                include_current = start_chapter <= chapter <= end_chapter
        elif current_chapter is not None and include_current:
            # 這是普通影片，且屬於要包含的章節
            if str(video_name) != "合計" and not str(video_name).strip().isdigit():
                result_videos.append(row)
    
    return pd.DataFrame(result_videos)

def get_user_chapter_choice():
    """獲取用戶的章節選擇"""
    print("\n========== 成績結算範圍選擇 ==========")
    return get_custom_chapter_range()

def get_custom_chapter_range():
    """獲取用戶自定義的章節範圍"""
    # 獲取可用章節
    sheetname = [n for n in score_xlsx.sheet_names if "總表" in n or "影片" in n][0]
    temp_video_df = score_xlsx.parse(sheetname)
    temp_video_df = temp_video_df[["影片", "影片長度(分)"]].drop_duplicates()
    temp_video_df = temp_video_df.rename(columns={"影片": "Video", "影片長度(分)": "VideoLength"})
    
    available_chapters = get_available_chapters(temp_video_df)
    print(f"\n可用的章節: {available_chapters}")
    print("輸入說明:")
    print("- 輸入 0: 計算全部章節")
    print("- 輸入範圍: 例如 3-5, 2-4")
    
    while True:
        range_input = input("\n請輸入章節範圍: ").strip()
        
        try:
            # 處理輸入 0 的情況（全部章節）
            if range_input == "0":
                print("✅ 已選擇全部章節")
                return None, None, "全部章節"
            
            if "-" not in range_input:
                print("❌ 請使用正確格式，例如: 1-7 或輸入 0")
                continue
                
            parts = range_input.split("-")
            if len(parts) != 2:
                print("❌ 請使用正確格式，例如: 1-7 或輸入 0")
                continue
                
            start_ch = int(parts[0].strip())
            end_ch = int(parts[1].strip())
            
            # 驗證章節是否存在
            if start_ch not in available_chapters:
                print(f"❌ 起始章節 {start_ch} 不存在，可用章節: {available_chapters}")
                continue
                
            if end_ch not in available_chapters:
                print(f"❌ 結束章節 {end_ch} 不存在，可用章節: {available_chapters}")
                continue
                
            if start_ch > end_ch:
                print("❌ 起始章節必須小於或等於結束章節")
                continue
                
            # 顯示所選範圍的詳細內容
            print(f"\n📋 所選範圍內容詳覽 (第{start_ch}章到第{end_ch}章):")
            print("-" * 50)
            
            current_chapter = None
            chapter_videos = {}
            
            # 初始化所選章節字典
            for ch in range(start_ch, end_ch + 1):
                chapter_videos[ch] = []
            
            # 收集所選範圍內的影片
            chapter_video_counter = {}  # 每章節的影片計數器
            for _, row in temp_video_df.iterrows():
                video_name = row["Video"]
                video_length = row["VideoLength"]
                chapter = get_chapter_from_video_name(video_name)
                
                if chapter is not None:
                    current_chapter = chapter
                    if start_ch <= chapter <= end_ch:
                        print(f"\n{row['Video']}")
                        # 初始化該章節的計數器
                        chapter_video_counter[current_chapter] = 1
                elif current_chapter is not None and start_ch <= current_chapter <= end_ch:
                    if str(video_name) != "合計" and not str(video_name).strip().isdigit():
                        # 使用該章節的計數器
                        video_num = chapter_video_counter.get(current_chapter, 1)
                        print(f"   {video_num}. {video_name} ({video_length}分)")
                        chapter_videos[current_chapter].append({
                            'name': video_name,
                            'length': video_length
                        })
                        # 該章節計數器 +1
                        chapter_video_counter[current_chapter] = video_num + 1
                        
            return start_ch, end_ch, f"第{start_ch}章到第{end_ch}章"
            
        except ValueError:
            print("❌ 請輸入有效的數字範圍，例如: 1-7 或輸入 0")

# 獲取用戶選擇
start_chapter, end_chapter, range_description = get_user_chapter_choice()

# 讀取資料
user_df = pd.read_excel("openedu_all_videos.xlsx")
sheetname = [n for n in score_xlsx.sheet_names if "總表" in n or "影片" in n][0]
video_df = score_xlsx.parse(sheetname)
video_df = video_df[["影片", "影片長度(分)"]].drop_duplicates()
video_df = video_df.rename(columns={"影片": "Video", "影片長度(分)": "VideoLength"})

# 根據用戶選擇過濾影片
video_df = filter_videos_by_chapters(video_df, start_chapter, end_chapter)

user_df["WatchCount"] = pd.to_numeric(user_df["觀看數"], errors="coerce").fillna(0)
user_df["WatchMinutes"] = pd.to_numeric(
    user_df["觀看時間(分)"], errors="coerce"
).fillna(0)
user_df["FinishRate"] = (
    user_df["完看率"].astype(str).str.replace("%", "").astype(float).fillna(0)
)
user_df = user_df.rename(
    columns={"帳戶名稱": "Account", "影片名稱": "Video", "觀看數": "StudentWatchCount"}
)

all_accounts = user_df["Account"].unique()
all_videos = video_df["Video"].unique()

view_count_denominator = len(video_df["Video"])  # 影觀公式的分母(實際影片數量)
view_percent_denominator = video_df[
    "VideoLength"
].sum()  # 影觀全公式的分母(實際影片總長度)

print(f"\n========== 成績計算資訊 ==========")
print(f"計算範圍: {range_description}")
print(f"影片數量: {view_count_denominator}")
print(f"影片總長度: {view_percent_denominator} 分鐘")

final_result = []

for account in all_accounts:
    rows = []
    udata = user_df[user_df["Account"] == account].set_index("Video")
    for _, vrow in video_df.iterrows():
        vname = vrow["Video"]
        vlen = vrow["VideoLength"]
        if vname in udata.index:
            row = udata.loc[vname]["FinishRate"]
            if isinstance(row, pd.Series):
                finish_rate = float(row.min())
            else:
                finish_rate = float(row)
        else:
            finish_rate = 0.0
        watched = vlen * finish_rate / 100
        rows.append(
            {
                "Video": vname,
                "VideoLength": vlen,
                "FinishRate": finish_rate,
                "WatchedRecord": watched,
            }
        )
    user_videos = pd.DataFrame(rows)

    # 影觀
    user_account_data = user_df[user_df["Account"] == account]
    # 過濾條件：影片名稱在篩選後的章節範圍內 且 學生觀看次數不為空(>0)
    filtered_videos = user_account_data[
        (user_account_data["Video"].isin(all_videos)) & 
        (user_account_data["StudentWatchCount"] > 0)
    ]
    view_count = len(filtered_videos) / view_count_denominator  

    # 影觀全
    view_percent = user_videos["WatchedRecord"].sum() / view_percent_denominator
    # 確保成績不超過100%，取計算結果和100%中的最小值
    view_count_percent = min(view_count * 100, 100)
    view_percent_full = min(view_percent * 100, 100)

    final_result.append(
        {
            "帳戶名稱": account,
            "影觀": f"{view_count_percent:.2f}%",
            "影觀全": f"{view_percent_full:.2f}%",
        }
    )

final_df = pd.DataFrame(final_result)
print(f"\n========== 成績結算結果 ==========")
print(final_df.head())

if not os.path.exists("./scores"):
    os.makedirs("./scores")

# 根據章節範圍生成檔案名稱
if start_chapter is None and end_chapter is None:
    filename = "openedu_all_videos_score_all.xlsx"
else:
    filename = f"openedu_videos_score_ch{start_chapter}_to_ch{end_chapter}.xlsx"

file_path = f"./scores/{filename}"
final_df.to_excel(file_path, index=False)

print(f"✅ 成績已儲存至: {file_path}")
print(f"✅ 計算範圍: {range_description}")
print(f"✅ 共處理 {len(final_df)} 個帳戶的成績")
