import streamlit as st
import pandas as pd
import re
import calendar
import io
from datetime import datetime, timedelta

# ==========================================
# 核心邏輯函式
# ==========================================
def parse_minguo_datetime(dt_str):
    """解析民國年字串 轉為 datetime"""
    if pd.isna(dt_str): return None
    try:
        nums = re.findall(r'\d+', str(dt_str))
        if len(nums) >= 5:
            year = int(nums[0]) + 1911
            return datetime(year, int(nums[1]), int(nums[2]), int(nums[3]), int(nums[4]))
    except: pass
    return None

def parse_duration_to_days(duration_str):
    """解析 '1日4時' 轉為天數"""
    if pd.isna(duration_str): return 0.0
    days = 0.0
    day_match = re.search(r'(\d+)日', str(duration_str))
    hour_match = re.search(r'(\d+)時', str(duration_str))
    if day_match: days += float(day_match.group(1))
    if hour_match: days += float(hour_match.group(1)) / 8.0
    return days

def get_leave_lookup_table(leave_files, target_month, min_leave_days):
    """解析請假資料"""
    MEAL_CHECKPOINTS = {'早': 7, '中': 12, '晚': 17}
    leave_set = set()
    
    for f_file in leave_files:
        if f_file is None: continue
        try:
            try: 
                df = pd.read_excel(f_file)
            except: 
                f_file.seek(0) 
                df = pd.read_excel(f_file, engine='xlrd')

            for _, row in df.iterrows():
                if parse_duration_to_days(row.get('共計', '')) < min_leave_days: 
                    continue
                
                name = str(row.get('姓名', '')).strip()
                start_dt = parse_minguo_datetime(row.get('差假開始日期', ''))
                end_dt = parse_minguo_datetime(row.get('差假結束日期', ''))
                
                if not start_dt or not end_dt: continue

                curr_date = start_dt.date()
                while curr_date <= end_dt.date():
                    if curr_date.month == target_month:
                        for m_name, hour in MEAL_CHECKPOINTS.items():
                            check_time = datetime(curr_date.year, curr_date.month, curr_date.day, hour, 0)
                            if start_dt <= check_time < end_dt:
                                leave_set.add((name, curr_date.day, m_name))
                    curr_date += timedelta(days=1)
        except Exception as e:
            st.warning(f"⚠️ 讀取請假檔失敗 ({f_file.name}): {e}")
    return leave_set

def process_comparison(meal_file, leave_lookup, target_month, meal_prices):
    """比對點餐表"""
    TARGET_SHEETS = [f"{i}組" for i in range(1, 10)] + ["日照"]
    if not meal_file: return []
    results = []
    excel = pd.ExcelFile(meal_file)

    for sheet in TARGET_SHEETS:
        if sheet not in excel.sheet_names: continue
        df = pd.read_excel(meal_file, sheet_name=sheet, header=2)
        df['姓名'] = df['姓名'].replace('nan', None).ffill()
        df.columns = [str(c).strip() for c in df.columns]

        for _, row in df.iterrows():
            name = str(row.get('姓名', '')).strip()
            raw_meal = str(row.get('餐別', '')).strip()
            if not name or name == 'nan' or not raw_meal: continue
            
            meal_key = raw_meal[0]
            for day in range(1, 32):
                day_col = str(day)
                if day_col in df.columns:
                    val = str(row[day_col]).upper().strip()
                    if val in ['V', 'N']:
                        if (name, day, meal_key) in leave_lookup:
                            results.append({
                                "組別": sheet, "姓名": name, "日期": f"{target_month}月{day}日",
                                "餐別": raw_meal, "費用": meal_prices.get(meal_key, 0),
                                "狀態": val, "異常說明": "請假期間仍有訂餐"
                            })
    return results

# ==========================================
# Streamlit UI 介面設計
# ==========================================
def main():
    st.set_page_config(page_title="點餐與差假比對系統", page_icon="🍱", layout="wide")
    
    # --- 注入 Custom CSS ---
    st.markdown("""
    <style>
    button[kind="primary"] {
        background-color: #FF9800 !important;
        border-color: #FF9800 !important;
        color: white !important;
        font-weight: bold !important;
        height: 3em !important;
        font-size: 1.2rem !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("🍱 點餐與差假交叉比對系統")

    # --- 操作說明 (精簡版) ---
    st.info("""**📖 操作說明**
1. **左側設定**：確認年月（預設為上個月）與餐點費用。
2. **上傳檔案**：於右側分別上傳點餐檔與差假記錄檔。
3. **開始比對**：點擊下方橘色按鈕，比對結果將自動顯示並提供下載。""")

    # --- 參數計算 ---
    today = datetime.now()
    last_month_end = today.replace(day=1) - timedelta(days=1)
    
    # --- 側邊欄：左側設定 ---
    with st.sidebar:
        st.header("⚙️ 1. 左側設定")
        target_year_minguo = st.number_input("目標民國年份", 100, 200, last_month_end.year-1911)
        target_month = st.number_input("目標月份", 1, 12, last_month_end.month)
        min_leave_days = st.number_input("異常門檻 (請假大於幾天)", 0.0, 30.0, 1.0, 0.5)
        
        st.divider()
        st.subheader("💰 餐點費用設定")
        p_b = st.number_input("早餐費用", 0, 200, 40)
        p_l = st.number_input("中餐費用", 0, 200, 75)
        p_d = st.number_input("晚餐費用", 0, 200, 65)
        meal_prices = {'早': p_b, '中': p_l, '晚': p_d}

    # --- 主畫面：右側上傳 ---
    st.subheader("📁 2. 上傳檔案")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**點餐系統檔案 (.xlsx)**")
        meal_file = st.file_uploader("上傳點餐檔", type=['xlsx', 'xlsm'], label_visibility="collapsed")

    with col2:
        st.markdown("**差假紀錄檔案 (可多選或分開上傳)**")
        leave_h1 = st.file_uploader("上傳差假檔1", type=['xls', 'xlsx'], label_visibility="collapsed")
        leave_h2 = st.file_uploader("上傳差假檔2", type=['xls', 'xlsx'], label_visibility="collapsed")

    # --- 執行按鈕 ---
    st.divider()
    if st.button("🚀 3. 開始交叉比對", type="primary", use_container_width=True):
        if not meal_file or (not leave_h1 and not leave_h2):
            st.error("❌ 請確認已上傳點餐檔與至少一份差假紀錄檔！")
            return

        with st.spinner("系統比對中..."):
            leave_lookup = get_leave_lookup_table([leave_h1, leave_h2], target_month, min_leave_days)
            mismatch = process_comparison(meal_file, leave_lookup, target_month, meal_prices)
            
            if mismatch:
                df = pd.DataFrame(mismatch)
                # 排序
                df['d_rank'] = df['日期'].map(lambda x: int(re.search(r'月(\d+)日', x).group(1)))
                df = df.sort_values(['組別', 'd_rank', '姓名']).drop(columns=['d_rank'])
                
                st.success(f"✅ 比對完成！偵測到 {len(df)} 筆異常。")
                st.warning(f"💰 異常餐點總計：{df['費用'].sum()} 元")
                st.dataframe(df, use_container_width=True)
                
                # 下載按鈕
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                st.download_button("📥 下載比對結果 Excel", buf.getvalue(), 
                                 f"比對結果_{target_year_minguo}年{target_month}月.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 type="primary")
            else:
                st.success("✨ 未發現任何異常資料！")

if __name__ == "__main__":
    main()

