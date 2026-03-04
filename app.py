import streamlit as st
import pandas as pd
import re
import calendar
import io
from datetime import datetime, timedelta

# ==========================================
# 常數設定區
# ==========================================
MEAL_CHECKPOINTS = {'早': 7, '中': 12, '晚': 17}
TARGET_SHEETS = [f"{i}組" for i in range(1, 10)] + ["日照"]

# ==========================================
# 核心邏輯函式 (保持不變)
# ==========================================
def parse_minguo_datetime(dt_str):
    if pd.isna(dt_str): return None
    try:
        nums = re.findall(r'\d+', str(dt_str))
        if len(nums) >= 5:
            year = int(nums[0]) + 1911
            return datetime(year, int(nums[1]), int(nums[2]), int(nums[3]), int(nums[4]))
    except: pass
    return None

def parse_duration_to_days(duration_str):
    if pd.isna(duration_str): return 0.0
    days = 0.0
    day_match = re.search(r'(\d+)日', str(duration_str))
    hour_match = re.search(r'(\d+)時', str(duration_str))
    if day_match: days += float(day_match.group(1))
    if hour_match: days += float(hour_match.group(1)) / 8.0
    return days

def get_leave_lookup_table(leave_files, target_month, min_leave_days):
    leave_set = set()
    for f_file in leave_files:
        if f_file is None: continue
        try:
            try: df = pd.read_excel(f_file)
            except: 
                f_file.seek(0)
                df = pd.read_excel(f_file, engine='xlrd')
            for _, row in df.iterrows():
                if parse_duration_to_days(row.get('共計', '')) < min_leave_days: continue
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

def get_meal_price(meal_str: str, meal_prices: dict) -> int:
    if not meal_str: return 0
    meal_key = meal_str.strip()[0] 
    return meal_prices.get(meal_key, 0)

def process_comparison(meal_file, leave_lookup, target_month, meal_prices):
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
                                "餐別": raw_meal, "費用": get_meal_price(raw_meal, meal_prices),
                                "狀態": val, "異常說明": "請假期間仍有訂餐"
                            })
    return results

# ==========================================
# Streamlit UI 介面設計
# ==========================================
def main():
    st.set_page_config(page_title="點餐與差假比對系統", page_icon="🍱", layout="wide")
    
    # CSS 優化：橘色按鈕與區塊美化
    st.markdown("""
    <style>
    div.stButton > button:first-child {
        background-color: #FF9800 !important;
        border-color: #FF9800 !important;
        color: white !important;
        font-weight: bold !important;
        height: 3em !important;
        font-size: 1.2rem !important;
    }
    div.stButton > button:hover {
        background-color: #F57C00 !important;
        border-color: #F57C00 !important;
    }
    .instruction-box {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #FF9800;
        margin-bottom: 25px;
    }
    </style>
    """, unsafe_allow_html=True)

    # 標題
    st.title("🍱 點餐與差假交叉比對系統")

    # --- 1. 操作說明 (放在最顯眼處) ---
    st.markdown("""
    <div class="instruction-box">
        <h4 style='margin-top:0;'>📖 操作說明</h4>
        <ol>
            <li><b>左側設定</b>：確認<b>年月</b>（預設為上個月）與<b>餐點費用</b>。</li>
            <li><b>上傳檔案</b>：於右側分別上傳<b>點餐檔</b>與<b>差假記錄檔</b>。</li>
            <li><b>開始比對</b>：點擊下方<b>橘色按鈕</b>，比對結果將自動顯示並提供 Excel 下載。</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

    # 計算預設年月
    today = datetime.now()
    last_month_end = today.replace(day=1) - timedelta(days=1)
    default_year_minguo = last_month_end.year - 1911
    default_month = last_month_end.month

    # --- 2. 側邊欄：參數設定 ---
    with st.sidebar:
        st.header("⚙️ 第一步：參數設定")
        target_year_minguo = st.number_input("目標民國年份", min_value=100, max_value=200, value=default_year_minguo)
        target_month = st.number_input("目標月份", min_value=1, max_value=12, value=default_month)
        min_leave_days = st.number_input("異常門檻 (請假大於幾天)", min_value=0.0, max_value=30.0, value=1.0, step=0.5)
        
        st.divider()
        st.header("💰 餐點費用 (元)")
        p_b = st.number_input("早餐費用", min_value=0, value=40)
        p_l = st.number_input("中餐費用", min_value=0, value=75)
        p_d = st.number_input("晚餐費用", min_value=0, value=65)
        meal_prices = {'早': p_b, '中': p_l, '晚': p_d}
        
        ad_year = target_year_minguo + 1911
        st.info(f"📅 設定範圍：{ad_year}年{target_month}月")

    # --- 3. 主畫面：檔案上傳 ---
    st.subheader("第二步：上傳檔案")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 📥 點餐系統檔案")
        meal_file = st.file_uploader("請上傳點餐 Excel (.xlsx, .xlsm)", type=['xlsx', 'xlsm'], key="meal")

    with col2:
        st.markdown("##### 📥 差假紀錄檔案 (可選多份)")
        leave_h1 = st.file_uploader("上傳上半月/下半月差假檔 (.xls, .xlsx)", type=['xls', 'xlsx'], key="h1")
        leave_h2 = st.file_uploader("若有第二份差假檔請上傳", type=['xls', 'xlsx'], key="h2")

    # --- 4. 執行按鈕 ---
    st.divider()
    if st.button("🚀 開始執行比對", type="primary", use_container_width=True):
        if not meal_file:
            st.error("⚠️ 請先上傳點餐系統檔案。")
            return
        if not leave_h1 and not leave_h2:
            st.error("⚠️ 請至少上傳一份差假紀錄檔案。")
            return

        with st.spinner("🔍 正在進行資料比對，請稍候..."):
            leave_lookup = get_leave_lookup_table([leave_h1, leave_h2], target_month, min_leave_days)
            mismatch_data = process_comparison(meal_file, leave_lookup, target_month, meal_prices)
            
            if mismatch_data:
                df_final = pd.DataFrame(mismatch_data)
                
                # 排序優化
                meal_order = {'早': 1, '中': 2, '晚': 3}
                df_final['rank'] = df_final['餐別'].str[0].map(lambda x: meal_order.get(x, 9))
                df_final['日期_rank'] = df_final['日期'].map(lambda x: int(re.search(r'月(\d+)日', x).group(1)) if re.search(r'月(\d+)日', x) else 0)
                df_final = df_final.sort_values(by=['組別', '日期_rank', '姓名', 'rank']).drop(columns=['rank', '日期_rank'])
                
                st.success(f"🎉 比對完成！共發現 {len(df_final)} 筆異常。")
                st.metric("異常總金額", f"{df_final['費用'].sum()} 元")
                st.dataframe(df_final, use_container_width=True)
                
                # 下載按鈕
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                st.download_button(
                    label="📥 下載比對結果 Excel",
                    data=buf.getvalue(),
                    file_name=f"點餐異常比對_{target_year_minguo}{target_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.success("✨ 恭喜！未偵測到任何異常訂餐資料。")
                st.balloons()

if __name__ == "__main__":
    main()
