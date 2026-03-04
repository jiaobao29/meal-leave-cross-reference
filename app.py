import streamlit as st
import pandas as pd
import re
import calendar
import io
from datetime import datetime, timedelta

# ==========================================
# 常數設定區 (移除 MEAL_PRICES，改為動態傳入)
# ==========================================
MEAL_CHECKPOINTS = {'早': 7, '中': 12, '晚': 17}
TARGET_SHEETS = [f"{i}組" for i in range(1, 10)] +["日照"]

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
    leave_set = set()
    
    for f_file in leave_files:
        if f_file is None: 
            continue
            
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
                
                if not start_dt or not end_dt: 
                    continue

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
    """依餐別字串首字取得對應費用 (動態接收費用設定)"""
    if not meal_str: return 0
    meal_key = meal_str.strip()[0] 
    return meal_prices.get(meal_key, 0)

def process_comparison(meal_file, leave_lookup, target_month, meal_prices):
    """比對點餐表"""
    if not meal_file: return []
    results =[]
    
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
                                "組別": sheet,
                                "姓名": name,
                                "日期": f"{target_month}月{day}日",
                                "餐別": raw_meal,
                                "費用": get_meal_price(raw_meal, meal_prices), # 動態費用
                                "狀態": val,
                                "異常說明": "請假期間仍有訂餐"
                            })
    return results

# ==========================================
# Streamlit UI 介面設計
# ==========================================
def main():
    st.set_page_config(page_title="點餐與差假比對系統", page_icon="🍱", layout="wide")
    
    # --- 自動計算「上個月」的年月值 ---
    today = datetime.now()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_last_month = first_day_of_current_month - timedelta(days=1)
    
    default_year_minguo = last_day_of_last_month.year - 1911
    default_month = last_day_of_last_month.month
    
    # --- 注入 Custom CSS ---
    st.markdown("""
    <style>
    button[kind="primary"] {
        background-color: #FF9800 !important;
        border-color: #FF9800 !important;
        color: white !important;
        font-weight: bold !important;
    }
    button[kind="primary"]:hover {
        background-color: #F57C00 !important;
        border-color: #F57C00 !important;
    }
    /* 調整操作說明文字的大小與間距 */
    .instruction-box {
        background-color: #f8f9fa;
        padding: 12px 18px;
        border-radius: 8px;
        border-left: 5px solid #FF9800;
        margin-bottom: 20px;
        font-size: 0.95rem;
        color: #444;
        line-height: 1.6;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("🍱 點餐與差假交叉比對工具")

    # --- 新增的操作說明區塊 ---
    st.markdown("""
    <div class="instruction-box">
        <b>💡 操作說明：</b><br>
        1. <b>左側設定</b>：請輸入<b>欲比對的資料月份</b>與餐點費用。<br>
        2. <b>上傳檔案</b>：於右側分別上傳「點餐檔」與「差假記錄檔」。<br>
        3. <b>開始比對</b>：點擊下方橘色按鈕，結果將自動顯示並提供 Excel 下載。
    </div>
    """, unsafe_allow_html=True)
    # --- 側邊欄：參數設定 ---
    with st.sidebar:
        st.header("⚙️ 參數設定")
        
        # 使用我們計算出來的 default_year_minguo 和 default_month
        target_year_minguo = st.number_input(
            "目標民國年份", min_value=100, max_value=200, 
            value=default_year_minguo, step=1
        )
        target_month = st.number_input(
            "目標月份", min_value=1, max_value=12, 
            value=default_month, step=1
        )
        min_leave_days = st.number_input("異常門檻 (請假大於幾天)", min_value=0.0, max_value=30.0, value=1.0, step=0.5)
        
        ad_year = target_year_minguo + 1911
        _, days_in_month = calendar.monthrange(ad_year, target_month)
        st.info(f"📅 目前設定：西元 {ad_year} 年 {target_month} 月\n\n當月天數：{days_in_month} 天")
        
        st.divider()
        
        st.header("💰 餐點費用設定 (元)")
        price_breakfast = st.number_input("早 餐 費 用", min_value=0, value=40, step=5)
        price_lunch = st.number_input("中 餐 費 用", min_value=0, value=75, step=5)
        price_dinner = st.number_input("晚 餐 費 用", min_value=0, value=65, step=5)
        
        # 封裝成字典供後續使用
        meal_prices = {'早': price_breakfast, '中': price_lunch, '晚': price_dinner}

    # --- 主畫面：檔案上傳區 ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ 上傳點餐系統檔案")
        # 利用 Markdown 的 h5 讓字體變大，並隱藏原本 uploader 的 label
        st.markdown("##### 📁 請上傳點餐系統 Excel (.xlsx, .xlsm)")
        meal_file = st.file_uploader("", type=['xlsx', 'xlsm'], key="meal", label_visibility="collapsed")

    with col2:
        st.subheader("2️⃣ 上傳差假紀錄檔案")
        st.markdown("##### 📁 請上傳上半月差假紀錄 (.xls, .xlsx)")
        leave_h1 = st.file_uploader("", type=['xls', 'xlsx'], key="h1", label_visibility="collapsed")
        
        st.markdown("##### 📁 請上傳下半月差假紀錄 (.xls, .xlsx)")
        leave_h2 = st.file_uploader("", type=['xls', 'xlsx'], key="h2", label_visibility="collapsed")

    # --- 執行按鈕與結果處理 ---
    st.divider()
    # 按鈕設定為 primary，CSS 樣式會自動將其變為橘色
    if st.button("🚀 開始交叉比對", type="primary", use_container_width=True):
        
        if not meal_file:
            st.error("❌ 請務必上傳「點餐系統」檔案！")
            return
        if not leave_h1 and not leave_h2:
            st.error("❌ 請至少上傳一份「差假紀錄」檔案 (上半月或下半月)！")
            return

        with st.spinner("系統比對中，請稍候..."):
            # 1. 取得請假名單表
            leave_lookup = get_leave_lookup_table([leave_h1, leave_h2], 
                target_month=target_month, 
                min_leave_days=min_leave_days
            )
            
            # 2. 比對點餐資料 (傳入動態設定的 meal_prices)
            mismatch_data = process_comparison(
                meal_file, 
                leave_lookup, 
                target_month=target_month,
                meal_prices=meal_prices
            )
            
            # 3. 產出結果
            if mismatch_data:
                df_final = pd.DataFrame(mismatch_data)
                
                # 排序邏輯
                meal_order = {'早': 1, '中': 2, '晚': 3}
                df_final['rank'] = df_final['餐別'].str[0].map(lambda x: meal_order.get(x, 9))
                df_final['組別_rank'] = df_final['組別'].map(
                    lambda x: (0, int(re.search(r'\d+', x).group())) if re.search(r'\d+', x) else (1, 0)
                )
                df_final['日期_rank'] = df_final['日期'].map(
                    lambda x: int(re.search(r'月(\d+)日', x).group(1)) if re.search(r'月(\d+)日', x) else 0
                )
                df_final = df_final.sort_values(by=['組別_rank', '日期_rank', '姓名', 'rank']).drop(columns=['rank', '組別_rank', '日期_rank'])
                
                total_cost = df_final['費用'].sum()
                
                st.success(f"✅ 比對完成！偵測到 **{len(df_final)}** 筆異常。")
                st.warning(f"💰 異常餐點總計費用：**{total_cost}** 元")
                
                # 在網頁顯示結果
                st.dataframe(df_final, use_container_width=True)
                
                # 製作 Excel 下載按鈕
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                # 此按鈕同樣套用 primary，會變成橘色
                st.download_button(
                    label="📥 下載比對結果 Excel",
                    data=output_buffer.getvalue(),
                    file_name=f"比對結果_{target_year_minguo}年{target_month}月.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.success("✨ 掃描完成，未發現任何異常資料！")
                st.balloons()

if __name__ == "__main__":
    main()



