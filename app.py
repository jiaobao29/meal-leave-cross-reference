import streamlit as st
import pandas as pd
import re
import calendar
import io
import unicodedata
from datetime import datetime, timedelta

# ==========================================
# 常數設定區 (移除 MEAL_PRICES，改為動態傳入)
# ==========================================
MEAL_CHECKPOINTS = {'早': 7, '中': 12, '晚': 17}
TARGET_SHEETS =[f"{i}組" for i in range(1, 10)] + ["日照"]

# ==========================================
# 核心邏輯函式
# ==========================================
def parse_minguo_datetime(dt_str):
    """解析民國年字串 轉為 datetime"""
    if pd.isna(dt_str): return None
    try:
        nums = re.findall(r'\d+', str(dt_str))
        
        # BLUEPRINT RULE 1: Structural Guard Before Regex
        if len(nums) < 5:
            st.warning(f"⚠️ 日期欄位格式不符 (需含年/月/日/時/分，實際解析到 {len(nums)} 個數字): '{dt_str}'，已略過此筆。")
            st.session_state['has_warning'] = True  # 標記發生警告
            return None
            
        if len(nums) >= 5:
            year = int(nums[0]) + 1911
            return datetime(year, int(nums[1]), int(nums[2]), int(nums[3]), int(nums[4]))
    except Exception as e:
        # Never silently ignore exceptions
        st.warning(f"⚠️ 解析民國日期字串發生錯誤 ({dt_str}): {e}")
        st.session_state['has_warning'] = True
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

def get_leave_lookup_table(leave_files, target_year, target_month, min_leave_days):
    """解析請假資料"""
    leave_set = set()
    # BLUEPRINT V2 STEP 1: Track Metrics & Detect Months
    parsed_leave_count = 0
    found_months = set()
    
    for f_file in leave_files:
        if f_file is None: 
            continue
            
        try:
            try: 
                df = pd.read_excel(f_file)
            except Exception as e_default: 
                f_file.seek(0) 
                try:
                    df = pd.read_excel(f_file, engine='xlrd')
                except Exception as e_xlrd:
                    raise ValueError(f"無法解析 Excel 檔案 ({f_file.name})\n預設引擎錯誤: {e_default}\nxlrd引擎錯誤: {e_xlrd}")

            # BLUEPRINT RULE 2: Required Column Guard (Fail-Fast on Schema)
            LEAVE_REQUIRED_COLS = {"姓名", "共計", "差假開始日期", "差假結束日期"}
            missing_cols = LEAVE_REQUIRED_COLS - set(df.columns)
            if missing_cols:
                st.warning(f"⚠️ 差假檔 '{f_file.name}' 缺少必要欄位: {missing_cols}，已略過此檔案。")
                st.session_state['has_warning'] = True  # 標記發生警告
                continue  # skip this file entirely — fail fast at the file level

            for _, row in df.iterrows():
                if parse_duration_to_days(row.get('共計', '')) < min_leave_days: 
                    continue
                
                name = str(row.get('姓名', '')).strip()
                
                # BLUEPRINT RULE 3: Value-Level Sentinel Check on Name
                if not name or name.lower() == 'nan':
                    continue
                    
                start_dt = parse_minguo_datetime(row.get('差假開始日期', ''))
                end_dt = parse_minguo_datetime(row.get('差假結束日期', ''))
                
                if not start_dt or not end_dt: 
                    continue

                curr_date = start_dt.date()
                while curr_date <= end_dt.date():
                    # Blueprint V2: Collect all unique months present in the file
                    found_months.add(curr_date.month)
                    
                    if curr_date.year == target_year and curr_date.month == target_month:
                        for m_name, hour in MEAL_CHECKPOINTS.items():
                            check_time = datetime(curr_date.year, curr_date.month, curr_date.day, hour, 0)
                            if start_dt <= check_time < end_dt:
                                leave_set.add((name, curr_date.day, m_name))
                                parsed_leave_count += 1 # Blueprint V2: Tracking count
                    curr_date += timedelta(days=1)
        except Exception as e:
            st.warning(f"⚠️ 讀取或處理請假檔失敗 ({f_file.name}): {e}")
            st.session_state['has_warning'] = True
            
    # Blueprint V2: Return changes
    return leave_set, parsed_leave_count, found_months

def get_meal_price(meal_str: str, meal_prices: dict) -> int:
    """依餐別字串首字取得對應費用 (動態接收費用設定)"""
    if not meal_str: return 0
    meal_key = meal_str.strip()[0] 
    return meal_prices.get(meal_key, 0)

def process_comparison(meal_file, leave_lookup, target_month, meal_prices):
    """比對點餐表"""
    # BLUEPRINT V2 STEP 2: Add Tracking Metrics
    metrics = {'processed_sheets': 0, 'checked_meal_entries': 0}
    if not meal_file: return [], metrics
    
    results =[]
    
    try:
        excel = pd.ExcelFile(meal_file)
    except Exception as e:
        st.error(f"⚠️ 無法讀取點餐系統檔案: {e}")
        st.session_state['has_warning'] = True
        return[], metrics

    for sheet in TARGET_SHEETS:
        if sheet not in excel.sheet_names: continue
        
        try:
            df = pd.read_excel(meal_file, sheet_name=sheet, header=2)
            df.columns =[str(c).strip() for c in df.columns]

            REQUIRED_COLUMNS = {"姓名", "餐別"}
            missing = REQUIRED_COLUMNS - set(df.columns)
            if missing:
                st.warning(f"⚠️ 工作表 '{sheet}' 缺少必要欄位: {missing}，已自動略過該表。")
                st.session_state['has_warning'] = True
                continue

            # BLUEPRINT RULE 4: Day-Column Integer Validation
            day_cols_found =[str(d) for d in range(1, 32) if str(d) in df.columns]
            if not day_cols_found:
                st.warning(f"⚠️ 工作表 '{sheet}' 中找不到任何日期欄 (1–31)，請確認欄位標題格式，已略過該表。")
                st.session_state['has_warning'] = True
                continue
                
            # Blueprint V2: Valid sheet counted
            metrics['processed_sheets'] += 1

            df['姓名'] = df['姓名'].replace('nan', None).ffill()

            for _, row in df.iterrows():
                name = str(row.get('姓名', '')).strip()
                raw_meal = str(row.get('餐別', '')).strip()
                if not name or name == 'nan' or not raw_meal: continue
                
                meal_key = raw_meal[0]
                for day in range(1, 32):
                    day_col = str(day)
                    if day_col in df.columns:
                        val = str(row[day_col]).upper().strip()
                        
                        # BLUEPRINT RULE 5: Cell Value Normalisation Before Comparison
                        val = unicodedata.normalize('NFKC', val).upper().strip()
                        
                        if val in ['V', 'N']:
                            # Blueprint V2: Checked meal entries counted
                            metrics['checked_meal_entries'] += 1
                            
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
        except Exception as e:
            st.warning(f"⚠️ 處理工作表 '{sheet}' 時發生錯誤: {e}")
            st.session_state['has_warning'] = True

    return results, metrics

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
    if st.button("🚀 開始交叉比對", type="primary", use_container_width=True):
        
        st.session_state['has_warning'] = False
        
        if not meal_file:
            st.error("❌ 請務必上傳「點餐系統」檔案！")
            return
        if not leave_h1 and not leave_h2:
            st.error("❌ 請至少上傳一份「差假紀錄」檔案 (上半月或下半月)！")
            return

        # BLUEPRINT RULE 6: File Size Guard Before Processing
        MAX_FILE_SIZE_MB = 100
        for label, f in[("點餐檔", meal_file), ("上半月差假", leave_h1), ("下半月差假", leave_h2)]:
            if f and f.size > MAX_FILE_SIZE_MB * 1024 * 1024:
                st.error(f"❌ 檔案「{label}」({f.name}) 超過 {MAX_FILE_SIZE_MB}MB 上限，請確認後重新上傳。")
                return

        # BLUEPRINT V2 STEP 3: Replace st.spinner with st.status & Add Month-Mismatch Guard
        with st.status("🔍 開始執行交叉比對作業...", expanded=True) as status:
            st.write("📂 正在讀取並解析差假資料...")
            # 1. 取得請假名單表
            leave_lookup, leave_count, found_months = get_leave_lookup_table(
                [leave_h1, leave_h2], 
                target_year=ad_year,
                target_month=target_month, 
                min_leave_days=min_leave_days
            )
            
            # Crucial Guard: 檢查設定月份與檔案中包含的月份是否相符
            if found_months and target_month not in found_months:
                st.warning(f"🚨 **月份不一致警告**：您設定要比對 {target_month} 月，但差假檔案內實際包含的月份為 {sorted(list(found_months))} 月！這可能導致比對結果不準確。")
                st.session_state['has_warning'] = True
                
            st.write(f"✅ 成功載入 {leave_count} 筆符合 {target_month} 月的差假紀錄")
            
            st.write("🍱 正在讀取並比對點餐資料...")
            # 2. 比對點餐資料
            mismatch_data, scan_metrics = process_comparison(
                meal_file, 
                leave_lookup, 
                target_month=target_month,
                meal_prices=meal_prices
            )
            
            st.write(f"✅ 共讀取 {scan_metrics['processed_sheets']} 個工作表，掃描 {scan_metrics['checked_meal_entries']} 筆點餐紀錄")
            status.update(label="比對作業完成！", state="complete", expanded=False)
            
        # BLUEPRINT V2 STEP 4: Enhance the Final Output Logic
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
            
            st.dataframe(df_final, use_container_width=True)
            
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 下載比對結果 Excel",
                data=output_buffer.getvalue(),
                file_name=f"比對結果_{target_year_minguo}年{target_month}月.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            # Empty Data Guards & True Success Conditions
            if scan_metrics['checked_meal_entries'] == 0:
                st.error("⚠️ 掃描完成，但**未能找到任何有效的點餐紀錄 (V/N)**！請確認點餐檔案是否為空白。")
            elif leave_count == 0 and not (found_months and target_month not in found_months):
                # 只有在沒有觸發上方的強烈月份警告時，才提示此訊息
                st.warning(f"⚠️ 掃描完成，但**未載入任何 {target_month} 月的有效差假紀錄**。若該月確實無人請假請忽略此訊息，否則請檢查檔案。")
            elif scan_metrics['checked_meal_entries'] > 0 and leave_count > 0:
                st.success(f"✨ 掃描完成！本次共深入檢查了 **{scan_metrics['processed_sheets']}** 個工作表、**{scan_metrics['checked_meal_entries']}** 筆點餐紀錄，未發現任何異常狀況！")
                st.balloons()
            elif st.session_state.get('has_warning', False):
                st.info("ℹ️ 掃描結束。因發生上述警告（如檔案格式有誤或略過工作表），部分資料未能完整比對。請修正錯誤後重新執行。")
            else:
                st.success("✨ 掃描完成，未發現任何異常資料！")
                st.balloons()

if __name__ == "__main__":
    main()
