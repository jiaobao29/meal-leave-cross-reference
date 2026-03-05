import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime, timedelta

# ==========================================
# 常數設定區
# ==========================================
MEAL_CHECKPOINTS = {'早': 7, '中': 12, '晚': 17}
TARGET_SHEETS = [f"{i}組" for i in range(1, 10)] + ["日照"]

# ==========================================
# 工具函式
# ==========================================
def parse_minguo_datetime(dt_str):
    """解析民國年字串 轉為 datetime"""
    if pd.isna(dt_str): return None
    try:
        nums = re.findall(r'\d+', str(dt_str))
        if len(nums) < 5:
            st.warning(f"⚠️ 日期欄位格式不符 (需含年/月/日/時/分，實際解析到 {len(nums)} 個數字): '{dt_str}'，已略過此筆。")
            st.session_state['has_warning'] = True
            return None
            
        if len(nums) >= 5:
            year = int(nums[0]) + 1911
            return datetime(year, int(nums[1]), int(nums[2]), int(nums[3]), int(nums[4]))
    except Exception as e:
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

def get_meal_price(meal_str: str, meal_prices: dict) -> int:
    """依餐別字串首字取得對應費用"""
    if not meal_str: return 0
    meal_key = meal_str.strip()[0] 
    return meal_prices.get(meal_key, 0)

# ==========================================
# 核心業務邏輯
# ==========================================
def get_leave_lookup_table(leave_files, target_year, target_month, min_leave_days):
    """解析請假資料並建立快速查詢表"""
    leave_set = set()
    parsed_leave_count = 0
    found_months = set()
    
    for f_file in leave_files:
        if f_file is None: continue
            
        try:
            try: 
                df = pd.read_excel(f_file)
            except Exception as e_default: 
                f_file.seek(0) 
                try:
                    df = pd.read_excel(f_file, engine='xlrd')
                except Exception as e_xlrd:
                    raise ValueError(f"無法解析 Excel ({f_file.name})\n預設: {e_default}\nxlrd: {e_xlrd}")

            LEAVE_REQUIRED_COLS = {"姓名", "共計", "差假開始日期", "差假結束日期"}
            missing_cols = LEAVE_REQUIRED_COLS - set(df.columns)
            if missing_cols:
                st.warning(f"⚠️ 差假檔 '{f_file.name}' 缺少必要欄位: {missing_cols}，已略過此檔案。")
                st.session_state['has_warning'] = True
                continue

            for _, row in df.iterrows():
                if parse_duration_to_days(row.get('共計', '')) < min_leave_days: continue
                
                name = str(row.get('姓名', '')).strip()
                if not name or name.lower() == 'nan': continue
                    
                start_dt = parse_minguo_datetime(row.get('差假開始日期', ''))
                end_dt = parse_minguo_datetime(row.get('差假結束日期', ''))
                if not start_dt or not end_dt: continue

                curr_date = start_dt.date()
                while curr_date <= end_dt.date():
                    found_months.add(curr_date.month)
                    if curr_date.year == target_year and curr_date.month == target_month:
                        for m_name, hour in MEAL_CHECKPOINTS.items():
                            check_time = datetime(curr_date.year, curr_date.month, curr_date.day, hour, 0)
                            if start_dt <= check_time < end_dt:
                                leave_set.add((name, curr_date.day, m_name))
                                parsed_leave_count += 1
                    curr_date += timedelta(days=1)
        except Exception as e:
            st.warning(f"⚠️ 讀取或處理請假檔失敗 ({f_file.name}): {e}")
            st.session_state['has_warning'] = True
            
    return leave_set, parsed_leave_count, found_months

def process_comparison(meal_file, leave_lookup, target_month, meal_prices):
    """比對點餐表並產出異常清單"""
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
            df.columns = [str(c).strip() for c in df.columns]

            if {"姓名", "餐別"} - set(df.columns):
                st.warning(f"⚠️ 工作表 '{sheet}' 缺少必要欄位，已略過該表。")
                st.session_state['has_warning'] = True
                continue

            day_cols_found =[str(d) for d in range(1, 32) if str(d) in df.columns]
            if not day_cols_found:
                st.warning(f"⚠️ 工作表 '{sheet}' 找不到日期欄(1–31)，已略過該表。")
                st.session_state['has_warning'] = True
                continue
                
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
                        val = unicodedata.normalize('NFKC', str(row[day_col])).upper().strip()
                        if val in ['V', 'N']:
                            metrics['checked_meal_entries'] += 1
                            if (name, day, meal_key) in leave_lookup:
                                results.append({
                                    "組別": sheet,
                                    "姓名": name,
                                    "日期": f"{target_month}月{day}日",
                                    "餐別": raw_meal,
                                    "費用": get_meal_price(raw_meal, meal_prices),
                                    "狀態": val,
                                    "異常說明": "請假期間仍有訂餐"
                                })
        except Exception as e:
            st.warning(f"⚠️ 處理工作表 '{sheet}' 時發生錯誤: {e}")
            st.session_state['has_warning'] = True

    return results, metrics
