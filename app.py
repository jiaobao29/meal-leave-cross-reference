import streamlit as st
import pandas as pd
import calendar
import io
import re
from datetime import datetime, timedelta

# 引入核心邏輯模組
from logic import get_leave_lookup_table, process_comparison

def main():
    st.set_page_config(page_title="點餐與差假比對系統", page_icon="🍱", layout="wide")
    
    # --- 自動計算「上個月」的年月值 ---
    today = datetime.now()
    first_day_current = today.replace(day=1)
    last_day_last = first_day_current - timedelta(days=1)
    
    default_year_minguo = last_day_last.year - 1911
    default_month = last_day_last.month
    
    # --- 注入 Custom CSS ---
    st.markdown("""
    <style>
    button[kind="primary"] { background-color: #FF9800 !important; border-color: #FF9800 !important; color: white !important; font-weight: bold !important; }
    button[kind="primary"]:hover { background-color: #F57C00 !important; border-color: #F57C00 !important; }
    .instruction-box { background-color: #f8f9fa; padding: 12px 18px; border-radius: 8px; border-left: 5px solid #FF9800; margin-bottom: 20px; font-size: 0.95rem; color: #444; line-height: 1.6; }
    </style>
    """, unsafe_allow_html=True)

    st.title("🍱 點餐與差假交叉比對工具")
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
        target_year_minguo = st.number_input("目標民國年份", min_value=100, max_value=200, value=default_year_minguo, step=1)
        target_month = st.number_input("目標月份", min_value=1, max_value=12, value=default_month, step=1)
        min_leave_days = st.number_input("異常門檻 (請假大於幾天)", min_value=0.0, max_value=30.0, value=1.0, step=0.5)
        
        ad_year = target_year_minguo + 1911
        _, days_in_month = calendar.monthrange(ad_year, target_month)
        st.info(f"📅 目前設定：西元 {ad_year} 年 {target_month} 月\n\n當月天數：{days_in_month} 天")
        st.divider()
        
        st.header("💰 餐點費用設定 (元)")
        meal_prices = {
            '早': st.number_input("早 餐 費 用", min_value=0, value=40, step=5),
            '中': st.number_input("中 餐 費 用", min_value=0, value=75, step=5),
            '晚': st.number_input("晚 餐 費 用", min_value=0, value=65, step=5)
        }

    # --- 主畫面：檔案上傳區 ---
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ 上傳點餐系統檔案")
        meal_file = st.file_uploader("📁 請上傳點餐系統 Excel (.xlsx, .xlsm)", type=['xlsx', 'xlsm'], key="meal")
    with col2:
        st.subheader("2️⃣ 上傳差假紀錄檔案")
        leave_h1 = st.file_uploader("📁 請上傳上半月差假紀錄 (.xls, .xlsx)", type=['xls', 'xlsx'], key="h1")
        leave_h2 = st.file_uploader("📁 請上傳下半月差假紀錄 (.xls, .xlsx)", type=['xls', 'xlsx'], key="h2")

    st.divider()

    # --- 執行按鈕與結果處理 ---
    if st.button("🚀 開始交叉比對", type="primary", use_container_width=True):
        st.session_state['has_warning'] = False
        
        if not meal_file:
            st.error("❌ 請務必上傳「點餐系統」檔案！")
            return
        if not leave_h1 and not leave_h2:
            st.error("❌ 請至少上傳一份「差假紀錄」檔案 (上半月或下半月)！")
            return

        MAX_FILE_SIZE_MB = 100
        for label, f in[("點餐檔", meal_file), ("上半月差假", leave_h1), ("下半月差假", leave_h2)]:
            if f and f.size > MAX_FILE_SIZE_MB * 1024 * 1024:
                st.error(f"❌ 檔案「{label}」超過 {MAX_FILE_SIZE_MB}MB 上限，請重新上傳。")
                return

        # 執行掃描 (呼叫 logic.py 中的函式)
        with st.status("🔍 開始執行交叉比對作業...", expanded=True) as status:
            st.write("📂 正在讀取並解析差假資料...")
            leave_lookup, leave_count, found_months = get_leave_lookup_table([leave_h1, leave_h2], ad_year, target_month, min_leave_days
            )
            
            if found_months and target_month not in found_months:
                st.warning(f"🚨 **月份不一致警告**：您設定要比對 {target_month} 月，但差假檔案內實際包含的月份為 {sorted(list(found_months))} 月！")
                st.session_state['has_warning'] = True
                
            st.write(f"✅ 成功載入 {leave_count} 筆符合 {target_month} 月的差假紀錄")
            st.write("🍱 正在讀取並比對點餐資料...")
            
            mismatch_data, scan_metrics = process_comparison(
                meal_file, leave_lookup, target_month, meal_prices
            )
            st.write(f"✅ 共讀取 {scan_metrics['processed_sheets']} 個工作表，掃描 {scan_metrics['checked_meal_entries']} 筆點餐紀錄")
            status.update(label="比對作業完成！", state="complete", expanded=False)
            
        # 顯示結果與下載
        if mismatch_data:
            df_final = pd.DataFrame(mismatch_data)
            
            # DataFrame 排序格式化
            meal_order = {'早': 1, '中': 2, '晚': 3}
            df_final['rank'] = df_final['餐別'].str[0].map(lambda x: meal_order.get(x, 9))
            df_final['組別_rank'] = df_final['組別'].map(lambda x: (0, int(re.search(r'\d+', x).group())) if re.search(r'\d+', x) else (1, 0))
            df_final['日期_rank'] = df_final['日期'].map(lambda x: int(re.search(r'月(\d+)日', x).group(1)) if re.search(r'月(\d+)日', x) else 0)
            df_final = df_final.sort_values(by=['組別_rank', '日期_rank', '姓名', 'rank']).drop(columns=['rank', '組別_rank', '日期_rank'])
            
            st.success(f"✅ 比對完成！偵測到 **{len(df_final)}** 筆異常。")
            st.warning(f"💰 異常餐點總計費用：**{df_final['費用'].sum()}** 元")
            st.dataframe(df_final, use_container_width=True)
            
            # 準備 Excel 下載
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
            if scan_metrics['checked_meal_entries'] == 0:
                st.error("⚠️ 掃描完成，但**未能找到任何有效的點餐紀錄 (V/N)**！請確認點餐檔案是否為空白。")
            elif leave_count == 0 and not (found_months and target_month not in found_months):
                st.warning(f"⚠️ 掃描完成，但**未載入任何 {target_month} 月的有效差假紀錄**。若無人請假請忽略此訊息。")
            elif scan_metrics['checked_meal_entries'] > 0 and leave_count > 0:
                st.success(f"✨ 掃描完成！本次共深入檢查了 **{scan_metrics['processed_sheets']}** 個工作表，未發現任何異常狀況！")
                st.balloons()
            elif st.session_state.get('has_warning', False):
                st.info("ℹ️ 掃描結束。因發生上述警告，部分資料未能完整比對。請修正錯誤後重新執行。")
            else:
                st.success("✨ 掃描完成，未發現任何異常資料！")
                st.balloons()

if __name__ == "__main__":
    main()
