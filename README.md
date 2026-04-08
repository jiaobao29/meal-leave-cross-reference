```
# 🍱 人事室專用：點餐與差假交叉比對系統 (Meal-Leave Audit Tool)

> **🤖 AI Developer Context Note:**
> 本文件旨在提供開發者與 AI 助手最精確的系統脈絡。本專案為一個無資料庫的靜態分析工具，依賴 Streamlit 作為 UI、Pandas 處理記憶體內運算。
> **核心業務邏輯**：比對「差假區間」是否覆蓋「餐點檢核點 (早7/中12/晚17)」，若覆蓋且員工該餐有打勾 (V/N)，則視為溢領異常。

---

## 📂 系統架構 (Architecture & Boundaries)

系統採 MVC 解耦設計，除錯時請依據以下邊界定位問題：

| 模組檔案 | 職責定義 (Responsibility) | 潛在除錯方向 (Debugging Vectors) |
| :--- | :--- | :--- |
| **`app.py`** | **View / Controller**<br>處理 Streamlit UI、檔案接收大小限制 (100MB)、報表產出、**最終 DataFrame 的排序邏輯**。 | UI 顯示異常、Excel 下載失敗、最終結果排序錯誤 (如數字與文字部門混排問題)。 |
| **`logic.py`** | **Service / Business Logic**<br>處理民國年解析、請假時間集 (Set) 建立、動態遍歷 Excel Sheets、核心交叉比對演算法。 | 日期解析報錯 (Regex)、比對漏抓/多抓、工作表被錯誤略過 (Skipped)。 |
| **`schema.py`** | **Model / Validation**<br>定義來源 Excel 的標題列常數、提供欄位驗證 (Validation) 函式。 | 來源 Excel 欄位名稱變更導致抓不到資料時，**只需改此檔**。 |

---

## 🔄 資料流與重要假設 (Data Flow & Assumptions)

在進行任何維護或重構前，請確保不破壞以下假設：

### 1. 點餐檔案 (Meal Records)
*   **動態偵測機制 (Dynamic Sheet Detection)**：程式**不再寫死**目標 Sheet 名稱。只要工作表符合 `schema.py` 定義（包含「姓名」、「餐別」及「1~31」日期數字欄位），就會自動排入處理。不符合的 (例如「注意事項頁」) 會被靜默略過並記錄。
*   **標題列位置**：點餐系統的標題列固定在第三行，因此讀取時強制使用 `header=2` (跳過前兩列)。
*   **有效標記**：單元格內不分大小寫、不管全半角，只要是 `V` 或 `N` 就視為有訂餐。

### 2. 差假檔案 (Leave Records)
*   **日期格式**：來源為民國年字串 (如 `112/01/01 08:00`)，透過 `logic.py` 中的正則表達式萃取數字，並加上 `1911` 轉換為 Python `datetime`。
*   **時間交疊邏輯**：
    *   早餐檢核點：`07:00`
    *   午餐檢核點：`12:00`
    *   晚餐檢核點：`17:00`
    *   *規則*：`差假開始時間 <= 檢核點時間 < 差假結束時間` 且經過「請假天數門檻」篩選，才算成功請假。

---

## 🛠️ 開發與維護指南 (Maintenance Guide)

### 常見情境 1：使用者說「XX室的點餐紀錄沒有被比對出來！」
1. 檢查該 Sheet 的第三行 (Excel Row 3) 是否確實包含 `姓名`、`餐別`。
2. 檢查是否包含 `1` 到 `31` 的數字標題。
3. 如果使用者改了欄位名稱（例如變成 `員工姓名`），請到 `schema.py` 修改 `COL_MEAL_NAME = "員工姓名"`。

### 常見情境 2：日期解析出現 Warning 報錯
*   這通常是因為人事室匯出的 Excel 中，日期欄位含有奇怪格式或為空。
*   檢查 `logic.py` 中的 `parse_minguo_datetime` 函式，確認正則表達式 `re.findall(r'\d+', str(dt_str))` 是否需要增加例外處理。

### 常見情境 3：輸出的 Excel 排序亂掉
*   排序邏輯位於 `app.py` 的 `df_final.sort_values` 區塊。
*   目前的 `組別_rank` 邏輯能區分「帶數字部門 (1組)」與「純文字部門 (行政室)」。若未來新增特殊命名部門，請在此處微調 Regex 權重。

---

## 🚀 部署與啟動

```bash
# 1. 安裝環境依賴 (需支援 openpyxl 與 xlrd 以處理新舊版 Excel)
pip install -r requirements.txt

# 2. 啟動 Streamlit 服務
streamlit run app.py
```

---

## 📝 版本演進 (Changelog)
*   **v1.0**: 初始腳本版本。
*   **v1.1**: 導入 MVC 架構，分離 `schema.py`，解決欄位寫死導致容易崩潰的問題。
*   **v1.2 (Current)**: 實作 **動態工作表偵測 (Dynamic Sheet Detection)**，移除硬編碼的 `TARGET_SHEETS`，支援任意名稱之部門 (如行政室、人事室)，並優化混合字串排序與靜默過濾防呆機制。
```
