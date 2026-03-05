# meal-leave-cross-reference
將每個月的點餐系統 與 上半月、下半月差假紀錄進行交叉比對 得出 "請假同時又有訂餐的人"

***

# 🍱 Meal-Leave Cross-Reference Tool 
### (點餐與差假交叉比對系統)

A lightweight, high-accuracy auditing tool built with **Streamlit** and **Pandas**. This tool automates the tedious process of verifying meal orders against leave records, ensuring that no meal costs are incurred while staff are on leave.

## 🌟 Key Features
- **Smart Cross-Referencing**: Automatically detects if a meal (Breakfast, Lunch, Dinner) was marked as ordered ("V" or "N") during a staff member's leave period.
- **Minguo Calendar Support**: Native parsing of Taiwan Minguo dates (e.g., `112/10/05 08:00`).
- **Month-Mismatch Guard**: Built-in safety check to ensure the uploaded leave records actually match the target month selected.
- **Dynamic Pricing**: Custom meal cost settings (Breakfast/Lunch/Dinner) with real-time total anomaly cost calculation.
- **Detailed Audit Metrics**: Provides transparency on how many sheets were scanned and how many entries were verified.
- **Exportable Results**: One-click download of all detected anomalies into a formatted Excel file.

## 🛠️ Project Structure
To maintain high maintainability, the logic is separated from the UI:
- `app.py`: Streamlit interface, file upload handling, and result rendering.
- `logic.py`: Core data processing, regex date parsing, and cross-reference algorithms.

## 🚀 Getting Started

### Prerequisites
- Python 3.9+
- The following packages: `streamlit`, `pandas`, `openpyxl`, `xlrd`

### Installation
1. **Clone or download** this repository to your local machine.
2. **Install dependencies**:
   ```bash
   pip install streamlit pandas openpyxl xlrd
   ```

### Running the App
Navigate to the project folder and run:
```bash
streamlit run app.py
```

## 📖 How to Use
1. **Configure Settings**: Set the Target Year (Minguo), Target Month, and current Meal Prices in the sidebar.
2. **Upload Files**:
   - **Meal File**: The Excel file containing group sheets (e.g., 1組, 2組...日照).
   - **Leave Records**: Upload the H1 (1st-15th) and H2 (16th-31st) leave Excel files.
3. **Execute**: Click **"🚀 開始交叉比對"**.
4. **Review**: Check the anomaly table and the calculated total wasted costs.
5. **Download**: Export the results to Excel for administrative filing.

## ⚠️ Important Notes
- **File Schema**: The Leave Record must contain columns: `姓名`, `共計`, `差假開始日期`, `差假結束日期`.
- **Meal Table Format**: The Meal System file must have the `姓名` and `餐別` columns with daily columns numbered `1` through `31`.
- **Normalization**: The tool automatically handles full-width/half-width characters and case sensitivity (e.g., `v` vs `V`).

## 👨‍💻 Maintainability
This project follows a "Concise Logic" pattern. If the Excel layout changes in the future, modify the `TARGET_SHEETS` or `MEAL_CHECKPOINTS` constants in `logic.py`.

***
