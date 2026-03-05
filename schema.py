# schema.py

# ==========================================
# 差假紀錄 (Leave Records) 欄位與常數
# ==========================================
COL_LEAVE_NAME = "姓名"
COL_LEAVE_TOTAL = "共計"
COL_LEAVE_START = "差假開始日期"
COL_LEAVE_END = "差假結束日期"

LEAVE_REQUIRED_COLS = {COL_LEAVE_NAME, COL_LEAVE_TOTAL, COL_LEAVE_START, COL_LEAVE_END}

# ==========================================
# 點餐紀錄 (Meal Records) 欄位與常數
# ==========================================
COL_MEAL_NAME = "姓名"
COL_MEAL_TYPE = "餐別"

MEAL_REQUIRED_COLS = {COL_MEAL_NAME, COL_MEAL_TYPE}

# ==========================================
# 點餐記號 (Valid Meal Marks) 常數
# ==========================================
VALID_MEAL_MARKS = ['V', 'N']

# ==========================================
# 驗證函式 (Validation Functions)
# ==========================================
def validate_leave_schema(df_columns):
    """
    驗證差假紀錄是否包含所有必要欄位。
    回傳: (is_valid: bool, missing_cols: set)
    """
    missing_cols = LEAVE_REQUIRED_COLS - set(df_columns)
    is_valid = len(missing_cols) == 0
    return is_valid, missing_cols

def validate_meal_schema(df_columns):
    """
    驗證點餐紀錄是否包含所有必要欄位。
    回傳: (is_valid: bool, missing_cols: set)
    """
    missing_cols = MEAL_REQUIRED_COLS - set(df_columns)
    is_valid = len(missing_cols) == 0
    return is_valid, missing_cols

def validate_meal_days(df_columns):
    """
    驗證點餐紀錄中是否有包含任何日期欄位 (1-31)。
    回傳: (has_days: bool, day_cols_found: list)
    """
    day_cols_found =[str(d) for d in range(1, 32) if str(d) in df_columns]
    has_days = len(day_cols_found) > 0
    return has_days, day_cols_found
