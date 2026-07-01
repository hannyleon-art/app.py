import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl

# 1. הגדרת תצורת עמוד ותמיכה מלאה מימין לשמאל (RTL)
st.set_page_config(page_title="ניהול מלאי מאובטח - אפסנכל", layout="wide", page_icon="🔒")

st.markdown("""
    <style>
    body, div, p, h1, h2, h3, label, input {
        text-align: right;
        direction: rtl;
    }
    .stSelectbox, .stTextInput, .stNumberInput, .stPassword, .stDataFrame {
        direction: rtl !important;
        text-align: right !important;
    }
    .block-container-box {
        background-color: #fdfdfd;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 30px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .block-header {
        background-color: #2c3e50;
        color: white;
        padding: 6px 12px;
        border-radius: 4px;
        font-size: 14px;
        margin-bottom: 10px;
        display: inline-block;
    }
    .rollback-container {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 8px;
        border-right: 5px solid #ffc107;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# 2. הגדרות קבצים ופרופילים
FILENAME = "mlai_afsankol.xlsx"
LOG_FILE = "log_שינויי_מלאי.csv"

USER_PROFILES = {
    "מנהל מלאי": {"password": "123", "role": "admin"},
    "אפסנאי ראשי": {"password": "456", "role": "user"},
    "משתמש קצה": {"password": "789", "role": "user"}
}

# 3. פונקציה לשמירה אוטומטית ותיעוד ביומן
def log_and_save_excel(user, sheet_name, df_to_save, item, size, old_qty, new_qty, month):
    try:
        with pd.ExcelWriter(FILENAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        st.error(f"שגיאה בשמירה אוטומטית לקובץ האקסל: {e}")
        return False

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_data = {
        "תאריך ושעה": [now],
        "שם המשתמש": [user],
        "חודש עדכון": [month],
        "לשונית/מיקום": [sheet_name],
        "שם פריט": [item],
        "מידה": [size],
        "כמות קודמת": [old_qty],
        "כמות חדשה": [new_qty]
    }
    new_log_df = pd.DataFrame(log_data)
    if os.path.exists(LOG_FILE):
        new_log_df.to_csv(LOG_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
    else:
        new_log_df.to_csv(LOG_FILE, mode='w', header=True, index=False, encoding='utf-8-sig')
    return True

# 4. פונקציית טעינה ופירוק הבלוקים הרוחביים (הסרת השורה העליונה הריקה/המיותרת)
def load_excel_data(file_path):
    if not os.path.exists(file_path):
        return None, f"שגיאה: הקובץ '{file_path}' לא נמצא."
    try:
        xls = pd.ExcelFile(file_path)
        all_sheets = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # אם השורה הראשונה מכילה ערכים כמו "מכולה", "מחסן פנימי" או שהיא כמעט כולה ריקה - נסיר אותה
            if df.shape[0] > 0:
                first_row_str = "".join([str(x) for x in df.iloc[0].values])
                if "מכולה" in first_row_str or "מחסן" in first_row_str or df.iloc[0].isna().sum() > (df.shape[1] / 2):
                    df = df.iloc[1:].reset_index(drop=True)
            
            df = df.fillna("")
            all_sheets[sheet_name.strip()] = df
        return all_sheets, None
    except Exception as e:
        return None, f"שגיאה בקריאת הקובץ: {e}"

# 5. עיצוב צבעים ייחודי לעמודות ה-RTL
def apply_block_styling(df):
    style_df = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in df.columns:
        col_str = str(col).strip()
        if "שם פריט" in col_str:
            style_df[col] = 'background-color: #e6f2ff; color: #000000; text-align: right;' # כחול בהיר לשם הפריט
        elif "מידה" in col_str:
            style_df[col] = 'background-color: #f2f2f2; color: #000000; text-align: right;' # אפור בהיר למידה
        elif "כמות" in col_str:
            style_df[col] = 'background-color: #e2f0d9; color: #000000; font-weight: bold; text-align: right;' # ירוק בהיר לכמות
    return style_df

# --- מנגנון התחברות ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_user = None
if 'history' not in st.session_state:
    st.session_state.history = []

if not st.session_state.authenticated:
    st.header("🔒 כניסה למערכת המלאי - אפסנכל")
    selected_profile = st.selectbox("בחר פרופיל משתמש:", list(USER_PROFILES.keys()))
    password_input = st.text_input("הזן סיסמה:", type="password")
    if st.button("התחבר"):
        if USER_PROFILES[selected_profile]["password"] == password_input:
            st.session_state.authenticated = True
            st.session_state.current_user = selected_profile
            st.rerun()
        else:
            st.error("הסיסמה שגויה.")
else:
    st.sidebar.markdown(f"👤 משתמש: **{st.session_state.current_user}**")
    if st.sidebar.button("התנתק 🔓"):
        st.session_state.authenticated = False
        st.rerun()

    if 'sheets_dict' not in st.session_state:
        sheets_dict, error = load_excel_data(FILENAME)
        if not error:
            st.session_state.sheets_dict = sheets_dict
        else:
            st.error(error)

    if 'sheets_dict' in st.session_state:
        st.sidebar.header("📅 תקופת דיווח")
        selected_month = st.sidebar.selectbox("חודש עדכון נוכחי:", ["יולי 2026", "אוגוסט 2026", "ספטמבר 2026"])
        
        tab1, tab2, tab3 = st.tabs(["📋 צפייה בבלוקים של מלאי", "✏️ עדכון ספירה מהיר", "📜 יומן פעולות וביטול (Rollback)"])
        
        # טאב 1 - תצוגה מופרדת של בלוקים (כל בלוק בנפרד, עמודות מסודרות מימין לשמאל)
        with tab1:
            st.subheader("🔎 רשימות מלאי מופרדות לפי בלוקים (RTL)")
            sheet_list = list(st.session_state.sheets_dict.keys())
            chosen_sheet = st.selectbox("בחר לשונית לצפייה מחסן/מכולה:", sheet_list, key="view_sheet_select")
            
            df_sheet = st.session_state.sheets_dict[chosen_sheet]
            
            # זיהוי ופירוק הבלוקים הרוחביים (כל 4 עמודות הן בלוק נפרד)
            block_index = 1
            for c in range(0, df_sheet.shape[1], 4):
                if c + 2 < df_sheet.shape[1]:
                    # שליפת 3 העמודות של הבלוק הנוכחי
                    block_df = df_sheet.iloc[:, c:c+3].copy()
                    
                    # קביעת שמות העמודות על בסיס השורה הראשונה האמיתית (שורת הכותרת "שם פריט", "מידה", "כמות")
                    if block_df.shape[0] > 0:
                        new_headers = [str(block_df.iloc[0, 0]), str(block_df.iloc[0, 1]), str(block_df.iloc[0, 2])]
                        # ניקוי כותרות דינמי
                        new_headers = [h if h != "nan" and h.strip() != "" else f"עמודה_{i}" for i, h in enumerate(new_headers)]
                        block_df.columns = new_headers
                        block_df = block_df.iloc[1:].reset_index(drop=True)
                    
                    # סינון שורות ריקות לחלוטין בתוך הבלוק כדי לשמור על קובץ נקי
                    block_df = block_df[block_df.iloc[:, 0].astype(str).str.strip() != ""]
                    
                    if not block_df.empty:
                        st.markdown(f"<div class='block-container-box'>", unsafe_allow_html=True)
                        st.markdown(f"<div class='block-header'>📦 מכולה: {chosen_sheet} | קבוצת פריטים #{block_index}</div>", unsafe_allow_html=True)
                        
                        # סידור פיזי של העמודות מימין לשמאל: שם פריט -> מידה -> כמות
                        desired_order = []
                        for col in block_df.columns:
                            if "שם פריט" in str(col): desired_order.append(col)
                        for col in block_df.columns:
                            if "מידה" in str(col): desired_order.append(col)
                        for col in block_df.columns:
                            if "כמות" in str(col): desired_order.append(col)
                        
                        # הגנת גיבוי במקרה והשמות שונים
                        if len(desired_order) == 3:
                            block_df = block_df[desired_order]
                        
                        styled_block = block_df.style.apply(apply_block_styling, axis=None)
                        st.dataframe(styled_block, use_container_width=True, hide_index=True)
                        st.markdown("</div>", unsafe_allow_html=True)
                        block_index += 1

        # טאב 2 - עדכון ספירה דינמי לפי הבלוקים המופרדים
        with tab2:
            st.subheader("✏️ הזנת נתוני ספירת מלאי")
            chosen_sheet_edit = st.selectbox("1. בחר לשונית / מחסן לעדכון:", sheet_list, key="edit_sheet_select")
            df_target = st.session_state.sheets_dict[chosen_sheet_edit]
            
            row_options = []
            row_coords = {}
            
            # מיפוי השורות והבלוקים לבחירה מהירה וידידותית
            for idx, row in df_target.iterrows():
                # דילוג על שורת הכותרות המובנית של האקסל
                if idx == 0:
                    continue
                for c in range(0, df_target.shape[1], 4):
                    if c + 2 < df_target.shape[1]:
                        item_val = str(row.iloc[c]).strip()
                        size_val = str(row.iloc[c+1]).strip()
                        qty_val = str(row.iloc[c+2]).strip()
                        
                        if item_val and "שם פריט" not in item_val:
                            label = f"{item_val} | מידה: {size_val} | כמות נוכחית: {qty_val if qty_val else '0'}"
                            row_options.append(label)
                            row_coords[label] = (idx, c, item_val, size_val, qty_val)
            
            if row_options:
                chosen_item_label = st.selectbox("2. בחר את הפריט המדויק לעדכון ספירה:", sorted(list(set(row_options))))
                t_row, t_col, name, size, old_qty = row_coords[chosen_item_label]
                
                try:
                    current_qty_num = int(float(old_qty)) if old_qty else 0
                except:
                    current_qty_num = 0
                
                new_qty = st.number_input(f"3. הזן כמות ספורה חדשה עבור {name} (מידה {size}):", min_value=0, value=current_qty_num, step=1)
                
                if st.button("💾 שמור ספירה אוטומטית לקובץ"):
                    # שמירת מצב היסטורי לצורך ביטול פעולה
                    history_snapshot = {k: v.copy() for k, v in st.session_state.sheets_dict.items()}
                    st.session_state.history.append((history_snapshot, name, size, chosen_sheet_edit))
                    
                    # עדכון מקומי בזיכרון האפליקציה
                    st.session_state.sheets_dict[chosen_sheet_edit].iloc[t_row, t_col+2] = int(new_qty)
                    
                    # שמירה אוטומטית מלאה ישירות לקובץ האקסל הפיזי
                    success = log_and_save_excel(
                        user=st.session_state.current_user,
                        sheet_name=chosen_sheet_edit,
                        df_to_save=st.session_state.sheets_dict[chosen_sheet_edit],
                        item=name,
                        size=size,
                        old_qty=current_qty_num,
                        new_qty=int(new_qty),
                        month=selected_month
                    )
                    if success:
                        st.success("✅ הספירה עודכנה ונשמרה אוטומטית ישירות באקסל המרכזי!")
                        st.rerun()
            else:
                st.info("לא נמצאו פריטים תקינים בלשונית זו.")

        # טאב 3 - מנגנון שחזור לאחור (Rollback) ויומן שינויים
        with tab3:
            st.subheader("⏪ מנגנון ביטול פעולות (Rollback)")
            if st.session_state.history:
                st.markdown("<div class='rollback-container'>", unsafe_allow_html=True)
                last_change = st.session_state.history[-1]
                st.warning(f"⚠️ פעולה אחרונה שבוצעה בריצה זו: עדכון פריט **{last_change[1]}** (מידה {last_change[2]}) בלשונית **{last_change[3]}**.")
                
                if st.button("⏪ בטל פעולה אחרונה ושחזר את קובץ האקסל"):
                    previous_state, p_name, p_size, p_sheet = st.session_state.history.pop()
                    st.session_state.sheets_dict = previous_state
                    
                    try:
                        with pd.ExcelWriter(FILENAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            st.session_state.sheets_dict[p_sheet].to_excel(writer, sheet_name=p_sheet, index=False)
                        st.success("⏪ המערכת ביצעה שחזור (Rollback) מוצלח ושמרה את המצב הקודם באקסל!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"שגיאה בעת ניסיון שחזור פיזי של קובץ האקסל: {ex}")
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.info("אין פעולות קודמות הזמינות לביטול בריצה הנוכחית.")
                
            st.write("---")
            st.subheader("📜 יומן שינויים היסטורי מלא")
            if os.path.exists(LOG_FILE):
                st.dataframe(pd.read_csv(LOG_FILE).sort_index(ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("יומן השינויים ריק כרגע.")
