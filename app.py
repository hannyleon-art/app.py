import streamlit as st
import pandas as pd
import os
from datetime import datetime

# הגדרת תצורת עמוד ותמיכה מימין לשמאל (RTL)
st.set_page_config(page_title="ניהול מלאי מאובטח - אפסנכל", layout="wide", page_icon="🔒")

st.markdown("""
    <style>
    body, div, p, h1, h2, h3, th, td, label, input {
        text-align: right;
        direction: rtl;
    }
    .stSelectbox, .stTextInput, .stNumberInput, .stPassword {
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

# 1. ניהול פרופילים והרשאות מוגדרים מראש
USER_PROFILES = {
    "מנהל מלאי": {"password": "123", "role": "admin"},
    "אפסנאי ראשי": {"password": "456", "role": "user"},
    "משתמש קצה": {"password": "789", "role": "user"}
}

# 2. שם הקובץ החדש והנקי שלך
FILENAME = "mlai_afsankol.xlsx"
LOG_FILE = "log_שינויי_מלאי.csv"

# 3. פונקציה לתיעוד שינויים (LOG)
def log_change(user, location, item, size, old_qty, new_qty, month):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_data = {
        "תאריך ושעה": [now],
        "שם המשתמש": [user],
        "חודש עדכון": [month],
        "מיקום/מכולה": [location],
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

# 4. פונקציה לטעינה וסידור המלאי מאקסל - תוקנה לחלוטין להגנה מפני ערכי float/NaN
def load_and_parse_data(file_path):
    if not os.path.exists(file_path):
        return None, f"שגיאה: הקובץ '{file_path}' לא נמצא בתיקיית האפליקציה. אנא ודא שהעלית אותו."
    
    try:
        xls = pd.ExcelFile(file_path)
        all_data = {}
        
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            parsed_chunks = []
            
            # סריקת בלוקים רוחביים של 4 עמודות (שם פריט, מידה, כמות, ריק)
            for i in range(0, df.shape[1], 4):
                chunk = df.iloc[:, i:i+4]
                header_row_idx = None
                
                # מעבר על השורות למציאת שורת הכותרת "שם פריט" ו"מידה"
                for idx, row in chunk.iterrows():
                    # המרה בטוחה של כל ערכי השורה למחרוזת טקסט כדי למנוע שגיאות float לא ניתנות למעבר
                    row_str = [str(val) if pd.notna(val) else "" for val in row.values]
                    
                    if any("שם פריט" in s for s in row_str) and any("מידה" in s for s in row_str):
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    clean_chunk = chunk.iloc[header_row_idx+1:].copy()
                    clean_chunk.columns = ['שם פריט', 'מידה', 'כמות', 'ריק'][:chunk.shape[1]]
                    columns_to_keep = [col for col in clean_chunk.columns if col in ['שם פריט', 'מידה', 'כמות']]
                    clean_chunk = clean_chunk[columns_to_keep].dropna(subset=['שם פריט'])
                    parsed_chunks.append(clean_chunk)
            
            if parsed_chunks:
                sheet_df = pd.concat(parsed_chunks, ignore_index=True)
                # ניקוי והגנה על הערכים בתוך הטבלה המאוחדת
                sheet_df['שם פריט'] = sheet_df['שם פריט'].fillna('').astype(str).str.strip()
                sheet_df['מידה'] = sheet_df['מידה'].fillna('').astype(str).str.strip()
                sheet_df['כמות'] = pd.to_numeric(sheet_df['כמות'], errors='coerce').fillna(0).astype(int)
                sheet_df = sheet_df[(sheet_df['שם פריט'] != 'שם פריט') & (sheet_df['שם פריט'] != '')]
                all_data[sheet.strip()] = sheet_df
                
        return all_data, None
    except Exception as e:
        return None, f"שגיאה בקריאת הקובץ: {e}"

# --- מסך הזדהות וכניסה ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_user = None

if not st.session_state.authenticated:
    st.header("🔒 כניסה למערכת המלאי - אפסנכל")
    st.write("אנא הזדהה באמצעות הפרופיל שלך כדי להמשיך:")
    
    selected_profile = st.selectbox("בחר פרופיל משתמש:", list(USER_PROFILES.keys()))
    password_input = st.text_input("הזן סיסמה:", type="password")
    
    if st.button("התחבר"):
        if USER_PROFILES[selected_profile]["password"] == password_input:
            st.session_state.authenticated = True
            st.session_state.current_user = selected_profile
            st.success(f"ברוך הבא, {selected_profile}!")
            st.rerun()
        else:
            st.error("הסיסמה שהוזנה אינה נכונה.")
else:
    # המשתמש מחובר
    st.sidebar.markdown(f"👤 מחובר בתור: **{st.session_state.current_user}**")
    if st.sidebar.button("התנתק 🔓"):
        st.session_state.authenticated = False
        st.session_state.current_user = None
        st.rerun()

    # טעינת המידע ל-Session State
    if 'data_dict' not in st.session_state:
        data_dict, error = load_and_parse_data(FILENAME)
        if not error:
            st.session_state.data_dict = data_dict
        else:
            st.error(error)

    if 'data_dict' in st.session_state:
        st.sidebar.header("⚙️ הגדרות")
        selected_month = st.sidebar.selectbox("📅 חודש עדכון נוכחי:", ["יולי 2026", "אוגוסט 2026", "ספטמבר 2026"])
        locations = list(st.session_state.data_dict.keys())
        selected_location = st.sidebar.selectbox("🎯 בחר מיקום / מכולה:", locations)
        
        current_df = st.session_state.data_dict[selected_location].copy().reset_index(drop=True)
        
        # יצירת טאבים
        tab1, tab2, tab3 = st.tabs(["🔍 צפייה במלאי", "✏️ עדכון כמויות חודשי", "📜 יומן תיעוד שינויים (Log)"])
        
        # טאב 1 - צפייה
        with tab1:
            st.subheader(f"📋 רשימת מלאי נוכחית: {selected_location}")
            search_query = st.text_input("חפש פריט מהיר:", "", key="view_search")
            filtered_df = current_df.copy()
            if search_query:
                filtered_df['שם פריט'] = filtered_df['שם פריט'].astype(str)
                filtered_df = filtered_df[filtered_df['שם פריט'].str.contains(search_query, case=False, na=False)]
            st.dataframe(filtered_df, use_container_width=True, hide_index=True)
            
        # טאב 2 - עדכון
        with tab2:
            st.subheader(f"✍️ עדכון מלאי תקופתי - חודש {selected_month}")
            st.warning(f"שים לב: כל שינוי שיתבצע יירשם באופן מאובטח תחת המשתמש שלך (**{st.session_state.current_user}**).")
            
            available_items = sorted(list(current_df['שם פריט'].unique()))
            edit_item = st.selectbox("בחר פריט לעדכון:", available_items)
            
            available_sizes = sorted(list(current_df[current_df['שם פריט'] == edit_item]['מידה'].unique()))
            edit_size = st.selectbox("בחר מידה:", available_sizes)
            
            current_qty_row = current_df[(current_df['שם פריט'] == edit_item) & (current_df['מידה'] == edit_size)]
            current_qty = int(current_qty_row['כמות'].values[0]) if not current_qty_row.empty else 0
            
            st.write(f"הכמות הנוכחית הרשומה במערכת: **{current_qty}**")
            new_qty = st.number_input("הזן כמות ספורה חדשה:", min_value=0, value=current_qty, step=1)
            
            if st.button("💾 אשר ושמור שינוי במלאי"):
                if new_qty == current_qty:
                    st.info("לא בוצע שינוי בכמות, אין מה לעדכן.")
                else:
                    idx_to_update = current_df[(current_df['שם פריט'] == edit_item) & (current_df['מידה'] == edit_size)].index
                    if not idx_to_update.empty:
                        st.session_state.data_dict[selected_location].iloc[idx_to_update[0], 2] = new_qty
                        
                        log_change(
                            user=st.session_state.current_user,
                            location=selected_location,
                            item=edit_item,
                            size=edit_size,
                            old_qty=current_qty,
                            new_qty=new_qty,
                            month=selected_month
                        )
                        
                        st.success(f"✅ המלאי עודכן בהצלחה ונתוני השינוי נרשמו ביומן המערכת!")
                        st.rerun()

        # טאב 3 - יומן תיעוד
        with tab3:
            st.subheader("📜 היסטוריית פעולות ועדכוני מלאי")
            if os.path.exists(LOG_FILE):
                log_df = pd.read_csv(LOG_FILE)
                st.dataframe(log_df.sort_index(ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("טרם בוצעו שינויים במלאי. יומן התיעוד ריק כעת.")
