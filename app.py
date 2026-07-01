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
    .location-container {
        background-color: #f8f9fa;
        padding: 12px 15px;
        border-radius: 8px;
        margin-top: 25px;
        margin-bottom: 10px;
        border-right: 6px solid #007bff;
    }
    .highlight-box {
        background-color: #eef1f6;
        padding: 10px;
        border-radius: 5px;
        border-right: 4px solid #28a745;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# 1. ניהול פרופילים והרשאות מוגדרים מראש
USER_PROFILES = {
    "מנהל מלאי": {"password": "123", "role": "admin"},
    "אפסנאי ראשי": {"password": "456", "role": "user"},
    "משתמש קצה": {"password": "789", "role": "user"}
}

# 2. הגדרות קבצים
FILENAME = "mlai_afsankol.xlsx"
LOG_FILE = "log_שינויי_מלאי.csv"

# 3. פונקציה לתיעוד פעולות (LOG)
def log_change(user, location, item, size, old_qty, new_qty, month, action_type="עדכון כמות"):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_data = {
        "תאריך ושעה": [now],
        "שם המשתמש": [user],
        "סוג הפעולה": [action_type],
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

# 4. פונקציה לטעינה וסידור המלאי המלא מאקסל (שומרת על כל הרשימות המקוריות מהקובץ)
def load_and_parse_data(file_path):
    if not os.path.exists(file_path):
        return None, f"שגיאה: הקובץ '{file_path}' לא נמצא בתיקיית האפליקציה. אנא ודא שהעלית אותו ל-GitHub."
    
    try:
        xls = pd.ExcelFile(file_path)
        all_data = {}
        
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            parsed_chunks = []
            
            # מעבר על הבלוקים הרוחביים באקסל (מבנה ה-4 עמודות שלך)
            for i in range(0, df.shape[1], 4):
                chunk = df.iloc[:, i:i+4]
                header_row_idx = None
                
                for idx, row in chunk.iterrows():
                    row_str = [str(val) if pd.notna(val) else "" for val in row.values]
                    if any("שם פריט" in s for s in row_str) and any("מידה" in s for s in row_str):
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    clean_chunk = chunk.iloc[header_row_idx+1:].copy()
                    clean_chunk.columns = ['שם פריט', 'מידה', 'כמות', 'ריק'][:chunk.shape[1]]
                    columns_to_keep = [col for col in clean_chunk.columns if col in ['שם פריט', 'מידה', 'כמות']]
                    clean_chunk = clean_chunk[columns_to_keep]
                    parsed_chunks.append(clean_chunk)
            
            if parsed_chunks:
                sheet_df = pd.concat(parsed_chunks, ignore_index=True)
                sheet_df['שם פריט'] = sheet_df['שם פריט'].fillna('').astype(str).str.strip()
                sheet_df['מידה'] = sheet_df['מידה'].fillna('').astype(str).str.strip()
                # המרת כמויות למספר (תאים ריקים הופכים ל-0 באופן אוטומטי כדי שתוכל לעדכן אותם)
                sheet_df['כמות'] = pd.to_numeric(sheet_df['כמות'], errors='coerce').fillna(0).astype(int)
                
                # סינון שורות כותרת כפולות ושורות ריקות לחלוטין
                sheet_df = sheet_df[(sheet_df['שם פריט'] != 'שם פריט') & (sheet_df['שם פריט'] != '')]
                all_data[sheet.strip()] = sheet_df.reset_index(drop=True)
                
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
    # מחובר בהצלחה
    st.sidebar.markdown(f"👤 מחובר בתור: **{st.session_state.current_user}**")
    if st.sidebar.button("התנתק 🔓"):
        st.session_state.authenticated = False
        st.session_state.current_user = None
        st.rerun()

    # טעינת המידע מתוך קובץ האקסל המקורי לזיכרון האפליקציה
    if 'data_dict' not in st.session_state:
        data_dict, error = load_and_parse_data(FILENAME)
        if not error:
            st.session_state.data_dict = data_dict
        else:
            st.error(error)

    if 'data_dict' in st.session_state:
        st.sidebar.header("⚙️ הגדרות עדכון")
        selected_month = st.sidebar.selectbox("📅 חודש עדכון נוכחי:", ["יולי 2026", "אוגוסט 2026", "ספטמבר 2026"])
        
        # בניית הטאבים החדשים לפי הדרישות שלך
        tab0, tab1, tab2 = st.tabs([
            "📋 צפייה ברשימות המלאי המקוריות", 
            "✏️ ביצוע עדכון ספירה מהיר", 
            "📜 יומן תיעוד שינויים (Log)"
        ])
        
        # טאב 0 - תצוגת רשימת הפריטים המלאה כפי שהיא מופיעה באקסל
        with tab0:
            st.subheader("🔍 רשימות הפריטים המלאות מתוך קובץ האקסל")
            st.write("להלן מבנה המוצרים והמידות הקבוע כפי שנטען מהקובץ המקורי. ניתן לבצע חיפוש מהיר:")
            
            global_search = st.text_input("🔍 סנן מוצר בכל המכולות במקביל (למשל: נעל, חולצה):", "")
            
            for loc, df_loc in st.session_state.data_dict.items():
                st.markdown(f"<div class='location-container'><h3>📦 {loc}</h3></div>", unsafe_allow_html=True)
                
                filtered_df = df_loc.copy()
                if global_search:
                    filtered_df = filtered_df[filtered_df['שם פריט'].str.contains(global_search, case=False, na=False)]
                
                if not filtered_df.empty:
                    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
                else:
                    st.info(f"אין פריטים להצגה במכולה זו התואמים לחיפוש.")

        # טאב 1 - ביצוע עדכון ספירה בכל רגע נתון
        with tab1:
            st.subheader("✏️ עדכון ספירת מלאי נוכחית")
            st.write("בחר את המיקום, סמן את הפריט שספרת והזן את הכמות החדשה ברגע זה:")
            
            locations_list = list(st.session_state.data_dict.keys())
            selected_loc = st.selectbox("1. בחר מיקום / מכולה לעדכון:", locations_list, key="edit_loc_select")
            
            # טעינת הטבלה המלאה של המיקום הנבחר כדי שהמשתמש יראה את כל האפשרויות שלו מול העיניים
            current_df = st.session_state.data_dict[selected_loc].copy()
            
            st.markdown("<div class='highlight-box'><strong>💡 טיפ לחסכון בזמן:</strong> רשימת הפריטים מטה מסודרת בדיוק לפי סדר הופעתה באקסל שלך.</div>", unsafe_allow_html=True)
            
            # בחירת הפריט והמידה מתוך הרשימה הקיימת באקסל
            available_items = sorted(list(current_df['שם פריט'].unique()))
            edit_item = st.selectbox("2. בחר את שם הפריט מתוך הרשימה הקיימת:", available_items)
            
            available_sizes = sorted(list(current_df[current_df['שם פריט'] == edit_item]['מידה'].unique()))
            edit_size = st.selectbox("3. בחר מידה לפריט זה:", available_sizes)
            
            # שליפת הכמות הנוכחית השמורה במערכת
            current_qty_row = current_df[(current_df['שם פריט'] == edit_item) & (current_df['מידה'] == edit_size)]
            current_qty = int(current_qty_row['כמות'].values[0]) if not current_qty_row.empty else 0
            
            st.info(f"📋 כמות רשומה נוכחית במערכת עבור פריט זה: **{current_qty}**")
            
            # הזנת הספירה החדשה
            new_qty = st.number_input("4. הזן את כמות הספירה החדשה בפועל:", min_value=0, value=current_qty, step=1)
            
            if st.button("💾 אשר, עדכן פריט ושמור ספירה"):
                if new_qty == current_qty:
                    st.info("הכמות שהזנת זהה לכמות הקיימת, לא בוצע שינוי.")
                else:
                    # מציאת האינדקס המדויק בתוך ה-Dataframe הראשי ועדכון הכמות שלו
                    master_df = st.session_state.data_dict[selected_loc]
                    idx_to_update = master_df[(master_df['שם פריט'] == edit_item) & (master_df['מידה'] == edit_size)].index
                    
                    if not idx_to_update.empty:
                        st.session_state.data_dict[selected_loc].iloc[idx_to_update[0], 2] = new_qty
                        
                        # תיעוד השינוי ביומן
                        log_change(
                            user=st.session_state.current_user,
                            location=selected_loc,
                            item=edit_item,
                            size=edit_size,
                            old_qty=current_qty,
                            new_qty=new_qty,
                            month=selected_month,
                            action_type="עדכון כמות בספירה"
                        )
                        
                        st.success(f"✅ הספירה עודכנה! הפריט '{edit_item}' (מידה {edit_size}) עודכן לכמות החדשה: {new_qty}.")
                        st.rerun()

        # טאב 2 - יומן תיעוד שינויים
        with tab2:
            st.subheader("📜 יומן תיעוד פעולות מלאי")
            if os.path.exists(LOG_FILE):
                log_df = pd.read_csv(LOG_FILE)
                st.dataframe(log_df.sort_index(ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("טרם בוצעו עדכוני ספירה במערכת. יומן התיעוד ריק.")
