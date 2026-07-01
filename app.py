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
    .sheet-title {
        background-color: #2c3e50;
        color: white;
        padding: 10px 15px;
        border-radius: 6px;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    .update-box {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 8px;
        border-right: 5px solid #28a745;
        margin-top: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# 1. ניהול פרופילים והרשאות
USER_PROFILES = {
    "מנהל מלאי": {"password": "123", "role": "admin"},
    "אפסנאי ראשי": {"password": "456", "role": "user"},
    "משתמש קצה": {"password": "789", "role": "user"}
}

# 2. הגדרות קבצים (שונה לשם הקובץ המדויק שלך)
FILENAME = "mlai_afsankol.xlsx"
LOG_FILE = "log_שינויי_מלאי.csv"

# 3. פונקציה לתיעוד פעולות (LOG)
def log_change(user, location, item, size, old_qty, new_qty, month):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_data = {
        "תאריך ושעה": [now],
        "שם המשתמש": [user],
        "חודש עדכון": [month],
        "מיקום/לשונית": [location],
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

# 4. פונקציית טעינה מדויקת המשמרת את הגיליון כפי שהוא
def load_original_sheets(file_path):
    if not os.path.exists(file_path):
        return None, f"שגיאה: הקובץ '{file_path}' לא נמצא בתיקיית האפליקציה."
    
    try:
        xls = pd.ExcelFile(file_path)
        all_sheets = {}
        
        for sheet_name in xls.sheet_names:
            # קריאת הגיליון הגולמי ללא סינונים אגרסיביים
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # ניקוי בסיסי של שורות ועמודות ריקות לחלוטין מסביב
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # מילוי ערכים ריקים בטקסט נקי כדי שלא יקרוס
            df = df.fillna("")
            
            all_sheets[sheet_name.strip()] = df
            
        return all_sheets, None
    except Exception as e:
        return None, f"שגיאה בקריאת הקובץ: {e}"

# --- מסך הזדהות וכניסה ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_user = None

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
    #sidebar
    st.sidebar.markdown(f"👤 משתמש: **{st.session_state.current_user}**")
    if st.sidebar.button("התנתק 🔓"):
        st.session_state.authenticated = False
        st.rerun()

    # טעינת המידע פעם אחת לזיכרון הריצה
    if 'sheets_dict' not in st.session_state:
        sheets_dict, error = load_original_sheets(FILENAME)
        if not error:
            st.session_state.sheets_dict = sheets_dict
        else:
            st.error(error)

    if 'sheets_dict' in st.session_state:
        st.sidebar.header("📅 תקופת דיווח")
        selected_month = st.sidebar.selectbox("חודש עדכון נוכחי:", ["יולי 2026", "אוגוסט 2026", "ספטמבר 2026"])
        
        tab1, tab2, tab3 = st.tabs(["📋 צפייה בלשוניות המקוריות", "✏️ עדכון ספירה מהיר", "📜 יומן שינויים"])
        
        # טאב 1 - תצוגה מוחלטת של הגיליונות
        with tab1:
            st.subheader("🔎 צפייה במבנה הגיליונות המקורי")
            st.write("להלן הטבלאות המלאות כפי שהן מופיעות בלשוניות קובץ האקסל:")
            
            for sheet, df_sheet in st.session_state.sheets_dict.items():
                st.markdown(f"<div class='sheet-title'>📄 לשונית: {sheet}</div>", unsafe_allow_html=True)
                st.dataframe(df_sheet, use_container_width=True, hide_index=True)
                
        # טאב 2 - ביצוע עדכון ספירה בכל רגע נתון
        with tab2:
            st.subheader("✏️ הזנת ספירת מלאי פריטים")
            
            sheet_list = list(st.session_state.sheets_dict.keys())
            chosen_sheet = st.selectbox("1. בחר את הלשונית/המכולה הרלוונטית:", sheet_list)
            
            df_target = st.session_state.sheets_dict[chosen_sheet]
            
            st.write("2. בחר את הפריט המדויק מתוך הלשונית לעדכון:")
            
            # יצירת רשימת שורות ידידותית לבחירה (הקוד מחבר את שם הפריט והמידה מהשורות הקיימות)
            row_options = []
            row_mapping = {}
            
            # סריקה דינמית של השורות בלשונית הנבחרת
            for idx, row in df_target.iterrows():
                row_vals = [str(v) for v in row.values if str(v).strip() != ""]
                if row_vals:
                    # יצירת תיאור שורה מזהה (למשל: "נעל חיר - מידה 42")
                    label = f"שורה {idx + 1}: " + " | ".join(row_vals[:3])
                    row_options.append(label)
                    row_mapping[label] = idx
            
            if row_options:
                chosen_row_label = st.selectbox("בחר שורה לעדכון ספירה:", row_options)
                target_idx = row_mapping[chosen_row_label]
                
                st.markdown("<div class='update-box'>", unsafe_allow_html=True)
                st.markdown(f"**פריט נבחר:** {chosen_row_label}")
                
                # הזנת הכמות החדשה בספירה
                new_qty = st.number_input("הזן כמות ספורה נוכחית (בכל רגע נתון):", min_value=0, value=0, step=1)
                
                if st.button("💾 שמור ספירה ועדכן מערכת"):
                    # עדכון ישיר בתוך מבנה הטבלה המקורית של הלשונית (בדרך כלל עמודה שלישית או לפי זיהוי)
                    # נחפש את העמודה שבה רשום מספר או עמודה אינדקס 2
                    col_index = 2 if df_target.shape[1] > 2 else df_target.shape[1] - 1
                    
                    old_value = st.session_state.sheets_dict[chosen_sheet].iloc[target_idx, col_index]
                    st.session_state.sheets_dict[chosen_sheet].iloc[target_idx, col_index] = new_qty
                    
                    # רישום ביומן הפעולות
                    log_change(
                        user=st.session_state.current_user,
                        location=chosen_sheet,
                        item=chosen_row_label,
                        size="-",
                        old_qty=old_value,
                        new_qty=new_qty,
                        month=selected_month
                    )
                    
                    st.success("✅ הספירה נשמרה ועודכנה בלשונית בהצלחה!")
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.info("לא נמצאו פריטים ניתנים לעדכון בלשונית זו.")

        # טאב 3 - יומן
        with tab3:
            st.subheader("📜 היסטוריית עדכוני ספירה")
            if os.path.exists(LOG_FILE):
                st.dataframe(pd.read_csv(LOG_FILE).sort_index(ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("טרם בוצעו ספירות מלאי.")
