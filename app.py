import streamlit as st
import pandas as pd
import os
from datetime import datetime
import openpyxl

# 1. Page Configuration & Global RTL Layout styling
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
    /* Custom Styling for the Rollback Section */
    .rollback-container {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 8px;
        border-right: 5px solid #ffc107;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Files & User Profiles Setup
FILENAME = "mlai_afsankol.xlsx"
LOG_FILE = "log_שינויי_מלאי.csv"

USER_PROFILES = {
    "מנהל מלאי": {"password": "123", "role": "admin"},
    "אפסנאי ראשי": {"password": "456", "role": "user"},
    "משתמש קצה": {"password": "789", "role": "user"}
}

# 3. Helper: Log and Automatically Save back to the Live Excel File
def log_and_save_excel(user, sheet_name, df, item, size, old_qty, new_qty, month):
    # Save the updated structure back to the Excel file live
    try:
        with pd.ExcelWriter(FILENAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        st.error(f"שגיאה בשמירה אוטומטית לקובץ האקסל: {e}")
        return False

    # Append to the CSV action log
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

# 4. Data Loader: Parses repeated groups of 4 columns across your spreadsheets
def load_excel_data(file_path):
    if not os.path.exists(file_path):
        return None, f"שגיאה: הקובץ '{file_path}' לא נמצא."
    try:
        xls = pd.ExcelFile(file_path)
        all_sheets = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df = df.fillna("")
            all_sheets[sheet_name.strip()] = df
        return all_sheets, None
    except Exception as e:
        return None, f"שגיאה בקריאת הקובץ: {e}"

# 5. Styling DataFrame Columns distinctively (RTL, Color Mappings)
def color_columns(val, col_type):
    if col_type == 'name':
        return 'background-color: #e6f2ff; color: #000000; text-align: right;' # Light Blue for Product Name
    elif col_type == 'size':
        return 'background-color: #f2f2f2; color: #000000; text-align: right;' # Light Grey for Size
    elif col_type == 'qty':
        return 'background-color: #e2f0d9; color: #000000; font-weight: bold; text-align: right;' # Light Green for Quantity
    return 'text-align: right;'

def apply_sheet_styling(df):
    # Dynamically find column classifications based on row values or column indexes
    style_df = pd.DataFrame('', index=df.index, columns=df.columns)
    for col_idx, col_name in enumerate(df.columns):
        col_str = str(col_name).lower()
        # Fallback to column index groupings of 4 from your file structure
        rem = col_idx % 4
        if "שם פריט" in col_str or rem == 1:
            style_df[col_name] = style_df[col_name].apply(lambda x: color_columns(x, 'name'))
        elif "מידה" in col_str or rem == 2:
            style_df[col_name] = style_df[col_name].apply(lambda x: color_columns(x, 'size'))
        elif "כמות" in col_str or rem == 3:
            style_df[col_name] = style_df[col_name].apply(lambda x: color_columns(x, 'qty'))
    return style_df

# --- Authentication Logic ---
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_user = None
if 'history' not in st.session_state:
    st.session_state.history = []  # Stores deep-copies of sheets_dict states for custom rollbacks

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
    # Sidebar control panel
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
        
        tab1, tab2, tab3 = st.tabs(["📋 צפייה בלשוניות האקסל (RTL)", "✏️ עדכון ספירה ושינוי כמויות", "📜 יומן פעולות וביטול מהלכים (Rollback)"])
        
        # TAB 1 - Colored, Custom Visual Sheet Display
        with tab1:
            st.subheader("🔎 מבט על מלאי מקורי צבוע ומסודר (RTL)")
            for sheet, df_sheet in st.session_state.sheets_dict.items():
                st.markdown(f"### 📄 לשונית: {sheet}")
                # Render using custom Styler object for colors and true RTL alignment
                styled_df = df_sheet.style.apply(apply_sheet_styling, axis=None)
                st.dataframe(styled_df, use_container_width=True, hide_index=True)

        # TAB 2 - Live Editing & Verification
        with tab2:
            st.subheader("✏️ הזנת נתוני ספירה עדכניים")
            sheet_list = list(st.session_state.sheets_dict.keys())
            chosen_sheet = st.selectbox("1. בחר לשונית / מחסן:", sheet_list)
            
            df_target = st.session_state.sheets_dict[chosen_sheet]
            
            # Formulating drop-down choice selectors derived explicitly from the custom file row content
            row_options = []
            row_coords = {} # Maps user choice directly back to matrix coordinate (row index, column index)
            
            for idx, row in df_target.iterrows():
                # Scan groupings of 4 columns across the row for data content
                for c in range(1, df_target.shape[1], 4):
                    if c+2 < df_target.shape[1]:
                        item_val = str(row.iloc[c]).strip()
                        size_val = str(row.iloc[c+1]).strip()
                        qty_val = str(row.iloc[c+2]).strip()
                        
                        if item_val and item_val != "שם פריט" and item_val != "מכולה":
                            label = f"שורה {idx+1} | {item_val} | מידה: {size_val} | כמות נוכחית: {qty_val}"
                            row_options.append(label)
                            row_coords[label] = (idx, c, item_val, size_val, qty_val)
            
            if row_options:
                chosen_item_label = st.selectbox("2. בחר את השורה הספציפית שברצונך לעדכן:", row_options)
                t_row, t_col, name, size, old_qty = row_coords[chosen_item_label]
                
                try:
                    current_qty_num = int(float(old_qty)) if old_qty else 0
                except:
                    current_qty_num = 0
                
                new_qty = st.number_input(f"3. הזן כמות ספורה חדשה עבור {name} (מידה {size}):", min_value=0, value=current_qty_num, step=1)
                
                if st.button("💾 שמור שינויים באופן אוטומטי"):
                    # Save a deep copy of the state into our history tracker before doing modifications
                    history_snapshot = {k: v.copy() for k, v in st.session_state.sheets_dict.items()}
                    st.session_state.history.append((history_snapshot, name, size, chosen_sheet))
                    
                    # Apply changes instantly to local application state
                    st.session_state.sheets_dict[chosen_sheet].iloc[t_row, t_col+2] = int(new_qty)
                    
                    # Instantly save live updates directly back to the physical Excel sheet and logs
                    success = log_and_save_excel(
                        user=st.session_state.current_user,
                        sheet_name=chosen_sheet,
                        df=st.session_state.sheets_dict[chosen_sheet],
                        item=name,
                        size=size,
                        old_qty=current_qty_num,
                        new_qty=int(new_qty),
                        month=selected_month
                    )
                    if success:
                        st.success("✅ השינוי נשמר אוטומטית ישירות בקובץ האקסל וביומן המערכת!")
                        st.rerun()
            else:
                st.info("לא נמצאו פריטים תקינים הניתנים לעדכון בלשונית זו.")

        # TAB 3 - Rollback Mechanism Interface
        with tab3:
            st.subheader("⏪ מנגנון ביטול פעולות (Rollback)")
            if st.session_state.history:
                st.markdown("<div class='rollback-container'>", unsafe_allow_html=True)
                last_change = st.session_state.history[-1]
                st.warning(f"⚠️ הפעולה האחרונה שבוצעה: עדכון פריט **{last_change[1]}** (מידה {last_change[2]}) בלשונית **{last_change[3]}**.")
                
                if st.button("⏪ בטל פעולה אחרונה ושחזר מצב קודם"):
                    # Revert to last snapshot state
                    previous_state, p_name, p_size, p_sheet = st.session_state.history.pop()
                    st.session_state.sheets_dict = previous_state
                    
                    # Persist the rolled-back data structural state back directly inside Excel
                    try:
                        with pd.ExcelWriter(FILENAME, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            st.session_state.sheets_dict[p_sheet].to_excel(writer, sheet_name=p_sheet, index=False)
                        st.success("⏪ המערכת ביצעה שחזור (Rollback) מוצלח למצב המקורי ושמרה את קובץ האקסל!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"שגיאה בעת ניסיון שחזור פיזי של קובץ האקסל: {ex}")
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.info("אין פעולות קודמות הזמינות לביטול/שחזור בריצה הנוכחית.")
                
            st.write("---")
            st.subheader("📜 יומן שינויים היסטורי מלא")
            if os.path.exists(LOG_FILE):
                st.dataframe(pd.read_csv(LOG_FILE).sort_index(ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("יומן השינויים ריק כרגע.")
