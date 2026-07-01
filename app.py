import streamlit as st
import pandas as pd
import os

# הגדרת תצורת עמוד ותמיכה מימין לשמאל (RTL)
st.set_page_config(page_title="ניהול מלאי חודשי - אפסנכל", layout="wide", page_icon="📦")

st.markdown("""
    <style>
    body, div, p, h1, h2, h3, th, td, label {
        text-align: right;
        direction: rtl;
    }
    .stSelectbox, .stTextInput, .stNumberInput {
        direction: rtl;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 מערכת ניהול, מעקב ועדכון מלאי - אפסנכל")

FILENAME = "טבלת מעקב מלאים אפסנכל יולי 2026.xlsx"

# פונקציה לטעינה וסידור ראשוני של הנתונים
def load_and_parse_data(file_path):
    if not os.path.exists(file_path):
        return None, "הקובץ לא נמצא"
    
    xls = pd.ExcelFile(file_path)
    all_data = {}
    
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        parsed_chunks = []
        
        # מעבר על הבלוקים הרוחביים של 4 עמודות בקובץ
        for i in range(0, df.shape[1], 4):
            chunk = df.iloc[:, i:i+4]
            header_row_idx = None
            for idx, row in chunk.iterrows():
                row_str = row.astype(str).values
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
            sheet_df['שם פריט'] = sheet_df['שם פריט'].astype(str).str.strip()
            sheet_df['מידה'] = sheet_df['מידה'].astype(str).str.strip()
            sheet_df['כמות'] = pd.to_numeric(sheet_df['כמות'], errors='coerce').fillna(0).astype(int)
            sheet_df = sheet_df[sheet_df['שם פריט'] != 'שם פריט']
            all_data[sheet.strip()] = sheet_df
            
    return all_data, None

# טעינת המידע לתוך ה-Session State כדי שיישמר בין לחיצות כפתורים
if 'data_dict' not in st.session_state:
    data_dict, error = load_and_parse_data(FILENAME)
    if not error:
        st.session_state.data_dict = data_dict
    else:
        st.error(error)

if 'data_dict' in st.session_state:
    # תפריט ניווט צדדי לקביעת החודש/פעולה
    st.sidebar.header("⚙️ הגדרות מערכת")
    
    # בחירת חודש עבודה (מאפשר לעקוב רעיונית לפי חודשים)
    selected_month = st.sidebar.selectbox("📅 חודש עדכון נוכחי:", 
        ["יולי 2026", "אוגוסט 2026", "ספטמבר 2026", "אוקטובר 2026"])
    
    locations = list(st.session_state.data_dict.keys())
    selected_location = st.sidebar.selectbox("🎯 בחר מיקום / מכולה:", locations)
    
    current_df = st.session_state.data_dict[selected_location].copy().reset_index(drop=True)
    
    # יצירת טאבים להפרדה בין צפייה לעדכון
    tab1, tab2 = st.tabs(["🔍 צפייה וחיפוש במלאי", "✏️ עדכון כמויות מהיר"])
    
    # טאב 1: צפייה וחיפוש
    with tab1:
        st.subheader(f"📋 רשימת מלאי עבור חודש {selected_month} - {selected_location}")
        search_query = st.text_input("חפש פריט מהיר (לדוגמה: חולצה, נעל...):", "", key="search")
        
        filtered_df = current_df.copy()
        if search_query:
            filtered_df = filtered_df[filtered_df['שם פריט'].str.contains(search_query, case=False, na=False)]
        
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        
    # טאב 2: עדכון כמויות חודשי שוטף
    with tab2:
        st.subheader(f"✍️ טופס דיווח ועדכון מלאי - {selected_month}")
        st.write("בחר פריט ומידה והזן את הכמות העדכנית שנמצאת במחסן ברגע זה:")
        
        # בחירת הפריט מתוך הרשימה הקיימת במכולה
        available_items = sorted(list(current_df['שם פריט'].unique()))
        edit_item = st.selectbox("בחר פריט לעדכון:", available_items)
        
        # סינון המידות הזמינות עבור אותו הפריט בלבד
        available_sizes = sorted(list(current_df[current_df['שם פריט'] == edit_item]['מידה'].unique()))
        edit_size = st.selectbox("בחר מידה:", available_sizes)
        
        # שליפת הכמות הנוכחית הקיימת במערכת
        current_qty_row = current_df[(current_df['שם פריט'] == edit_item) & (current_df['מידה'] == edit_size)]
        current_qty = int(current_qty_row['כמות'].values[0]) if not current_qty_row.empty else 0
        
        st.info(f"הכמות הרשומה כרגע במערכת: **{current_qty}** יחידות.")
        
        # הזנת הכמות החדשה
        new_qty = st.number_input("הזן כמות חדשה (ספירה עדכנית):", min_value=0, value=current_qty, step=1)
        
        # כפתור שמירה ועדכון
        if st.button("💾 שמור עדכון מלאי חודשי"):
            # עדכון בתוך ה-Session State (הממשק הזמני)
            idx_to_update = current_df[(current_df['שם פריט'] == edit_item) & (current_df['מידה'] == edit_size)].index
            
            if not idx_to_update.empty:
                st.session_state.data_dict[selected_location].iloc[idx_to_update[0], 2] = new_qty
                st.success(f"העדכון בוצע בהצלחה! פריט '{edit_item}' במידה '{edit_size}' עודכן ל-{new_qty} יחידות עבור {selected_month}.")
                
                # כאן ניתן להוסיף קוד שישמור את השינויים חזרה לקובץ האקסל המקורי:
                # with pd.ExcelWriter(FILENAME) as writer:
                #     for sheet_name, df_sheet in st.session_state.data_dict.items():
                #         df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                
                st.rerun()
            else:
                st.error("הפריט או המידה לא נמצאו במערכת המבנית הנוכחית.")
