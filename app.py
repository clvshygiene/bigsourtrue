import streamlit as st
import pandas as pd
from datetime import datetime

# --- 頁面設定 ---
st.set_page_config(page_title="校園掃區檢核系統", page_icon="🧹", layout="centered")
st.title("🧹 114-2 校園大掃除檢核系統")

# --- 1. 讀取資料函式 (Excel 下載模式) ---
@st.cache_data(ttl=600)
def load_data():
    try:
        # 👇【請修改】貼上您的 Google 試算表連結
        google_sheet_url = "https://docs.google.com/spreadsheets/d/1jqpj-DOe1X2cf6cToWmtW19_0FdN3REioa34aXn4boA/edit?usp=sharing"
        
        # --- 自動將連結轉換為下載 Excel 的格式 ---
        # 主要是把 /edit 換成 /export?format=xlsx
        if "/edit" in google_sheet_url:
            excel_url = google_sheet_url.replace("/edit", "/export?format=xlsx")
            # 移除後面的參數確保乾淨
            excel_url = excel_url.split("?")[0] + "?format=xlsx"
        else:
            excel_url = google_sheet_url

        # 直接讀取 Excel (一次讀取所有工作表)
        # sheet_name=None 代表讀取全部，會回傳一個 Dictionary
        all_sheets = pd.read_excel(excel_url, sheet_name=None, dtype=str)
        
        # 檢查是否有缺分頁
        required_sheets = ['班級清單', '地點資料庫', '掃區分配總表', '檢查標準']
        for sheet in required_sheets:
            if sheet not in all_sheets:
                st.error(f"❌ 找不到工作表：「{sheet}」。請確認 Google 試算表的分頁名稱是否正確！")
                return None, None, None

        # 取出各個 DataFrame
        df_classes = all_sheets['班級清單']
        df_locations = all_sheets['地點資料庫']
        df_assign = all_sheets['掃區分配總表']
        df_standards = all_sheets['檢查標準']
        
        # --- 資料串接 (邏輯不變) ---
        df_full = pd.merge(df_assign, df_locations, on='地點ID', how='left')
        df_full = pd.merge(df_full, df_classes, left_on='負責班級', right_on='班級代碼', how='left')
        df_full = df_full.dropna(subset=['負責班級'])
        
        return df_classes, df_full, df_standards
        
    except Exception as e:
        st.error("❌ 資料讀取失敗！")
        st.warning(f"錯誤訊息: {e}")
        st.info("💡 請檢查：Google 試算表連結是否正確，權限是否已開啟「知道連結者可檢視」。")
        return None, None, None

# 執行資料載入
df_classes, df_tasks, df_standards = load_data()

if df_tasks is not None:

    # --- 2. 側邊欄：登入選單 ---
    st.sidebar.header("📍 班級登入")
    
    # 確保年級欄位存在且去重
    if '年級' in df_classes.columns:
        all_grades = sorted(df_classes['年級'].astype(str).unique())
        selected_grade = st.sidebar.selectbox("請選擇年級", all_grades)
        
        # 篩選班級
        classes_filter = df_classes[df_classes['年級'] == selected_grade]
    else:
        st.error("❌ 班級清單中找不到「年級」欄位，請檢查試算表。")
        st.stop()
    
    # 建立選單
    class_options = {
        f"{row['班級代碼']} - {row['顯示名稱']}": row['班級代碼'] 
        for index, row in classes_filter.iterrows()
    }
    
    if not class_options:
        st.warning("此年級無班級資料。")
        st.stop()

    selected_option = st.sidebar.selectbox("請選擇班級", list(class_options.keys()))
    current_class_id = class_options[selected_option]

    # --- 3. 主畫面 ---
    st.info(f"👋 歡迎 **{selected_option}**！ 請完成今日掃區檢查。")
    st.caption(f"📅 日期：{datetime.now().strftime('%Y-%m-%d')}")

    my_tasks = df_tasks[df_tasks['負責班級'] == current_class_id]
    standards_grouped = df_standards.groupby('檢查類型')

    if my_tasks.empty:
        st.warning("❓ 這個班級目前沒有分配到任何掃區。")
    else:
        with st.form(key='cleaning_form'):
            all_checked = True 
            
            for index, row in my_tasks.iterrows():
                # 處理顯示名稱
                bldg = row['大樓'] if pd.notna(row['大樓']) else ""
                floor = row['樓層'] if pd.notna(row['樓層']) else ""
                detail = row['詳細位置'] if pd.notna(row['詳細位置']) else ""
                full_name = f"{bldg} {floor} {detail}".strip()
                
                check_type = row['檢查類型']
                note = row['特別注意事項']
                location_id = row['地點ID']
                
                st.subheader(f"📍 {full_name}")
                
                if pd.notna(note) and str(note).strip() != "":
                    st.warning(f"💡 **注意：** {note}")
                
                if check_type in standards_grouped.groups:
                    type_df = standards_grouped.get_group(check_type)
                    # 依子分類分組
                    if '子分類' in type_df.columns:
                        sub_groups = type_df.groupby('子分類', sort=False)
                        for sub_cat, items_df in sub_groups:
                            if pd.notna(sub_cat):
                                st.markdown(f"**🔹 {sub_cat}**")
                            
                            cols = st.columns(2)
                            for idx, item_row in enumerate(items_df.itertuples()):
                                key_str = f"{current_class_id}_{location_id}_{item_row.檢查細項}"
                                with cols[idx % 2]:
                                    if not st.checkbox(item_row.檢查細項, key=key_str):
                                        all_checked = False
                            st.write("")
                    else:
                        # 如果沒有子分類欄位，就直接顯示
                        for item_row in type_df.itertuples():
                             if not st.checkbox(item_row.檢查細項, key=f"{current_class_id}_{location_id}_{item_row.檢查細項}"):
                                 all_checked = False
                else:
                    st.error(f"⚠️ 找不到類型「{check_type}」的檢查標準。")
                
                st.markdown("---") 

            feedback = st.text_area("📝 特殊狀況回報 (若無免填)")
            
            if st.form_submit_button("✅ 完成檢查並提交"):
                if all_checked:
                    st.balloons()
                    st.success("🎉 檢查完成，資料已送出！")
                else:
                    st.error("⚠️ 還有項目未勾選喔！")
