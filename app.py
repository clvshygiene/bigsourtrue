import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection  # 👈 新增這行

# ... (前面的頁面設定 set_page_config 維持不變) ...

# --- 1. 讀取與合併資料 (Google Sheets 版本) ---
@st.cache_data(ttl=600)  # ttl=600 代表資料會快取 10 分鐘，避免一直頻繁讀取 Google
def load_data():
    try:
        # 👇 請將這裡換成您剛剛複製的 Google 試算表連結
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1jqpj-DOe1X2cf6cToWmtW19_0FdN3REioa34aXn4boA/edit?usp=sharing"
        
        # 建立連線
        conn = st.connection("gsheets", type=GSheetsConnection)

        # 讀取四個分頁 (worksheet 對應您的分頁名稱)
        # usecols=None 代表讀取所有欄位, dtype=str 強制轉為文字格式避免 001 變 1
        df_classes = conn.read(spreadsheet=spreadsheet_url, worksheet="班級清單", dtype=str)
        df_locations = conn.read(spreadsheet=spreadsheet_url, worksheet="地點資料庫", dtype=str)
        df_assign = conn.read(spreadsheet=spreadsheet_url, worksheet="掃區分配總表", dtype=str)
        df_standards = conn.read(spreadsheet=spreadsheet_url, worksheet="檢查標準") # 標準這頁通常不用強制轉字串
        
        # --- 資料串接 (邏輯跟原本一樣) ---
        # 1. 以「地點ID」為準，將「地點資料庫」的資訊合併到「分配表」
        df_full = pd.merge(df_assign, df_locations, on='地點ID', how='left')
        
        # 2. 以「負責班級」為準，將「班級清單」的資訊合併進來
        # 注意：Google Sheets 讀進來有時候會有空白行，這裡多做一個 dropna 保險
        df_classes = df_classes.dropna(how='all')
        
        df_full = pd.merge(df_full, df_classes, left_on='負責班級', right_on='班級代碼', how='left')
        
        # 過濾掉沒有分配班級的地點
        df_full = df_full.dropna(subset=['負責班級'])
        
        return df_classes, df_full, df_standards
        
    except Exception as e:
        st.error(f"❌ 資料讀取失敗: {e}")
        st.info("請檢查：1. Google 試算表連結是否正確 2. 是否已開啟「知道連結者可檢視」權限 3. 分頁名稱是否正確")
        return None, None, None


    # --- 2. 側邊欄：登入選單 ---
    st.sidebar.header("📍 班級登入")
    
    # 步驟 1: 選擇年級 (排序)
    # unique() 抓出來可能是字串或數字，統一轉字串排序
    all_grades = sorted(df_classes['年級'].astype(str).unique())
    selected_grade = st.sidebar.selectbox("請選擇年級", all_grades)
    
    # 步驟 2: 選擇班級
    # 篩選該年級的班級
    classes_filter = df_classes[df_classes['年級'] == selected_grade]
    
    # 建立選單選項： "班級代碼 - 顯示名稱" (例如: 101 - 餐飲科)
    # 使用字典來對照： { "選項文字": "真實代碼" }
    class_options = {
        f"{row['班級代碼']} - {row['顯示名稱']}": row['班級代碼'] 
        for index, row in classes_filter.iterrows()
    }
    
    selected_option = st.sidebar.selectbox("請選擇班級", list(class_options.keys()))
    
    # 取得使用者選到的真實「班級代碼」
    current_class_id = class_options[selected_option]

    # --- 3. 主畫面：顯示檢核表 ---
    
    st.info(f"👋 歡迎 **{selected_option}**！ 請完成今日掃區檢查。")
    st.caption(f"📅 日期：{datetime.now().strftime('%Y-%m-%d')}")

    # 篩選出這個班級的所有任務
    my_tasks = df_tasks[df_tasks['負責班級'] == current_class_id]

    # 將檢查標準依照「檢查類型」分組，轉成字典方便查詢
    # key=類型, value=該類型的所有資料(DataFrame)
    standards_grouped = df_standards.groupby('檢查類型')

    if my_tasks.empty:
        st.warning("❓ 這個班級目前沒有分配到任何掃區，請確認分配表。")
    else:
        with st.form(key='cleaning_form'):
            all_checked = True # 預設全部都有勾
            
            # --- 逐一顯示每個掃區 ---
            for index, row in my_tasks.iterrows():
                
                # 1. 處理地點名稱 (如果欄位是 NaN 轉成空字串)
                bldg = row['大樓'] if pd.notna(row['大樓']) else ""
                floor = row['樓層'] if pd.notna(row['樓層']) else ""
                detail = row['詳細位置'] if pd.notna(row['詳細位置']) else ""
                
                full_name = f"{bldg} {floor} {detail}".strip()
                location_id = row['地點ID']
                check_type = row['檢查類型']
                note = row['特別注意事項']
                
                # 顯示標題
                st.subheader(f"📍 {full_name}")
                
                # 2. 顯示特別注意事項 (如果有寫的話)
                if pd.notna(note) and str(note).strip() != "":
                    st.warning(f"💡 **注意：** {note}")
                
                # 3. 抓取對應的檢查項目
                if check_type in standards_grouped.groups:
                    # 取得該類型的所有檢查項目
                    type_df = standards_grouped.get_group(check_type)
                    
                    # --- 支援「子分類」顯示 ---
                    # 依照子分類再分組一次 (例如：地面、窗戶、黑板)
                    # sort=False 讓它依照 Excel 的順序顯示，不要亂依照筆畫排
                    sub_groups = type_df.groupby('子分類', sort=False)
                    
                    for sub_cat, items_df in sub_groups:
                        # 顯示子分類小標題 (如果子分類不是空的)
                        if pd.notna(sub_cat):
                            st.markdown(f"**🔹 {sub_cat}**")
                        
                        # 顯示檢查細項 (使用兩欄併排，節省空間)
                        cols = st.columns(2)
                        for idx, item_row in enumerate(items_df.itertuples()):
                            item_name = item_row.檢查細項
                            # 製作唯一的 key，避免元件衝突
                            key_str = f"{current_class_id}_{location_id}_{item_name}"
                            
                            # 奇數放左邊，偶數放右邊
                            with cols[idx % 2]:
                                if not st.checkbox(item_name, key=key_str):
                                    all_checked = False
                        
                        st.write("") # 分隔一下子分類
                        
                else:
                    # 如果 Excel 裡有寫類型，但檢查標準表找不到
                    st.error(f"⚠️ 系統找不到類型「{check_type}」的檢查項目，請檢查 Excel 設定。")
                
                st.markdown("---") # 分隔線

            # --- 提交區 ---
            feedback = st.text_area("📝 特殊狀況回報 (例如：掃具損壞、設備故障，若無免填)")
            
            submit_btn = st.form_submit_button("✅ 完成檢查並提交")
            
            if submit_btn:
                if all_checked:
                    st.balloons()
                    st.success(f"🎉 太棒了！{selected_option} 檢查完成，資料已送出！")
                    # TODO: 這裡可以加入將結果寫入 Google Sheets 的程式碼
                else:
                    st.error("⚠️ 還有項目未勾選喔！請確認都有做到再提交。")