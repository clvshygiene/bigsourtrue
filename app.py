import streamlit as st
import pandas as pd
from datetime import datetime
import io
from docx import Document 
from docx.shared import Pt, Inches, RGBColor  # 👈 新增 RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# --- 頁面設定 ---
st.set_page_config(page_title="校園掃區檢核系統", page_icon="🧹", layout="centered")
st.title("🧹 114-2 校園大掃除檢核系統")

# --- 1. 讀取資料函式 ---
@st.cache_data(ttl=600)
def load_data():
    try:
        # 👇 請確認這裡填寫的是正確的 Google 試算表連結
        google_sheet_url = "https://docs.google.com/spreadsheets/d/您的ID/edit?usp=sharing"
        
        # 自動轉換為 Excel 下載連結
        if "/edit" in google_sheet_url:
            excel_url = google_sheet_url.replace("/edit", "/export?format=xlsx")
            excel_url = excel_url.split("?")[0] + "?format=xlsx"
        else:
            excel_url = google_sheet_url

        # 讀取 Excel
        all_sheets = pd.read_excel(excel_url, sheet_name=None, dtype=str)
        
        # 檢查必要分頁
        required_sheets = ['班級清單', '地點資料庫', '掃區分配總表', '檢查標準']
        for sheet in required_sheets:
            if sheet not in all_sheets:
                st.error(f"❌ 找不到工作表：「{sheet}」")
                return None, None, None

        df_classes = all_sheets['班級清單']
        df_locations = all_sheets['地點資料庫']
        df_assign = all_sheets['掃區分配總表']
        df_standards = all_sheets['檢查標準']
        
        # 資料合併
        df_full = pd.merge(df_assign, df_locations, on='地點ID', how='left')
        df_full = pd.merge(df_full, df_classes, left_on='負責班級', right_on='班級代碼', how='left')
        df_full = df_full.dropna(subset=['負責班級'])
        
        return df_classes, df_full, df_standards
        
    except Exception as e:
        st.error(f"❌ 資料讀取失敗！錯誤訊息：{e}")
        return None, None, None

# --- 2. 產生 Word 文件的函式 ---
def generate_docx(class_name, tasks_df, standards_df):
    doc = Document()
    
    # 設定中文字型
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    # 標題
    title = doc.add_heading(f'{class_name} 大掃除檢核表', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.add_run(f"列印日期：{datetime.now().strftime('%Y-%m-%d')}\n").bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 準備標準字典
    standards_grouped = standards_df.groupby('檢查類型')

    # 遍歷任務
    for index, row in tasks_df.iterrows():
        # 地點名稱
        bldg = str(row['大樓']) if pd.notna(row['大樓']) else ""
        floor = str(row['樓層']) if pd.notna(row['樓層']) else ""
        detail = str(row['詳細位置']) if pd.notna(row['詳細位置']) else ""
        full_name = f"{bldg} {floor} {detail}".strip()
        
        # 加入地點標題
        doc.add_heading(f"📍 {full_name}", level=2)
        
        # 注意事項
        note = row['特別注意事項']
        if pd.notna(note) and str(note).strip() != "":
            p = doc.add_paragraph()
            run = p.add_run(f"⚠️ 注意：{note}")
            # 👇 【修正】這裡改成正確的 RGBColor
            run.font.color.rgb = RGBColor(255, 0, 0) 
        
        # 建立檢查表格
        check_type = row['檢查類型']
        if check_type in standards_grouped.groups:
            type_df = standards_grouped.get_group(check_type)
            
            # 建立表格
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '子分類'
            hdr_cells[1].text = '檢查項目'
            hdr_cells[2].text = '檢查確認(打勾)'
            
            # 填入資料
            if '子分類' in type_df.columns:
                type_df_sorted = type_df.sort_values(by=['子分類'], na_position='first')
                for item_row in type_df_sorted.itertuples():
                    row_cells = table.add_row().cells
                    sub_cat = str(item_row.子分類) if pd.notna(item_row.子分類) else "-"
                    row_cells[0].text = sub_cat
                    row_cells[1].text = item_row.檢查細項
                    row_cells[2].text = "□"
            else:
                 for item_row in type_df.itertuples():
                    row_cells = table.add_row().cells
                    row_cells[0].text = "-"
                    row_cells[1].text = item_row.檢查細項
                    row_cells[2].text = "□"
        else:
            doc.add_paragraph(f"(未找到類型 {check_type} 的檢查標準)")
            
        doc.add_paragraph("\n") # 空行

    # --- 簽名區塊 ---
    doc.add_page_break()
    doc.add_heading("簽名確認區", level=1)
    
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = 'Table Grid'
    
    for row in sig_table.rows:
        row.height = Inches(0.8)
    
    sig_table.cell(0, 0).text = "衛生股長 (1)"
    sig_table.cell(0, 1).text = "衛生股長 (2)"
    sig_table.cell(1, 0).text = "衛生糾察 (1)"
    sig_table.cell(1, 1).text = "衛生糾察 (2)"
    sig_table.cell(2, 0).text = "導師簽名"
    a = sig_table.cell(2, 0)
    b = sig_table.cell(2, 1)
    a.merge(b)

    return doc

# --- 主程式 ---
df_classes, df_tasks, df_standards = load_data()

if df_tasks is not None:
    st.sidebar.header("📍 班級登入")
    
    if '年級' in df_classes.columns:
        all_grades = sorted(df_classes['年級'].astype(str).unique())
        selected_grade = st.sidebar.selectbox("請選擇年級", all_grades)
        classes_filter = df_classes[df_classes['年級'] == selected_grade]
    else:
        st.error("班級清單缺少「年級」欄位")
        st.stop()
    
    class_options = {
        f"{row['班級代碼']} - {row['顯示名稱']}": row['班級代碼'] 
        for index, row in classes_filter.iterrows()
    }
    
    if not class_options:
        st.stop()

    selected_option = st.sidebar.selectbox("請選擇班級", list(class_options.keys()))
    current_class_id = class_options[selected_option]

    st.info(f"👋 歡迎 **{selected_option}**")
    
    my_tasks = df_tasks[df_tasks['負責班級'] == current_class_id]
    
    # --- Word 下載按鈕 ---
    if not my_tasks.empty:
        st.markdown("### 🖨️ 紙本檢核表下載")
        st.write("點擊下方按鈕下載 Word 檔，印出後完成簽名。")
        
        doc = generate_docx(selected_option, my_tasks, df_standards)
        bio = io.BytesIO()
        doc.save(bio)
        
        st.download_button(
            label="📥 下載 Word 檢核表 (.docx)",
            data=bio.getvalue(),
            file_name=f"{selected_option}_大掃除檢核表.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.markdown("---")

    # --- 數位預覽區 ---
    st.markdown("### 📱 數位預覽 (僅供參考)")
    standards_grouped = df_standards.groupby('檢查類型')

    if my_tasks.empty:
        st.warning("目前無分配掃區。")
    else:
        with st.form(key='preview_form'):
            for index, row in my_tasks.iterrows():
                bldg = str(row['大樓']) if pd.notna(row['大樓']) else ""
                floor = str(row['樓層']) if pd.notna(row['樓層']) else ""
                detail = str(row['詳細位置']) if pd.notna(row['詳細位置']) else ""
                full_name = f"{bldg} {floor} {detail}".strip()
                
                check_type = row['檢查類型']
                note = row['特別注意事項']
                location_id = row['地點ID']
                
                st.subheader(f"📍 {full_name}")
                if pd.notna(note) and str(note).strip() != "":
                    st.warning(f"注意：{note}")
                
                if check_type in standards_grouped.groups:
                    type_df = standards_grouped.get_group(check_type)
                    
                    if '子分類' in type_df.columns:
                        sub_groups = type_df.groupby('子分類', sort=False)
                        for sub_cat, items_df in sub_groups:
                            if pd.notna(sub_cat):
                                st.markdown(f"**🔹 {sub_cat}**")
                            
                            cols = st.columns(2)
                            for idx, item_row in enumerate(items_df.itertuples()):
                                unique_key = f"{current_class_id}_{location_id}_{sub_cat}_{item_row.檢查細項}_{idx}"
                                with cols[idx % 2]:
                                    st.checkbox(item_row.檢查細項, key=unique_key)
                            st.write("")
                    else:
                        for idx, item_row in enumerate(type_df.itertuples()):
                             unique_key = f"{current_class_id}_{location_id}_{item_row.檢查細項}_{idx}"
                             st.checkbox(item_row.檢查細項, key=unique_key)
                
                st.markdown("---")
            
            st.form_submit_button("數位送出 (測試用)")
