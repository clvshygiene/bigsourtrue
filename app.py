import streamlit as st
import pandas as pd
from datetime import datetime
import io
from docx import Document 
from docx.shared import Pt, Inches, RGBColor, Cm
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
        google_sheet_url = "https://docs.google.com/spreadsheets/d/1jqpj-DOe1X2cf6cToWmtW19_0FdN3REioa34aXn4boA/edit?usp=sharing
"
        
        if "/edit" in google_sheet_url:
            excel_url = google_sheet_url.replace("/edit", "/export?format=xlsx")
            excel_url = excel_url.split("?")[0] + "?format=xlsx"
        else:
            excel_url = google_sheet_url

        all_sheets = pd.read_excel(excel_url, sheet_name=None, dtype=str)
        
        required_sheets = ['班級清單', '地點資料庫', '掃區分配總表', '檢查標準']
        for sheet in required_sheets:
            if sheet not in all_sheets:
                st.error(f"❌ 找不到工作表：「{sheet}」")
                return None, None, None

        df_classes = all_sheets['班級清單']
        df_locations = all_sheets['地點資料庫']
        df_assign = all_sheets['掃區分配總表']
        df_standards = all_sheets['檢查標準']
        
        df_full = pd.merge(df_assign, df_locations, on='地點ID', how='left')
        df_full = pd.merge(df_full, df_classes, left_on='負責班級', right_on='班級代碼', how='left')
        df_full = df_full.dropna(subset=['負責班級'])
        
        return df_classes, df_full, df_standards
        
    except Exception as e:
        st.error(f"❌ 資料讀取失敗！錯誤訊息：{e}")
        return None, None, None

# --- 輔助函式：建立簽名區 (2x2 矩陣) ---
def add_signature_block(doc):
    doc.add_paragraph("\n") # 隔開一點距離
    
    # 建立 2x2 表格 (衛生股長, 衛生糾察 / 導師, 衛生組)
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    
    # 設定列高 (簽名要有空間)
    for row in sig_table.rows:
        row.height = Cm(2.0) # 設定約 2 公分高，夠簽名
    
    # 填入標題 (左上角小字或是直接置中)
    # 這裡採用：標題 + 換行預留空間的方式
    
    # 第一列
    c1 = sig_table.cell(0, 0)
    c1.text = "衛生股長"
    c1.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    c2 = sig_table.cell(0, 1)
    c2.text = "衛生糾察"
    
    # 第二列
    c3 = sig_table.cell(1, 0)
    c3.text = "導師簽名"
    
    c4 = sig_table.cell(1, 1)
    c4.text = "衛生組核章"

# --- 輔助函式：建立任務清單區 ---
def add_task_section(doc, tasks_df, standards_grouped, title_text):
    # 標題
    heading = doc.add_heading(title_text, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 【修正】移除列印日期
    # p = doc.add_paragraph() ... (已刪除)

    for index, row in tasks_df.iterrows():
        bldg = str(row['大樓']) if pd.notna(row['大樓']) else ""
        floor = str(row['樓層']) if pd.notna(row['樓層']) else ""
        detail = str(row['詳細位置']) if pd.notna(row['詳細位置']) else ""
        full_name = f"{bldg} {floor} {detail}".strip()
        
        doc.add_heading(f"📍 {full_name}", level=2)
        
        note = row['特別注意事項']
        if pd.notna(note) and str(note).strip() != "":
            p = doc.add_paragraph()
            run = p.add_run(f"⚠️ 注意：{note}")
            run.font.color.rgb = RGBColor(255, 0, 0)
        
        check_type = row['檢查類型']
        if check_type in standards_grouped.groups:
            type_df = standards_grouped.get_group(check_type)
            
            # 【視覺改良】表格設定
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            # ⛔ 重要：關閉自動調整，這樣我們設定的寬度才會生效
            table.allow_autofit = False 
            
            # 設定表頭
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '檢查項目'
            hdr_cells[1].text = '確認'
            
            # 置中表頭
            hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 【關鍵】設定欄寬
            # 總寬度約 18.5cm (A4 21cm - 左右邊界 1.27*2)
            # 設定確認欄只要 1.5 cm，剩下給項目欄
            table.columns[0].width = Cm(17.0) 
            table.columns[1].width = Cm(1.5) 
            
            # 確保第一列的儲存格寬度也被鎖定 (python-docx 的特性)
            hdr_cells[0].width = Cm(17.0)
            hdr_cells[1].width = Cm(1.5)

            if '子分類' in type_df.columns:
                type_df_sorted = type_df.sort_values(by=['子分類'], na_position='first')
            else:
                type_df_sorted = type_df

            for item_row in type_df_sorted.itertuples():
                row_cells = table.add_row().cells
                row_cells[0].text = item_row.檢查細項
                
                # 調整欄寬 (每一列都要設定，確保整齊)
                row_cells[0].width = Cm(17.0)
                row_cells[1].width = Cm(1.5)
                
                # 確認格置中
                p = row_cells[1].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run("□")
                run.font.size = Pt(14) # 方框稍微大一點點比較好勾
        else:
            doc.add_paragraph(f"(未找到類型 {check_type} 的檢查標準)")
            
        doc.add_paragraph("") 

    add_signature_block(doc)


# --- 2. 產生 Word 文件的核心函式 ---
def generate_docx(display_name, tasks_df, standards_df):
    doc = Document()
    
    # 版面邊界設為「窄」
    section = doc.sections[0]
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    standards_grouped = standards_df.groupby('檢查類型')

    df_indoor = tasks_df[tasks_df['檢查類型'] == '內掃教室']
    df_outdoor = tasks_df[tasks_df['檢查類型'] != '內掃教室']

    # --- 第一部分：內掃教室 ---
    if not df_indoor.empty:
        add_task_section(doc, df_indoor, standards_grouped, f"{display_name} - 內掃教室")
    
    # --- 分頁 ---
    if not df_indoor.empty and not df_outdoor.empty:
        doc.add_page_break()
    
    # --- 第二部分：外掃區 ---
    if not df_outdoor.empty:
        add_task_section(doc, df_outdoor, standards_grouped, f"{display_name} - 外掃區域")

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
    
    if " - " in selected_option:
        current_display_name = selected_option.split(" - ")[-1]
    else:
        current_display_name = selected_option

    st.info(f"👋 歡迎 **{current_display_name}**")
    
    my_tasks = df_tasks[df_tasks['負責班級'] == current_class_id]
    
    # --- Word 下載按鈕 ---
    if not my_tasks.empty:
        st.markdown("### 🖨️ 紙本檢核表下載")
        st.write("點擊下方按鈕下載 Word 檔。檔案已自動分為「內掃」與「外掃」兩頁。")
        
        doc = generate_docx(current_display_name, my_tasks, df_standards)
        bio = io.BytesIO()
        doc.save(bio)
        
        st.download_button(
            label="📥 下載 Word 檢核表 (.docx)",
            data=bio.getvalue(),
            file_name=f"{current_display_name}_大掃除檢核表.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        st.markdown("---")

    # --- 數位預覽區 ---
    st.markdown("### 📱 數位預覽")
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
