import io
import docx
import pandas as pd
import streamlit as st

# ==========================================
# 1. 輔助函式：將 DataFrame 寫入 Word
# ==========================================
def add_df_to_word(doc, df_data, title=""):
    if df_data is None or (isinstance(df_data, pd.DataFrame) and df_data.empty):
        return

    if title:
        doc.add_heading(title, level=2)

    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)

    # 內容
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = "" if pd.isna(val) else str(val)

    doc.add_paragraph()


# ==========================================
# 2. 輔助函式：產生 Word 報告 (必須放在最上方)
# ==========================================
def build_word_report(df_dispatch, df_efficiency, df_diversity, df_region):
    doc = docx.Document()
    doc.add_heading('服務統計分析報告', level=1)

    add_df_to_word(doc, df_dispatch, '一、 派案統計')
    add_df_to_word(doc, df_efficiency, '二、 服務時效追蹤')
    add_df_to_word(doc, df_diversity, '三、 多元服務數量追蹤')
    add_df_to_word(doc, df_region, '四、 服務區域與個管師案量統計')

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ==========================================
# 3. Streamlit 主畫面與檔案處理
# ==========================================
st.title("Excel 自動化報表產製系統")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    # 讀取 Excel 檔案邏輯...
    # 請在此處執行您的資料處理，並產出對應的 df_dispatch, df_efficiency, df_diversity, df_region
    
    # 範例預設（若處理失敗預設為 None）
    df_dispatch = locals().get('df_dispatch', None)
    df_efficiency = locals().get('df_efficiency', None)
    df_diversity = locals().get('df_diversity', None)
    df_region = locals().get('df_region', None)

    # --- 下載 Word 按鈕 ---
    st.markdown('---')
    try:
        word_bytes = build_word_report(
            df_dispatch, df_efficiency, df_diversity, df_region
        )
        
        st.download_button(
            label="📝 下載 Word 統計報告",
            data=word_bytes,
            file_name="服務統計分析報告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"報告生成時發生錯誤：{str(e)}")
