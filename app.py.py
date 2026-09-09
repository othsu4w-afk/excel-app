import io
import docx
from docx.shared import Pt, RGBColor
import pandas as pd
import streamlit as st

# ==========================================
# 輔助函式：將 DataFrame 安全寫入 Word
# ==========================================
def add_df_to_word(doc, df_data, title=""):
    # 防護：如果資料為 None 或 Empty 則不處理
    if df_data is None or df_data.empty:
        return

    if title:
        doc.add_heading(title, level=2)

    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 設定表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)

    # 填充內容
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = "" if pd.isna(val) else str(val)

    doc.add_paragraph()  # 空行分隔


# ==========================================
# 輔助函式：產生 Word 報告
# ==========================================
def build_word_report(df_dispatch, df_efficiency, df_diversity, df_region):
    doc = docx.Document()
    doc.add_heading('服務統計分析報告', level=1)

    # 只寫入非 None 且非空的表格
    add_df_to_word(doc, df_dispatch, '一、 派案統計')
    add_df_to_word(doc, df_efficiency, '二、 服務時效追蹤（含 3 天及 5 天時效）')
    add_df_to_word(doc, df_diversity, '三、 多元服務數量追蹤')
    add_df_to_word(doc, df_region, '四、 服務區域與個管師案量統計')

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# ==========================================
# 輔助函式：修復 PyArrow 轉型問題
# ==========================================
def sanitize_dataframe_for_streamlit(df):
    """將容易出錯的欄位或所有 object 欄位轉為字串，防止 PyArrow 崩潰"""
    if df is None or df.empty:
        return df
    
    df_clean = df.copy()
    for col in df_clean.columns:
        # 將 object 欄位及含有換行/混合型態的欄位強制轉為字串
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).replace('nan', '')
    return df_clean


# ==========================================
# 主程式邏輯 (範例結構)
# ==========================================
st.title("Excel 自動化報表產製系統")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    # 讀取 Excel (請根據您原有的讀取邏輯呈現)
    # ... 您的資料處理邏輯 ...
    
    # 假設計算完畢後的 DataFrames 如下：
    # df_dispatch, df_efficiency, df_diversity, df_region
    
    # 【範例展示】渲染至 Streamlit 前先清理資料型態
    # st.dataframe(sanitize_dataframe_for_streamlit(df_dispatch))

    # --- 下載 Word 按鈕（加強防護機制） ---
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
