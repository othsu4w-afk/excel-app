import io
import docx
import pandas as pd
import streamlit as st

# ==========================================
# 1. 輔助函式：清理資料型態 (防止 Streamlit 渲染崩潰)
# ==========================================
def sanitize_dataframe_for_streamlit(df):
    """將容易出錯的欄位或所有 object 欄位轉為純字串，避免 PyArrow 報錯"""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return df
    
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).replace('nan', '')
    return df_clean

# ==========================================
# 2. 輔助函式：將 DataFrame 寫入 Word
# ==========================================
def add_df_to_word(doc, df_data, title=""):
    if df_data is None or not isinstance(df_data, pd.DataFrame) or df_data.empty:
        return

    if title:
        doc.add_heading(title, level=2)

    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 設定表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)

    # 填入內容
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = "" if pd.isna(val) else str(val)

    doc.add_paragraph()

# ==========================================
# 3. 輔助函式：產生 Word 報告
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
# 4. Streamlit 主程式
# ==========================================
st.title("Excel 自動化報表產製系統")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    # 這裡放您的 Excel 計算邏輯（範例變數名稱如下）
    # df_dispatch = process_dispatch_sheet(...)
    # df_efficiency = process_efficiency_sheet(...)
    # df_diversity, df_region = process_new_cases_sheet(...)

    # 取得變數（若未定義則預設為 None）
    df_dispatch = locals().get('df_dispatch', None)
    df_efficiency = locals().get('df_efficiency', None)
    df_diversity = locals().get('df_diversity', None)
    df_region = locals().get('df_region', None)

    # --- 網頁線上預覽區塊 ---
    st.markdown("## 📊 統計結果預覽")

    if df_dispatch is not None:
        st.subheader("一、 派案統計")
        st.dataframe(sanitize_dataframe_for_streamlit(df_dispatch), width='stretch')

    if df_efficiency is not None:
        st.subheader("二、 服務時效追蹤")
        st.dataframe(sanitize_dataframe_for_streamlit(df_efficiency), width='stretch')

    if df_diversity is not None:
        st.subheader("三、 多元服務數量追蹤")
        st.dataframe(sanitize_dataframe_for_streamlit(df_diversity), width='stretch')

    if df_region is not None:
        st.subheader("四、 服務區域與個管師案量統計")
        st.dataframe(sanitize_dataframe_for_streamlit(df_region), width='stretch')

    # --- 下載 Word 按鈕區塊 ---
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
