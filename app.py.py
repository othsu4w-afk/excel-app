import io
import docx
import pandas as pd
import streamlit as st

# ==========================================
# 1. 輔助函式區
# ==========================================
def sanitize_dataframe_for_streamlit(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return None
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).replace('nan', '')
    return df_clean

def add_df_to_word(doc, df_data, title=""):
    if df_data is None or not isinstance(df_data, pd.DataFrame) or df_data.empty:
        return
    if title:
        doc.add_heading(title, level=2)
    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = "" if pd.isna(val) else str(val)
    doc.add_paragraph()

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
# 2. 主程式
# ==========================================
st.title("Excel 自動化報表產製系統")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file:
    # 預先初始化變數
    df_dispatch = None
    df_efficiency = None
    df_diversity = None
    df_region = None

    try:
        # 讀取 Excel 檔案的所有 Sheet
        xls = pd.ExcelFile(uploaded_file)
        
        # ---------------------------------------------------------
        # 請在此處呼叫您原有的 Excel 資料處理邏輯/計算函式：
        # 範例：
        # df_dispatch = process_dispatch(pd.read_excel(xls, '派案'))
        # df_efficiency = process_efficiency(pd.read_excel(xls, '時效'))
        # df_diversity = process_diversity(pd.read_excel(xls, '多元'))
        # df_region = process_region(pd.read_excel(xls, '區域'))
        # ---------------------------------------------------------

    except Exception as e:
        st.error(f"Excel 資料讀取或計算失敗：{str(e)}")

    # --- 顯示預覽區塊 ---
    st.markdown("## 📊 統計結果預覽")

    has_data = False
    
    if df_dispatch is not None:
        st.subheader("一、 派案統計")
        st.dataframe(sanitize_dataframe_for_streamlit(df_dispatch), use_container_width=True)
        has_data = True

    if df_efficiency is not None:
        st.subheader("二、 服務時效追蹤")
        st.dataframe(sanitize_dataframe_for_streamlit(df_efficiency), use_container_width=True)
        has_data = True

    if df_diversity is not None:
        st.subheader("三、 多元服務數量追蹤")
        st.dataframe(sanitize_dataframe_for_streamlit(df_diversity), use_container_width=True)
        has_data = True

    if df_region is not None:
        st.subheader("四、 服務區域與個管師案量統計")
        st.dataframe(sanitize_dataframe_for_streamlit(df_region), use_container_width=True)
        has_data = True

    if not has_data:
        st.warning("⚠️ Excel 讀取成功，但尚未產生計算結果。請確認資料處理函式是否已將結果賦值給變數。")

    # --- 下載按鈕 ---
    st.markdown('---')
    try:
        word_bytes = build_word_report(df_dispatch, df_efficiency, df_diversity, df_region)
        st.download_button(
            label="📝 下載 Word 統計報告",
            data=word_bytes,
            file_name="服務統計分析報告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"報告生成時發生錯誤：{str(e)}")
