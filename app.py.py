import io
import docx
import pandas as pd
import streamlit as st

# ==========================================
# 1. 頁面標題與檔案上傳
# ==========================================
st.title("Excel 自動化報表產製系統")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xls"])

# 初始化變數，確保就算沒算出來也不會報名詞未定義錯誤
df_dispatch, df_efficiency, df_diversity, df_region = None, None, None, None

if uploaded_file:
    # 讀取 Excel 分頁 (根據您實際的檔案工作表調整)
    xls = pd.ExcelFile(uploaded_file)
    
    # 假設您有各個對應的處理函式，在這裡進行資料處理與賦值：
    # df_dispatch = process_dispatch_sheet(pd.read_excel(xls, '派案'))
    # df_efficiency = process_efficiency_sheet(pd.read_excel(xls, '時效'))
    # df_diversity = process_diversity_sheet(pd.read_excel(xls, '多元'))
    # df_region = process_region_sheet(pd.read_excel(xls, '區域'))
    
    # ------------------------------------------
    # 2. 只有上傳檔案且資料處理完成後，才產製 Word
    # ------------------------------------------
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
