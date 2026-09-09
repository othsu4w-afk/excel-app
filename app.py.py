import io
import docx
from docx.shared import Pt, RGBColor
import pandas as pd
import streamlit as st


def create_word_report(data_df):
    doc = docx.Document()

    # 設定標題
    title = doc.add_heading('每月 A 派案及服務統計報告', level=1)

    # 加入說明內文
    doc.add_paragraph('本報告由 Streamlit 系統自動生成，以下為最新月份之服務品質與個案統計數據：')

    # 加入表格 (例如：個案多元服務量統計)
    doc.add_heading('服務品質追蹤情形(個案多元服務量)', level=2)

    table = doc.add_table(rows=1, cols=len(data_df.columns))
    table.style = 'Table Grid'

    # 填入表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(data_df.columns):
        hdr_cells[i].text = str(col_name)

    # 填入表格資料
    for index, row in data_df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)

    # 儲存至記憶體緩衝區 (BytesIO)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- Streamlit 介面範例 ---
st.title('長照 A 單位報告生成器')

# 假設這是您在 Streamlit 處理好的 DataFrame
sample_data = {
    '多元服務量': [
        '0項服務',
        '1項服務',
        '2項服務',
        '3項服務',
        '4項服務',
        '5項(含)以上服務',
    ],
    '7月人數': [7, 507, 793, 580, 239, 88],
    '占比': ['0.32%', '22.90%', '35.82%', '26.20%', '10.79%', '3.97%'],
}
df = pd.DataFrame(sample_data)

st.write('### 預覽資料表')
st.dataframe(df)

# 下載按鈕
docx_bytes = create_word_report(df)
st.download_button(
    label='📥 下載 Word 報告檔 (.docx)',
    data=docx_bytes,
    file_name='每月A派案分析報告.docx',
    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
)
