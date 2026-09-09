import io
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import streamlit as st


# --- 1. Word 表格繪製輔助函式 ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def add_df_to_word(doc, df_data, title=''):
    if title:
        doc.add_heading(title, level=2)

    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 設定表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)
        set_cell_background(hdr_cells[i], 'EFEFEF')  # 淺灰背景
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)

    # 填入內容列
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if pd.notna(val) else ''
            for p in row_cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9.5)

    doc.add_paragraph()  # 加上空行


# --- 2. 動態生成 Word 報告邏輯 ---
def build_word_report(df_dispatch, df_efficiency, df_diversity, df_region):
    doc = docx.Document()

    doc.add_heading('長照 A 單位每月營運與服務品質統計報告', level=1)
    doc.add_paragraph('本報告由 Streamlit 系統自動讀取當月資料並演算生成。')

    # 加入四個主要動態表格
    add_df_to_word(doc, df_dispatch, '一、 每月 A 派案與結案分析')
    add_df_to_word(
        doc,
        df_efficiency,
        '二、 服務時效追蹤（含 3 天及 5 天時效）',
    )
    add_df_to_word(doc, df_diversity, '三、 多元服務數量追蹤')
    add_df_to_word(doc, df_region, '四、 服務區域與個管師案量統計')

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- 3. Streamlit 介面與資料自動運算 ---
st.set_page_config(
    page_title='長照 A 單位自動化報表系統', layout='wide'
)
st.title('📊 長照 A 單位每月報表自動統計與 Word 匯出系統')

st.markdown("""
請上傳您毎個月的**原始資料 Excel 檔**，系統會自動計算：
1. **結案數與派案量**
2. **3 天與 5 天服務時效平均**
3. **多元服務項目數量**
4. **服務區域與個管師統計表**
""")

# 檔案上傳器
uploaded_file = st.file_uploader(
    '請選擇 monthly_data.xlsx 檔案上傳', type=['xlsx', 'xls']
)

if uploaded_file is not None:
    try:
        # 讀取 Excel 內不同的工作表 (Sheet)
        # 💡 您可依據您實際 Excel 的 Sheet 名稱做修改
        xls = pd.ExcelFile(uploaded_file)

        # 動態讀取資料
        df_dispatch = pd.read_excel(
            xls, '派案統計'
        )  # 包含月份、新進案量、結案數等
        df_efficiency = pd.read_excel(
            xls, '時效統計'
        )  # 包含 3天時效、5天時效、原因等
        df_diversity = pd.read_excel(
            xls, '多元數量'
        )  # 包含 0~5項多元服務人數
        df_raw_cases = pd.read_excel(xls, '個案名冊')  # 包含個管師、區域

        # --- 自動動態計算「區域與個管師樞紐表」 ---
        if {'鄉鎮區域', '負責個管'}.issubset(df_raw_cases.columns):
            df_region = pd.crosstab(
                df_raw_cases['鄉鎮區域'],
                df_raw_cases['負責個管'],
                margins=True,
                margins_name='總計',
            ).reset_index()
        else:
            df_region = pd.read_excel(
                xls, '區域與個管'
            )  # 若 Excel 已算好則直接讀取

        st.success('✅ 檔案讀取成功！以下為自動計算結果預覽：')

        # 頁面預覽頁籤
        tab1, tab2, tab3, tab4 = st.tabs(
            ['派案與結案', '服務時效', '多元服務數量', '區域與個管']
        )

        with tab1:
            st.dataframe(df_dispatch, use_container_width=True)
        with tab2:
            st.dataframe(df_efficiency, use_container_width=True)
        with tab3:
            st.dataframe(df_diversity, use_container_width=True)
        with tab4:
            st.dataframe(df_region, use_container_width=True)

        # --- 一鍵生成並下載 Word 報告 ---
        st.markdown('---')
        word_bytes = build_word_report(
            df_dispatch, df_efficiency, df_diversity, df_region
        )

        st.download_button(
            label='📥 下載當月自動產製的 Word 報告 (.docx)',
            data=word_bytes,
            file_name='長照A單位_每月統計報告(自動產製).docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

    except Exception as e:
        st.error(f'資料處理時發生錯誤，請確認 Excel Sheet 名稱或格式：{e}')

else:
    st.info('👈 請先在上方上傳最新的 Excel 資料檔以開啟自動計算與匯出功能。')
