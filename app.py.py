import io
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import streamlit as st


# ==========================================
# 1. Word 表格排版輔助函式
# ==========================================
def set_cell_background(cell, fill_hex):
    """設定 Word 表格儲存格背景顏色"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def add_df_to_word(doc, df_data, title=''):
    """將 Pandas DataFrame 轉繪為 Word 格式化表格"""
    if df_data is None or df_data.empty:
        return

    if title:
        doc.add_heading(title, level=2)

    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 設定表頭 (灰底 + 粗體)
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)
        set_cell_background(hdr_cells[i], 'EFEFEF')
        for p in hdr_cells[i].paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.size = Pt(10)

    # 填入表格資料列
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if pd.notna(val) else ''
            for p in row_cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9.5)

    doc.add_paragraph()  # 段落空行


# ==========================================
# 2. 生成 Word 報告主 logic
# ==========================================
def build_word_report(df_3days, df_5days, df_close, df_diversity, df_region):
    doc = docx.Document()

    # 報告主標題
    doc.add_heading('長照 A 單位每月營運與服務品質統計報告', level=1)
    doc.add_paragraph('本報告由 Streamlit 自動化系統讀取最新 Excel 數據匯出。')

    # 1. 結案數與派案統計
    if df_close is not None and not df_close.empty:
        add_df_to_word(doc, df_close, '一、 結案與派案統計')

    # 2. 服務時效追蹤 (3 天與 5 天)
    if df_3days is not None and not df_3days.empty:
        add_df_to_word(doc, df_3days, '二、 3 天服務時效統計')

    if df_5days is not None and not df_5days.empty:
        add_df_to_word(doc, df_5days, '三、 5 天服務時效統計')

    # 3. 多元服務數量
    if df_diversity is not None and not df_diversity.empty:
        add_df_to_word(doc, df_diversity, '四、 多元服務數量統計')

    # 4. 服務區域與個管師統計
    if df_region is not None and not df_region.empty:
        add_df_to_word(doc, df_region, '五、 服務區域與個管師案量統計')

    # 將 Word 檔案寫入記憶體 Buffer
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 3. Streamlit 畫面配置與資料處理
# ==========================================
st.set_page_config(
    page_title='長照 A 單位自動化報表系統', layout='wide'
)
st.title('📊 長照 A 單位每月報表自動統計與 Word 匯出系統')

st.markdown("""
請上傳您的月統計 Excel 檔，系統會自動對接工作表：
* **`結案`** ➔ 結案數與派案分析
* **`三天時效` / `五天時效`** ➔ 服務時效追蹤
* **`多元總表`** ➔ 多元服務數量統計
* **`總表`** ➔ 自動計算「居住地」與「A個管」區域交叉表
""")

# 檔案上傳元件
uploaded_file = st.file_uploader(
    '請上傳 Excel 統計檔案 (.xlsx)', type=['xlsx', 'xls']
)

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # 預先宣告變數 (防止 NameError)
        df_3days = None
        df_5days = None
        df_close = None
        df_diversity = None
        df_region = None

        # 1. 讀取「三天時效」
        if '三天時效' in sheet_names:
            df_3days = pd.read_excel(xls, '三天時效')

        # 2. 讀取「五天時效」
        if '五天時效' in sheet_names:
            df_5days = pd.read_excel(xls, '五天時效')

        # 3. 讀取「結案」
        if '結案' in sheet_names:
            df_close = pd.read_excel(xls, '結案')

        # 4. 讀取「多元總表」
        if '多元總表' in sheet_names:
            df_diversity = pd.read_excel(xls, '多元總表')

        # 5. 從「總表」自動算區域與個管交叉表
        if '總表' in sheet_names:
            df_raw = pd.read_excel(xls, '總表')
            # 自動判定是否包含欄位「居住地」與「A個管」
            if {'居住地', 'A個管'}.issubset(df_raw.columns):
                df_region = pd.crosstab(
                    df_raw['居住地'],
                    df_raw['A個管'],
                    margins=True,
                    margins_name='總計',
                ).reset_index()

        st.success('✅ 檔案讀取成功！資料預覽如下：')

        # 分頁顯示畫面預覽
        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            [
                '結案統計',
                '三天時效',
                '五天時效',
                '多元總表',
                '區域與個管(自動計算)',
            ]
        )

        with tab1:
            if df_close is not None:
                st.dataframe(df_close, use_container_width=True)
            else:
                st.info('未找到「結案」工作表')

        with tab2:
            if df_3days is not None:
                st.dataframe(df_3days, use_container_width=True)
            else:
                st.info('未找到「三天時效」工作表')

        with tab3:
            if df_5days is not None:
                st.dataframe(df_5days, use_container_width=True)
            else:
                st.info('未找到「五天時效」工作表')

        with tab4:
            if df_diversity is not None:
                st.dataframe(df_diversity, use_container_width=True)
            else:
                st.info('未找到「多元總表」工作表')

        with tab5:
            if df_region is not None:
                st.dataframe(df_region, use_container_width=True)
            else:
                st.info('未找到「總表」或缺乏「居住地/A個管」欄位')

        st.markdown('---')

        # 一鍵生成 Word 檔案
        word_bytes = build_word_report(
            df_3days, df_5days, df_close, df_diversity, df_region
        )

        st.download_button(
            label='📥 下載完整 Word 報告 (.docx)',
            data=word_bytes,
            file_name='長照A單位_月報統計報告.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

    except Exception as e:
        st.error(f'資料處理時發生錯誤：{e}')

else:
    st.info('👈 請在上方上傳 Excel 檔案以開始分析。')
