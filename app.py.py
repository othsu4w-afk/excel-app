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
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names  # 取得所有工作表名稱

        # --- 1. 預先宣告變數 (防止 NameError) ---
        df_dispatch = pd.DataFrame()
        df_efficiency = pd.DataFrame()
        df_diversity = pd.DataFrame()
        df_region = None

        # --- 2. 讀取「結案 / 派案」資料 ---
        if '結案' in sheet_names:
            df_dispatch = pd.read_excel(xls, '結案')
        elif '派案統計' in sheet_names:
            df_dispatch = pd.read_excel(xls, '派案統計')

        # --- 3. 讀取「服務時效」資料 ---
        if '三天時效' in sheet_names:
            df_efficiency = pd.read_excel(xls, '三天時效')
        elif '時效統計' in sheet_names:
            df_efficiency = pd.read_excel(xls, '時效統計')

        # --- 4. 讀取「多元服務」資料 ---
        if '多元總表' in sheet_names:
            df_diversity = pd.read_excel(xls, '多元總表')
        elif '多元數量' in sheet_names:
            df_diversity = pd.read_excel(xls, '多元數量')

        # --- 5. 計算「區域與個管」交叉表 (從總表) ---
        if '總表' in sheet_names:
            df_raw = pd.read_excel(xls, '總表')
            if {'居住地', 'A個管'}.issubset(df_raw.columns):
                df_region = pd.crosstab(
                    df_raw['居住地'],
                    df_raw['A個管'],
                    margins=True,
                    margins_name='總計',
                ).reset_index()

        st.success('✅ 檔案讀取成功！')

        # --- 6. 分頁預覽 (加入條件判斷) ---
        tab1, tab2, tab3, tab4 = st.tabs(
            ['派案與結案', '服務時效', '多元服務數量', '區域與個管']
        )

        with tab1:
            if not df_dispatch.empty:
                st.dataframe(df_dispatch, use_container_width=True)
            else:
                st.info('尚無派案/結案資料')

        with tab2:
            if not df_efficiency.empty:
                st.dataframe(df_efficiency, use_container_width=True)
            else:
                st.info('尚無時效資料')

        with tab3:
            if not df_diversity.empty:
                st.dataframe(df_diversity, use_container_width=True)
            else:
                st.info('尚無多元服務資料')

        with tab4:
            if df_region is not None:
                st.dataframe(df_region, use_container_width=True)
            else:
                st.info('尚無區域與個管資料')

        # --- 7. 下載 Word 按鈕 (僅傳入有資料的表格) ---
        st.markdown('---')
        word_bytes = build_word_report(
            df_dispatch, df_efficiency, df_diversity, df_region
        )

        st.download_button(
            label='📥 下載當月 Word 報告 (.docx)',
            data=word_bytes,
            file_name='長照A單位_每月統計報告.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

    except Exception as e:
        st.error(f'資料處理時發生錯誤：{e}')
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
