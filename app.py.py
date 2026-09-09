import io
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import streamlit as st


# ==========================================
# 1. Word 排版與表格輔助函式
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
    """將 DataFrame 寫入 Word 並呈現美化表格"""
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

    # 填入內容資料列
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if pd.notna(val) else ''
            for p in row_cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9.5)

    doc.add_paragraph()


# ==========================================
# 2. 自動計算輔助函式 (時效平均值與統計)
# ==========================================
def calculate_efficiency_stats(df, days_limit_name='三天'):
    """自動計算時效平均值與總筆數統計"""
    if df is None or df.empty:
        return None, f'無{days_limit_name}時效資料'

    # 搜尋可能包含天數數值的欄位
    num_cols = df.select_dtypes(include=['number']).columns
    if len(num_cols) > 0:
        target_col = num_cols[0]
        avg_days = round(df[target_col].mean(), 2)
        total_count = df[target_col].count()
        summary_text = (
            f'{days_limit_name}時效分析：總計評估 {total_count} 案，平均耗時 '
            f'**{avg_days}** 天。'
        )

        # 建立附帶平均值的統計表
        df_summary = df.copy()
        # 於下方新增一列顯示平均值
        summary_row = {col: '' for col in df.columns}
        summary_row[df.columns[0]] = '平均值 / 總計'
        summary_row[target_col] = f'平均 {avg_days} 天 (共 {total_count} 案)'
        df_summary = pd.concat(
            [df_summary, pd.DataFrame([summary_row])], ignore_index=True
        )
        return df_summary, summary_text
    else:
        return df, f'{days_limit_name}時效：資料未包含數值天數欄位'


# ==========================================
# 3. 生成 Word 報告主 Logic
# ==========================================
def build_word_report(
    df_3days,
    text_3days,
    df_5days,
    text_5days,
    df_close,
    df_close_summary,
    df_diversity,
    df_diversity_summary,
    df_region,
):
    doc = docx.Document()

    # 主標題
    doc.add_heading('長照 A 單位每月營運與服務品質統計報告', level=1)
    doc.add_paragraph(
        '本報告由 Streamlit 自動化系統讀取「多元總表」及各時效工作表演算生成。'
    )

    # 1. 結案數與類別統計
    if df_close is not None and not df_close.empty:
        doc.add_heading('一、 結案與派案統計', level=2)
        if df_close_summary is not None and not df_close_summary.empty:
            doc.add_paragraph('【結案類別與原因統計表】')
            add_df_to_word(doc, df_close_summary)
        doc.add_paragraph('【結案明細資料】')
        add_df_to_word(doc, df_close)

    # 2. 服務時效追蹤與平均值 (3天 & 5天)
    doc.add_heading('二、 服務時效統計與平均值分析', level=2)
    if df_3days is not None and not df_3days.empty:
        doc.add_paragraph(f'1. {text_3days}')
        add_df_to_word(doc, df_3days, '3 天時效明細與平均')

    if df_5days is not None and not df_5days.empty:
        doc.add_paragraph(f'2. {text_5days}')
        add_df_to_word(doc, df_5days, '5 天時效明細與平均')

    # 3. 多元服務數量與類別統計
    if df_diversity is not None and not df_diversity.empty:
        doc.add_heading('三、 多元服務數量與類別統計', level=2)
        if df_diversity_summary is not None and not df_diversity_summary.empty:
            doc.add_paragraph('【多元服務類別項目彙整】')
            add_df_to_word(doc, df_diversity_summary)
        doc.add_paragraph('【多元總表原始資料】')
        add_df_to_word(doc, df_diversity)

    # 4. 服務區域與個管師交叉統計表
    if df_region is not None and not df_region.empty:
        add_df_to_word(
            doc, df_region, '四、 服務區域與個管師案量交叉統計表'
        )

    # 寫入記憶體 Buffer
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 4. Streamlit 介面與資料統計處理
# ==========================================
st.set_page_config(
    page_title='長照 A 單位自動化報表系統', layout='wide'
)
st.title('📊 長照 A 單位每月報表自動統計與 Word 匯出系統')

st.markdown("""
請上傳您的月統計 Excel 檔，系統會自動對接工作表並計算平均值與各類別統計：
* **`多元總表`** ➔ 自動計算 **服務區域與個管統計** + **多元服務類別統計**
* **`三天時效` / `五天時效`** ➔ 自動計算 **平均天數 (平均值)**
* **`結案`** ➔ 自動進行 **結案原因與類別統計**
""")

uploaded_file = st.file_uploader(
    '請上傳 Excel 統計檔案 (.xlsx)', type=['xlsx', 'xls']
)

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # 預先宣告變數
        df_3days = None
        df_5days = None
        df_close = None
        df_diversity = None
        df_region = None
        df_close_summary = None
        df_diversity_summary = None

        # ----------------------------------------------------
        # A. 讀取「結案」並做類別統計
        # ----------------------------------------------------
        if '結案' in sheet_names:
            df_close = pd.read_excel(xls, '結案')
            # 尋找類別或原因欄位做統計
            possible_cat_cols = [
                c
                for c in df_close.columns
                if '原因' in c or '類別' in c or '狀態' in c
            ]
            if possible_cat_cols:
                cat_col = possible_cat_cols[0]
                df_close_summary = (
                    df_close[cat_col]
                    .value_counts()
                    .reset_index()
                    .rename(columns={'index': cat_col, cat_col: '結案人次'})
                )

        # ----------------------------------------------------
        # B. 讀取「三天時效」與「五天時效」並計算平均值
        # ----------------------------------------------------
        df_3days_raw = (
            pd.read_excel(xls, '三天時效') if '三天時效' in sheet_names else None
        )
        df_3days, text_3days = calculate_efficiency_stats(
            df_3days_raw, '3天'
        )

        df_5days_raw = (
            pd.read_excel(xls, '五天時效') if '五天時效' in sheet_names else None
        )
        df_5days, text_5days = calculate_efficiency_stats(
            df_5days_raw, '5天'
        )

        # ----------------------------------------------------
        # C. 讀取「多元總表」（替代原本的總表）
        # ----------------------------------------------------
        if '多元總表' in sheet_names:
            df_diversity = pd.read_excel(xls, '多元總表')

            # 1. 區域與個管師樞紐統計 (居住地 vs A個管)
            if {'居住地', 'A個管'}.issubset(df_diversity.columns):
                df_region = pd.crosstab(
                    df_diversity['居住地'],
                    df_diversity['A個管'],
                    margins=True,
                    margins_name='總計',
                ).reset_index()

            # 2. 各類別服務統計 (若有多元服務項目欄位)
            div_cols = [
                c
                for c in df_diversity.columns
                if '項目' in c or '類別' in c or '服務' in c
            ]
            if div_cols:
                df_diversity_summary = (
                    df_diversity[div_cols[0]]
                    .value_counts()
                    .reset_index()
                    .rename(
                        columns={
                            'index': div_cols[0],
                            div_cols[0]: '使用人次統計',
                        }
                    )
                )

        st.success('✅ 檔案讀取與自動統計計算完成！')

        # 顯示平均值重點指標 (KPI)
        col1, col2 = st.columns(2)
        with col1:
            st.info(f'📌 {text_3days}')
        with col2:
            st.info(f'📌 {text_5days}')

        # 分頁預覽
        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            [
                '結案與類別統計',
                '三天時效(含平均)',
                '五天時效(含平均)',
                '多元總表類別統計',
                '區域與個管統計(多元總表計算)',
            ]
        )

        with tab1:
            if df_close_summary is not None:
                st.write('**結案類別人次統計：**')
                st.dataframe(df_close_summary, use_container_width=True)
            if df_close is not None:
                st.write('**結案明細：**')
                st.dataframe(df_close, use_container_width=True)

        with tab2:
            if df_3days is not None:
                st.dataframe(df_3days, use_container_width=True)

        with tab3:
            if df_5days is not None:
                st.dataframe(df_5days, use_container_width=True)

        with tab4:
            if df_diversity_summary is not None:
                st.write('**多元服務類別項目統計：**')
                st.dataframe(df_diversity_summary, use_container_width=True)
            if df_diversity is not None:
                st.write('**多元總表原始資料：**')
                st.dataframe(df_diversity, use_container_width=True)

        with tab5:
            if df_region is not None:
                st.dataframe(df_region, use_container_width=True)
            else:
                st.warning(
                    '⚠️ 多元總表中找不到「居住地」或「A個管」欄位，無法計算區域交叉表。'
                )

        st.markdown('---')

        # 一鍵生成 Word 檔案
        word_bytes = build_word_report(
            df_3days,
            text_3days,
            df_5days,
            text_5days,
            df_close,
            df_close_summary,
            df_diversity,
            df_diversity_summary,
            df_region,
        )

        st.download_button(
            label='📥 下載當月完整 Word 報告 (.docx)',
            data=word_bytes,
            file_name='長照A單位_月報統計與平均值分析報告.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

    except Exception as e:
        st.error(f'資料處理時發生錯誤：{e}')

else:
    st.info('👈 請上傳包含「多元總表、三天時效、五天時效、結案」的 Excel 檔案。')
