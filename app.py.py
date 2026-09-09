import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st


# 設定表格邊框與美化輔助函式
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def build_table(doc, df_data, title=''):
    if title:
        doc.add_heading(title, level=2)

    # 建立表格
    table = doc.add_table(rows=1, cols=len(df_data.columns))
    table.style = 'Table Grid'

    # 設定表頭
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(df_data.columns):
        hdr_cells[i].text = str(col_name)
        # 設定表頭背景為淺灰色、文字加粗
        set_cell_background(hdr_cells[i], 'EFEFEF')
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

    doc.add_paragraph()  # 空列隔開


def generate_word_report():
    doc = docx.Document()

    # 主標題
    h1 = doc.add_heading('屏東市 A 單位每月會議與服務統計報告', level=1)
    doc.add_paragraph('本報告由 Streamlit 系統自動匯出，彙整當月 A 派案、服務時效、多元服務及個管分工數據。')

    # ----------------------------------------------------
    # 1. 派案與結案數分析
    # ----------------------------------------------------
    df_dispatch = pd.DataFrame({
        '月份': ['7月', '8月', '9月', '10月', '11月', '12月'],
        '新進案量': [16, '', '', '', '', ''],
        '結案數': [75, '', '', '', '', ''],
    })
    build_table(doc, df_dispatch, '一、 派案與結案統計')

    # ----------------------------------------------------
    # 2. 服務時效分析（包含 3 天與 5 天時效）
    # ----------------------------------------------------
    df_efficiency = pd.DataFrame({
        '月份': ['7月', '8月', '9月', '10月', '11月', '12月'],
        '訪案及計畫擬定平均天數 (≦3天)': [2.12, '', '', '', '', ''],
        '照專簽審後至照會單位平均天數 (≦2天)': [1.37, '', '', '', '', ''],
        '照會後第1次服務輸送到達平均天數 (≦5天)': [5.34, '', '', '', '', ''],
        '未符合時效內原因說明': ['配合案家時間(9)\n住院(2)', '', '', '', '', ''],
    })
    build_table(doc, df_efficiency, '二、 服務時效追蹤（3天及5天時效）')

    # ----------------------------------------------------
    # 3. 多元服務數量追蹤
    # ----------------------------------------------------
    df_diversity = pd.DataFrame({
        '多元服務量': [
            '0項服務',
            '1項服務',
            '2項服務',
            '3項服務',
            '4項服務',
            '5項(含)以上服務',
            '總計(人數)',
            '占比(%)',
        ],
        '7月人數/數據': [
            '7',
            '507',
            '793',
            '580',
            '239',
            '88',
            '2214',
            '76.78%',
        ],
        '8月': ['', '', '', '', '', '', '', ''],
        '9月': ['', '', '', '', '', '', '', ''],
        '10月': ['', '', '', '', '', '', '', ''],
        '11月': ['', '', '', '', '', '', '', ''],
        '12月': ['', '', '', '', '', '', '', ''],
    })
    build_table(doc, df_diversity, '三、 多元服務數量統計')

    # ----------------------------------------------------
    # 4. 服務區域與個管師案量交叉分析表
    # ----------------------------------------------------
    df_region = pd.DataFrame({
        '區域 / 個管': [
            '屏東市',
            '萬丹鄉',
            '長治鄉',
            '麟洛鄉',
            '鹽埔鄉',
            '高樹鄉',
            '九如鄉',
            '里港鄉',
            '內埔鄉',
            '潮州鎮',
            '萬巒鄉',
            '竹田鄉',
            '總計',
        ],
        '金菊': [37, 0, 90, 0, 0, 0, 0, 0, 0, 0, 0, 0, 127],
        '怡然': [66, 35, 18, 0, 0, 0, 0, 0, 0, 0, 0, 0, 119],
        '依凡': [119, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 119],
        '文琪': [116, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 116],
        '心慧': [116, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 116],
        '謹誼': [115, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 115],
        '秀珍': [37, 0, 0, 78, 0, 0, 0, 0, 0, 0, 0, 0, 115],
        '讚美': [113, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 113],
        '區域總計': [
            738,
            249,
            148,
            78,
            158,
            162,
            89,
            63,
            218,
            166,
            74,
            71,
            2214,
        ],
    })
    build_table(doc, df_region, '四、 服務區域與個管個案統計')

    # 將生成的 Word 檔寫入記憶體
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# --- Streamlit 畫面配置 ---
st.title('長照 A 單位報表生成與匯出系統')
st.markdown(
    '點擊下方按鈕，即可將包含 **結案數、3天與5天時效、多元服務數量、服務區域與個管表格** 的完整數據自動生成為 Word 報告檔！'
)

# 生成 Word 檔案
word_file = generate_word_report()

# 下載按鈕
st.download_button(
    label='📥 下載完整 Word 報告 (.docx)',
    data=word_file,
    file_name='屏東市A單位會議紀錄_統計分析報告.docx',
    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
)
