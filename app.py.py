import io
import re
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import streamlit as st


# ==========================================
# 1. Word 表格與美化輔助函式
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
    """將 DataFrame 轉寫為 Word 表格"""
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

    # 填入內容
    for _, row in df_data.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if pd.notna(val) else ''
            for p in row_cells[i].paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9.5)

    doc.add_paragraph()


# ==========================================
# 2. 強化版字串清理與智慧歸類函式
# ==========================================
def clean_text_advanced(text):
    """將全形轉半形、移除所有換行/空格/Tab，並轉大寫"""
    if pd.isna(text):
        return ''
    text = str(text)
    # 全形符號轉半形
    text = (
        text.replace('（', '(')
        .replace('）', ')')
        .replace('－', '-')
        .replace('—', '-')
    )
    # 移除所有空白、換行 (\r, \n)、Tab
    text = re.sub(r'[\r\n\t\s]+', '', text)
    return text.upper()


def map_to_new_case_category(cleaned_val):
    """智慧歸類：只要包含關鍵字就歸納，不遺漏任何個案"""
    if not cleaned_val:
        return None

    # 1. 優先判斷特例/組合條件 (帶有非複評舊案或A轉A)
    if '非複評' in cleaned_val and '照管' in cleaned_val:
        return '非複評舊案-照管中心(無時效)'
    elif '非複評' in cleaned_val and ('A轉A' in cleaned_val or '轉A' in cleaned_val):
        return '非複評舊案-A轉A(無時效)'
    elif 'A轉A' in cleaned_val or '轉A' in cleaned_val:
        return '非複評舊案-A轉A(無時效)'

    # 2. 判斷各項新案開發與來源
    elif 'A開發' in cleaned_val or 'A-開發' in cleaned_val:
        return '新-A開發'
    elif 'B開發' in cleaned_val or 'B-開發' in cleaned_val:
        return '新-B開發'
    elif '舊案下放' in cleaned_val or '下放' in cleaned_val:
        return '新-舊案下放'
    elif '照管' in cleaned_val:
        return '新-照管中心'
    elif '輔具' in cleaned_val:
        return '純輔具(無時效)'
    elif '出服' in cleaned_val or '出備' in cleaned_val:
        return '出服'

    return None


# ==========================================
# 3. 處理 3 天時效 (包含指定之類別 + 去重取最新)
# ==========================================
def process_3days_sheet(df_3days_raw):
    if df_3days_raw is None or df_3days_raw.empty:
        return None, None, ''

    id_col = next(
        (c for c in df_3days_raw.columns if any(k in str(c).upper() for k in ['ID', '身分證', '個案'])),
        df_3days_raw.columns[0],
    )
    type_col = next(
        (c for c in df_3days_raw.columns if any(k in str(c) for k in ['類別', '項目', '類型', '來源'])),
        None,
    )
    date_col = next(
        (c for c in df_3days_raw.columns if any(k in str(c) for k in ['日', '時間', 'DATE'])),
        None,
    )
    days_col = next(
        (c for c in df_3days_raw.columns if any(k in str(c) for k in ['天', '耗時'])),
        None,
    )

    df_filtered = df_3days_raw.copy()

    if type_col:
        target_cats = [
            '新-A開發',
            '新-B開發',
            '新-照管中心',
            '新-舊案下放',
            '出服',
            '初評',
            '複評',
            'AA01',
            '跨月追蹤',
            '跨月計畫追蹤',
        ]
        cleaned_cats = [clean_text_advanced(c) for c in target_cats]
        df_filtered['_clean_type'] = df_filtered[type_col].apply(clean_text_advanced)
        df_filtered = df_filtered[df_filtered['_clean_type'].isin(cleaned_cats)].drop(
            columns=['_clean_type']
        )

    if id_col in df_filtered.columns:
        if date_col:
            df_filtered[date_col] = pd.to_datetime(df_filtered[date_col], errors='coerce')
            df_filtered = df_filtered.sort_values(by=date_col, ascending=True)
        df_filtered = df_filtered.drop_duplicates(subset=[id_col], keep='last')

    if days_col:
        df_filtered[days_col] = pd.to_numeric(df_filtered[days_col], errors='coerce')
        avg_3 = round(df_filtered[days_col].mean(), 2)
        cnt_3 = df_filtered[days_col].count()
        text_3days = f'三天時效（初評類/複評/AA01/跨月追蹤，ID去重取最新）：共 {cnt_3} 案，平均耗時 {avg_3} 天。'

        summary_row = {c: '' for c in df_filtered.columns}
        summary_row[df_filtered.columns[0]] = '平均值 / 總計'
        summary_row[days_col] = f'平均 {avg_3} 天 (共 {cnt_3} 案)'
        df_display = pd.concat([df_filtered, pd.DataFrame([summary_row])], ignore_index=True)
    else:
        text_3days = f'三天時效符合條件案數：共 {len(df_filtered)} 案。'
        df_display = df_filtered

    return df_filtered, df_display, text_3days


# ==========================================
# 4. 處理新案統計（來源 + 區域個管由大到小排序）
# ==========================================
def process_new_cases(df_3days_raw, df_diversity):
    new_case_types = [
        '新-A開發',
        '新-B開發',
        '新-照管中心',
        '新-舊案下放',
        '非複評舊案-照管中心(無時效)',
        '非複評舊案-A轉A(無時效)',
        '出服',
        '純輔具(無時效)',
    ]

    # 優先從三天時效工作表讀取，若無則從多元總表讀取
    df_source = None
    if df_3days_raw is not None and not df_3days_raw.empty:
        df_source = df_3days_raw.copy()
    elif df_diversity is not None and not df_diversity.empty:
        df_source = df_diversity.copy()

    if df_source is None or df_source.empty:
        return None, None

    # 自動識別類別欄位
    type_col = next(
        (c for c in df_source.columns if any(k in str(c) for k in ['類別', '項目', '來源', '類型', '新案'])),
        df_source.columns[0],
    )

    # 進行清理與模糊分類
    df_source['_clean_val'] = df_source[type_col].apply(clean_text_advanced)
    df_source['_mapped_category'] = df_source['_clean_val'].apply(map_to_new_case_category)

    # 篩選出符合 8 類新案的資料
    df_new_filtered = df_source[df_source['_mapped_category'].notna()].copy()

    if df_new_filtered.empty:
        return None, None

    # 彙整 圖 1 新案來源統計
    counts = df_new_filtered['_mapped_category'].value_counts()
    total_new = len(df_new_filtered)

    summary_list = []
    for cat in new_case_types:
        cnt = counts.get(cat, 0)
        pct = f'{round((cnt / total_new) * 100, 1)}%' if total_new > 0 else '0%'
        summary_list.append({
            '新案來源類別': cat,
            '案數': cnt,
            '比例': pct,
        })

    df_new_summary = pd.DataFrame(summary_list)
    total_row = pd.DataFrame([{
        '新案來源類別': '總計',
        '案數': total_new,
        '比例': '100.0%',
    }])
    df_new_summary = pd.concat([df_new_summary, total_row], ignore_index=True)

    # 彙整 圖 2 區域 x 個管師交叉統計
    region_col = next(
        (c for c in df_new_filtered.columns if any(k in str(c) for k in ['居住', '區域', '鄉鎮', '縣市', '地址', '區', '鄉', '鎮', '市'])),
        None,
    )
    manager_col = next(
        (c for c in df_new_filtered.columns if any(k in str(c) for k in ['個管', '主責', '負責', '人員', '專員', '主管'])),
        None,
    )

    df_new_region = None
    if region_col and manager_col:
        ct = pd.crosstab(
            df_new_filtered[region_col],
            df_new_filtered[manager_col],
            margins=True,
            margins_name='總計',
        )

        # 最下方「總計」列：個管欄位由大(左)到小(右)排序，最右邊固定為「總計」
        total_row_ct = ct.loc['總計']
        manager_cols = [c for c in ct.columns if c != '總計']
        sorted_managers = total_row_ct[manager_cols].sort_values(ascending=False).index.tolist()
        final_cols = sorted_managers + ['總計']

        df_new_region = ct[final_cols].reset_index()
        df_new_region.rename(columns={region_col: '列標籤'}, inplace=True)

    return df_new_summary, df_new_region


# ==========================================
# 5. 處理五天時效與照會
# ==========================================
def process_5days_sheet(df_5days_raw):
    if df_5days_raw is None or df_5days_raw.empty:
        return None, None, None

    ref_col = next((c for c in df_5days_raw.columns if any(k in str(c) for k in ['照會', '派案'])), None)
    item_col = next((c for c in df_5days_raw.columns if any(k in str(c) for k in ['項目', '類別'])), None)
    note_col = next((c for c in df_5days_raw.columns if any(k in str(c) for k in ['備註', '原因'])), None)

    df_1day_referral = None
    df_2days_referral_notes = None

    if ref_col:
        df_5days_raw[ref_col] = pd.to_numeric(df_5days_raw[ref_col], errors='coerce')

        total_cnt = len(df_5days_raw)
        df_within_1 = df_5days_raw[df_5days_raw[ref_col] <= 1]
        cnt_within_1 = len(df_within_1)
        avg_days_within_1 = round(df_within_1[ref_col].mean(), 2) if cnt_within_1 > 0 else 0

        df_over_2 = df_5days_raw[df_5days_raw[ref_col] > 2]
        cnt_over_2 = len(df_over_2)

        df_1day_referral = pd.DataFrame([{
            '總個案數': total_cnt,
            '1天內完成照會人數': cnt_within_1,
            '1天內照會達成率': f'{round(cnt_within_1 / total_cnt * 100, 1)}%' if total_cnt > 0 else '0%',
            '1天內完成照會之平均天數': f'{avg_days_within_1} 天',
            '照會時效2天以上(>2天)人數': f'{cnt_over_2} 人',
        }])

        if note_col and not df_over_2.empty:
            cols = [c for c in [item_col, ref_col, note_col] if c is not None]
            df_2days_referral_notes = df_over_2[cols].dropna(subset=[note_col])

    return df_1day_referral, df_2days_referral_notes, df_5days_raw


# ==========================================
# 6. 重疊涵蓋率計算
# ==========================================
def calculate_overlap_stats(df_3days_filtered, df_5days_raw):
    if df_3days_filtered is None or df_3days_filtered.empty or df_5days_raw is None or df_5days_raw.empty:
        return None

    id_col_3 = next(
        (c for c in df_3days_filtered.columns if any(k in str(c).upper() for k in ['ID', '身分證', '個案'])),
        df_3days_filtered.columns[0],
    )
    id_col_5 = next(
        (c for c in df_5days_raw.columns if any(k in str(c).upper() for k in ['ID', '身分證', '個案'])),
        df_5days_raw.columns[0],
    )

    den_ids = set(df_3days_filtered[id_col_3].dropna().astype(str).str.strip().unique())
    den_cnt = len(den_ids)

    item_col_5 = next(
        (c for c in df_5days_raw.columns if any(k in str(c) for k in ['項目', '類別', '服務'])),
        None,
    )

    if item_col_5:
        df_5_c = df_5days_raw[df_5days_raw[item_col_5].astype(str).str.contains('專業服務')]
    else:
        df_5_c = df_5days_raw

    c_ids = set(df_5_c[id_col_5].dropna().astype(str).str.strip().unique())

    num_ids = den_ids.intersection(c_ids)
    num_cnt = len(num_ids)

    rate = f'{round((num_cnt / den_cnt) * 100, 1)}%' if den_cnt > 0 else '0%'

    return pd.DataFrame([{
        '分母 (三天時效個案去重人數)': f'{den_cnt} 人',
        '分子 (五天時效有專業服務且與三天時效重疊個案數)': f'{num_cnt} 人',
        '重疊涵蓋率': rate,
    }])


# ==========================================
# 7. 多元服務項數轉置與前 3 名分欄統計
# ==========================================
def analyze_diversity_sheet(df_diversity):
    if df_diversity is None or df_diversity.empty:
        return None, None

    total_cases = len(df_diversity)
    if total_cases == 0:
        return None, None

    col_start = 13  # N欄
    col_end = 31  # AE欄
    actual_max_cols = len(df_diversity.columns)
    effective_end = min(col_end, actual_max_cols)

    nae_cols = []
    if actual_max_cols > col_start:
        nae_cols = list(df_diversity.columns[col_start:effective_end])

    nae_top3_df = pd.DataFrame()

    if nae_cols:
        map_func = getattr(df_diversity[nae_cols], 'map', getattr(df_diversity[nae_cols], 'applymap', None))
        service_counts_per_case = map_func(
            lambda x: 1 if str(x).strip() and str(x).strip() not in ['nan', 'None'] else 0
        ).sum(axis=1)
        df_diversity['_service_count'] = service_counts_per_case
    else:
        df_diversity['_service_count'] = 1

    unique_counts = sorted([c for c in df_diversity['_service_count'].unique() if c > 0])

    transposed_rows = []

    for s_cnt in unique_counts:
        sub_df = df_diversity[df_diversity['_service_count'] == s_cnt]
        sub_total = len(sub_df)

        item_freq = {}
        if nae_cols:
            for c in nae_cols:
                cnt = sub_df[c].astype(str).str.strip().replace({'': None, 'nan': None, 'None': None}).dropna().count()
                if cnt > 0:
                    item_freq[c] = cnt

        sorted_items = sorted(item_freq.items(), key=lambda x: x[1], reverse=True)[:3]

        row_data = {
            '服務項數': f'{s_cnt}項服務',
            '總人數': f'{sub_total}人',
            '第 1 名': '-',
            '第 2 名': '-',
            '第 3 名': '-',
        }

        for rank, (itm, cnt) in enumerate(sorted_items, start=1):
            p = round((cnt / sub_total) * 100, 1) if sub_total > 0 else 0
            row_data[f'第 {rank} 名'] = f'{itm} ({cnt}人, {p}%)'

        transposed_rows.append(row_data)

    multi_summary_df = pd.DataFrame(transposed_rows)

    if nae_cols:
        nae_counts = {}
        for c in nae_cols:
            cnt = df_diversity[c].astype(str).str.strip().replace({'': None, 'nan': None, 'None': None}).dropna().count()
            if cnt > 0:
                nae_counts[c] = cnt
        sorted_nae = sorted(nae_counts.items(), key=lambda x: x[1], reverse=True)[:3]

        nae_list = []
        for rank, (item_name, count) in enumerate(sorted_nae, start=1):
            pct = round((count / total_cases) * 100, 1)
            nae_list.append({
                '名次': f'第 {rank} 名',
                '服務項目 (N~AE欄)': item_name,
                '使用人數': count,
                '比例': f'{pct}%',
            })
        nae_top3_df = pd.DataFrame(nae_list)

    return multi_summary_df, nae_top3_df


# ==========================================
# 8. 全區與個管師案量交叉統計 (由大到小排序)
# ==========================================
def process_region_manager_sheet(df_diversity):
    if df_diversity is None or df_diversity.empty:
        return None

    df_diversity.columns = [str(c).strip() for c in df_diversity.columns]

    region_col = next(
        (c for c in df_diversity.columns if any(k in str(c) for k in ['居住', '區域', '鄉鎮', '縣市', '地址', '區', '鄉', '鎮', '市'])),
        None,
    )
    manager_col = next(
        (c for c in df_diversity.columns if any(k in str(c) for k in ['個管', '主責', '負責', '人員', '專員', '主管'])),
        None,
    )

    if not region_col or not manager_col:
        return None

    ct = pd.crosstab(
        df_diversity[region_col],
        df_diversity[manager_col],
        margins=True,
        margins_name='總計',
    )

    total_row = ct.loc['總計']
    manager_cols = [c for c in ct.columns if c != '總計']

    sorted_managers = total_row[manager_cols].sort_values(ascending=False).index.tolist()
    final_cols = sorted_managers + ['總計']

    ct_sorted = ct[final_cols].reset_index()
    ct_sorted.rename(columns={region_col: '列標籤'}, inplace=True)

    return ct_sorted


# ==========================================
# 9. 生成 Word 報告主邏輯
# ==========================================
def build_word_report(
    df_3days,
    text_3days,
    df_new_summary,
    df_new_region,
    df_1day_referral,
    df_2days_referral_notes,
    df_overlap_result,
    df_close_summary,
    df_multi_summary,
    df_nae_top3,
    df_region,
):
    doc = docx.Document()

    doc.add_heading('長照 A 單位每月營運與服務品質統計報告', level=1)
    doc.add_paragraph('本報告由 Streamlit 自動化系統根據最新月報 Excel 數據分析生成。')

    if df_close_summary is not None and not df_close_summary.empty:
        add_df_to_word(doc, df_close_summary, '一、 結案原因統計分析 (已排除交接個案)')

    if df_new_summary is not None and not df_new_summary.empty:
        add_df_to_word(doc, df_new_summary, '二、 新案來源與類別統計 (圖1)')

    if df_new_region is not None and not df_new_region.empty:
        add_df_to_word(doc, df_new_region, '三、 新案區域與個管師統計 (圖2，總計由大至小)')

    if df_3days is not None and not df_3days.empty:
        doc.add_heading('四、 3 天服務時效統計 (去重取最新)', level=2)
        doc.add_paragraph(text_3days)
        add_df_to_word(doc, df_3days)

    doc.add_heading('五、 照會服務與專業服務個案交叉分析', level=2)
    if df_1day_referral is not None and not df_1day_referral.empty:
        doc.add_paragraph('【1 天內照會與 2 天以上時效統計彙整】')
        add_df_to_word(doc, df_1day_referral)

    if df_2days_referral_notes is not None and not df_2days_referral_notes.empty:
        doc.add_paragraph('【照會超過 2 天個案之備註說明統整】')
        add_df_to_word(doc, df_2days_referral_notes)

    if df_overlap_result is not None and not df_overlap_result.empty:
        doc.add_paragraph('【三天時效 與 五天時效(專業服務) 重疊涵蓋率】')
        add_df_to_word(doc, df_overlap_result)

    doc.add_heading('六、 多元服務項目數量與熱門前 3 名統計', level=2)
    if df_multi_summary is not None and not df_multi_summary.empty:
        doc.add_paragraph('【多元服務項數與熱門前 3 名彙整表】')
        add_df_to_word(doc, df_multi_summary)

    if df_nae_top3 is not None and not df_nae_top3.empty:
        doc.add_paragraph('【N~AE 欄位單項服務整體 Top 3】')
        add_df_to_word(doc, df_nae_top3)

    if df_region is not None and not df_region.empty:
        add_df_to_word(doc, df_region, '七、 全區總案量統計 (區域與個管師，個管由大至小排序)')

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 10. Streamlit 主介面
# ==========================================
st.set_page_config(page_title='長照 A 單位自動化報表系統', layout='wide')
st.title('📊 長照 A 單位每月報表自動統計與 Word 匯出系統')

uploaded_file = st.file_uploader('請上傳 Excel 統計檔案 (.xlsx)', type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        df_close_summary = None
        df_3days_filtered = None
        df_3days_display = None
        text_3days = ''
        df_new_summary = None
        df_new_region = None
        df_1day_referral = None
        df_2days_referral_notes = None
        df_5days_raw = None
        df_overlap_result = None
        df_multi_summary = None
        df_nae_top3 = None
        df_region = None

        # 1. 結案原因統計 (排除交接個案)
        if '結案' in sheet_names:
            df_close = pd.read_excel(xls, '結案')
            reason_cols = [c for c in df_close.columns if any(k in str(c) for k in ['原因', '類別', '狀態'])]
            if reason_cols:
                r_col = reason_cols[0]
                df_close_filtered = df_close[~df_close[r_col].astype(str).str.contains('交接')]
                vc = df_close_filtered[r_col].dropna().value_counts()
                df_close_summary = pd.DataFrame({
                    r_col: vc.index,
                    '結案人數': vc.values
                })
                total_close = df_close_summary['結案人數'].sum()
                df_close_summary['比例'] = (
                    (df_close_summary['結案人數'] / total_close * 100).round(1).astype(str) + '%'
                    if total_close > 0 else '0%'
                )

        # 2. 三天時效
        if '三天時效' in sheet_names:
            df_3days_raw = pd.read_excel(xls, '三天時效')
            df_3days_filtered, df_3days_display, text_3days = process_3days_sheet(df_3days_raw)

        # 3. 新案統計 (來源圖1 + 區域個管圖2)
        df_diversity_raw = pd.read_excel(xls, '多元總表') if '多元總表' in sheet_names else None
        df_3days_raw = pd.read_excel(xls, '三天時效') if '三天時效' in sheet_names else None
        df_new_summary, df_new_region = process_new_cases(df_3days_raw, df_diversity_raw)

        # 4. 五天時效與照會
        if '五天時效' in sheet_names:
            df_5days_raw = pd.read_excel(xls, '五天時效')
            df_1day_referral, df_2days_referral_notes, df_5days_raw = process_5days_sheet(df_5days_raw)

        # 5. 重疊涵蓋率
        if df_3days_filtered is not None and df_5days_raw is not None:
            df_overlap_result = calculate_overlap_stats(df_3days_filtered, df_5days_raw)

        # 6. 多元總表與全區總案量統計（區域與個管）
        if df_diversity_raw is not None:
            df_multi_summary, df_nae_top3 = analyze_diversity_sheet(df_diversity_raw)
            df_region = process_region_manager_sheet(df_diversity_raw)

        st.success('✅ Excel 資料統計分析完成！')

        tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
            '🆕 新案統計 (來源與區域個管)',
            '結案原因 (排除交接)',
            '三天時效 (篩選與去重)',
            '照會與重疊率',
            '多元服務 (名次分欄)',
            '全區總案量 (區域與個管)',
        ])

        with tab0:
            st.subheader('1. 新案來源統計 (自動修正換行/空格)')
            if df_new_summary is not None:
                st.dataframe(df_new_summary, use_container_width=True)
            else:
                st.info('未偵測到新案資料。')

            st.subheader('2. 新案依「區域」與「個管」統計 (總計由左至右大到小)')
            if df_new_region is not None:
                st.dataframe(df_new_region, use_container_width=True)
            else:
                st.info('未獲取到符合新案條件的區域/個管交叉數據。')

        with tab1:
            st.subheader('結案原因統計表 (已排除交接個案)')
            if df_close_summary is not None:
                st.dataframe(df_close_summary, use_container_width=True)

        with tab2:
            st.subheader('三天時效 (篩選與去重)')
            if df_3days_display is not None:
                st.caption(text_3days)
                st.dataframe(df_3days_display, use_container_width=True)

        with tab3:
            st.subheader('照會 1 天達成率、平均天數與 >2 天人數')
            if df_1day_referral is not None:
                st.dataframe(df_1day_referral, use_container_width=True)

            if df_2days_referral_notes is not None:
                st.subheader('⚠️ 照會超過 2 天備註明細')
                st.dataframe(df_2days_referral_notes, use_container_width=True)

            if df_overlap_result is not None:
                st.subheader('🔗 三天時效 與 五天時效(專業服務) 重疊涵蓋率')
                st.dataframe(df_overlap_result, use_container_width=True)

        with tab4:
            st.subheader('多元服務項數與前 3 名熱門項目 (轉置分欄)')
            if df_multi_summary is not None:
                st.dataframe(df_multi_summary, use_container_width=True)

            if df_nae_top3 is not None:
                st.subheader('N~AE 欄位整體單項服務 Top 3')
                st.dataframe(df_nae_top3, use_container_width=True)

        with tab5:
            st.subheader('全區總案量統計表 (區域與個管師，總案量由左至右從大到小)')
            if df_region is not None:
                st.dataframe(df_region, use_container_width=True)

        st.markdown('---')

        word_bytes = build_word_report(
            df_3days_display,
            text_3days,
            df_new_summary,
            df_new_region,
            df_1day_referral,
            df_2days_referral_notes,
            df_overlap_result,
            df_close_summary,
            df_multi_summary,
            df_nae_top3,
            df_region,
        )

        st.download_button(
            label='📥 下載當月精準統計 Word 報告 (.docx)',
            data=word_bytes,
            file_name='長照A單位_月報精準統計報告.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )

    except Exception as e:
        st.error(f'資料處理時發生錯誤：{e}')
