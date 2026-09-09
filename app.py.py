import io
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

    doc.add_paragraph()  # 段落空行


# ==========================================
# 2. 三天與五天時效「專業服務」重疊個案統計
# ==========================================
def calculate_overlap_stats(df_3days_raw, df_5days_raw):
    """計算三天時效(初評/複評/AA01/跨月計畫追蹤)與五天時效(專業服務)的重疊個案"""
    if (
        df_3days_raw is None
        or df_3days_raw.empty
        or df_5days_raw is None
        or df_5days_raw.empty
    ):
        return None

    # 尋找 ID 欄位
    id_col_3 = next(
        (c for c in df_3days_raw.columns if 'ID' in str(c).upper() or '身分證' in str(c) or '個案' in str(c)),
        df_3days_raw.columns[0],
    )
    id_col_5 = next(
        (c for c in df_5days_raw.columns if 'ID' in str(c).upper() or '身分證' in str(c) or '個案' in str(c)),
        df_5days_raw.columns[0],
    )

    # 計算分母：三天時效 (初評、複評、AA01、跨月計畫追蹤)
    target_categories = ['初評', '複評', 'AA01', '跨月計畫追蹤']
    type_col_3 = next(
        (c for c in df_3days_raw.columns if '類別' in str(c) or '項目' in str(c) or '類型' in str(c)),
        None,
    )

    if type_col_3:
        df_3_filtered = df_3days_raw[
            df_3days_raw[type_col_3]
            .astype(str)
            .str.strip()
            .isin(target_categories)
        ]
    else:
        df_3_filtered = df_3days_raw

    denominator_ids = set(df_3_filtered[id_col_3].dropna().astype(str).str.strip().unique())
    denominator_cnt = len(denominator_ids)

    # 計算分子：五天時效中有專業服務的個案
    item_col_5 = next(
        (c for c in df_5days_raw.columns if '項目' in str(c) or '類別' in str(c) or '服務' in str(c)),
        None,
    )

    if item_col_5:
        df_5_c = df_5days_raw[
            df_5days_raw[item_col_5].astype(str).str.contains('專業服務')
        ]
    else:
        df_5_c = df_5days_raw

    c_ids = set(df_5_c[id_col_5].dropna().astype(str).str.strip().unique())

    # 比對重疊個案
    numerator_ids = denominator_ids.intersection(c_ids)
    numerator_cnt = len(numerator_ids)

    rate = (
        f'{round((numerator_cnt / denominator_cnt) * 100, 1)}%'
        if denominator_cnt > 0
        else '0%'
    )

    return pd.DataFrame([{
        '統計項目': '三天時效個案中具專業服務之個案比率',
        '分母 (三天時效:初評/複評/AA01/跨月計畫追蹤)': f'{denominator_cnt} 人',
        '分子 (五天時效有專業服務且重疊個案)': f'{numerator_cnt} 人',
        '重疊涵蓋率': rate,
    }])


# ==========================================
# 3. 多元總表進階統計邏輯
# ==========================================
def analyze_diversity_sheet(df_diversity):
    if df_diversity is None or df_diversity.empty:
        return None, None

    total_cases = len(df_diversity)
    if total_cases == 0:
        return None, None

    col_start = 13  # N欄 (Index 13)
    col_end = 31  # AE欄 (Index 30)

    actual_max_cols = len(df_diversity.columns)
    effective_end = min(col_end, actual_max_cols)

    nae_cols = []
    if actual_max_cols > col_start:
        nae_cols = list(df_diversity.columns[col_start:effective_end])

    nae_top3_df = pd.DataFrame()

    if nae_cols:
        nae_counts = {}
        for c in nae_cols:
            valid_cnt = (
                df_diversity[c]
                .astype(str)
                .str.strip()
                .replace({'': None, 'nan': None, 'None': None})
                .dropna()
                .count()
            )
            if valid_cnt > 0:
                nae_counts[c] = valid_cnt

        sorted_nae = sorted(
            nae_counts.items(), key=lambda x: x[1], reverse=True
        )[:3]

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

        service_counts_per_case = (
            df_diversity[nae_cols]
            .astype(str)
            .applymap(
                lambda x: (
                    1
                    if x.strip() and x.strip() not in ['nan', 'None']
                    else 0
                )
            )
            .sum(axis=1)
        )
        df_diversity['_service_count'] = service_counts_per_case
    else:
        df_diversity['_service_count'] = 1

    unique_service_counts = sorted(
        [c for c in df_diversity['_service_count'].unique() if c > 0]
    )

    multi_summary_dict = {}

    for s_cnt in unique_service_counts:
        sub_df = df_diversity[df_diversity['_service_count'] == s_cnt]
        sub_total = len(sub_df)

        item_freq = {}
        if nae_cols:
            for c in nae_cols:
                cnt = (
                    sub_df[c]
                    .astype(str)
                    .str.strip()
                    .replace({'': None, 'nan': None, 'None': None})
                    .dropna()
                    .count()
                )
                if cnt > 0:
                    item_freq[c] = cnt

        sorted_items = sorted(
            item_freq.items(), key=lambda x: x[1], reverse=True
        )[:3]

        formatted_lines = [f'【總人數：{sub_total}人】']
        for rank, (itm, cnt) in enumerate(sorted_items, start=1):
            p = round((cnt / sub_total) * 100, 1) if sub_total > 0 else 0
            formatted_lines.append(f'Top {rank}: {itm} ({cnt}人, {p}%)')

        multi_summary_dict[f'{s_cnt}項服務'] = '\n'.join(formatted_lines)

    multi_summary_df = pd.DataFrame([multi_summary_dict])

    if not nae_top3_df.empty:
        nae_summary_text = '【整體單項Top 3】\n' + '\n'.join([
            f"{row['名次']}: {row['服務項目 (N~AE欄)']} ({row['使用人數']}人, {row['比例']})"
            for _, row in nae_top3_df.iterrows()
        ])
        multi_summary_df['單項服務總前三名 (N~AE欄)'] = nae_summary_text

    return multi_summary_df, nae_top3_df


# ==========================================
# 4. 生成 Word 報告主邏輯
# ==========================================
def build_word_report(
    df_3days,
    text_3days,
    df_5days_summary,
    df_5days_notes,
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
        add_df_to_word(doc, df_close_summary, '一、 結案原因統計分析')

    if df_3days is not None and not df_3days.empty:
        doc.add_heading('二、 3 天服務時效統計', level=2)
        doc.add_paragraph(text_3days)
        add_df_to_word(doc, df_3days)

    doc.add_heading('三、 5 天服務時效統計 (B碼與C碼分析)', level=2)
    if df_5days_summary is not None and not df_5days_summary.empty:
        doc.add_paragraph('【5天時效分類統計表（居家、日照、家託、專業服務）】')
        add_df_to_word(doc, df_5days_summary)

    if df_5days_notes is not None and not df_5days_notes.empty:
        doc.add_paragraph('【超過 5 天個案之備註說明統整】')
        add_df_to_word(doc, df_5days_notes)

    doc.add_heading('四、 照會服務與專業服務個案交叉分析', level=2)
    if df_1day_referral is not None and not df_1day_referral.empty:
        doc.add_paragraph('【1 天內照會統計彙整】')
        add_df_to_word(doc, df_1day_referral)

    if df_2days_referral_notes is not None and not df_2days_referral_notes.empty:
        doc.add_paragraph('【照會超過 2 天個案之備註說明統整】')
        add_df_to_word(doc, df_2days_referral_notes)

    if df_overlap_result is not None and not df_overlap_result.empty:
        doc.add_paragraph('【三天時效與專業服務重疊個案統計表】')
        add_df_to_word(doc, df_overlap_result)

    doc.add_heading('五、 多元服務項目數量與熱門前 3 名統計', level=2)
    if df_multi_summary is not None and not df_multi_summary.empty:
        doc.add_paragraph('【多元服務項數 (1項/2項...) 與熱門服務前 3 名彙整表】')
        add_df_to_word(doc, df_multi_summary)

    if df_nae_top3 is not None and not df_nae_top3.empty:
        doc.add_paragraph('【N~AE 欄位單項服務前 3 名明細】')
        add_df_to_word(doc, df_nae_top3)

    if df_region is not None and not df_region.empty:
        add_df_to_word(doc, df_region, '六、 服務區域與個管師案量交叉統計表')

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ==========================================
# 5. Streamlit 介面與資料處理
# ==========================================
st.set_page_config(page_title='長照 A 單位自動化報表系統', layout='wide')
st.title('📊 長照 A 單位每月報表自動統計與 Word 匯出系統')

uploaded_file = st.file_uploader('請上傳 Excel 統計檔案 (.xlsx)', type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        df_close_summary = None
        df_3days = None
        df_3days_raw = None
        df_5days_raw = None
        text_3days = ''
        df_5days_summary = None
        df_5days_notes = None
        df_1day_referral = None
        df_2days_referral_notes = None
        df_overlap_result = None
        df_multi_summary = None
        df_nae_top3 = None
        df_region = None

        # 1. 結案原因統計 (針對新舊版 Pandas 計算相容修正)
        if '結案' in sheet_names:
            df_close = pd.read_excel(xls, '結案')
            reason_cols = [c for c in df_close.columns if '原因' in str(c) or '類別' in str(c) or '狀態' in str(c)]
            if reason_cols:
                r_col = reason_cols[0]
                vc = df_close[r_col].dropna().value_counts()
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
            # 安全轉換數字欄位
            num_cols = []
            for c in df_3days_raw.columns:
                converted = pd.to_numeric(df_3days_raw[c], errors='coerce')
                if converted.notna().sum() > 0 and ('天' in str(c) or '耗時' in str(c) or '日' in str(c)):
                    df_3days_raw[c] = converted
                    num_cols.append(c)

            if num_cols:
                day_col = num_cols[0]
                avg_3 = round(df_3days_raw[day_col].mean(), 2)
                cnt_3 = df_3days_raw[day_col].count()
                text_3days = f'三天時效：總評估 {cnt_3} 案，平均耗時 {avg_3} 天。'

                summary_row = {c: '' for c in df_3days_raw.columns}
                summary_row[df_3days_raw.columns[0]] = '平均值 / 總計'
                summary_row[day_col] = f'平均 {avg_3} 天 (共 {cnt_3} 案)'
                df_3days = pd.concat([df_3days_raw, pd.DataFrame([summary_row])], ignore_index=True)

        # 3. 五天時效與照會
        if '五天時效' in sheet_names:
            df_5days_raw = pd.read_excel(xls, '五天時效')

            item_col = next((c for c in df_5days_raw.columns if '項目' in str(c) or '類別' in str(c)), None)
            days_col = next((c for c in df_5days_raw.columns if '天數' in str(c) or '耗時' in str(c) or '五天' in str(c)), None)
            ref_col = next((c for c in df_5days_raw.columns if '照會' in str(c) or '派案天' in str(c)), None)
            note_col = next((c for c in df_5days_raw.columns if '備註' in str(c) or '原因' in str(c)), None)

            # 強制轉換天數為數值
            if days_col:
                df_5days_raw[days_col] = pd.to_numeric(df_5days_raw[days_col], errors='coerce')
            if ref_col:
                df_5days_raw[ref_col] = pd.to_numeric(df_5days_raw[ref_col], errors='coerce')

            b_items = ['居家照顧', '日間照顧', '家庭托顧']
            c_items = ['專業服務']

            if item_col and days_col:
                df_b = df_5days_raw[df_5days_raw[item_col].astype(str).isin(b_items)]
                df_c = df_5days_raw[df_5days_raw[item_col].astype(str).str.contains('專業服務')]

                stats_data = []
                for code_name, sub_df in [('B碼 (居家/日照/家託)', df_b), ('C碼 (專業服務)', df_c)]:
                    if not sub_df.empty:
                        total_cnt = len(sub_df)
                        avg_d = round(sub_df[days_col].mean(), 2)
                        over_5 = len(sub_df[sub_df[days_col] > 5])
                        stats_data.append({
                            '類別代碼': code_name,
                            '總案數': total_cnt,
                            '平均耗時(天)': avg_d if pd.notna(avg_d) else 0,
                            '超過5天案數': over_5,
                            '符合5天率': f'{round((total_cnt - over_5) / total_cnt * 100, 1)}%',
                        })
                df_5days_summary = pd.DataFrame(stats_data)

            if days_col and note_col:
                df_over_5 = df_5days_raw[df_5days_raw[days_col] > 5]
                if not df_over_5.empty:
                    cols_to_show = [c for c in [item_col, days_col, note_col] if c is not None]
                    df_5days_notes = df_over_5[cols_to_show].dropna(subset=[note_col])

            if ref_col:
                total_cases = len(df_5days_raw)
                within_1day = len(df_5days_raw[df_5days_raw[ref_col] <= 1])
                over_2days = df_5days_raw[df_5days_raw[ref_col] > 2]

                df_1day_referral = pd.DataFrame([{
                    '總個案數': total_cases,
                    '1天內完成照會數': within_1day,
                    '1天內照會達成率': f'{round(within_1day / total_cases * 100, 1)}%' if total_cases > 0 else '0%',
                    '超過2天照會數': len(over_2days),
                }])

                if note_col and not over_2days.empty:
                    cols_ref = [c for c in [item_col, ref_col, note_col] if c is not None]
                    df_2days_referral_notes = over_2days[cols_ref].dropna(subset=[note_col])

        # 4. 計算三天與五天「專業服務」重疊個案比率
        if df_3days_raw is not None and df_5days_raw is not None:
            df_overlap_result = calculate_overlap_stats(df_3days_raw, df_5days_raw)

        # 5. 多元總表進階分析
        if '多元總表' in sheet_names:
            df_diversity = pd.read_excel(xls, '多元總表')
            df_multi_summary, df_nae_top3 = analyze_diversity_sheet(df_diversity)

            if {'居住地', 'A個管'}.issubset(df_diversity.columns):
                df_region = pd.crosstab(
                    df_diversity['居住地'],
                    df_diversity['A個管'],
                    margins=True,
                    margins_name='總計',
                ).reset_index()

        st.success('✅ Excel 資料統計分析與篩選完成！')

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            '結案原因統計',
            '三天時效',
            '五天與照會時效',
            '多元服務前3名統計',
            '區域與個管統計',
        ])

        with tab1:
            st.subheader('結案原因統計表')
            if df_close_summary is not None:
                st.dataframe(df_close_summary, use_container_width=True)

        with tab2:
            st.subheader('三天時效與平均天數')
            if df_3days is not None:
                st.caption(text_3days)
                st.dataframe(df_3days, use_container_width=True)

        with tab3:
            st.subheader('五天時效與照會統計')
            if df_5days_summary is not None:
                st.dataframe(df_5days_summary, use_container_width=True)

            if df_5days_notes is not None:
                st.subheader('⚠️ 五天時效超過 5 天備註')
                st.dataframe(df_5days_notes, use_container_width=True)

            if df_1day_referral is not None:
                st.subheader('照會 1 天達成率與超過 2 天備註')
                st.dataframe(df_1day_referral, use_container_width=True)
                if df_2days_referral_notes is not None:
                    st.dataframe(df_2days_referral_notes, use_container_width=True)

            if df_overlap_result is not None:
                st.subheader('🔗 三天時效 與 五天時效(專業服務) 重疊涵蓋率')
                st.dataframe(df_overlap_result, use_container_width=True)

        with tab4:
            st.subheader('多元服務項數與前 3 名熱門項目統計')
            if df_multi_summary is not None:
                st.dataframe(df_multi_summary, use_container_width=True)

            if df_nae_top3 is not None:
                st.subheader('N~AE 欄位整體單項服務 Top 3')
                st.dataframe(df_nae_top3, use_container_width=True)

        with tab5:
            st.subheader('區域與個管師統計表')
            if df_region is not None:
                st.dataframe(df_region, use_container_width=True)

        st.markdown('---')

        word_bytes = build_word_report(
            df_3days,
            text_3days,
            df_5days_summary,
            df_5days_notes,
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
