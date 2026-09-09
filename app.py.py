import re
import pandas as pd


# 1. 強化版字串清理函式：將全形轉半形、移除所有換行與空格，並轉大寫
def clean_text_advanced(text):
    if pd.isna(text):
        return ''
    text = str(text)
    # 全形轉半形
    text = (
        text.replace('（', '(')
        .replace('）', ')')
        .replace('－', '-')
        .replace('—', '-')
    )
    # 移除所有空白、換行、Tab
    text = re.sub(r'[\r\n\t\s]+', '', text)
    return text.upper()


# 2. 智慧歸類函式：只要包含關鍵字就歸納，不刪除任何資料
def map_to_new_case_category(cleaned_val):
    if not cleaned_val:
        return None

    if 'A開發' in cleaned_val:
        return '新-A開發'
    elif 'B開發' in cleaned_val:
        return '新-B開發'
    elif '舊案下放' in cleaned_val:
        return '新-舊案下放'
    elif '非複評舊案' in cleaned_val and '非複評舊案-照管中心' in cleaned_val:
        return '非複評舊案-照管中心(無時效)'
    elif '新-照管中心' in cleaned_val or '照管' in cleaned_val:
        return '新-照管中心'
    elif '非複評舊案-A轉A' in cleaned_val or 'A轉A' in cleaned_val:
        return '非複評舊案-A轉A(無時效)'
    elif '輔具' in cleaned_val:
        return '純輔具(無時效)'
    elif '出服' in cleaned_val:
        return '出服'in cleaned_val or '出備' in cleaned_val:

    return None


# 3. 修正後的 process_new_cases 函式
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

    # 優先從三天時效抓，如果沒有則從多元總表抓
    df_source = None
    if df_3days_raw is not None and not df_3days_raw.empty:
        df_source = df_3days_raw.copy()
    elif df_diversity is not None and not df_diversity.empty:
        df_source = df_diversity.copy()

    if df_source is None or df_source.empty:
        return None, None

    # 自動偵測 類別/來源 欄位
    type_col = next(
        (
            c
            for c in df_source.columns
            if any(
                k in str(c) for k in ['類別', '項目', '來源', '類型', '新案']
            )
        ),
        None,
    )
    if not type_col:
        # 如果沒找到，嘗試看第一欄
        type_col = df_source.columns[0]

    # 對資料進行清理與模糊歸類
    df_source['_clean_val'] = df_source[type_col].apply(clean_text_advanced)
    df_source['_mapped_category'] = df_source['_clean_val'].apply(
        map_to_new_case_category
    )

    # 篩選出屬於新案 8 類別的資料
    df_new_filtered = df_source[
        df_source['_mapped_category'].notna()
    ].copy()

    if df_new_filtered.empty:
        return None, None

    # 彙整圖 1 新案來源統計
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

    # 彙整圖 2：區域 x 個管師交叉統計 (放寬欄位搜尋邏輯)
    region_col = next(
        (
            c
            for c in df_new_filtered.columns
            if any(
                k in str(c)
                for k in [
                    '居住',
                    '區域',
                    '鄉鎮',
                    '區',
                    '鄉',
                    '鎮',
                    '市',
                ]
            )
        ),
        None,
    )
    manager_col = next(
        (
            c
            for c in df_new_filtered.columns
            if any(
                k in str(c) for k in ['個管', '主責', '負責', '人員']
            )
        ),
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

        # 排序：從大(左) -> 小(右)
        total_row_ct = ct.loc['總計']
        manager_cols = [c for c in ct.columns if c != '總計']
        sorted_managers = (
            total_row_ct[manager_cols].sort_values(ascending=False).index.tolist()
        )
        final_cols = sorted_managers + ['總計']

        df_new_region = ct[final_cols].reset_index()
        df_new_region.rename(columns={region_col: '列標籤'}, inplace=True)

    return df_new_summary, df_new_region
