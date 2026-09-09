import io
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import streamlit as st


# ==========================================
# 1. 新案統計處理函式 (來源統計 + 區域個管交叉排序)
# ==========================================
def process_new_cases_sheet(df_diversity):
    """計算新案：

    1. 來源依照類別 (圖一 8 種類別) 統計案數與比例
    2. 新案總案數依「區域與個管」交叉統計 (圖二)，最下排總計由左至右從大到小排列
    """
    if df_diversity is None or df_diversity.empty:
        return None, None

    df_diversity.columns = [str(c).strip() for c in df_diversity.columns]

    # 自動識別欄位名稱
    type_col = next(
        (
            c
            for c in df_diversity.columns
            if any(k in c for k in ['類別', '來源', '項目', '類型'])
        ),
        None,
    )
    region_col = next(
        (
            c
            for c in df_diversity.columns
            if any(k in c for k in ['居住', '區域', '鄉鎮', '縣市', '地址'])
        ),
        None,
    )
    manager_col = next(
        (
            c
            for c in df_diversity.columns
            if any(k in c for k in ['個管', 'A個管', '專員', '主管'])
        ),
        None,
    )

    if not type_col:
        return None, None

    # 圖一指定的新案類別列表
    new_case_categories = [
        '新-A開發',
        '新-B開發',
        '新-照管中心',
        '新-舊案下放',
        '非複評舊案-照管中心\n(無時效)',
        '非複評舊案-照管中心(無時效)',
        '非複評舊案-A轉A\n(無時效)',
        '非複評舊案-A轉A(無時效)',
        '出服',
        '純輔具\n(無時效)',
        '純輔具(無時效)',
    ]

    df_diversity['_clean_type'] = df_diversity[type_col].astype(str).str.strip()

    # 比對標記新案
    def is_new_case(val):
        val_clean = str(val).replace(' ', '').replace('\n', '')
        for target in new_case_categories:
            target_clean = target.replace(' ', '').replace('\n', '')
            if target_clean in val_clean or val_clean in target_clean:
                return True
        return False

    df_new = df_diversity[df_diversity['_clean_type'].apply(is_new_case)].copy()

    if df_new.empty:
        return None, None

    # --- A. 新案來源類別統計 (圖一) ---
    source_counts = df_new['_clean_type'].value_counts().reset_index()
    source_counts.columns = ['新案來源類別', '案數']

    total_new = len(df_new)
    source_counts['比例'] = (
        (source_counts['案數'] / total_new * 100).round(1).astype(str) + '%'
        if total_new > 0
        else '0%'
    )

    total_row_df = pd.DataFrame(
        [{'新案來源類別': '總計', '案數': total_new, '比例': '100.0%'}]
    )
    df_new_source_summary = pd.concat(
        [source_counts, total_row_df], ignore_index=True
    )

    # --- B. 新案依區域與個管統計 (圖二，總計由左至右、由大到小排序) ---
    df_new_region_manager = None
    if region_col and manager_col:
        ct = pd.crosstab(
            df_new[region_col],
            df_new[manager_col],
            margins=True,
            margins_name='總計',
        )

        # 擷取「總計」列數值並對個管進行從大到小排序
        total_row = ct.loc['總計']
        manager_cols = [c for c in ct.columns if c != '總計']

        # 從大到小排序個管欄位
        sorted_managers = (
            total_row[manager_cols].sort_values(ascending=False).index.tolist()
        )
        final_cols = sorted_managers + ['總計']

        ct_sorted = ct[final_cols].reset_index()
        ct_sorted.rename(columns={region_col: '列標籤'}, inplace=True)
        df_new_region_manager = ct_sorted

    return df_new_source_summary, df_new_region_manager


# ==========================================
# 2. Streamlit 介面展示範例
# ==========================================
# 假設 df_diversity 是讀入的多元服務或總表 DataFrame
# df_new_source, df_new_region = process_new_cases_sheet(df_diversity)

# st.subheader("圖一：新案來源類別統計")
# st.dataframe(df_new_source)

# st.subheader("圖二：新案區域與個管交叉統計 (總計由大至小排序)")
# st.dataframe(df_new_region)
