import streamlit as st
import pandas as pd
import plotly.express as px

# 設置頁面標題與寬度
st.set_page_config(page_title="長照報表自動化統計系統", layout="wide")

st.title("📊 長照服務個案統計與時效分析系統")
st.write("自動讀取分頁數據，並針對 **服務區域** 與 **個管員** 生成交叉統計與圖表")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("請上傳 Excel 報表 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 讀取 Excel 的所有分頁
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    
    st.success(f"成功讀取檔案！包含分頁：{', '.join(sheet_names)}")
    
    # 建立分頁頁籤
    tab_list = st.tabs(sheet_names)
    
    for i, sheet_name in enumerate(sheet_names):
        with tab_list[i]:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # 清理欄位名稱（去除換行符號與多餘空格）
            df.columns = [str(col).replace('\n', '').strip() for col in df.columns]
            
            # 自動識別「個管員」欄位名稱
            case_manager_col = None
            if 'A個管' in df.columns:
                case_manager_col = 'A個管'
            elif 'Unnamed: 0' in df.columns:
                case_manager_col = 'Unnamed: 0'
                
            # 自動識別「服務區域/居住地」欄位名稱
            area_col = None
            if '居住地(下拉)' in df.columns:
                area_col = '居住地(下拉)'
            elif '居住地' in df.columns:
                area_col = '居住地'
            
            # 顯示資料預覽
            st.subheader(f"📋 【{sheet_name}】資料預覽（共 {len(df)} 筆）")
            st.dataframe(df.head(), use_container_width=True)
            
            st.divider()
            
            # 統計圖表區
            st.subheader(f"📈 【{sheet_name}】統計分析")
            
            col1, col2 = st.columns(2)
            
            # 1. 服務區域統計圖表
            with col1:
                if area_col and area_col in df.columns:
                    area_counts = df[area_col].dropna().value_counts().reset_index()
                    area_counts.columns = ['服務區域', '個案數']
                    
                    st.write("### 🏠 各服務區域個案分布")
                    fig_area = px.bar(
                        area_counts, 
                        x='服務區域', 
                        y='個案數', 
                        text='個案數',
                        color='個案數',
                        color_continuous_scale='Viridis',
                        title=f"{sheet_name} - 各服務區域統計"
                    )
                    fig_area.update_traces(textposition='outside')
                    st.plotly_chart(fig_area, use_container_width=True)
                else:
                    st.info("此分頁無『居住地』相關欄位")

            # 2. 個管員統計圖表
            with col2:
                if case_manager_col and case_manager_col in df.columns:
                    cm_counts = df[case_manager_col].dropna().value_counts().reset_index()
                    cm_counts.columns = ['個管員', '個案數']
                    
                    st.write("### 👤 各個管員案件量統計")
                    fig_cm = px.bar(
                        cm_counts, 
                        x='個管員', 
                        y='個案數', 
                        text='個案數',
                        color='個案數',
                        color_continuous_scale='Blues',
                        title=f"{sheet_name} - 各個管員案件量"
                    )
                    fig_cm.update_traces(textposition='outside')
                    st.plotly_chart(fig_cm, use_container_width=True)
                else:
                    st.info("此分頁無『個管員』相關欄位")
            
            # 3. 交叉分析（個管員 x 服務區域）
            if area_col and case_manager_col and (area_col in df.columns) and (case_manager_col in df.columns):
                st.write("### 🔀 個管員與服務區域 交叉分析圖")
                cross_df = df.groupby([case_manager_col, area_col]).size().reset_index(name='個案數')
                cross_df.columns = ['個管員', '服務區域', '個案數']
                
                fig_cross = px.bar(
                    cross_df, 
                    x='個管員', 
                    y='個案數', 
                    color='服務區域', 
                    barmode='stack',
                    title=f"{sheet_name} - 個管員在各區域的案件分布"
                )
                st.plotly_chart(fig_cross, use_container_width=True)
