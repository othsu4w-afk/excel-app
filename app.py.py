import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 設置網頁標題
st.title("📊 Excel 自動化統計報表工具")
st.write("上傳你的 Excel 檔案，系統將自動為你生成數據總覽與圖表！")

# 2. 建立檔案上傳元件
uploaded_file = st.file_uploader("請選擇一個 Excel 檔案 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # 3. 讀取 Excel 數據
    df = pd.read_excel(uploaded_file)

    # 4. 顯示資料預覽
    st.subheader("📋 資料預覽（前 5 列）")
    st.dataframe(df.head())

    # 5. 自動計算關鍵統計指標
    st.subheader("📈 基礎數據統計")
    numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
    
    if numeric_columns:
        selected_column = st.selectbox("請選擇要分析的數值欄位：", numeric_columns)
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("總和", f"{df[selected_column].sum():,.2f}")
        col2.metric("平均值", f"{df[selected_column].mean():,.2f}")
        col3.metric("最大值", f"{df[selected_column].max():,.2f}")
        col4.metric("最小值", f"{df[selected_column].min():,.2f}")

        # 6. 自動繪製圖表
        st.subheader("📊 數據視覺化圖表")
        categorical_columns = df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        if categorical_columns:
            x_axis = st.selectbox("請選擇 X 軸欄位（分類類別）：", categorical_columns)
            fig = px.bar(df, x=x_axis, y=selected_column, title=f"{x_axis} 與 {selected_column} 分析圖")
            st.plotly_chart(fig, use_container_width=True)
        else:
            fig = px.line(df, y=selected_column, title=f"{selected_column} 趨勢圖")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ 該 Excel 檔案中沒有找到數值類型的欄位，無法進行統計。")

