import streamlit as st
import pandas as pd

st.title('Excel 格式驗證器')

st.write("""
## 為選定欄位排除特定的特殊字符
請選擇欄位並指定不應包含的特殊字符。
""")

uploaded_file = st.file_uploader("選擇文件", type=['xlsx'])
special_characters = ['/', '+', '-', '*', '@', '#', '$', '%', '^', '&', '(', ')']

if uploaded_file is not None:
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheet = st.selectbox('選擇工作表', sheet_names)

    if selected_sheet:
        data = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        all_columns = data.columns.tolist()
        
        st.write("### 選擇需要檢查的欄位和排除的特殊字符")
        column_char_exclude = {}
        for column in all_columns:
            with st.expander(f"設定 '{column}' 欄位"):
                char_checks = {char: st.checkbox(f"排除 '{char}'", key=f"{column}_{char}") for char in special_characters}
                column_char_exclude[column] = [char for char, checked in char_checks.items() if checked]

        st.write("### 驗證結果")
        for column, chars_to_exclude in column_char_exclude.items():
            if chars_to_exclude:
                data[f'IsValid_{column}'] = data[column].apply(lambda x: all(char not in str(x) for char in chars_to_exclude))
                valid_count = data[f'IsValid_{column}'].sum()
                total_count = len(data)

                st.write(f"處理欄位：{column}")
                st.write(f"排除字符：{', '.join(chars_to_exclude)}")
                st.write(f"總行數：{total_count}, 符合條件的行數：{valid_count}")
                st.dataframe(data[~data[f'IsValid_{column}']])
