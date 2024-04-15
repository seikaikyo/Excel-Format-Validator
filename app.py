import streamlit as st
import pandas as pd
import re

# 允許的字符正則表達式，包括中文、破折號、圓括號、底線和加號
allowed_chars_regex = r'^[A-Za-z0-9\s,.;:!/\-\(\)_\+\u4e00-\u9fff]+$'

def validate_data(row, format_rules, no_slash_fields):
    errors = []
    for column, value in row.items():
        value_str = str(value)
        if column in no_slash_fields and '/' in value_str:
            errors.append(f"{column}: 不應包含斜杠 (/) at column '{column}' with value '{value}'")
        elif not re.match(allowed_chars_regex, value_str):
            illegal_chars = set(re.findall(r'[^A-Za-z0-9\s,.;:!/\-\(\)_\+\u4e00-\u9fff]', value_str))
            if illegal_chars:
                illegal_char_str = ', '.join(illegal_chars)
                errors.append(f"{column}: 包含不允許的特殊字符 {illegal_char_str} at column '{column}' with value '{value}'")
    return errors

st.title('Excel 格式驗證器')

st.write("""
## 允許的字符
本應用程式只接受以下字符：
- 英文字母 (A-Z, a-z)
- 數字 (0-9)
- 空格 ( )
- 基本標點符號 (, . ; : ! / -)
- 圓括號 ()
- 底線 _
- 加號 +
- 中文字符
""")

st.write(f"正則表達式為：`{allowed_chars_regex}`")

uploaded_file = st.file_uploader("選擇文件", type=['xlsx'])
no_slash_fields = []

if uploaded_file is not None:
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheets = st.multiselect('選擇工作表', sheet_names, default=sheet_names)

    if selected_sheets:
        data = pd.read_excel(uploaded_file, sheet_name=selected_sheets[0])  # 使用第一個選擇的工作表來決定欄位
        all_columns = data.columns.tolist()
        no_slash_fields = st.multiselect('選擇不應包含斜杠 (/) 的欄位', all_columns, default=[])

        for sheet in selected_sheets:
            data = pd.read_excel(uploaded_file, sheet_name=sheet)
            format_rules = data.iloc[0].to_dict()
            data = data.drop(0).reset_index(drop=True)
            st.write("本程式驗證規則：")
            st.json(format_rules)
            error_list = []
            total_rows = len(data)
            validated_rows = 0
            progress_bar = st.progress(0)

            for index, row in data.iterrows():
                errors = validate_data(row, format_rules, no_slash_fields)
                if errors:
                    error_list.append(f"第 {index+1} 行錯誤: {errors}")
                validated_rows += 1
                progress_bar.progress(validated_rows / total_rows)

            if error_list:
                st.write(f"驗證數/驗證總筆數: {validated_rows}/{total_rows}")
                for error in error_list:
                    st.error(error)
            else:
                st.success("所有數據均符合格式要求！")
                st.write(f"驗證數/驗證總筆數: {validated_rows}/{total_rows}")
