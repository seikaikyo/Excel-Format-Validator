import streamlit as st
import pandas as pd
import re
import streamlit as st

# 允許的字符正則表達式
allowed_chars_regex = r'^[A-Za-z0-9\s,.;:!-]+$'

st.title('Excel 格式驗證器')

st.write("""
## 允許的字符
本應用程式只接受以下字符：
- 英文字母 (A-Z, a-z)
- 數字 (0-9)
- 空格 ( )
- 基本標點符號 (, . ; : ! -)
""")

st.write(f"正則表達式為：`{allowed_chars_regex}`")

def validate_data(row, format_rules):
    errors = []
    for column, value in row.items():
        rule = format_rules.get(column, "")
        # 確保 rule 是字符串
        if not isinstance(rule, str):
            rule = str(rule)
        # 提取碼數
        match = re.match(r'.*\((\d+)碼\)', rule)
        if match:
            expected_length = int(match.group(1))
            value_str = str(value)
            if len(value_str) != expected_length:
                errors.append(f"{column}: 長度應為 {expected_length}，實際為 {len(value_str)}")
        # 檢查特殊字符
        if not re.match(r'^[\w\s,.;:!-]+$', str(value)):
            errors.append(f"{column}: 包含不允許的特殊字符")
    return errors

def load_excel(file):
    data = pd.read_excel(file)
    # 使用第一列作為格式規則
    format_rules = data.iloc[0].to_dict()
    data = data.drop(0).reset_index(drop=True)
    return data, format_rules

st.title('Excel 格式驗證器')

uploaded_file = st.file_uploader("選擇文件", type=['xlsx'])
if uploaded_file is not None:
    st.write(f"本次驗證檔案檔名為: {uploaded_file.name}")
    data, format_rules = load_excel(uploaded_file)
    st.write("本程式驗證規則：")
    st.json(format_rules)

    error_list = []
    total_rows = len(data)
    validated_rows = 0

    progress_bar = st.progress(0)
    for index, row in data.iterrows():
        errors = validate_data(row, format_rules)
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
