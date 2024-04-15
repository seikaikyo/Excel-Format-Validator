import streamlit as st
import pandas as pd
import json
import os

config_file = 'config.json'

def save_config(config_data):
    with open(config_file, 'w') as f:
        json.dump(config_data, f)

def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            return json.load(f)
    return {}

def delete_config(config_name):
    configs = load_config()
    if config_name in configs:
        del configs[config_name]
        save_config(configs)
        return True
    return False

st.title('Excel 格式驗證器')

st.write("""
## 為選定欄位排除特定的特殊字符
請選擇欄位並指定不應包含的特殊字符。
""")

uploaded_file = st.file_uploader("選擇文件", type=['xlsx'])
special_characters = ['/', '+', '-', '*', '@', '#', '$', '%', '^', '&', '(', ')']

configurations = load_config()
config_names = list(configurations.keys())
selected_config_name = st.selectbox('選擇或創建新配置', ['創建新配置'] + config_names)

if selected_config_name != '創建新配置' and st.button('刪除此配置'):
    if delete_config(selected_config_name):
        st.success('配置已刪除')
        configurations.pop(selected_config_name, None)  # 更新配置列表
        selected_config_name = '創建新配置'  # 重設選項
    else:
        st.error('刪除配置失敗')

if uploaded_file is not None:
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheets = st.multiselect('選擇工作表', sheet_names)

    if selected_sheets:
        data = pd.read_excel(uploaded_file, sheet_name=selected_sheets[0])
        all_columns = data.columns.tolist()

        column_char_exclude = configurations.get(selected_config_name, {})

        for column in all_columns:
            with st.expander(f"設定 '{column}' 欄位"):
                default_chars = column_char_exclude.get(column, [])
                char_checks = {char: st.checkbox(f"排除 '{char}'", value=char in default_chars, key=f"{column}_{char}") for char in special_characters}
                column_char_exclude[column] = [char for char, checked in char_checks.items() if checked]

        config_name = st.text_input('配置名稱', value=selected_config_name if selected_config_name != '創建新配置' else '')
        if st.button('保存配置'):
            configurations[config_name] = column_char_exclude
            save_config(configurations)
            st.success('配置已保存！')

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
