import streamlit as st
import pandas as pd
import json
import os
import re

config_file = 'config.json'

def save_config(config_data):
    with open(config_file, 'w') as f:
        json.dump(config_data, f)

def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            config_data = json.load(f)
        # 確保所有配置都是字典格式
        updated_config_data = {}
        for key, value in config_data.items():
            if isinstance(value, dict):
                updated_config_data[key] = value
            else:
                # 如果不是字典，初始化為空的設定
                updated_config_data[key] = {col: {'exclude_chars': [], 'max_length': 255, 'disable_chinese': False} for col in (value if isinstance(value, list) else [])}
        return updated_config_data
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
## 為選定欄位排除特定的特殊字符包含中文及設置字元長度限制
請選擇欄位並指定不應包含的特殊字符包含中文，並設定字元長度限制。
設定的最大字元長度包括中英文和數字。注意：英文和數字通常計為1個字元，中文和其他全形字符可能計為2個或更多字元。
""")



uploaded_file = st.file_uploader("選擇文件", type=['xlsx'])
special_characters = ['/', '+', '-', '*', '@', '#', '$', '%', '^', '&', '(', ')']

configurations = load_config()
config_names = list(configurations.keys())
selected_config_name = st.selectbox('選擇或創建新配置', ['創建新配置'] + config_names)

if selected_config_name != '創建新配置' and st.button('刪除此配置'):
    if delete_config(selected_config_name):
        st.success('配置已刪除')
        configurations.pop(selected_config_name, None)
        selected_config_name = '創建新配置'
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
                default_chars = column_char_exclude.get(column, {'exclude_chars': [], 'max_length': 255, 'disable_chinese': False})
                # 這裡加入對default_chars的字典類型檢查
                if not isinstance(default_chars, dict):
                    default_chars = {'exclude_chars': [], 'max_length': 255, 'disable_chinese': False}
                char_checks = {char: st.checkbox(f"排除 '{char}'", value=char in default_chars['exclude_chars'], key=f"{column}_{char}") for char in special_characters}
                max_length = st.number_input('設定最大字元長度', value=default_chars['max_length'], min_value=1, max_value=1024, key=f"max_length_{column}")
                disable_chinese = st.checkbox('禁用中文字符', value=default_chars['disable_chinese'], key=f"disable_chinese_{column}")

                column_char_exclude[column] = {
                    'exclude_chars': [char for char, checked in char_checks.items() if checked],
                    'max_length': max_length,
                    'disable_chinese': disable_chinese
                }

        config_name = st.text_input('配置名稱', value=selected_config_name if selected_config_name != '創建新配置' else '')
        if st.button('保存配置'):
            configurations[config_name] = column_char_exclude
            save_config(configurations)
            st.success('配置已保存！')

        st.write("### 驗證結果")
        for column, settings in column_char_exclude.items():
            chars_to_exclude = settings['exclude_chars']
            max_length = settings['max_length']
            disable_chinese = settings['disable_chinese']

            data[f'IsValid_{column}'] = data[column].apply(lambda x: all(char not in str(x) for char in chars_to_exclude) and len(str(x)) <= max_length and (not disable_chinese or not re.search('[\u4e00-\u9fff]', str(x))))
            valid_count = data[f'IsValid_{column}'].sum()
            total_count = len(data)

            st.write(f"處理欄位：{column}")
            st.write(f"排除字符：{', '.join(chars_to_exclude)}")
            st.write(f"最大字元長度：{max_length}")
            st.write(f"禁用中文字符：{'是' if disable_chinese else '否'}")
            st.write(f"總行數：{total_count}, 符合條件的行數：{valid_count}")
            st.dataframe(data[~data[f'IsValid_{column}']])
