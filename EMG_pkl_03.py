import os
import pandas as pd

def collect_folder_and_txt_content(base_path):
    rows = []

    for folder_name in os.listdir(base_path):
        folder_path = os.path.join(base_path, folder_name)
        if not os.path.isdir(folder_path):
            continue  # 跳过非目录项

        txt_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.txt')]

        if len(txt_files) != 1:
            print(f"⚠️ {folder_name} 中的 txt 文件数量不是 1，跳过")
            continue

        txt_path = os.path.join(folder_path, txt_files[0])
        with open(txt_path, 'r', encoding='utf-8') as f:
            content = f.read().strip()

        rows.append([folder_name, content])

    # 创建 DataFrame
    df = pd.DataFrame(rows, columns=["文件夹名称", "TXT内容"])
    return df

# 设置路径
base_dir = r"C:\Users\xijia\Desktop\腰部肌电信号采集数据\B01处理_验证码\2022级生物医学工程-智能传感分析-2腰部疲劳肌电采集分析实验_word_\txt_files"

# 生成 DataFrame 并显示
df = collect_folder_and_txt_content(base_dir)
print(df)
