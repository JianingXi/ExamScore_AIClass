import os
from docx import Document

import pandas as pd


# 目标根目录
root_dir = r"C:\MyDocument\ToDoList\D20_ToHardDisk\D20250716_腰部肌电信号采集数据\B01处理_验证码\班级_23生工_春_第四次实验_腰部疲劳肌电采集分析实验_word_副本"

# 定义截取的起始和结束标记
start_marker = "3.请同学将领取的个人验证码（用于作业识别和签到）上传。操作步骤：把纸条的验证码输入到答案框，点击提交按钮上传。"
end_marker = "正确答案："

def docx_to_text(docx_path):
    """将docx文件读取并合并为纯文本"""
    doc = Document(docx_path)
    full_text = "\n".join([para.text for para in doc.paragraphs])
    return full_text

def clean_text(text):
    """去除空格、回车、Tab，只保留可显示字符"""
    return "".join(text.split())

def extract_section(text):
    """提取 start_marker 与 end_marker 之间的内容"""
    start_idx = text.find(start_marker)
    if start_idx == -1:
        return None
    start_idx += len(start_marker)  # 跳过起始标记

    end_idx = text.find(end_marker, start_idx)
    if end_idx == -1:
        return None

    return text[start_idx:end_idx]

def process_folder():
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for file in filenames:
            if file.lower().endswith(".docx"):
                docx_path = os.path.join(dirpath, file)
                print(f"处理文件：{docx_path}")

                # 读取 docx 内容
                raw_text = docx_to_text(docx_path)

                # 提取中间段
                section = extract_section(raw_text)
                if section:
                    # 清洗提取内容
                    cleaned = clean_text(section)

                    # 保存为 txt（与 docx 同名）
                    txt_path = os.path.join(dirpath, file.replace(".docx", ".txt"))
                    with open(txt_path, "w", encoding="utf-8") as f:
                        f.write(cleaned)
                    print(f"→ 已保存到：{txt_path}")
                else:
                    print("⚠ 未找到指定内容区间")



# 存储结果的列表
records = []

def clean_text(text):
    """删除 '学生答案：'，并去掉首尾空格"""
    return text.replace("学生答案：", "").strip()

def process_txt_files():
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # 文件夹名称（最后一级）
        folder_name = os.path.basename(dirpath)

        for file in filenames:
            if file.lower().endswith(".txt"):
                txt_path = os.path.join(dirpath, file)
                print(f"读取文件：{txt_path}")

                try:
                    with open(txt_path, "r", encoding="utf-8") as f:
                        content = f.read()
                except:
                    with open(txt_path, "r", encoding="gbk") as f:
                        content = f.read()

                cleaned = clean_text(content)

                # 保存记录
                records.append({"Folder_Name": folder_name, "Cleaned_Text": cleaned})

def save_to_excel():
    df = pd.DataFrame(records)
    output_path = os.path.join(root_dir, "汇总结果.xlsx")
    df.to_excel(output_path, index=False, encoding="utf-8")
    print(f"\n✅ 已生成Excel：{output_path}")

if __name__ == "__main__":
    process_folder()
    process_txt_files()
    save_to_excel()
