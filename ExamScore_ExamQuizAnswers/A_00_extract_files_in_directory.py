import os
import zipfile
import rarfile

import re


# 配置 rarfile 使用的解压工具路径
rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\UnRAR.exe"  # 确保路径正确


def extract_zip(file_path, extract_to):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)


def extract_rar(file_path, extract_to):
    with rarfile.RarFile(file_path, 'r') as rar_ref:
        rar_ref.extractall(extract_to)


def extract_files_in_directory(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_name, file_extension = os.path.splitext(file)

            if file_extension.lower() in ['.zip', '.rar']:
                extract_to = os.path.join(root, file_name)

                # 仅在文件是压缩包时创建解压缩目标目录
                if not os.path.exists(extract_to):
                    os.makedirs(extract_to)

                if file_extension.lower() == '.zip':
                    try:
                        print(f"Extracting {file_path} to {extract_to}")
                        extract_zip(file_path, extract_to)
                        os.remove(file_path)
                        print(f"Deleted {file_path}")
                    except Exception as e:
                        print(f"Error extracting {file_path}: {e}")
                elif file_extension.lower() == '.rar':
                    try:
                        print(f"Extracting {file_path} to {extract_to}")
                        extract_rar(file_path, extract_to)
                        os.remove(file_path)
                        print(f"Deleted {file_path}")
                    except Exception as e:
                        print(f"Error extracting {file_path}: {e}")


# 不修改你的函数，只负责调用
def is_compressed_file(filename):
    compressed_extensions = ['.zip', '.rar', '.7z', '.tar', '.gz']
    return any(filename.lower().endswith(ext) for ext in compressed_extensions)

# 将文件名中的非法字符替换为下划线
def sanitize_filename(filename):
    # 只保留中文、英文字母、数字、点、下划线和连字符，其他都换成 _
    return re.sub(r'[^\w\u4e00-\u9fff.-]', '_', filename)

def extract_all_in_directory(target_folder):
    for root, dirs, files in os.walk(target_folder):
        for file in files:
            if is_compressed_file(file):
                original_path = os.path.join(root, file)
                sanitized_name = sanitize_filename(file)
                sanitized_path = os.path.join(root, sanitized_name)

                # 如果名字变了，先重命名
                if sanitized_path != original_path:
                    try:
                        os.rename(original_path, sanitized_path)
                        print(f"已重命名：{original_path} -> {sanitized_path}")
                    except Exception as e:
                        print(f"重命名失败：{original_path}，错误：{e}")
                        continue  # 跳过无法处理的文件
    extract_files_in_directory(target_folder)

if __name__ == "__main__":
    folder_path = r"C:\Users\xijia\Desktop\批改web\S01_Raw\临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_"
    extract_all_in_directory(folder_path)