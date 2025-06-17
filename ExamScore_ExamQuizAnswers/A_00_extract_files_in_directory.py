
import os
import zipfile
import rarfile

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


# 示例用法
directory_path = r"C:\MyDocument\ToDoList\D20_DoingPlatform\D20_人工智能与大数据\新建文件夹\23临床药学-2024-2025春季期末考试A卷(word)"  # 替换为你的目录路径
extract_files_in_directory(directory_path)



