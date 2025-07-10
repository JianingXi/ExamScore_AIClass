import os

import time
import win32com.client

import subprocess

import re

def init_word():
    """
    初始化 Word 应用程序。
    """
    try:
        print("Initializing Word Application...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        return word
    except Exception as e:
        print(f"初始化 Word 应用程序时发生错误: {e}")
        return None


def close_word(word):
    """
    关闭 Word 应用程序。
    """
    try:
        if word:
            word.Quit()
            del word  # 确保完全释放 COM 对象
    except Exception as e:
        print(f"关闭 Word 应用程序时发生错误: {e}")


def convert_doc_to_docx(word, input_file, output_file):
    """
    将 .doc 文件转换为 .docx 文件。

    :param word: Word 应用程序对象。
    :param input_file: 输入的 .doc 文件路径。
    :param output_file: 输出的 .docx 文件路径。
    """
    doc = None
    try:
        # 将文件路径转换为绝对路径并处理中文字符
        input_file = os.path.abspath(input_file)
        output_file = os.path.abspath(output_file)

        # 检查输入文件是否存在
        if not os.path.exists(input_file):
            print(f"输入文件不存在: {input_file}")
            return

        print(f"正在转换文件: {input_file}")
        doc = word.Documents.Open(input_file)

        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_file), exist_ok=True)

        doc.SaveAs(output_file, FileFormat=16)  # 16 是 .docx 格式的文件格式代码
        print(f"转换成功！.docx 文件已保存至: {output_file}")
    except Exception as e:
        print(f"转换过程中发生错误: {e}")
    finally:
        if doc:
            try:
                doc.Close()
            except Exception as e:
                print(f"关闭文档时发生错误: {e}")


def traverse_and_convert_doc_to_docx(root_folder):
    """
    遍历给定的根文件夹，并将所有 .doc 文件转换为 .docx 文件。

    :param root_folder: 要遍历的根文件夹路径。
    """
    # 初始化 Word 应用程序
    word = init_word()
    if not word:
        print("无法初始化 Word 应用程序，退出。")
        return

    # 使用 os.walk 遍历根文件夹及其所有子文件夹
    for subdir, _, files in os.walk(root_folder):
        # 遍历当前文件夹中的所有文件
        for file in files:
            # 检查文件是否为 .doc 格式的 Word 文档
            if file.lower().endswith('.doc') and not file.lower().endswith('.docx'):
                # 构建输入文件的完整路径
                input_file = os.path.join(subdir, file)
                # 构建输出文件的完整路径
                output_file = os.path.join(subdir, file + 'x')
                # 调用转换函数，将 .doc 文件转换为 .docx 文件
                convert_doc_to_docx(word, input_file, output_file)

    # 关闭 Word 应用程序
    close_word(word)


def delete_doc_files(root_folder):
    """
    删除指定文件夹及子文件夹中的所有 .doc 文件，保留 .docx 文件。

    参数：
    root_folder (str): 根目录路径
    """
    deleted_files = []
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith(".doc") and not filename.lower().endswith(".docx"):
                file_path = os.path.join(dirpath, filename)
                try:
                    os.remove(file_path)
                    deleted_files.append(file_path)
                except Exception as e:
                    print(f"无法删除 {file_path}：{e}")
    print(f"共删除 {len(deleted_files)} 个 .doc 文件。")
    return deleted_files

def convert_docx_to_txt(input_file, output_file):
    """
    使用 pandoc 将 Word 文档 (.docx) 转换为文本文件 (.txt)。

    :param input_file: 输入的 Word 文档路径。
    :param output_file: 输出的文本文件路径。
    """
    try:
        # 运行 pandoc 命令，将文件转换为纯文本格式
        subprocess.run(['pandoc', input_file, '-t', 'plain', '-o', output_file], check=True)
        print(f"转换成功！文本文件已保存至: {output_file}")
    except subprocess.CalledProcessError as e:
        print(f"转换过程中发生错误: {e}")


def traverse_and_convert_docx_to_txt(root_folder, new_folder):
    """
    遍历给定的根文件夹，并将所有 Word 文档转换为文本文件，存储在当前目录下的新建文件夹 txt_files 中。

    :param root_folder: 要遍历的根文件夹路径。
    """
    # 定义输出文件夹路径
    output_root_folder = os.path.join(root_folder, new_folder)
    # 确保输出文件夹存在，若不存在则创建
    os.makedirs(output_root_folder, exist_ok=True)

    # 使用 os.walk 遍历根文件夹及其所有子文件夹
    for subdir, _, files in os.walk(root_folder):
        # 遍历当前文件夹中的所有文件
        for file in files:
            # 检查文件是否为 .docx 格式的 Word 文档
            if file.lower().endswith('.docx'):
                # 构建输入文件的完整路径
                input_file = os.path.join(subdir, file)
                # 生成输出文件的相对路径，并将后缀名从 .docx 改为 .txt
                relative_path = os.path.relpath(input_file, root_folder)
                # 确保输出目录存在，若不存在则创建
                output_file_dir = os.path.join(output_root_folder, os.path.dirname(relative_path))
                os.makedirs(output_file_dir, exist_ok=True)
                # 构建输出文件的完整路径
                output_file = os.path.join(output_file_dir, os.path.basename(relative_path).replace('.docx', '.txt'))
                # 调用转换函数，将 Word 文档转换为文本文件
                convert_docx_to_txt(input_file, output_file)
    return output_root_folder


def extract_student_answers_only(file_path):
    dir_path = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    pattern = re.compile(r"学生答案：(.*?)正确答案：", re.DOTALL)
    matches = pattern.findall(content)

    if not matches:
        print("⚠️ 没有找到匹配的学生答案段落。")
        return

    new_files = []
    for i, answer_block in enumerate(matches, 1):
        index_str = f"{i:03d}"
        output_file = os.path.join(dir_path, f"{base_name}-student_{index_str}.txt")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(answer_block.strip())
        new_files.append(output_file)

    print(f"✅ 成功提取并保存 {len(matches)} 段学生答案内容。")

    # 删除原始输入文件
    try:
        os.remove(file_path)
    except Exception as e:
        print(f"⚠️ 删除原始文件失败：{e}")


def split_all_txt_in_directory(folder_path):
    txt_files = [
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".txt") and os.path.isfile(os.path.join(folder_path, f))
    ]

    if not txt_files:
        print("📂 未找到任何 .txt 文件。")
        return

    for txt_file in txt_files:
        extract_student_answers_only(txt_file)



def delete_selected_student_txt(folder_path, target_nums):
    """
    删除 folder_path 中所有结尾为 -student_XXX.txt 的文件，且 XXX 在 target_nums 中。

    参数:
        folder_path (str): 要扫描的文件夹路径
        target_nums (list[int]): 要删除的编号列表，如 [1, 2, 3, 16]
    """
    target_indices = {f"{i:03d}" for i in target_nums}
    deleted_files = []

    for filename in os.listdir(folder_path):
        match = re.match(r".*-student_(\d{3})\.txt$", filename)
        if match:
            num = match.group(1)
            if num in target_indices:
                file_path = os.path.join(folder_path, filename)
                try:
                    os.remove(file_path)
                    deleted_files.append(filename)
                except Exception as e:
                    print(f"⚠️ 无法删除 {filename}：{e}")
    print(f"✅ 共删除 {len(deleted_files)} 个文件：")



def replace_exit_code_in_txt_files(root_folder: str, target_text: str):
    modified_files = []

    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith('.txt'):
                file_path = os.path.join(dirpath, filename)

                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()

                if target_text in content:
                    new_content = content.replace(target_text, '')

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)

                    modified_files.append(file_path)

    print(f"✅ 处理完成，共修改 {len(modified_files)} 个文件。")
    return modified_files


def select_ans_txt_files(root_folder: str, delete_num_array):
    traverse_and_convert_doc_to_docx(root_folder)
    delete_doc_files(root_folder)

    folder_path = traverse_and_convert_docx_to_txt(root_folder, 'txt_files')

    # split_all_txt_in_directory(folder_path)
    # delete_selected_student_txt(folder_path, target_nums=delete_num_array)
    for dirpath, dirnames, filenames in os.walk(folder_path):
        # 跳过根目录，只处理子文件夹
        if dirpath != folder_path:
            try:
                split_all_txt_in_directory(dirpath)
                delete_selected_student_txt(dirpath, target_nums=delete_num_array)
                print(f"已处理：{dirpath}")
            except Exception as e:
                print(f"处理失败：{dirpath}，错误：{e}")

    # replace_exit_code_in_txt_files(root_folder, "进程已结束,退出代码0")


if __name__ == "__main__":
    # 示例用法
    root_folder = r"C:\MyDocument\ToDoList"  # 替换为你的目录路径
    delete_num_array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 16]
    select_ans_txt_files(root_folder, delete_num_array)