import os
import shutil
from ExamScore_ExamQuizAnswers.A_02_convert_docx_to_txt import select_ans_txt_files

def is_only_one_subfolder(folder_path):
    """
    判断一个文件夹是否只含一个子文件夹，且无其他文件
    """
    items = os.listdir(folder_path)
    if len(items) != 1:
        return False, None
    only_item = items[0]
    only_item_path = os.path.join(folder_path, only_item)
    if os.path.isdir(only_item_path):
        return True, only_item_path
    return False, None

def flatten_recursively(folder_path):
    """
    递归地将某文件夹中只含一个子文件夹的结构向上提
    """
    while True:
        only_one, subfolder_path = is_only_one_subfolder(folder_path)
        if not only_one:
            break

        print(f"\n📁 扁平化：{folder_path} ← {subfolder_path}")

        for item in os.listdir(subfolder_path):
            src = os.path.join(subfolder_path, item)
            dst = os.path.join(folder_path, item)

            if os.path.exists(dst):
                print(f"⚠️ 跳过已有项：{dst}")
                continue

            shutil.move(src, dst)
            print(f"✅ 移动：{src} → {dst}")

        # 删除已空子文件夹
        if not os.listdir(subfolder_path):
            os.rmdir(subfolder_path)
            print(f"✅ 删除子目录：{subfolder_path}")
        else:
            print(f"❌ 子目录未空，未删除：{subfolder_path}")
            break  # 防止无限循环

    # 继续递归处理当前目录下的所有子目录
    for name in os.listdir(folder_path):
        sub_path = os.path.join(folder_path, name)
        if os.path.isdir(sub_path):
            flatten_recursively(sub_path)

def flatten_all_folders(root_path):
    """
    对 root_path 下的所有目录层级递归扁平化处理
    """
    flatten_recursively(root_path)


def print_tree(start_path, prefix=""):
    """
    打印从 start_path 开始的目录和文件树状结构。
    """
    # 获取当前目录下的所有项，并排序：文件夹在前
    entries = sorted(os.listdir(start_path), key=lambda x: (not os.path.isdir(os.path.join(start_path, x)), x.lower()))
    total = len(entries)

    for idx, entry in enumerate(entries):
        path = os.path.join(start_path, entry)
        connector = "└── " if idx == total - 1 else "├── "
        print(prefix + connector + entry)

        if os.path.isdir(path):
            extension = "    " if idx == total - 1 else "│   "
            print_tree(path, prefix + extension)









# 扫描文件夹的pkl是否名字和数量对的上


# 目标路径（修改为你的路径）
target_folder = r"C:\Users\xijia\Desktop\检查计算机应用基础这门课资料\评估资料-留学生计算机基础"

for i_ind in range(8):
    flatten_all_folders(target_folder)

print_tree(target_folder)

delete_num_array = [1, 2]
txt_folder_path = target_folder + r"\txt_files"
# select_ans_txt_files(target_folder, txt_folder_path, delete_num_array)