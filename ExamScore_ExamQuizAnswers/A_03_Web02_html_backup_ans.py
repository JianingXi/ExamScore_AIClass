import os
from pathlib import Path
import shutil

# html -> test file
def convert_html_batch(source_root: str, target_root: str):
    """
    扫描 source_root 下的所有子文件夹，
    若其中包含 page1.html、page2.html、page3.html，
    则将它们转为 .txt 文件，统一保存到 target_root / [顶层子目录名]。
    """

    source_root = Path(source_root)
    target_root = Path(target_root)

    # HTML -> 纯文本
    def html_to_txt(html_path):
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            print(f"❌ 解析失败：{html_path}，错误：{e}")
            return ""

    # 遍历所有子目录
    for root, dirs, files in os.walk(source_root):
        print(files)
        file_set = set(files)
        required_pages = {"page1.html", "page2.html", "page3.html"}

        if not required_pages.isdisjoint(file_set):
            current_folder = Path(root)
            relative = current_folder.relative_to(source_root)
            top_level = relative.parts[0] if len(relative.parts) > 0 else current_folder.name
            output_folder = target_root / top_level
            output_folder.mkdir(parents=True, exist_ok=True)

            print(f"\n✅ 处理：{current_folder} → 保存至：{output_folder}")

            for page in sorted(required_pages):
                html_path = current_folder / page
                txt = html_to_txt(html_path)

                if txt.strip():
                    output_path = output_folder / page.replace(".html", ".txt")
                    with open(output_path, "w", encoding="utf-8") as f:
                        f.write(txt)
                    print(f"   ✔️ 写入：{output_path}")
                else:
                    print(f"   ⚠️ 空内容跳过：{html_path}")






import os
from pathlib import Path
import shutil

def flatten_and_rename_txt_files(folder_path: str):
    base_path = Path(folder_path)

    # Step 1：递归查找所有 txt 文件
    for txt_file in base_path.rglob("*.txt"):
        if txt_file.is_file() and txt_file.parent != base_path:
            # 获取相对路径的所有目录部分（不包括文件名）
            relative_parts = txt_file.relative_to(base_path).parts[:-1]
            filename = txt_file.name
            flat_name = "_".join(relative_parts + (filename,))
            new_path = base_path / flat_name

            # 移动并重命名到根目录
            shutil.move(str(txt_file), str(new_path))
            print(f"✅ 移动并重命名：{txt_file} → {new_path}")

    # Step 2：再次遍历根目录，替换文件名中的 "-"
    for file in base_path.glob("*.txt"):
        if "-" in file.name:
            new_name = file.name.replace("-", "_")
            new_path = file.with_name(new_name)
            file.rename(new_path)
            print(f"🛠️ 文件名中 - → _：{file.name} → {new_name}")

    # Step 3：删除空文件夹
    for folder in sorted(base_path.rglob("*"), reverse=True):
        if folder.is_dir() and not any(folder.iterdir()):
            folder.rmdir()
            print(f"🧹 删除空文件夹：{folder}")


def convert_html_to_backup_answers(exam_folder_path_raw: str, exam_folder_path_html: str, exam_file_base_names: list):
    """
    将一批 HTML 题目转为备用答案 txt，并统一重命名。

    参数：
    - exam_folder_path: 根目录路径，例如 "C:\\Users\\xijia\\Desktop\\批改web"
    - exam_file_base_names: 文件夹名列表，例如 ["xxx_word_", "yyy_word_"]
    """

    for name in exam_file_base_names:
        print(exam_folder_path_raw + '\\' + name)
        convert_html_batch(
            source_root = exam_folder_path_raw + '\\' + name,
            target_root = exam_folder_path_html + '\\' + name
        )

        flatten_and_rename_txt_files(
            folder_path = exam_folder_path_html + '\\' + name
        )
