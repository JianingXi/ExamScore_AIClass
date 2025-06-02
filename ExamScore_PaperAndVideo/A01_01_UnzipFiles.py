import os
import zipfile
from pathlib import Path

import shutil
import rarfile


def extract_archives(input_folder: str):
    """
    Extract all .zip archives in the given folder and delete each archive after extraction.
    """
    input_dir = Path(input_folder)
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"Path '{input_folder}' is not a valid directory.")

    for archive_path in input_dir.iterdir():
        if archive_path.is_file() and archive_path.suffix.lower() == ".zip":
            target_dir = input_dir / archive_path.stem
            target_dir.mkdir(exist_ok=True)
            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                zip_ref.extractall(target_dir)
            # 删除原始压缩包
            archive_path.unlink()





def extract_all_archives(folder):
    """
    在 folder 下解压所有压缩包，并删除压缩包文件。
    支持 zip, rar, 7z 等格式。
    """
    for fname in os.listdir(folder):
        path = os.path.join(folder, fname)
        if os.path.isfile(path) and fname.lower().endswith(('.zip', '.rar', '.7z')):
            print(f"Extracting {path}...")
            try:
                # 尝试用 shutil（支持 zip, tar 等）
                shutil.unpack_archive(path, folder)
            except (shutil.ReadError, ValueError):
                # 如果是 rar/7z，使用 rarfile 库
                try:
                    with rarfile.RarFile(path) as rf:
                        rf.extractall(folder)
                except Exception as e:
                    print(f"  ❌  无法解压 {fname}: {e}")
                    continue
            # 解压成功后删除压缩包
            os.remove(path)

def flatten_subdirs(folder):
    """
    将 folder 下所有一级子目录的内容搬到 folder，
    并删除这些子目录（如果搬空成功）。
    """
    for name in os.listdir(folder):
        sub = os.path.join(folder, name)
        if os.path.isdir(sub):
            print(f"Flattening directory {sub}...")
            for item in os.listdir(sub):
                src = os.path.join(sub, item)
                dst = os.path.join(folder, item)
                # 如果同名冲突，可根据需要改名或覆盖
                shutil.move(src, dst)
            # 删除已空的子文件夹
            os.rmdir(sub)

def process_folder(folder):
    """
    针对单个文件夹，反复执行解压+扁平化，
    直到内部不含任何压缩包和子目录。
    """
    while True:
        # 先解压所有压缩包
        extract_all_archives(folder)
        # 再扁平化所有子目录
        flatten_subdirs(folder)
        # 检查是否已完成
        contents = os.listdir(folder)
        has_archive = any(f.lower().endswith(('.zip', '.rar', '.7z')) for f in contents)
        has_dir = any(os.path.isdir(os.path.join(folder, f)) for f in contents)
        if not has_archive and not has_dir:
            break

def scaning(root_dir):
    # 遍历根目录下的每个子文件夹
    for entry in os.listdir(root_dir):
        path = os.path.join(root_dir, entry)
        if os.path.isdir(path):
            print(f"\n▶ 处理子文件夹: {path}")
            process_folder(path)

def a01_01_unzip_files(ROOT: str):
    # ROOT = r"C:\MyPython\ExamScore_AIClass\ExamFiles"
    extract_archives(ROOT)
    scaning(ROOT)



