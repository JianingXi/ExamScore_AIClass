import os
from pathlib import Path
from PIL import Image


def delete_empty_folders(directory):
    """递归删除空文件夹，包括嵌套的空文件夹"""
    for root, dirs, files in os.walk(directory, topdown=False):
        for d in dirs:
            folder_path = Path(root) / d
            if not any(folder_path.iterdir()):
                folder_path.rmdir()
                print(f"已删除空文件夹：{folder_path}")


def merge_images_in_folder(folder_path):
    """将文件夹中的图片拼接成一张长图"""
    image_paths = list(folder_path.glob("*.png"))
    image_paths.extend(list(folder_path.glob("*.jpg")))
    image_paths.extend(list(folder_path.glob("*.jpeg")))
    image_paths.sort()

    if not image_paths:
        return None

    images = []
    for p in image_paths:
        try:
            img = Image.open(p)
            images.append(img)
        except Exception as e:
            print(f"无法打开文件 {p}: {e}")

    if not images:
        return None

    total_height = sum(img.height for img in images)
    max_width = max(img.width for img in images)

    merged_image = Image.new('RGB', (max_width, total_height))
    y_offset = 0
    for img in images:
        merged_image.paste(img, (0, y_offset))
        y_offset += img.height

    folder_name = Path(folder_path).name
    output_path = Path(folder_path).parent / f"{folder_name}.png"
    try:
        merged_image.save(output_path)
        print(f"已保存拼接图：{output_path}")
        return output_path
    except Exception as e:
        print(f"保存拼接图失败：{output_path}: {e}")
        return None


def process_directory(base_path):
    base_dir = Path(base_path)

    # 遍历整个目录，包括嵌套子文件夹
    for root, dirs, _ in os.walk(base_dir, topdown=False):
        for d in dirs:
            folder_path = Path(root) / d
            merged_image_path = merge_images_in_folder(folder_path)
            if merged_image_path:
                # 移动图片到上一级目录
                target_path = base_dir / merged_image_path.name
                try:
                    os.replace(merged_image_path, target_path)
                    print(f"已移动图片：{target_path}")
                except Exception as e:
                    print(f"移动图片失败：{target_path}: {e}")

            # 删除文件夹（若为空）
            if not any(folder_path.iterdir()):
                try:
                    folder_path.rmdir()
                    print(f"已删除空文件夹：{folder_path}")
                except Exception as e:
                    print(f"删除文件夹失败：{folder_path}: {e}")

    # 再次检查根目录中的空文件夹
    delete_empty_folders(base_dir)


if __name__ == "__main__":
    base_path = r"C:\MyPython\ExamScore_AIClass\作品\科普组\漫画类"
    process_directory(base_path)
