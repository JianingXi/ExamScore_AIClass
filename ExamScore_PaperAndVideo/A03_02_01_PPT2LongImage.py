import os
from pathlib import Path
import win32com.client as win32
from PIL import Image


# ==========================================================
# 1. PPTX → 每页 PNG（PowerPoint 原生导出）
# ==========================================================

def export_pptx_to_images(pptx_path: Path, out_dir: Path):
    """
    将 pptx 每一页导出为 PNG（PowerPoint 原生渲染）
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    powerpoint = win32.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True   # ✅ 必须是 True，不能 False

    presentation = powerpoint.Presentations.Open(
        str(pptx_path),
        WithWindow=False
    )

    # 导出全部幻灯片
    presentation.Export(
        str(out_dir),
        "PNG"
    )

    presentation.Close()
    powerpoint.Quit()


# ==========================================================
# 2. 图片统一压缩 + 纵向拼接
# ==========================================================

def merge_images_vertical(
    img_dir: Path,
    output_path: Path,
    max_width: int = 1200,
    quality: int = 90
):
    images = sorted(img_dir.glob("*.png"))
    if not images:
        return

    resized_images = []
    total_height = 0

    for img_path in images:
        img = Image.open(img_path).convert("RGB")

        if img.width > max_width:
            ratio = max_width / img.width
            img = img.resize(
                (max_width, int(img.height * ratio)),
                Image.LANCZOS
            )

        resized_images.append(img)
        total_height += img.height

    merged = Image.new("RGB", (max_width, total_height), (255, 255, 255))

    y = 0
    for img in resized_images:
        merged.paste(img, (0, y))
        y += img.height

    merged.save(output_path, quality=quality)
    print(f"✅ 生成长图: {output_path}")


# ==========================================================
# 3. 单个 PPTX → 长图
# ==========================================================

def pptx_to_long_image(pptx_path: Path):
    temp_dir = pptx_path.parent / f"__imgs_{pptx_path.stem}"
    output_img = pptx_path.with_suffix(".long.png")

    export_pptx_to_images(pptx_path, temp_dir)
    merge_images_vertical(temp_dir, output_img)

    # 清理中间图片
    for f in temp_dir.glob("*"):
        f.unlink()
    temp_dir.rmdir()


# ==========================================================
# 4. 批量处理（与你原扫描逻辑一致）
# ==========================================================

def batch_pptx_to_long_images(root_dir: str):
    root = Path(root_dir)

    for sub in root.iterdir():
        if not sub.is_dir():
            continue

        print(f"\n▶ 处理子文件夹: {sub}")

        pptx_files = list(sub.glob("*.pptx"))
        for pptx in pptx_files:
            try:
                pptx_to_long_image(pptx)
            except Exception as e:
                print(f"❌ 处理失败 {pptx}: {e}")


# ==========================================================
# 5. 主入口
# ==========================================================

if __name__ == "__main__":
    root_dir = r"D:\YourRootDir"
    batch_pptx_to_long_images(root_dir)
