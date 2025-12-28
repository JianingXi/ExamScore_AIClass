import os
import cv2
import json
import hashlib
from PIL import Image


# =========================================================
# 1. 等间隔抽取 16 帧（覆盖整个视频）
# =========================================================

def extract_uniform_frames(mp4_path, num_frames=16):
    cap = cv2.VideoCapture(mp4_path)
    if not cap.isOpened():
        raise RuntimeError(f"无法打开视频: {mp4_path}")

    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    duration = total_frames / fps if fps > 0 else 0

    if total_frames <= 0:
        cap.release()
        return [], {
            "fps": fps,
            "duration_sec": 0,
            "total_frames": 0
        }

    # 等间隔索引（0-based）
    indices = [
        round(i * (total_frames - 1) / (num_frames - 1))
        for i in range(num_frames)
    ]

    frames = []
    for idx in indices:
        cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
        ret, frame = cap.read()
        if ret:
            frames.append(frame)

    cap.release()

    return frames, {
        "fps": round(fps, 3),
        "duration_sec": round(duration, 2),
        "total_frames": total_frames,
        "sample_strategy": f"uniform_{num_frames}"
    }


# =========================================================
# 2. 拼接成长图（从上到下）
# =========================================================

def build_long_image_from_frames(
    frames,
    output_path,
    max_width=1000,
    quality=90
):
    if not frames:
        return False

    images = []
    total_height = 0

    for frame in frames:
        img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))

        if img.width > max_width:
            ratio = max_width / img.width
            img = img.resize(
                (max_width, int(img.height * ratio)),
                Image.LANCZOS
            )

        images.append(img)
        total_height += img.height

    long_img = Image.new("RGB", (max_width, total_height), (255, 255, 255))

    y = 0
    for img in images:
        long_img.paste(img, (0, y))
        y += img.height

    long_img.save(output_path, quality=quality)
    return True


# =========================================================
# 3. 单个 MP4 → long.jpg + meta.json
# =========================================================

def mp4_to_long_image(mp4_path, num_frames=16):
    base_dir = os.path.dirname(mp4_path)
    base_name = os.path.splitext(os.path.basename(mp4_path))[0]

    long_img_path = os.path.join(base_dir, base_name + ".long.jpg")
    meta_path = os.path.join(base_dir, base_name + ".meta.json")

    frames, video_meta = extract_uniform_frames(
        mp4_path,
        num_frames=num_frames
    )

    # 哈希去重（检测是否画面高度重复）
    hashes = [hashlib.md5(f.tobytes()).hexdigest() for f in frames]
    unique_frames = len(set(hashes))

    build_long_image_from_frames(frames, long_img_path)

    meta = {
        "source_mp4": os.path.basename(mp4_path),
        "output_long_image": os.path.basename(long_img_path),
        "fps": video_meta["fps"],
        "duration_sec": video_meta["duration_sec"],
        "total_frames": video_meta["total_frames"],
        "sampled_frames": len(frames),
        "unique_frames": unique_frames,
        "duplicate_ratio": round(
            1 - unique_frames / len(frames), 3
        ) if frames else 1.0,
        "sample_strategy": video_meta["sample_strategy"]
    }

    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    print(f"✅ 生成长图: {long_img_path}")
    print(f"✅ 生成 Meta: {meta_path}")


# =========================================================
# 4. 批量处理（递归扫描所有 MP4）
# =========================================================

def batch_mp4_to_long_images(root_dir, num_frames=16):
    """
    递归扫描 root_dir
    对每一个 mp4：
    - 等间隔抽取 16 帧
    - 拼接成长图
    - 生成 meta.json
    """
    for root, _, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith(".mp4"):
                mp4_path = os.path.join(root, file)
                try:
                    mp4_to_long_image(
                        mp4_path,
                        num_frames=num_frames
                    )
                except Exception as e:
                    print(f"❌ 处理失败 {mp4_path}: {e}")

