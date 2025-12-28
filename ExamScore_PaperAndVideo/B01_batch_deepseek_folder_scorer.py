import os
import re
import json
import pandas as pd
from dotenv import load_dotenv
import requests

# =========================================================
# 基础配置
# =========================================================

load_dotenv()
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")

API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL_NAME = "deepseek-chat"

SUPPORTED_IMAGE_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".webp")

# =========================================================
# 1. 姓名提取（从文件夹名）
# =========================================================

def extract_name_from_folder(folder_name: str):
    matches = re.findall(r'[\u4e00-\u9fa5]{2,4}', folder_name)
    return matches[0] if matches else ""

# =========================================================
# 2. 读取文件夹内容（文本 + 文件名摘要）
# =========================================================

def collect_folder_materials(folder_path: str):
    texts = []

    for root, _, files in os.walk(folder_path):
        for f in files:
            fp = os.path.join(root, f)

            # txt
            if f.lower().endswith(".txt"):
                try:
                    with open(fp, "r", encoding="utf-8", errors="ignore") as fr:
                        texts.append(fr.read())
                except Exception:
                    pass

            # json
            elif f.lower().endswith(".json"):
                try:
                    with open(fp, "r", encoding="utf-8", errors="ignore") as fr:
                        texts.append(fr.read())
                except Exception:
                    pass

            # image（只记录存在性，不传图）
            elif f.lower().endswith(SUPPORTED_IMAGE_EXT):
                texts.append(f"[图片文件] {f}")

    return "\n".join(texts)[:12000]  # 防止 prompt 过长

# =========================================================
# 3. DeepSeek 评分调用
# =========================================================

def deepseek_score(prompt: str):
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }

    system_prompt = (
        "你是一名高校学术竞赛评审专家。"
        "请根据给定材料进行评分，输出 JSON。"
    )

    user_prompt = f"""
请根据以下材料进行评分（总分100）：

评分维度：
- 作品创新性（30）
- 学术严谨性（30）
- 报告逻辑性（20）
- 表达能力（20）

要求：
1. 每项给整数分
2. 给出 50–80 字中文评语
3. 输出 JSON，格式如下：
{{
  "innovation": int,
  "rigor": int,
  "logic": int,
  "expression": int,
  "comment": "评语"
}}

材料如下：
----------------
{prompt}
----------------
"""

    data = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.4
    }

    resp = requests.post(API_URL, headers=headers, json=data, timeout=60)
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]

    return json.loads(content)

# =========================================================
# 4. 分数区间强制规范（65–95）
# =========================================================

def normalize_score_range(i, r, l, e, min_total=65, max_total=95):
    raw_total = i + r + l + e

    if raw_total == 0:
        return 16, 20, 14, 15, 65

    if min_total <= raw_total <= max_total:
        return i, r, l, e, raw_total

    target = min_total if raw_total < min_total else max_total
    scale = target / raw_total

    i2 = round(i * scale)
    r2 = round(r * scale)
    l2 = round(l * scale)
    e2 = round(e * scale)

    diff = target - (i2 + r2 + l2 + e2)
    i2 += diff  # 补到创新性

    return i2, r2, l2, e2, target

# =========================================================
# 5. 主批处理流程 → Excel
# =========================================================

def batch_score_to_excel(root_dir: str):
    records = []

    for d in os.listdir(root_dir):
        folder = os.path.join(root_dir, d)
        if not os.path.isdir(folder):
            continue

        print(f"▶ 评分中: {d}")

        name = extract_name_from_folder(d)
        materials = collect_folder_materials(folder)

        try:
            result = deepseek_score(materials)

            i, r, l, e, total = normalize_score_range(
                result["innovation"],
                result["rigor"],
                result["logic"],
                result["expression"]
            )

            records.append({
                "姓名": name,
                "文件夹": d,
                "创新性(30)": i,
                "学术严谨性(30)": r,
                "报告逻辑性(20)": l,
                "表达能力(20)": e,
                "总分": total,
                "评语": result["comment"]
            })

        except Exception as ex:
            print(f"❌ 失败: {d} -> {ex}")

    if records:
        df = pd.DataFrame(records)
        out_path = "./FinalScores.xlsx"
        df.to_excel(out_path, index=False)
        print(f"\n✅ 评分完成，已生成：{out_path}")

# =========================================================
# 6. 入口
# =========================================================

if __name__ == "__main__":
    ROOT_DIR = r"./"   # 改成你的根目录
    batch_score_to_excel(ROOT_DIR)
