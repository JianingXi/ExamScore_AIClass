# B01_batch_deepseek_folder_scorer.py

import os
import re
import json
import time
import requests
import pandas as pd
from dotenv import load_dotenv

from ExamScore_PaperAndVideo.B01_scoring_prompts import (
    STAGE1_SYSTEM_PROMPT,
    STAGE2_SYSTEM_PROMPT,
    COMMENT_ONLY_SYSTEM_PROMPT,
    build_stage1_user_prompt,
    build_stage2_user_prompt,
    build_comment_only_prompt
)

# =========================================================
# 基础配置
# =========================================================

load_dotenv()
API_KEY = os.getenv("DEEPSEEK_API_KEY")

API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL = "deepseek-chat"

SUPPORTED_TEXT = (".txt", ".json")
SUPPORTED_IMAGE = (".jpg", ".jpeg", ".png", ".bmp", ".webp")

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

# =========================================================
# 工具函数
# =========================================================

def extract_name(folder_name: str) -> str:
    m = re.findall(r'[\u4e00-\u9fa5]{2,4}', folder_name)
    return m[0] if m else folder_name

def safe_parse_json(text: str):
    if not text:
        raise ValueError("Empty response")

    text = text.strip()
    if text.startswith("{"):
        return json.loads(text)

    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e != -1:
        return json.loads(text[s:e + 1])

    raise ValueError("No valid JSON")

def call_llm(system_prompt, user_prompt, temperature=0.3, retry=2):
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": temperature
    }

    for i in range(retry + 1):
        try:
            r = requests.post(API_URL, headers=HEADERS, json=payload, timeout=60)
            r.raise_for_status()
            return safe_parse_json(
                r.json()["choices"][0]["message"]["content"]
            )
        except Exception:
            if i == retry:
                raise
            time.sleep(2)

# =========================================================
# Stage 1：单文件压缩理解
# =========================================================

def summarize_single_file(file_path: str) -> dict:
    ext = os.path.splitext(file_path)[1].lower()

    if ext in SUPPORTED_TEXT:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            material = f.read()[:4000]
    elif ext in SUPPORTED_IMAGE:
        material = f"图片文件：{os.path.basename(file_path)}，用于科研成果或学术汇报展示。"
    else:
        material = f"文件名：{os.path.basename(file_path)}（无法解析内容）"

    return call_llm(
        STAGE1_SYSTEM_PROMPT,
        build_stage1_user_prompt(material)
    )

# =========================================================
# Stage 2：作品级评分（锁定评分模板）
# =========================================================

def score_whole_work(summaries: list, locked_schema: list | None):
    merged = json.dumps(summaries, ensure_ascii=False)
    user_prompt = build_stage2_user_prompt(
        merged,
        locked_schema=locked_schema
    )

    raw = call_llm(STAGE2_SYSTEM_PROMPT, user_prompt)
    scores = raw["scores"]
    comment = raw["comment"]
    total = sum(scores.values())
    return scores, total, comment

# =========================================================
# 主流程（一次运行 = 一个批次）
# =========================================================

def batch_two_stage_score(root_dir: str):
    records = []
    locked_schema = None
    drift_warnings = []

    for folder in os.listdir(root_dir):
        folder_path = os.path.join(root_dir, folder)
        if not os.path.isdir(folder_path):
            continue

        print(f"▶ 评分中: {folder}")
        name = extract_name(folder)

        summary_dir = os.path.join(folder_path, "summaries")
        os.makedirs(summary_dir, exist_ok=True)

        summaries = []

        # ---------- Stage 1 ----------
        for f in os.listdir(folder_path):
            fp = os.path.join(folder_path, f)
            if not os.path.isfile(fp):
                continue
            if not f.lower().endswith(SUPPORTED_TEXT + SUPPORTED_IMAGE):
                continue

            cache = os.path.join(summary_dir, f + ".summary.json")
            if os.path.exists(cache):
                summaries.append(json.load(open(cache, encoding="utf-8")))
                continue

            try:
                s = summarize_single_file(fp)
                s["file"] = f
                summaries.append(s)
                json.dump(s, open(cache, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"  ⚠ 跳过文件: {f} -> {e}")

        if not summaries:
            continue

        # ---------- Stage 2 ----------
        scores, total, comment = score_whole_work(summaries, locked_schema)

        current_keys = list(scores.keys())

        if locked_schema is None:
            locked_schema = current_keys
        else:
            if set(current_keys) != set(locked_schema):
                drift_warnings.append({
                    "folder": folder,
                    "missing": list(set(locked_schema) - set(current_keys)),
                    "new": list(set(current_keys) - set(locked_schema))
                })

        row = {
            "姓名": name,
            "文件夹": folder,
            "总分": total,
            "评语": comment
        }

        # 🔒 只按 locked_schema 顺序展开
        for k in locked_schema:
            row[k] = scores.get(k, "")

        records.append(row)

    # ---------- 输出 ----------
    if drift_warnings:
        warn_path = os.path.join(root_dir, "ScoreItemDriftWarning.json")
        json.dump(drift_warnings, open(warn_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"\n⚠️ 评分项命名漂移报警（未改分）：{warn_path}")

    if records:
        out = os.path.join(root_dir, "FinalScores.xlsx")
        pd.DataFrame(records).to_excel(out, index=False)
        print(f"\n✅ 已生成 {out}")
