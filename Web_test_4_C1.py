import re
import pandas as pd
from pathlib import Path


# 每题评分函数
def score_q5(code: str, keywords: list, label: str):
    if not code.strip():
        return 0, f"{label}：未作答"

    hit = [kw for kw in keywords if kw.lower() in code.lower()]
    missed = [kw for kw in keywords if kw.lower() not in code.lower()]

    if len(hit) == 0:
        return 3, f"{label}：要素缺失"
    elif len(hit) == len(keywords):
        return 5, f"{label}：全部要素正确"
    else:
        return 4, f"{label}：部分缺失 -> 缺少：{missed}"

def score_q4(code: str, keywords: list, label: str):
    if not code.strip():
        return 0, f"{label}：未作答"

    hit = [kw for kw in keywords if kw.lower() in code.lower()]
    missed = [kw for kw in keywords if kw.lower() not in code.lower()]

    if len(hit) == 0:
        return 2, f"{label}：要素缺失"
    elif len(hit) == len(keywords):
        return 4, f"{label}：全部要素正确"
    else:
        return 3, f"{label}：部分缺失 -> 缺少：{missed}"


# 七个评分点关键词（源自注释）
Q_KEYWORDS = [
    (["<div", "header", "<img", "<p"], "header结构"),
    (["top", "logo", "<p"], "top结构"),
    (["logo", "width: 200px", "height: 130px", "margin"], "logo样式"),
    (["content", "box"], "content区结构（含3个box）"),
    (["box", "imgbox", "<img", "<p"], "box结构"),
    (["imgbox", "width: 300px", "height: 300px", "overflow", "100%"], "imgbox样式"),
    (["box p", "width: 300px", "text-align", "center"], "box p段落样式"),
]


# 主批改函数
def grade_page1_batch(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page1.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 无法读取 {file}")
            continue

        result = {"文件名": file.name}
        total = 0
        for idx, (keywords, label) in enumerate(Q_KEYWORDS, start=1):
            if idx <= 2:
                score, comment = score_q5(code, keywords, label)
            else:
                score, comment = score_q4(code, keywords, label)
            result[f"题{idx}得分"] = score
            result[f"题{idx}评语"] = comment
            total += score

        result["总分30"] = total
        records.append(result)

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，结果导出至：{output_excel}")




exam_file_base_name = \
    ["精食药临药web1班4-2-2024-2025-2_web期末考试（C卷）(word)",
     "公管心法Web1班4-1-2024-2025-2_web期末考试（C卷）(word)",
    ]
html_or_ordered = \
    ["txt_files_1_ordered",
     "txt_files_2_html"
    ]
txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
for name in exam_file_base_name:
    for dir_temp in html_or_ordered:
        fold_1 = txt_base_folder + "\\" + dir_temp + "\\" + name
        file_excel = txt_base_folder + "\\" + dir_temp + "\\" + name + "_评分结果page1.xlsx"
        grade_page1_batch(fold_1, file_excel)
