import re
import pandas as pd
from pathlib import Path

# ========== 评分函数 ==========
def score_question_1_structure(html_code: str):
    if html_code.strip() == "":
        return 0, "完全空白文件，无法评分。"
    expected_ids = ['header', 'logo_img', 'logo_title', 'navbar', 'content']
    matched = sum(1 for i in expected_ids if f'id="{i}"' in html_code or f"id='{i}'" in html_code)
    if matched <= 2:
        return 3, "部分结构写出，起评分。"
    elif matched <= 4:
        return 5, "大部分结构正确，略有缺漏。"
    else:
        return 6, "结构完整，命名规范。"

def score_question_2_content(html_code: str):
    if html_code.strip() == "":
        return 0, "完全空白文件，无法评分。"
    score = 3
    if "logo_img" in html_code and re.search(r'<img.*?>', html_code):
        score += 1
    if "logo_title" in html_code and re.search(r'<div[^>]*logo_title[^>]*>.*?</div>', html_code, re.DOTALL):
        score += 1
    if "navbar" in html_code and "<ul" in html_code and "<li" in html_code and "<a " in html_code:
        score += 1
    if "content" in html_code and "<h2" in html_code and "<hr" in html_code and "<p>" in html_code:
        score += 1
    return min(score, 6), "内容结构整体评估完成。"

def score_question_3_layout_css(html_code: str):
    if html_code.strip() == "":
        return 0, "完全空白文件，无法评分。"
    center_keywords = ["margin: 0 auto", "text-align: center", "display: block", "width"]
    blocks = ["#header", "#navbar", "#content"]
    matched = sum(1 for block in blocks if any(block in html_code and kw in html_code for kw in center_keywords))
    if matched == 0:
        return 3, "未体现居中样式，仅给起评分。"
    elif matched == 1:
        return 4, "部分模块居中处理。"
    elif matched == 2:
        return 5, "大部分布局正确，仍有缺失。"
    else:
        return 6, "三大模块均居中，布局规范。"

def score_question_4_navbar_css(html_code: str):
    if html_code.strip() == "":
        return 0, "完全空白文件，无法评分。"
    score = 3
    if "list-style: none" in html_code:
        score += 1
    if "display: inline" in html_code or "float: left" in html_code:
        score += 1
    if "a" in html_code and "color: #fff" in html_code and "text-decoration: none" in html_code:
        score += 1
    if ":hover" in html_code and "#f35252" in html_code:
        score += 1
    return min(score, 6), "导航栏样式评分完成。"

def score_question_5_content_css(html_code: str):
    if html_code.strip() == "":
        return 0, "完全空白文件，无法评分。"
    matched = 0
    if "border" in html_code and "radius" in html_code:
        matched += 1
    if "font-size" in html_code or "line-height" in html_code:
        matched += 1
    if "text-align" in html_code:
        matched += 1
    if "img" in html_code and ("width" in html_code or "height" in html_code):
        matched += 1
    if matched == 0:
        return 3, "未体现主要样式，起评分。"
    elif matched <= 2:
        return 4, "部分样式写出，略显不足。"
    elif matched == 3:
        return 5, "大部分样式完成。"
    else:
        return 6, "样式完整清晰。"

# ========== 主函数 ==========
def grade_all_html_txts_in_folder(folder: str, output_excel: str):
    base_path = Path(folder)
    records = []

    for txt_file in base_path.rglob("*.txt"):
        try:
            html_code = txt_file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 无法读取：{txt_file}")
            continue

        s1, c1 = score_question_1_structure(html_code)
        s2, c2 = score_question_2_content(html_code)
        s3, c3 = score_question_3_layout_css(html_code)
        s4, c4 = score_question_4_navbar_css(html_code)
        s5, c5 = score_question_5_content_css(html_code)
        total = s1 + s2 + s3 + s4 + s5

        records.append({
            "文件名": txt_file.name,
            "结构分": s1,
            "结构评语": c1,
            "内容分": s2,
            "内容评语": c2,
            "布局分": s3,
            "布局评语": c3,
            "导航CSS分": s4,
            "导航CSS评语": c4,
            "内容CSS分": s5,
            "内容CSS评语": c5,
            "总分": total,
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 评分完成，共处理 {len(df)} 份试卷。结果保存于：{output_excel}")

grade_all_html_txts_in_folder(
    folder=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_",
    output_excel=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\儿生预web1班2-1-2024-2025-2_Web_A卷_-2_word_评分结果page1.xlsx"
)
