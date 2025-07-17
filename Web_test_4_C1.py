import pandas as pd
from pathlib import Path

# ========= 每题评分函数 =========
def score_layout_q1_header_structure(code):
    if code.strip() == "":
        return 0, "未编写任何 HTML 结构"
    feedback = []
    score = 3
    if "<div" in code and "header" in code:
        feedback.append("写出了 header 区域")
    else:
        feedback.append("缺少 header 主体结构")

    if "<img" in code:
        score += 1
        feedback.append("包含图片标签")
    else:
        feedback.append("缺少图片标签")

    if "<p" in code:
        score += 1
        feedback.append("包含文字段落")
    else:
        feedback.append("缺少段落文字")

    return min(score, 5), "，".join(feedback)



def score_layout_q2_content_box_structure(code):
    if code.strip() == "":
        return 0, "未编写任何 HTML 结构"
    score = 3
    feedback = []
    has_content = "class=\"content\"" in code
    box_count = code.count("class=\"box\"")

    if has_content:
        score += 1
        feedback.append("写出了 content 容器")
    else:
        feedback.append("未写出 content 容器")

    if box_count >= 3:
        score += 1
        feedback.append("包含 3 个及以上 box 区域")
    elif box_count > 0:
        feedback.append(f"仅写出 {box_count} 个 box")
    else:
        feedback.append("缺少 box 区域")

    return min(score, 5), "，".join(feedback)


def score_layout_css_block(code: str, keywords: list, full_score: int, section_name: str):
    if code.strip() == "":
        return 0, f"{section_name} 样式完全缺失"

    css = code.lower()
    matched = sum(1 for kw in keywords if kw in css)
    score = full_score if matched == len(keywords) else 3 + matched

    missing = [kw for kw in keywords if kw not in css]
    hit = [kw for kw in keywords if kw in css]
    feedback = f"{section_name} 样式匹配 {matched}/{len(keywords)}；包含：{hit}，缺少：{missing}"

    return min(score, full_score), feedback


# 题3：header 样式
kw_3 = ["header", "height", "130px", "background"]

# 题4：top 样式
kw_4 = ["top", "width: 1000px", "height: 130px", "margin: 0 auto"]

# 题5：box 样式
kw_5 = ["box", "width: 300px", "height: 350px", "float"]

# 题6：imgbox + img 样式
kw_6 = ["imgbox", "width: 300px", "height: 300px", "overflow", "100%"]

# 题7：段落样式
kw_7 = ["box p", "width: 300px", "text-align", "center"]



def grade_page1_layout_from_txt_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 无法读取 {file.name}")
            continue

        s1, c1 = score_layout_q1_header_structure(code)
        s2, c2 = score_layout_q2_content_box_structure(code)
        s3, c3 = score_layout_css_block(code, kw_3, 4, "header")
        s4, c4 = score_layout_css_block(code, kw_4, 4, "top")
        s5, c5 = score_layout_css_block(code, kw_5, 4, "box")
        s6, c6 = score_layout_css_block(code, kw_6, 4, "imgbox/img")
        s7, c7 = score_layout_css_block(code, kw_7, 4, "段落p")

        total = s1 + s2 + s3 + s4 + s5 + s6 + s7

        records.append({
            "文件名": file.name,
            "题1得分": s1, "题1评语": c1,
            "题2得分": s2, "题2评语": c2,
            "题3得分": s3, "题3评语": c3,
            "题4得分": s4, "题4评语": c4,
            "题5得分": s5, "题5评语": c5,
            "题6得分": s6, "题6评语": c6,
            "题7得分": s7, "题7评语": c7,
            "总分30": total
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 评分完成，已导出：{output_excel}")



grade_page1_layout_from_txt_to_excel(
    folder=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_",
    output_excel=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\公管心法Web1班4-1-2024-2025-2_web期末考试_C卷__word_评分结果page1.xlsx"
)

