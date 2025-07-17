import re

def grade_page1_layout_html_css(code: str):
    result = []

    # ---------- Q1: 结构嵌套 ----------
    q1_score = 0
    required_tags = [".header", ".content", ".box", ".footer"]
    for tag in required_tags:
        if tag.replace(".", "class=") in code or tag in code:
            q1_score += 1
    score1 = 6 if q1_score >= 3 else 3
    comment1 = f"结构包含 {q1_score}/4 个必要图层（.header/.content/.box/.footer）"

    # ---------- Q2: HTML内容补全 ----------
    q2_hits = 0
    if "lsqs.jpg" in code: q2_hits += 1
    if "绿水青山就是金山银山" in code: q2_hits += 1
    if "shengtai.jpg" in code and "生态优先" in code: q2_hits += 1
    if "ziran.jpg" in code and "和谐共生" in code: q2_hits += 1
    if "© 2025" in code: q2_hits += 1
    score2 = 6 if q2_hits >= 4 else 3
    comment2 = f"已补全 {q2_hits}/5 项内容：图片与文字"

    # ---------- Q3: header/footer 样式 ----------
    header_style = re.search(r"\.header\s*{[^}]+}", code, re.DOTALL)
    footer_style = re.search(r"\.footer\s*{[^}]+}", code, re.DOTALL)
    h_score = sum([
        "background" in header_style.group() if header_style else 0,
        "text-align" in header_style.group() if header_style else 0,
        "padding" in header_style.group() if header_style else 0
    ])
    f_score = sum([
        "background" in footer_style.group() if footer_style else 0,
        "color" in footer_style.group() if footer_style else 0,
        "font-size" in footer_style.group() if footer_style else 0
    ])
    score3 = 6 if h_score + f_score >= 5 else 3
    comment3 = f"样式设置正确项 {h_score+f_score}/6（header/footer）"

    # ---------- Q4: header img 和 header-title ----------
    css = re.findall(r"<style[^>]*>(.*?)</style>", code, re.DOTALL)
    css_code = css[0] if css else ""
    hits_q4 = sum([
        "header img" in css_code,
        "max-width" in css_code,
        "border-radius" in css_code,
        "box-shadow" in css_code,
        "font-size" in css_code,
        "color" in css_code
    ])
    score4 = 6 if hits_q4 >= 5 else 3
    comment4 = f"图片与标题样式设置正确项：{hits_q4}/6"

    # ---------- Q5: .box 与 box img 样式 ----------
    hits_q5 = sum([
        "box" in css_code,
        "width: 45%" in css_code,
        "background" in css_code,
        "border-radius" in css_code,
        "float" in css_code,
        "padding" in css_code,
        "box img" in css_code,
        "margin-bottom" in css_code
    ])
    score5 = 6 if hits_q5 >= 6 else 3
    comment5 = f".box 样式设置正确项：{hits_q5}/8"

    return {
        "文件名": "",
        "题1得分": score1, "题1评语": comment1,
        "题2得分": score2, "题2评语": comment2,
        "题3得分": score3, "题3评语": comment3,
        "题4得分": score4, "题4评语": comment4,
        "题5得分": score5, "题5评语": comment5,
        "总分30": sum([score1, score2, score3, score4, score5])
    }


import pandas as pd
from pathlib import Path

def batch_grade_page1_layout_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 无法读取 {file.name}，错误：{e}")
            continue

        result = grade_page1_layout_html_css(code)
        result["文件名"] = file.name
        records.append(result)

    df = pd.DataFrame(records)
    df = df[[
        "文件名",
        "题1得分", "题1评语",
        "题2得分", "题2评语",
        "题3得分", "题3评语",
        "题4得分", "题4评语",
        "题5得分", "题5评语",
        "总分30"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，结果已保存至：{output_excel}")



input_folder = r"C:\Users\xijia\Desktop\批改web\S03_test_files_ordered\口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_"  # ← 替换为你的实际文件夹
output_excel = r"C:\Users\xijia\Desktop\批改web\S03_test_files_ordered\口麻影Web1班4-4-2024-2025-2_Web期末考试_影像_麻醉_口腔__word_评分结果page1.xlsx"                 # ← 导出路径

batch_grade_page1_layout_to_excel(input_folder, output_excel)
