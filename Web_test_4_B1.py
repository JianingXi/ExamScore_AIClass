import re

def grade_page1_weather_layout_v2(code: str):
    result = []

    # Q1 header 区：包含 top，设置高度40px和背景
    header_present = "<div class=\"header\">" in code
    top_in_header = "class=\"top\"" in code
    style_header = re.search(r"\.header\s*{[^}]*}", code, re.DOTALL)
    header_hits = 0
    if style_header:
        s = style_header.group()
        header_hits += "height: 40px" in s
        header_hits += "background" in s
    if header_present and top_in_header and header_hits >= 1:
        score1 = 5
        comment1 = "header 区结构正确，已设置高度或背景。"
    elif header_present:
        score1 = 4
        comment1 = "header 区存在但结构不完整或样式缺失。"
    else:
        score1 = 3
        comment1 = "未设置 header 区或结构明显缺失。"
    result.append((score1, comment1))

    # Q2 top 区宽高为 900×40，含 logo2.png
    top_style = re.search(r"\.top\s*{[^}]*}", code, re.DOTALL)
    top_hits = 0
    if top_style:
        s = top_style.group()
        top_hits += "width: 900px" in s
        top_hits += "height: 40px" in s
    has_logo = "logo2.png" in code
    if top_hits == 2 and has_logo:
        score2 = 5
        comment2 = "top 区宽高设置正确，已包含 logo2.png 图片。"
    elif top_hits >= 1 and has_logo:
        score2 = 4
        comment2 = "top 样式部分设置正确，已包含 logo 图片。"
    else:
        score2 = 3
        comment2 = "top 区样式设置不完整或图片缺失。"
    result.append((score2, comment2))

    # Q3 banner 区宽900px、居中，含图片和段落文字
    banner_div = "class=\"banner\"" in code
    banner_img = "banner.jpg" in code
    banner_text = "<p>" in code
    if banner_div and banner_img and banner_text:
        score3 = 5
        comment3 = "banner 区结构、图像、文字完整。"
    elif banner_div and (banner_img or banner_text):
        score3 = 4
        comment3 = "banner 内容部分完成，建议检查图像和文字。"
    else:
        score3 = 3
        comment3 = "banner 区结构或内容缺失，请补全。"
    result.append((score3, comment3))

    # Q4 content 区包含 left 和 right，设置下边框
    content_ok = "class=\"content\"" in code
    left_ok = "class=\"left\"" in code
    right_ok = "class=\"right\"" in code
    border_ok = "border-bottom" in code or "border" in code
    if content_ok and left_ok and right_ok and border_ok:
        score4 = 5
        comment4 = "content 区结构、子区和下边框设置齐全。"
    elif content_ok and (left_ok or right_ok):
        score4 = 4
        comment4 = "content 区结构存在，建议补充下边框或补全子区。"
    else:
        score4 = 3
        comment4 = "content 区结构不完整，未体现两个子区或样式缺失。"
    result.append((score4, comment4))

    # Q5 left 区宽250高140，包含立夏1.jpg 图片
    left_style = re.search(r"\.left\s*{[^}]*}", code, re.DOTALL)
    left_hits = 0
    if left_style:
        s = left_style.group()
        left_hits += "width: 250px" in s
        left_hits += "height: 140px" in s
    has_small_img = "立夏1.jpg" in code
    if left_hits == 2 and has_small_img:
        score5 = 5
        comment5 = "left 区结构正确，图像、尺寸设置完整。"
    elif left_hits >= 1 and has_small_img:
        score5 = 4
        comment5 = "left 区部分设置正确，建议检查宽高或图像。"
    else:
        score5 = 3
        comment5 = "left 区内容或样式缺失。"
    result.append((score5, comment5))

    # Q6 right 区宽600，包含 h3 和 p
    right_style = re.search(r"\.right\s*{[^}]*}", code, re.DOTALL)
    right_width = "width: 600px" in right_style.group() if right_style else False
    h3_ok = "<h3>" in code
    p_ok = "<p>" in code
    if right_width and h3_ok and p_ok:
        score6 = 5
        comment6 = "right 区内容与宽度设置正确。"
    elif h3_ok or p_ok:
        score6 = 4
        comment6 = "right 区内容部分存在，建议补充宽度样式或结构。"
    else:
        score6 = 3
        comment6 = "right 区结构或样式明显缺失。"
    result.append((score6, comment6))

    return {
        "文件名": "",
        "题1得分": result[0][0], "题1评语": result[0][1],
        "题2得分": result[1][0], "题2评语": result[1][1],
        "题3得分": result[2][0], "题3评语": result[2][1],
        "题4得分": result[3][0], "题4评语": result[3][1],
        "题5得分": result[4][0], "题5评语": result[4][1],
        "题6得分": result[5][0], "题6评语": result[5][1],
        "总分30": sum(r[0] for r in result)
    }



import pandas as pd
from pathlib import Path

def batch_grade_page1_weather_v2_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 跳过：{file.name}（错误：{e}）")
            continue

        result = grade_page1_weather_layout_v2(code)
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
        "题6得分", "题6评语",
        "总分30"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，Excel已导出：{output_excel}")



def main():
    input_folder = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_"
    output_excel = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_评分结果page1.xlsx"
    batch_grade_page1_weather_v2_to_excel(input_folder, output_excel)

if __name__ == "__main__":
    main()

