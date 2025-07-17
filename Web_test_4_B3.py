import re

def grade_page3_season_toggle(code: str):
    result = []

    # Q1: <ul> 列表 + li 鼠标事件绑定 foo(n)
    li_event = re.findall(r"<li[^>]*onmouseover\s*=\s*[\"']?foo\(\d\)", code)
    if len(li_event) >= 4:
        score1 = 5
        comment1 = "成功设置4个 li，并绑定鼠标事件 foo(n)。"
    elif len(li_event) >= 1:
        score1 = 4
        comment1 = f"部分 li 列表项绑定了事件（共 {len(li_event)} 项），建议补全。"
    elif "<li" in code:
        score1 = 2
        comment1 = "li 存在但未绑定事件，建议添加 onmouseover='foo(n)'。"
    else:
        score1 = 2
        comment1 = "未添加 ul/li 列表结构和交互函数。"
    result.append((score1, comment1))

    # Q2: content 中添加 picbox 和 p 段落（含 id）
    has_img = re.search(r"<img[^>]*id\s*=\s*[\"']?pic[\"']?", code)
    has_p = re.search(r"<p[^>]*id\s*=\s*[\"']?info[\"']?", code)
    in_content = "class=\"content\"" in code
    if has_img and has_p and in_content:
        score2 = 5
        comment2 = "content 区已添加图片与文字，id 使用规范。"
    elif has_img or has_p:
        score2 = 4
        comment2 = "已添加图片或文字，但结构或 id 使用不完整。"
    else:
        score2 = 2
        comment2 = "未补充 content 区图片和文字内容。"
    result.append((score2, comment2))

    # Q3: li 样式 + 鼠标悬停变色
    style_li = re.search(r"li\s*{[^}]*}", code, re.DOTALL)
    style_hover = re.search(r"li:hover\s*{[^}]*}", code, re.DOTALL)
    base_ok = 0
    hover_ok = 0
    if style_li:
        s = style_li.group()
        base_ok = sum([
            "width: 150px" in s,
            "height: 40px" in s,
            "border-radius" in s,
            "box-shadow" in s or "shadow" in s
        ])
    if style_hover:
        h = style_hover.group()
        hover_ok = "color" in h and "background" in h
    if base_ok >= 3 and hover_ok:
        score3 = 4
        comment3 = "li 样式设置齐全，包含圆角、阴影和悬停变色效果。"
    elif base_ok >= 2:
        score3 = 3
        comment3 = "li 样式部分设置正确，建议补充悬停样式。"
    else:
        score3 = 2
        comment3 = "li 样式设置不完整或缺失。"
    result.append((score3, comment3))

    # Q4: picbox 宽高 800×440 且图片 100%
    style_picbox = re.search(r"\.picbox\s*{[^}]*}", code, re.DOTALL)
    ok = 0
    if style_picbox:
        s = style_picbox.group()
        ok += "width: 800px" in s
        ok += "height: 440px" in s
    img_percent = re.search(r"img\s*{[^}]*width:\s*100%", code)
    if ok == 2 and img_percent:
        score4 = 4
        comment4 = "picbox 样式设置正确，图片适配良好。"
    elif ok >= 1:
        score4 = 3
        comment4 = "picbox 样式部分设置，建议补全宽高与图片自适应。"
    else:
        score4 = 2
        comment4 = "picbox 或图片样式设置缺失。"
    result.append((score4, comment4))

    # Q5: content 区下 p 标签样式与图片对齐
    p_style = re.search(r"\.content\s+p\s*{[^}]*}", code, re.DOTALL)
    if p_style:
        score5 = 4
        comment5 = "content 区 p 样式已设置，排版效果较好。"
    elif "<p" in code:
        score5 = 3
        comment5 = "p 段落存在但未设置样式。"
    else:
        score5 = 2
        comment5 = "p 段落或样式均未设置。"
    result.append((score5, comment5))

    # Q6: 定义图片数组与颜色数组
    js = re.findall(r"<script[^>]*>(.*?)</script>", code, re.DOTALL)
    script = js[0] if js else ""
    has_pics = "pics" in script and "[" in script
    has_colors = "colors" in script and "[" in script
    if has_pics and has_colors:
        score6 = 4
        comment6 = "已正确定义图片名数组和颜色数组。"
    elif has_pics or has_colors:
        score6 = 3
        comment6 = "只定义了 pics 或 colors，建议补全。"
    else:
        score6 = 2
        comment6 = "未正确定义所需数组。"
    result.append((score6, comment6))

    # Q7: 实现 foo(id) 更换图片文字和颜色
    has_foo = re.search(r"function\s+foo\(", script)
    updates_pic = "document.getElementById(\"pic\").src" in script
    updates_text = "document.getElementById(\"info\").innerHTML" in script
    updates_color = "style.color" in script
    if has_foo and updates_pic and updates_text and updates_color:
        score7 = 4
        comment7 = "foo(id) 实现完整，能切换图片、文字与颜色。"
    elif has_foo and (updates_pic or updates_text):
        score7 = 3
        comment7 = "foo(id) 部分功能实现，建议补全颜色或内容切换。"
    else:
        score7 = 2
        comment7 = "未正确实现切换逻辑。"
    result.append((score7, comment7))

    return {
        "文件名": "",
        "题1得分": result[0][0], "题1评语": result[0][1],
        "题2得分": result[1][0], "题2评语": result[1][1],
        "题3得分": result[2][0], "题3评语": result[2][1],
        "题4得分": result[3][0], "题4评语": result[3][1],
        "题5得分": result[4][0], "题5评语": result[4][1],
        "题6得分": result[5][0], "题6评语": result[5][1],
        "题7得分": result[6][0], "题7评语": result[6][1],
        "总分30": sum(r[0] for r in result)
    }





import pandas as pd
from pathlib import Path

def batch_grade_page3_season_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 跳过：{file.name}，错误：{e}")
            continue

        result = grade_page3_season_toggle(code)
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
        "题7得分", "题7评语",
        "总分30"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 成绩表已导出：{output_excel}")




def main():
    input_folder = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_"
    output_excel = r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临2六中web2班2-3-2024-2025-2_Web期末考试_B卷__word_评分结果page3.xlsx"
    batch_grade_page3_season_to_excel(input_folder, output_excel)

if __name__ == "__main__":
    main()
