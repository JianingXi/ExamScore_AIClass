
import re

def grade_page2_movie_rating(code: str):
    result = []

    # Q1：添加按钮行、跨2列、按钮居中
    tr_match = re.search(r"<tr>.*?<input[^>]*button[^>]*>.*?<input[^>]*reset[^>]*>.*?</tr>", code, re.DOTALL)
    colspan_ok = re.search(r"<td[^>]*colspan=[\"']?2[\"']?", code)
    if tr_match and colspan_ok:
        score1 = 4
        comment1 = "成功添加含按钮的表行，且使用 colspan 合并2列。"
    elif tr_match:
        score1 = 3
        comment1 = "按钮行已添加，但未设置 colspan=2。"
    elif "button" in code or "reset" in code:
        score1 = 2
        comment1 = "按钮未正确放入表格行或格式混乱。"
    else:
        score1 = 2
        comment1 = "未添加所需按钮和表格结构。"
    result.append((score1, comment1))

    # Q2：确认按钮绑定 foo()
    onclick_ok = re.search(r"<input[^>]*onclick=[\"']?foo\(\)", code)
    if onclick_ok:
        score2 = 4
        comment2 = "确认按钮成功绑定 foo() 事件处理函数。"
    elif "foo()" in code:
        score2 = 3
        comment2 = "foo() 函数存在，但按钮未正确绑定事件。"
    else:
        score2 = 2
        comment2 = "按钮未绑定事件或函数名错误。"
    result.append((score2, comment2))

    # Q3：为 #result 添加圆角和边框样式
    style_block = re.search(r"#result\s*{[^}]*}", code, re.DOTALL)
    if style_block:
        s = style_block.group()
        has_radius = "border-radius" in s
        has_border = "border" in s
        if has_radius and has_border:
            score3 = 4
            comment3 = "#result 样式设置完整，含圆角和边框。"
        elif has_radius or has_border:
            score3 = 3
            comment3 = "#result 样式部分正确，建议补全圆角或边框设置。"
        else:
            score3 = 2
            comment3 = "#result 样式缺失，请添加边框与圆角。"
    else:
        score3 = 2
        comment3 = "未设置 #result 样式。"
    result.append((score3, comment3))

    # Q4：foo 函数获取年龄和分级，含判断逻辑
    script_block = re.findall(r"<script[^>]*>(.*?)</script>", code, re.DOTALL)
    js_code = script_block[0] if script_block else ""
    has_get_age = "document.getElementById(\"age\").value" in js_code
    has_get_classify = "document.getElementById(\"classify\").value" in js_code
    has_if = "if" in js_code and ("G" in js_code or "PG" in js_code or "NC-17" in js_code)
    if has_get_age and has_get_classify and has_if:
        score4 = 4
        comment4 = "foo() 函数能正确获取输入值，并包含判断逻辑。"
    elif (has_get_age or has_get_classify) and has_if:
        score4 = 3
        comment4 = "获取部分数据并包含部分判断逻辑。"
    elif "function foo()" in js_code:
        score4 = 2
        comment4 = "函数框架存在，但获取值或判断不完整。"
    else:
        score4 = 2
        comment4 = "未有效实现 foo() 函数核心逻辑。"
    result.append((score4, comment4))

    # Q5：将结果输出到 #result
    writes_result = "document.getElementById(\"result\").value" in js_code
    if writes_result:
        score5 = 4
        comment5 = "结果已正确输出至 #result 框中。"
    elif ".value" in js_code or "#result" in js_code:
        score5 = 3
        comment5 = "部分输出语句存在，可能格式或引用不完整。"
    else:
        score5 = 2
        comment5 = "未将结果输出到指定文本框。"
    result.append((score5, comment5))

    return {
        "文件名": "",
        "题1得分": result[0][0], "题1评语": result[0][1],
        "题2得分": result[1][0], "题2评语": result[1][1],
        "题3得分": result[2][0], "题3评语": result[2][1],
        "题4得分": result[3][0], "题4评语": result[3][1],
        "题5得分": result[4][0], "题5评语": result[4][1],
        "总分20": sum(r[0] for r in result)
    }


import pandas as pd
from pathlib import Path

def batch_grade_page2_movie_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page2.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 跳过：{file.name}，错误：{e}")
            continue

        result = grade_page2_movie_rating(code)
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
        "总分20"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 已导出评分结果至：{output_excel}")




exam_file_base_name = \
    ["临2六中web2班2-3-2024-2025-2_Web期末考试（B卷）(word)",
    ]
html_or_ordered = \
    ["txt_files_1_ordered",
     "txt_files_2_html"
    ]
txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
for name in exam_file_base_name:
    for dir_temp in html_or_ordered:
        fold_1 = txt_base_folder + "\\" + dir_temp + "\\" + name
        file_excel = txt_base_folder + "\\" + dir_temp + "\\" + name + "_评分结果page2.xlsx"
        batch_grade_page2_movie_to_excel(fold_1, file_excel)
