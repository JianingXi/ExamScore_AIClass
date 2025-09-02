import re

def grade_page3_font_change_detailed(code: str):
    result = []

    # Q11：6个li+onmouseover
    li_tags = re.findall(r'<li\s+onmouseover=', code)
    li_count = len(li_tags)
    if li_count >= 6:
        score11 = 6
        comment11 = "已正确编写6个<li>并绑定onmouseover事件。"
    elif li_count >= 1:
        score11 = 4
        comment11 = f"仅编写了{li_count}个<li>，请补全剩余项。"
    elif "<li" in code:
        score11 = 3
        comment11 = "部分<li>标签未绑定onmouseover事件，请检查。"
    else:
        score11 = 3
        comment11 = "未编写<li>列表项，建议参考示例补全6个选项并绑定事件。"
    result.append((score11, comment11))

    # Q12：.box样式
    box_style = re.search(r"\.box\s*{([^}]+)}", code, re.DOTALL)
    box_hits = 0
    checklist = ["width: 800px", "height: 450px", "border", "border-radius", "font-size: 80px"]
    if box_style:
        css = box_style.group(1)
        box_hits = sum(1 for item in checklist if item in css)
    if box_hits >= 4:
        score12 = 6
        comment12 = "已设置.box大部分核心样式，页面结构完整。"
    elif box_hits >= 2:
        score12 = 4
        comment12 = f".box样式设置不完整，仅正确设置了{box_hits}/5项。"
    else:
        score12 = 3
        comment12 = ".box样式基本未设置，请根据注释完成宽高、边框和字体等样式。"
    result.append((score12, comment12))

    # Q13：.nav、ul 居中，去掉列表符号
    nav_style = re.search(r"\.nav\s*{([^}]+)}", code, re.DOTALL)
    ul_style = re.search(r"ul\s*{([^}]+)}", code, re.DOTALL)
    nav_hits = 0
    if nav_style and "text-align: center" in nav_style.group(1):
        nav_hits += 1
    if ul_style:
        s = ul_style.group(1)
        nav_hits += sum([
            "list-style: none" in s or "list-style-type: none" in s,
            "margin: 0 auto" in s or "text-align: center" in s,
            "width" in s
        ])
    if nav_hits >= 3:
        score13 = 6
        comment13 = "ul与nav居中显示、去除列表符号设置正确。"
    elif nav_hits >= 1:
        score13 = 4
        comment13 = f"ul或nav样式部分正确，建议补充设置居中、宽度与list-style。"
    else:
        score13 = 3
        comment13 = "未正确设置nav与ul样式，页面排版可能不整齐。"
    result.append((score13, comment13))

    # Q14：li的宽、居中、圆角、横排
    li_style = re.search(r"\.nav\s+ul\s+li\s*{([^}]+)}", code, re.DOTALL)
    li_hits = 0
    if li_style:
        s = li_style.group(1)
        li_hits = sum([
            "width: 150px" in s,
            "text-align: center" in s,
            "border-radius" in s,
            "float" in s or "display: inline-block" in s
        ])
    if li_hits >= 3:
        score14 = 6
        comment14 = "li样式设置齐全，已实现横排居中显示。"
    elif li_hits >= 1:
        score14 = 4
        comment14 = f"li样式设置部分正确，当前命中 {li_hits}/4 项，建议完善圆角与布局。"
    else:
        score14 = 3
        comment14 = "li缺少样式，页面可能未水平排列或显示异常。"
    result.append((score14, comment14))

    # Q15：changeText函数
    js = re.findall(r"<script[^>]*>(.*?)</script>", code, re.DOTALL)
    js_code = js[0] if js else ""
    has_func = "function changeText" in js_code
    has_text = "innerHTML" in js_code or "textContent" in js_code
    has_color = "style.color" in js_code or ".color" in js_code
    hits = sum([has_func, has_text, has_color])
    if hits >= 2:
        score15 = 6
        comment15 = "changeText函数基本实现，能修改文字和颜色。"
    elif hits == 1:
        score15 = 4
        comment15 = "changeText函数结构不完整，建议检查文字与颜色设置是否缺失。"
    else:
        score15 = 3
        comment15 = "未实现changeText函数或函数无效，页面交互缺失。"
    result.append((score15, comment15))

    return {
        "文件名": "",
        "题11得分": result[0][0], "题11评语": result[0][1],
        "题12得分": result[1][0], "题12评语": result[1][1],
        "题13得分": result[2][0], "题13评语": result[2][1],
        "题14得分": result[3][0], "题14评语": result[3][1],
        "题15得分": result[4][0], "题15评语": result[4][1],
        "总分30": sum(r[0] for r in result)
    }


import pandas as pd
from pathlib import Path

def batch_grade_page3_fontchange_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page3.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"❌ 读取失败：{file.name}，原因：{e}")
            continue

        result = grade_page3_font_change_detailed(code)
        result["文件名"] = file.name
        records.append(result)

    df = pd.DataFrame(records)
    df = df[[
        "文件名",
        "题11得分", "题11评语",
        "题12得分", "题12评语",
        "题13得分", "题13评语",
        "题14得分", "题14评语",
        "题15得分", "题15评语",
        "总分30"
    ]]
    df.to_excel(output_excel, index=False)
    print(f"✅ 批改完成，结果导出至：{output_excel}")





if __name__ == "__main__":
    exam_file_base_name = [
        "护理检验web1班4-3-2024-2025-2_Web期末考试（护理、检验）(word)",
        "口麻影Web1班4-4-2024-2025-2_Web期末考试（影像、麻醉、口腔）(word)",
    ]
    html_or_ordered = [
        "txt_files_1_ordered",
        "txt_files_2_html"
    ]
    txt_base_folder = r"C:\Users\xijia\Desktop\批改web\S02_TXT"
    for name in exam_file_base_name:
        for subdir in html_or_ordered:
            folder = fr"{txt_base_folder}\{subdir}\{name}"
            output_excel = fr"{folder}_评分结果page3.xlsx"
            batch_grade_page3_fontchange_to_excel(folder, output_excel)

