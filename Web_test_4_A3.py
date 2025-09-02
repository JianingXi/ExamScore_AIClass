import pandas as pd
from pathlib import Path
import re

# === 每题评分逻辑（确保非空最低得 3 分）===

def score_question_11_html_elements(code: str):
    if code.strip() == "":
        return 0, "完全未写网页元素。"
    has_textbox = re.search(r'<input[^>]+type=["\']text["\']', code)
    has_checkbox = re.search(r'<input[^>]+type=["\']checkbox["\']', code)
    has_button = re.search(r'(type=["\']button["\'])|(<button)', code)
    matched = sum([bool(has_textbox), bool(has_checkbox), bool(has_button)])
    if matched == 3:
        return 6, "文本框、复选框和按钮均已正确写出。"
    elif matched == 2:
        return 5, "已写出两个控件，基本完整。"
    elif matched == 1:
        return 4, "仅写出一个控件，内容不全。"
    else:
        return 3, "网页结构不完整但非空，给予起评分。"

def score_question_12_bind_event(code: str):
    if code.strip() == "":
        return 0, "未编写按钮事件。"
    if re.search(r'onclick\s*=\s*["\']ticket\(\)', code):
        return 6, "计算按钮成功绑定 ticket() 事件处理函数。"
    elif "onclick" in code and "ticket" in code:
        return 5, "事件绑定语法略有问题但基本正确。"
    elif "button" in code or "input" in code:
        return 4, "按钮存在但未绑定事件。"
    else:
        return 3, "存在控件但无事件绑定逻辑，最低送分。"

def score_question_13_define_ticket(code: str):
    if code.strip() == "":
        return 0, "未定义 ticket 函数。"
    if "function ticket()" in code and "alert" in code:
        return 6, "ticket 函数定义完整，含弹窗输出。"
    elif "function ticket()" in code:
        return 5, "ticket 函数结构完整但缺少结果展示。"
    elif "ticket" in code:
        return 4, "ticket 函数名出现但定义结构不完整。"
    else:
        return 3, "代码非空但结构不规范，给予基本分。"

def score_question_14_calc_logic(code: str):
    if code.strip() == "":
        return 0, "未编写任何票价计算逻辑。"
    logic_keywords = ["160", "80", "100"]
    matched_price = all(kw in code for kw in logic_keywords)
    has_total = re.search(r'[\+\*]', code) and "alert" in code
    if matched_price and has_total:
        return 6, "票价计算逻辑完整，包含总金额并弹窗显示。"
    elif matched_price:
        return 5, "票价设定正确，但未合并计算或未输出。"
    elif any(kw in code for kw in logic_keywords):
        return 4, "部分票价逻辑存在但不完整。"
    else:
        return 3, "尝试实现逻辑但未正确表达，给予起评分。"

def score_question_15_css_style(code: str):
    if code.strip() == "":
        return 0, "CSS 样式完全缺失。"
    score = 3
    css_keywords = [
        ("width: 250px", "height: 180px"),
        ("margin", "auto"),
        ("border-radius",),
        ("cadetblue",),
        ("input", "width: 50px"),
        ("button", "margin"),
    ]
    matched = 0
    for group in css_keywords:
        if all(g in code for g in group):
            matched += 1
    if matched >= 5:
        score = 6
        comment = "样式完整，页面美观布局合理。"
    elif matched >= 3:
        score = 5
        comment = "大部分样式编写正确，布局基本实现。"
    elif matched >= 1:
        score = 4
        comment = "部分样式编写成功，布局仍需加强。"
    else:
        score = 3
        comment = "CSS 结构不规范但代码非空，给予最低得分。"
    return score, comment

# === 主评分函数 ===

def grade_page3_from_txt_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*page3.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 无法读取：{file.name}, 错误：{e}")
            continue

        s11, c11 = score_question_11_html_elements(code)
        s12, c12 = score_question_12_bind_event(code)
        s13, c13 = score_question_13_define_ticket(code)
        s14, c14 = score_question_14_calc_logic(code)
        s15, c15 = score_question_15_css_style(code)
        total = s11 + s12 + s13 + s14 + s15

        records.append({
            "文件名": file.name,
            "题11得分": s11, "题11评语": c11,
            "题12得分": s12, "题12评语": c12,
            "题13得分": s13, "题13评语": c13,
            "题14得分": s14, "题14评语": c14,
            "题15得分": s15, "题15评语": c15,
            "总分30": total
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 批量评分完成，结果已导出至：{output_excel}")

# === MAIN 批量入口 ===

if __name__ == "__main__":
    exam_file_base_name = [
        "儿生预web1班2-1-2024-2025-2_Web（A卷）-2(word)",
        "临I南山Web2班2-2-2024-2025-2_Web（A卷）-2(word)",
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
            grade_page3_from_txt_to_excel(folder, output_excel)
