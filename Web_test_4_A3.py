import pandas as pd
from pathlib import Path

# 题11–15评分函数

def score_question_11_html_elements(code: str):
    if code.strip() == "":
        return 0, "完全未写网页元素。"
    has_textbox = '<input' in code and 'type="text"' in code
    has_checkbox = '<input' in code and 'type="checkbox"' in code
    has_button = '<input' in code and 'type="button"' in code or "<button" in code
    matched = sum([has_textbox, has_checkbox, has_button])
    if matched == 0:
        return 3, "控件基本未写，仅起评分。"
    elif matched == 1:
        return 4, "写出一个控件。"
    elif matched == 2:
        return 5, "写出两个控件，基本结构正确。"
    else:
        return 6, "所有控件写出，结构合理。"

def score_question_12_bind_event(code: str):
    if code.strip() == "":
        return 0, "完全未写按钮绑定事件。"
    if 'onclick="ticket()"' in code or "onclick='ticket()'" in code:
        return 6, "成功绑定 ticket() 事件。"
    elif "onclick" in code and "ticket" in code:
        return 5, "事件绑定基本正确，语法略有问题。"
    elif "<input" in code and "button" in code:
        return 4, "按钮写出但未正确绑定事件。"
    else:
        return 3, "控件或事件缺失，仅起评分。"

def score_question_13_define_ticket(code: str):
    if code.strip() == "":
        return 0, "完全未定义 ticket() 函数。"
    if "function ticket()" in code and "alert" in code:
        return 6, "ticket 函数定义完整，包含弹窗。"
    elif "function ticket()" in code:
        return 5, "函数定义结构完整但缺内容。"
    elif "ticket" in code:
        return 4, "出现函数名但结构不完整。"
    else:
        return 3, "只有提示性文字或空函数体。"

def score_question_14_calc_logic(code: str):
    if code.strip() == "":
        return 0, "未实现任何计算逻辑。"
    logic_ok = all(kw in code for kw in ["160", "80", "100"])
    has_sum = "sum" in code or "+" in code
    if logic_ok and has_sum and "alert" in code:
        return 6, "逻辑完整，结果可弹窗显示。"
    elif logic_ok and has_sum:
        return 5, "逻辑较完整但未展示结果。"
    elif logic_ok:
        return 4, "票价逻辑存在但未合并计算。"
    else:
        return 3, "尝试编写逻辑但不成系统。"

def score_question_15_css_style(code: str):
    if code.strip() == "":
        return 0, "未写任何CSS样式。"
    matched = 0
    if "width: 250px" in code and "height: 180px" in code:
        matched += 1
    if "margin" in code and "auto" in code:
        matched += 1
    if "border-radius" in code:
        matched += 1
    if "cadetblue" in code:
        matched += 1
    if "input" in code and "width: 50px" in code:
        matched += 1
    if "button" in code and "margin" in code:
        matched += 1
    if matched >= 5:
        return 6, "样式完整，符合要求。"
    elif matched >= 3:
        return 5, "大部分样式写出，略有缺失。"
    elif matched >= 1:
        return 4, "部分样式出现但不完整。"
    else:
        return 3, "仅起评分，样式不规范。"

# 主函数：评分并导出为 Excel

def grade_page3_from_txt_to_excel(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for file in folder_path.rglob("*.txt"):
        try:
            code = file.read_text(encoding="utf-8")
        except:
            print(f"⚠️ 无法读取：{file}")
            continue

        s11, c11 = score_question_11_html_elements(code)
        s12, c12 = score_question_12_bind_event(code)
        s13, c13 = score_question_13_define_ticket(code)
        s14, c14 = score_question_14_calc_logic(code)
        s15, c15 = score_question_15_css_style(code)
        total = s11 + s12 + s13 + s14 + s15

        records.append({
            "文件名": file.name,
            "题11得分": s11,
            "题11评语": c11,
            "题12得分": s12,
            "题12评语": c12,
            "题13得分": s13,
            "题13评语": c13,
            "题14得分": s14,
            "题14评语": c14,
            "题15得分": s15,
            "题15评语": c15,
            "总分30": total
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 评分完成，结果已导出至 {output_excel}")


grade_page3_from_txt_to_excel(
    folder=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_",
    output_excel=r"C:\Users\xijia\Desktop\批改web\S03_test_files_html\临I南山Web2班2-2-2024-2025-2_Web_A卷_-2_word_评分结果page3.xlsx"
)
