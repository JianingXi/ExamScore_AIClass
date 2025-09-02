
import pandas as pd
from pathlib import Path
import re

# === 各题评分函数（送分加强版） ===

def score_question_6_table_structure(code: str):
    if code.strip() == "":
        return 0, "完全未写表格结构。"
    row_count = code.count("<tr>")
    colspan_2_count = code.count('colspan="2"') + code.count("colspan='2'")
    if row_count >= 4 and colspan_2_count >= 2:
        return 4, "表格结构完整，首尾行正确设置跨列合并。"
    elif row_count >= 3 and colspan_2_count >= 1:
        return 3.5, "表格结构部分正确，部分行或合并设置遗漏。"
    elif row_count >= 2:
        return 3, "表格结构不完整，但写了一定内容。"
    else:
        return 3, "表格结构不完整，但写了一定内容。"

def score_question_7_table_content(code: str):
    if code.strip() == "":
        return 0, "完全未写表格内容。"
    matched = 0
    if "小区停车" in code and "时长" in code and "收费" in code:
        matched += 1
    if re.search(r'<input[^>]+type=["\']text["\']', code):
        matched += 1
    if re.search(r'<input[^>]+type=["\']button["\']', code):
        matched += 1
    if matched == 3:
        return 4, "插入文字、文本框、按钮均正确。"
    elif matched == 2:
        return 3.5, "控件或文字部分缺失，仍较完整。"
    elif matched == 1:
        return 3, "仅实现部分控件或文字内容。"
    else:
        return 2, "实现控件或文字内容不足。"

def score_question_8_button_event(code: str):
    if code.strip() == "":
        return 0, "完全未写按钮或事件绑定。"
    if 'onclick="parking()"' in code or "onclick='parking()'" in code:
        return 4, "按钮正确绑定了 parking() 函数事件。"
    elif "onclick" in code and "parking" in code:
        return 3.5, "事件绑定基本正确，语法略有瑕疵。"
    elif "<input" in code and "button" in code:
        return 3, "写了按钮，但未正确绑定事件。"
    else:
        return 2, "按钮和未正确绑定事件不足。"

def score_question_9_function_defined(code: str):
    if code.strip() == "":
        return 0, "完全未定义函数。"
    if "function parking()" in code:
        body = code.split("function parking()")[1]
        if "{" in body and "}" in body:
            return 4, "函数定义完整。"
        else:
            return 3.5, "函数定义语法不完整或缺失。"
    elif "parking" in code:
        return 3, "函数名存在但结构不完整。"
    else:
        return 2, "函数名和结构均不足。"

def score_question_10_parking_logic(code: str):
    if code.strip() == "":
        return 0, "完全未写逻辑。"
    keywords = ["15", "24", "Math", "parseInt", "parseFloat", "ceil", "收费", "if"]
    match_count = sum(1 for k in keywords if k in code)
    if match_count >= 5:
        return 4, "计费逻辑完整，含免费时长、分段计费和封顶。"
    elif match_count >= 3:
        return 3.5, "逻辑结构部分正确，含关键逻辑但不完整。"
    elif match_count >= 1:
        return 3, "仅实现部分逻辑内容，思路初显。"
    else:
        return 3, "仅实现部分逻辑内容，思路初显。"

# === 批量评分主函数 ===

def grade_html_js_page2_basic(folder: str, output_excel: str):
    folder_path = Path(folder)
    records = []

    for txt_file in folder_path.rglob("*page2.txt"):
        try:
            code = txt_file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"⚠️ 无法读取文件 {txt_file.name}: {e}")
            continue

        s6, c6 = score_question_6_table_structure(code)
        s7, c7 = score_question_7_table_content(code)
        s8, c8 = score_question_8_button_event(code)
        s9, c9 = score_question_9_function_defined(code)
        s10, c10 = score_question_10_parking_logic(code)
        total = s6 + s7 + s8 + s9 + s10

        records.append({
            "文件名": txt_file.name,
            "题6得分": s6, "题6评语": c6,
            "题7得分": s7, "题7评语": c7,
            "题8得分": s8, "题8评语": c8,
            "题9得分": s9, "题9评语": c9,
            "题10得分": s10, "题10评语": c10,
            "总分": total
        })

    df = pd.DataFrame(records)
    df.to_excel(output_excel, index=False)
    print(f"✅ 批量评分完成，结果保存至：{output_excel}")



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
            output_excel = fr"{folder}_评分结果page2.xlsx"
            grade_html_js_page2_basic(folder, output_excel)
